from dotenv import load_dotenv
load_dotenv()
import os
import io
import json
import re  
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import pandas as pd
import pdfplumber

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- EXPANDED SCHEMA: Conditional Motor Fields ---
class PolicyExtraction(BaseModel):
    policy_no: str = Field(description="The unique policy or certificate number")
    insurer_company: str = Field(description="Name of the insurance company")
    customer_name: str = Field(description="Full name of the policyholder or primary insured")
    gross_premium: str = Field(description="The base/gross premium value before tax")
    gst: str = Field(description="The total GST or tax amount applied")
    product_name: str = Field(description="The specific name of the insurance plan or product")
    policy_start_date: str = Field(description="The risk commencement or start date of the policy")
    policy_end_date: str = Field(description="The expiry or end date of the policy")
    
    # Motor Specific Context
    is_motor_policy: bool = Field(description="True if this is a motor, car, or vehicle insurance policy, False otherwise")
    rto_location: str = Field(description="RTO Location (Motor only), leave blank if not applicable", default="")
    vehicle_make_model: str = Field(description="Vehicle Make and Model (Motor only), leave blank if not applicable", default="")
    fuel_type: str = Field(description="Fuel Type (Motor only), leave blank if not applicable", default="")
    cubic_capacity: str = Field(description="Cubic Capacity or CC (Motor only), leave blank if not applicable", default="")
    mfg_or_reg_date: str = Field(description="Manufacturing or Registration Month and Year (Motor only), leave blank if not applicable", default="")

client = genai.Client()

@app.get("/metadata")
def get_metadata():
    try:
        with open("rm_list.json", "r") as f:
            rms = json.load(f)
        with open("agent_list.json", "r") as f:
            agents = json.load(f)
        return {"rms": rms, "agents": agents}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed loading internal rosters: {str(e)}")

@app.post("/extract-batch")
async def extract_multiple_policies(files: List[UploadFile] = File(...)):
    combined_results = []
    
    for file in files:
        if not file.filename.endswith('.pdf'):
            continue  
            
        try:
            pdf_bytes = await file.read()
            
            extracted_text = ""
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    # --- OPTIMIZATION 1: Page Limiting ---
                    # Only scan the first 5 pages where the Policy Schedule lives.
                    # This instantly drops heavy T&C boilerplate.
                    max_pages = min(5, len(pdf.pages))
                    for i in range(max_pages):
                        page_text = pdf.pages[i].extract_text()
                        if page_text:
                            extracted_text += page_text + "\n"
            except Exception as e:
                print(f"Local parsing skipped/failed: {e}")
            
            if len(extracted_text.strip()) > 500:
                # --- OPTIMIZATION 2: Regex Token Compression ---
                # Strip excessive blank lines and multi-spaces that consume useless tokens
                extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)
                extracted_text = re.sub(r'[ \t]{2,}', ' ', extracted_text)
                
                # --- OPTIMIZATION 3: Hard Token Ceiling ---
                # Cap the maximum string length to 10,000 characters (~2,500 input tokens max)
                extracted_text = extracted_text[:10000]
                
                prompt_contents = [
                    f"Analyze this raw text extracted from an insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract.\n\nRAW TEXT:\n{extracted_text}"
                ]
            else:
                # Fallback for scanned (image) PDFs
                prompt_contents = [
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    "Analyze this scanned insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract."
                ]

            response = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt_contents,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=PolicyExtraction,
                    temperature=0.0,
                ),
            )
            
            extracted_dict = json.loads(response.text)
            extracted_dict["source_file"] = file.filename
            combined_results.append(extracted_dict)
            
        except Exception as e:
            # Error Payload Tracking
            combined_results.append({
                "policy_no": "ERROR", "insurer_company": "Failed to parse file",
                "customer_name": "Check error log below", "gross_premium": "0", "gst": "0",
                "product_name": "N/A", "policy_start_date": "N/A", "policy_end_date": "N/A",
                "is_motor_policy": False, "rto_location": "", "vehicle_make_model": "",
                "fuel_type": "", "cubic_capacity": "", "mfg_or_reg_date": "",
                "source_file": file.filename,
                "parsing_error": str(e)
            })
            
    return {"success": True, "results": combined_results}

@app.post("/export-excel-batch")
async def export_batch_to_excel(data: List[dict]):
    try:
        df = pd.DataFrame(data)
        
        preferred_order = [
            "business_month", "business_year", "source_file", "policy_no", 
            "insurer_company", "customer_name", "product_name", "gross_premium", "gst",
            "policy_start_date", "policy_end_date", "relationship_manager", 
            "agent_id", "agent_name", "commissionable_premium", 
            "brokerage_rate_percent", "calculated_brokerage",
            "is_motor_policy", "rto_location", "vehicle_make_model", 
            "fuel_type", "cubic_capacity", "mfg_or_reg_date"
        ]
        
        clean_cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
        df = df[clean_cols]
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Batch Operations')
        output.seek(0)
        
        headers = {'Content-Disposition': 'attachment; filename="consolidated_policies_report.xlsx"'}
        return StreamingResponse(output, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
