from dotenv import load_dotenv
load_dotenv()  # Automatically injects the GEMINI_API_KEY from Render environment settings
import os
import io
import json
from typing import List
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import pandas as pd
import pdfplumber  # <-- NEW HYBRID LIBRARY IMPORT

app = FastAPI()

# Enable CORS so your local desktop HTML file can talk to the Render Cloud server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Define the data contract for extraction
class PolicyExtraction(BaseModel):
    policy_no: str = Field(description="The unique policy or certificate number")
    insurer_company: str = Field(description="Name of the insurance company")
    customer_name: str = Field(description="Full name of the policyholder or primary insured")
    gross_premium: str = Field(description="The base/gross premium value before tax")
    gst: str = Field(description="The total GST or tax amount applied")
    product_name: str = Field(description="The specific name of the insurance plan or product")
    policy_start_date: str = Field(description="The risk commencement or start date of the policy")
    policy_end_date: str = Field(description="The expiry or end date of the policy")

# Initialize the Gemini Client
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
    """Accepts multiple files, intelligently pre-parses text, and packages structured responses"""
    combined_results = []
    
    for file in files:
        if not file.filename.endswith('.pdf'):
            continue  
            
        try:
            pdf_bytes = await file.read()
            
            # --- 🚀 HYBRID EXTRACTION LOGIC (SPEED & COST OPTIMIZATION) ---
            extracted_text = ""
            try:
                # Attempt to extract raw text streams locally using python
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            extracted_text += page_text + "\n"
            except Exception as e:
                print(f"Local parsing skipped/failed: {e}")
            
            # Determine Payload Architecture based on local extraction success
            if len(extracted_text.strip()) > 500:
                # DIGITAL PDF ROUTE: Highly efficient text-based prompt (Uses 80% fewer tokens, fixes IndusInd table issues)
                prompt_contents = [
                    f"Analyze this raw text extracted from an insurance policy and accurately extract the requested fields mapping them exactly to the schema contract provided.\n\nRAW TEXT:\n{extracted_text}"
                ]
            else:
                # SCANNED IMAGE ROUTE: Fallback to heavier Visual Byte extraction for physical paper scans
                prompt_contents = [
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    "Analyze this scanned insurance policy image and accurately extract the requested fields mapping them exactly to the schema contract provided."
                ]
            # --------------------------------------------------------------

            # Call Gemini API with the optimized payload
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
            combined_results.append({
                "policy_no": "ERROR", 
                "insurer_company": "Failed to parse file",
                "customer_name": f"Error message: {str(e)}", 
                "gross_premium": "0", "gst": "0",
                "product_name": "N/A", "policy_start_date": "N/A", "policy_end_date": "N/A",
                "source_file": file.filename
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
            "brokerage_rate_percent", "calculated_brokerage"
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
