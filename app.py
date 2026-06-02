from dotenv import load_dotenv
load_dotenv()
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

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class PolicyExtraction(BaseModel):
    policy_no: str = Field(description="The unique policy or certificate number")
    insurer_company: str = Field(description="Name of the insurance company")
    customer_name: str = Field(description="Full name of the policyholder or primary insured")
    gross_premium: str = Field(description="The base/gross premium value before tax")
    gst: str = Field(description="The total GST or tax amount applied")
    product_name: str = Field(description="The specific name of the insurance plan or product")
    policy_start_date: str = Field(description="The risk commencement or start date of the policy")
    policy_end_date: str = Field(description="The expiry or end date of the policy")

client = genai.Client()

# 🌐 NEW ENDPOINTS: Read Dynamic Rosters from Git Files local path
@app.get("/metadata")
def get_metadata():
    try:
        with open("rm_list.json", "r") as f:
            rms = json.load(f)
        with open("agent_list.json", "r") as f:
            agents = json.load(f)
        return {"rms": rms, "agents": agents}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed loading lists: {str(e)}")

@app.post("/extract-batch")
async def extract_multiple_policies(files: List[UploadFile] = File(...)):
    combined_results = []
    for file in files:
        if not file.filename.endswith('.pdf'):
            continue
        try:
            pdf_bytes = await file.read()
            response = client.models.generate_content(
                model="gemini-2.5-flash-lite", # Highly optimal for structured extractions
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    "Analyze this insurance policy copy and accurately extract the requested fields mapping them exactly to the schema contract provided."
                ],
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
                "policy_no": "ERROR", "insurer_company": "Failed to parse file",
                "customer_name": f"Error: {str(e)}", "gross_premium": "0", "gst": "0",
                "product_name": "N/A", "policy_start_date": "N/A", "policy_end_date": "N/A",
                "source_file": file.filename
            })
    return {"success": True, "results": combined_results}

@app.post("/export-excel-batch")
async def export_batch_to_excel(data: List[dict]):
    try:
        df = pd.DataFrame(data)
        
        # Explicitly order columns for clean corporate reporting
        order = [
            "business_month", "business_year", "source_file", "policy_no", 
            "insurer_company", "customer_name", "gross_premium", "gst",
            "relationship_manager", "agent_id", "agent_name", 
            "commissionable_premium", "brokerage_rate_percent", "calculated_brokerage"
        ]
        # Fallback to keep whatever columns exist safely
        cols = [c for c in order if c in df.columns] + [c for c in df.columns if c not in order]
        df = df[cols]

        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Batch Operations Report')
        output.seek(0)
        
        return StreamingResponse(
            output, 
            headers={'Content-Disposition': 'attachment; filename="consolidated_operational_report.xlsx"'}, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
