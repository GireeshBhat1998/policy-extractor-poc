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

@app.post("/extract-batch")
async def extract_multiple_policies(files: List[UploadFile] = File(...)):
    """Accepts multiple files, iterates through them, and packages structured responses"""
    combined_results = []
    
    for file in files:
        if not file.filename.endswith('.pdf'):
            continue  # Skip non-PDF items gracefully
            
        try:
            pdf_bytes = await file.read()
            
            # Call Gemini API to extract data fields natively
            response = client.models.generate_content(
                model="gemini-2.5-flash",
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
            
            # Parse response string into a structured dictionary element
            extracted_dict = json.loads(response.text)
            extracted_dict["source_file"] = file.filename  # Attaches filename tracking
            combined_results.append(extracted_dict)
            
        except Exception as e:
            # Fallback error row payload if an individual file processing step defaults
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
    """Accepts JSON extraction list arrays and builds a clean vertical stacked Excel sheet"""
    try:
        df = pd.DataFrame(data)
        
        # Format table headers to be user-friendly spaces
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Batch Extraction')
        output.seek(0)
        
        headers = {
            'Content-Disposition': 'attachment; filename="consolidated_policies_report.xlsx"'
        }
        return StreamingResponse(output, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    # Render cloud binds dynamically using environment ports
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
