from dotenv import load_dotenv
load_dotenv()  # This automatically finds the .env file and injects the key!
import os
import io
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
import pandas as pd

app = FastAPI()

# Enable CORS so our frontend can communicate with the backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 1. Define the exact schema we need extracted
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
# It will automatically look for the GEMINI_API_KEY environment variable
client = genai.Client()

@app.post("/extract")
async def extract_policy_data(file: UploadFile = File(...)):
    if not file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are supported.")
    
    try:
        # Read the uploaded PDF file bytes
        pdf_bytes = await file.read()
        
        # 2. Call Gemini API to perform the OCR-less structured data extraction
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                "Analyze this insurance policy copy and accurately extract the requested fields mapping them exactly to the schema contract provided."
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=PolicyExtraction,
                temperature=0.0, # Low temperature ensures strict factual data matching
            ),
        )
        
        # The response.text is guaranteed to be valid JSON matching our Pydantic class
        return {"success": True, "data": response.text}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/export-excel")
async def export_to_excel(data: dict):
    try:
        # Convert the received data dictionary into a DataFrame
        df = pd.DataFrame([data])
        
        # Rename columns for presentation in Excel
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        # Save Excel to an in-memory buffer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Extracted Policy')
        output.seek(0)
        
        # 3. Stream the file back to the browser for direct download
        headers = {
            'Content-Disposition': 'attachment; filename="extracted_policy.xlsx"'
        }
        return StreamingResponse(output, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)