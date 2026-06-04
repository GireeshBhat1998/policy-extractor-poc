from dotenv import load_dotenv
load_dotenv()
import os
import io
import json
import re  
import zipfile
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
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

# ==========================================
# UNIVERSAL FUZZY MAPPING DICTIONARY
# (Expanded aliases to catch any variations without scrambling)
# ==========================================
INSURER_MAPPINGS = {
    "Bajaj Allianz": {"pol": ["POLICY_REFERENCE", "Policy No", "Policy Number"], "cust": ["CUSTOMER NAME", "Insured Name"], "prod": ["PRODUCT", "Product Name"], "prem": ["NET PREMIUM", "Premium", "Gross Premium"], "comm": ["TOTAL COMMISSION", "Commission"], "date": ["POLICY DATE", "Policy Date"]},
    "Care Health": {"pol": ["Policy No", "Policy Number"], "cust": ["Customer Name", "Insured Name"], "prod": ["Type", "Product Name", "Product"], "prem": ["Premium", "Gross Premium"], "comm": ["Total Amount", "Commission", "Total Commission"], "date": ["Effective Date/Policy Start date", "Policy Date", "Effective Date"]},
    "ICICI Lombard": {"pol": ["POLICY_NUMBER", "Policy Number", "Policy No"], "cust": ["INSURED_CUSTOMER_NAME", "Customer Name", "Insured Name"], "prod": ["PRODUCT_NAME", "Product Name"], "prem": ["TOTAL_PREMIUM_RECEIVED", "Premium"], "comm": ["ACTUAL_COMMISSION", "Commission", "Total Billed"], "date": ["POLICY_START_DATE", "Policy Date"]},
    "IndusInd": {"pol": ["PolicyNumber", "Policy No"], "cust": ["InsuredName", "Customer Name"], "prod": ["ProductCode", "Product Name"], "prem": ["PremiumAmount", "Premium"], "comm": ["FinalIRDAComm", "Commission"], "date": ["Month", "Policy Date"]},
    "Liberty": {"pol": ["POLICY/ENDORSEMENT NO.", "Policy No"], "cust": ["INSURED NAME", "Customer Name"], "prod": ["PRODUCT NAME", "Product"], "prem": ["GWP", "Premium", "Gross Premium"], "comm": ["FINAL COMM TO BE PAID", "Commission", "Total Commission"], "date": ["POLICY START DATE", "Policy Date"]},
    "National": {"pol": ["Policy #-Endo#", "Policy No", "Policy Number"], "cust": ["Insured Name", "Customer Name"], "prod": ["Prdt Code", "Product Name"], "prem": ["Premium Amount", "Premium"], "comm": ["Commission Amount", "Commission"], "date": ["Effective Date", "Policy Date"]},
    "Royal Sundaram": {"pol": ["POLICY ID", "Policy No"], "cust": ["CLIENT NAME", "Customer Name"], "prod": ["PRODUCT CATEGORY 2", "Product Name"], "prem": ["GROSS WRITTEN PREMIUM", "Premium"], "comm": ["TOTAL COMMISSION", "Commission"], "date": ["POLICY ENTRY DATE", "Policy Date"]},
    "Go Digit Life": {"pol": ["Policy Number", "Policy No"], "cust": ["Policy Holder Name", "Policy Holder", "Customer Name"], "prod": ["Product Name", "Product Code"], "prem": ["Net Premium", "Premium"], "comm": ["Total Commission Amount", "Commission"], "date": ["Policy Start Date", "Policy Issue Date", "Policy Date"]},
    "Go Digit General": {"pol": ["Policy Number", "Policy No.", "Policy No"], "cust": ["Customer Name", "Insured Name"], "prod": ["Product Name", "Product"], "prem": ["Gross Premium", "Net Premium", "Premium"], "comm": ["Total Commission", "Commission"], "date": ["Policy Issue Date", "Policy Date", "Policy Start Date"]},
    "Tata AIG": {"pol": ["Policy No", "Policy Number"], "cust": ["Insured Name", "Customer Name"], "prod": ["Product", "Product Name"], "prem": ["Premium", "Gross Premium"], "comm": ["Commission", "Total Commission"], "date": ["Policy Date", "Policy Start Date"]},
    "HDFC Ergo": {"pol": ["Policy Number", "Policy No.", "Policy No"], "cust": ["Customer Name", "Insured Name"], "prod": ["Product Name", "Product"], "prem": ["Premium", "Gross Premium", "Net Premium"], "comm": ["Commission", "Total Commission", "Brokerage"], "date": ["Policy Date", "Policy Start Date", "Transaction Date"]}
}

SUPPORTED_INSURERS = list(INSURER_MAPPINGS.keys()) + ["Unknown"]

class PolicyExtraction(BaseModel):
    policy_no: str = Field(description="The unique policy or certificate number")
    insurer_company: str = Field(description="Name of the insurance company")
    customer_name: str = Field(description="Full name of the policyholder or primary insured")
    gross_premium: str = Field(description="The base/gross premium value before tax")
    gst: str = Field(description="The total GST or tax amount applied")
    product_name: str = Field(description="The specific name of the insurance plan or product")
    policy_start_date: str = Field(description="The risk commencement or start date of the policy")
    policy_end_date: str = Field(description="The expiry or end date of the policy")
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
                    max_pages = min(5, len(pdf.pages))
                    for i in range(max_pages):
                        page_text = pdf.pages[i].extract_text()
                        if page_text:
                            extracted_text += page_text + "\n"
            except Exception:
                pass
            
            if len(extracted_text.strip()) > 500:
                extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)
                extracted_text = re.sub(r'[ \t]{2,}', ' ', extracted_text)
                extracted_text = extracted_text[:20000]
                prompt_contents = [f"Analyze this raw text extracted from an insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract.\n\nRAW TEXT:\n{extracted_text}"]
            else:
                prompt_contents = [
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    "Analyze this scanned insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract."
                ]

            response = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt_contents,
                config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=PolicyExtraction, temperature=0.0),
            )
            extracted_dict = json.loads(response.text)
            extracted_dict["source_file"] = file.filename
            combined_results.append(extracted_dict)
        except Exception as e:
            combined_results.append({"policy_no": "ERROR", "insurer_company": "Failed to parse file", "customer_name": "Check error log below", "gross_premium": "0", "gst": "0", "product_name": "N/A", "policy_start_date": "N/A", "policy_end_date": "N/A", "is_motor_policy": False, "rto_location": "", "vehicle_make_model": "", "fuel_type": "", "cubic_capacity": "", "mfg_or_reg_date": "", "source_file": file.filename, "parsing_error": str(e)})
    return {"success": True, "results": combined_results}

@app.post("/export-excel-batch")
async def export_batch_to_excel(data: List[dict]):
    try:
        df = pd.DataFrame(data)
        preferred_order = ["business_month", "business_year", "policy_number", "insurer_company", "customer_name", "product_name", "gross_premium", "gst", "policy_start_date", "policy_end_date", "relationship_manager", "agent_id", "agent_name", "agent_commission_rate", "commissionable_premium", "brokerage_rate_percent", "calculated_brokerage", "is_motor_policy", "rto_location", "vehicle_make_model", "fuel_type", "cubic_capacity", "mfg_or_reg_date", "source_file"]
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

# ==========================================
# PHASE 2: COMMISSION PROCESSING (UPGRADED CONFIDENCE SCORING)
# ==========================================

def safe_float(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return 0.0
    val_str = str(val).strip()
    if val_str.lower() in ['nil', 'na', 'n/a', '-']:
        return 0.0
    clean_val = re.sub(r'[^\d.-]', '', val_str)
    try:
        return float(clean_val) if clean_val else 0.0
    except ValueError:
        return 0.0

def get_col_name(columns, aliases):
    clean_cols = [str(c).strip().lower().replace('\n', ' ').replace('\r', '') for c in columns]
    # Pass 1: Exact Match
    for alias in aliases:
        clean_alias = alias.lower().strip()
        if clean_alias in clean_cols:
            return columns[clean_cols.index(clean_alias)]
    # Pass 2: Containment
    for alias in aliases:
        clean_alias = alias.lower().strip()
        for i, c in enumerate(clean_cols):
            if clean_alias in c:
                return columns[i]
    return None

def read_file_to_dfs(filename, contents):
    dfs = []
    filename_lower = filename.lower()
    
    if filename_lower.endswith('.csv'):
        try:
            dfs.append(pd.read_csv(io.BytesIO(contents), dtype=str))
        except Exception:
            try:
                dfs.append(pd.read_csv(io.BytesIO(contents), encoding='cp1252', dtype=str, on_bad_lines='skip'))
            except Exception:
                pass
    elif filename_lower.endswith('.pdf'):
        all_data = []
        try:
            with pdfplumber.open(io.BytesIO(contents)) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_data.extend(table)
            if len(all_data) > 1:
                dfs.append(pd.DataFrame(all_data[1:], columns=all_data[0]))
        except Exception:
            pass
    else:
        engine = 'pyxlsb' if filename_lower.endswith('.xlsb') else ('openpyxl' if filename_lower.endswith('.xlsx') else None)
        try:
            xls = pd.ExcelFile(io.BytesIO(contents), engine=engine)
            for sheet in xls.sheet_names:
                dfs.append(pd.read_excel(xls, sheet_name=sheet, dtype=str))
        except zipfile.BadZipFile:
            try:
                dfs.append(pd.read_csv(io.BytesIO(contents), dtype=str))
            except Exception:
                pass
        except Exception:
            pass
    return dfs

def guess_insurer_from_df(df):
    """Calculates a confidence score to prevent false-positive mapping"""
    df.columns = df.columns.astype(str).str.strip()
    
    def score_cols(cols, mapping):
        score = 0
        if get_col_name(cols, mapping['pol']): score += 1
        if get_col_name(cols, mapping['cust']): score += 1
        if get_col_name(cols, mapping['prem']): score += 1
        if get_col_name(cols, mapping['comm']): score += 1
        return score

    best_insurer = "Unknown"
    max_score = 0
    
    # Scan standard headers
    for insurer, mapping in INSURER_MAPPINGS.items():
        score = score_cols(df.columns, mapping)
        if score > max_score:
            max_score = score
            best_insurer = insurer
            
    if max_score >= 3:
        return best_insurer
        
    # Scan hidden headers (first 15 rows)
    for idx, row in df.head(15).iterrows():
        row_strs = [str(x).strip() for x in row.values]
        for insurer, mapping in INSURER_MAPPINGS.items():
            score = score_cols(row_strs, mapping)
            if score > max_score:
                max_score = score
                best_insurer = insurer
        if max_score >= 3:
            return best_insurer
            
    return "Unknown"

@app.post("/analyze-commission-files")
async def analyze_commission_files(files: List[UploadFile] = File(...)):
    results = []
    for file in files:
        try:
            contents = await file.read()
            dfs = read_file_to_dfs(file.filename, contents)
            
            detected = "Unknown"
            for df in dfs:
                detected = guess_insurer_from_df(df)
                if detected != "Unknown":
                    break
            
            if detected == "Unknown":
                fname_lower = file.filename.lower()
                for insurer in SUPPORTED_INSURERS:
                    if insurer != "Unknown" and insurer.lower() in fname_lower:
                        detected = insurer
                        break
                        
            results.append({"filename": file.filename, "detected_insurer": detected})
        except Exception:
            results.append({"filename": file.filename, "detected_insurer": "Error Reading File"})
            
    return {"success": True, "results": results}

@app.post("/process-commission-batch")
async def process_commission_batch(files: List[UploadFile] = File(...), insurers: str = Form(...)):
    try:
        insurer_list = json.loads(insurers)
        standardized_data = []
        
        for i, file in enumerate(files):
            final_insurer = insurer_list[i]
            if final_insurer == "Unknown" or final_insurer == "Error Reading File":
                continue 
                
            contents = await file.read()
            dfs = read_file_to_dfs(file.filename, contents)

            mapping = INSURER_MAPPINGS.get(final_insurer)
            if not mapping:
                continue

            for target_df in dfs:
                target_df.columns = target_df.columns.astype(str).str.strip()
                pol_col = get_col_name(target_df.columns, mapping['pol'])
                
                # Upgraded Smart Header Finder
                if not pol_col:
                    for idx, row in target_df.head(20).iterrows():
                        row_strs = [str(x).strip() for x in row.values]
                        found_pol = get_col_name(row_strs, mapping['pol'])
                        found_prem = get_col_name(row_strs, mapping['prem'])
                        found_comm = get_col_name(row_strs, mapping['comm'])
                        
                        # Only reset headers if it finds Pol Num AND Premium/Comm
                        if found_pol and (found_prem or found_comm):
                            # Ensure unique column names to prevent Pandas crashes
                            unique_cols = []
                            seen = set()
                            for c in row_strs:
                                new_c = c
                                count = 1
                                while new_c in seen:
                                    new_c = f"{c}_{count}"
                                    count += 1
                                seen.add(new_c)
                                unique_cols.append(new_c)
                                
                            target_df.columns = unique_cols
                            target_df = target_df.iloc[idx+1:].reset_index(drop=True)
                            pol_col = found_pol
                            break
                
                if pol_col:
                    cust_col = get_col_name(target_df.columns, mapping['cust'])
                    prod_col = get_col_name(target_df.columns, mapping['prod'])
                    prem_col = get_col_name(target_df.columns, mapping['prem'])
                    comm_col = get_col_name(target_df.columns, mapping['comm'])
                    date_col = get_col_name(target_df.columns, mapping['date'])
                    
                    for _, row in target_df.iterrows():
                        pol_val = str(row.get(pol_col, '')).strip()
                        if pd.isna(row.get(pol_col)) or pol_val == "" or pol_val.lower() == "nan" or "total" in pol_val.lower():
                            continue
                            
                        standardized_data.append({
                            "insurer_company": final_insurer,
                            "policy_number": pol_val,
                            "customer_name": str(row.get(cust_col, '')) if cust_col else "",
                            "product_name": str(row.get(prod_col, '')) if prod_col else "",
                            "gross_premium": safe_float(row.get(prem_col, 0)) if prem_col else 0.0,
                            "commission_received": safe_float(row.get(comm_col, 0)) if comm_col else 0.0,
                            "policy_date": str(row.get(date_col, '')) if date_col else "",
                            "source_file": file.filename
                        })
                    break 
                    
        if not standardized_data:
            return {"success": True, "total_records": 0, "data": []}
            
        clean_df = pd.DataFrame(standardized_data)
        clean_df['match_key'] = clean_df['policy_number'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        final_results = clean_df.to_dict(orient='records')
        
        return {"success": True, "total_records": len(final_results), "data": final_results}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/export-commission-excel")
async def export_commission_to_excel(data: List[dict]):
    try:
        df = pd.DataFrame(data)
        preferred_order = ["insurer_company", "policy_number", "match_key", "customer_name", "product_name", "gross_premium", "commission_received", "policy_date", "source_file"]
        clean_cols = [c for c in preferred_order if c in df.columns]
        df = df[clean_cols]
        df.columns = [col.replace('_', ' ').title() for col in df.columns]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Standardized Commissions')
        output.seek(0)
        headers = {'Content-Disposition': 'attachment; filename="standardized_commissions_report.xlsx"'}
        return StreamingResponse(output, headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
