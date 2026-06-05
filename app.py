from dotenv import load_dotenv
load_dotenv()
import os
import io
import json
import re  
import zipfile
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException, WebSocket, WebSocketDisconnect
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
# WEBSOCKET LOGGING MANAGER
# Streams live server actions to the UI
# ==========================================
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message: str):
        for connection in self.active_connections:
            try:
                await connection.send_text(message)
            except Exception:
                pass

manager = ConnectionManager()

@app.websocket("/ws/status")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# ==========================================
# HARDCODED SAFETY FALLBACK 
# ==========================================
DEFAULT_MAPPINGS = {
    "Bajaj Allianz": {"pol": ["POLICY_REFERENCE", "Policy No", "Policy Number"], "cust": ["CUSTOMER NAME", "Insured Name"], "prod": ["PRODUCT", "Product Name"], "prem": ["NET PREMIUM", "Premium", "Gross Premium"], "comm": ["TOTAL COMMISSION", "Commission"], "date": ["POLICY DATE", "Policy Date"]},
    "Care Health": {"pol": ["Policy No", "Policy Number"], "cust": ["Customer Name", "Insured Name"], "prod": ["Type", "Product Name", "Product"], "prem": ["Premium", "Gross Premium"], "comm": ["Total Amount", "Commission", "Total Commission"], "date": ["Effective Date/Policy Start date", "Policy Date", "Effective Date"]},
    "Go Digit Life": {"pol": ["Policy Number", "Policy No"], "cust": ["Policy Holder Name", "Customer Name"], "prod": ["Product Name", "Product Code"], "prem": ["Net Premium", "Premium"], "comm": ["Total Commission Amount", "Commission"], "date": ["Policy Start Date", "Policy Date"]},
    "Go Digit General": {"pol": ["Policy Number", "Policy No.", "Policy No"], "cust": ["Customer Name", "Insured Name"], "prod": ["Product Name", "Product"], "prem": ["Gross Premium", "Net Premium", "Premium"], "comm": ["Total Commission", "Commission"], "date": ["Policy Issue Date", "Policy Date"]}
}

# ==========================================
# LIVE RULE MANAGER
# ==========================================
EXCEL_MASTER_FILE = "Insurer_Master_Mapping.xlsx"
CSV_MASTER_FILE = "Insurer_Master_Mapping.csv"
cached_mappings = {}
last_modified_time = 0

def get_latest_rules():
    global cached_mappings, last_modified_time
    file_to_read = None
    if os.path.exists(EXCEL_MASTER_FILE): file_to_read = EXCEL_MASTER_FILE
    elif os.path.exists(CSV_MASTER_FILE): file_to_read = CSV_MASTER_FILE
    if not file_to_read: return DEFAULT_MAPPINGS

    current_mtime = os.path.getmtime(file_to_read)
    if current_mtime > last_modified_time:
        try:
            if file_to_read.endswith('.csv'): df = pd.read_csv(file_to_read)
            else: df = pd.read_excel(file_to_read)
            df.columns = df.columns.astype(str).str.strip().str.lower()
            new_rules = {}
            for _, row in df.iterrows():
                company_col = next((c for c in df.columns if 'insurer' in c or 'company' in c), None)
                if not company_col: continue
                company = str(row.get(company_col, '')).strip()
                if pd.isna(company) or company == "nan" or not company: continue
                def clean_list(keyword):
                    col = next((c for c in df.columns if keyword in c), None)
                    if not col: return []
                    val = row.get(col, '')
                    if pd.isna(val) or str(val).strip() == "": return []
                    return [x.strip() for x in str(val).split(',') if x.strip()]

                new_rules[company] = {
                    "pol": clean_list('pol'), "cust": clean_list('cust'), "prod": clean_list('prod'),
                    "prem": clean_list('prem'), "comm": clean_list('comm'), "date": clean_list('date')
                }
            if new_rules:
                cached_mappings = new_rules
                last_modified_time = current_mtime
        except Exception as e:
            print(f"File read error, falling back to defaults: {e}")
            return DEFAULT_MAPPINGS
    return cached_mappings if cached_mappings else DEFAULT_MAPPINGS

@app.get("/api/insurers")
def get_insurers():
    rules = get_latest_rules()
    insurers = list(rules.keys())
    insurers.sort()
    insurers.append("Unknown")
    return {"insurers": insurers}

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
        with open("rm_list.json", "r") as f: rms = json.load(f)
        with open("agent_list.json", "r") as f: agents = json.load(f)
        return {"rms": rms, "agents": agents}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed loading internal rosters: {str(e)}")

# ==========================================
# PHASE 1: POLICY DATA PROCESSING ROUTES
# ==========================================
@app.post("/extract-batch")
async def extract_multiple_policies(files: List[UploadFile] = File(...)):
    combined_results = []
    total = len(files)
    await manager.broadcast(f"POLICY|Received batch of {total} documents.")
    
    for idx, file in enumerate(files):
        if not file.filename.lower().endswith('.pdf'): continue  
        try:
            await manager.broadcast(f"POLICY|[{idx+1}/{total}] Reading {file.filename} locally...")
            pdf_bytes = await file.read()
            extracted_text = ""
            try:
                with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                    max_pages = min(2, len(pdf.pages))
                    for i in range(max_pages):
                        page_text = pdf.pages[i].extract_text()
                        if page_text: extracted_text += page_text + "\n"
            except Exception: pass
            
            if len(extracted_text.strip()) > 500:
                await manager.broadcast(f"POLICY|[{idx+1}/{total}] Local text extracted. Analyzing via AI...")
                extracted_text = re.sub(r'\n{2,}', '\n', extracted_text)
                extracted_text = re.sub(r'[ \t]{2,}', ' ', extracted_text)[:20000]
                prompt_contents = [f"Analyze this raw text extracted from an insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract.\n\nRAW TEXT:\n{extracted_text}"]
            else:
                await manager.broadcast(f"POLICY|[{idx+1}/{total}] Locked PDF detected. Processing via AI Vision...")
                prompt_contents = [types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), "Analyze this scanned insurance policy. If it is a motor/vehicle policy, set is_motor_policy to true and extract the vehicle details. If it is Health/Other, set it to false and leave vehicle fields blank. Map exactly to the schema contract."]

            # Using Gemini 3.5 Flash for high-frequency processing speed
            response = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt_contents,
                config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=PolicyExtraction, temperature=0.0),
            )
            extracted_dict = json.loads(response.text)
            extracted_dict["source_file"] = file.filename
            combined_results.append(extracted_dict)
            await manager.broadcast(f"POLICY|[{idx+1}/{total}] Successfully processed {file.filename}.")
        except Exception as e:
            await manager.broadcast(f"POLICY|❌ Error processing {file.filename}.")
            combined_results.append({"policy_no": "ERROR", "insurer_company": "Failed to parse", "customer_name": str(e), "gross_premium": "0", "gst": "0", "product_name": "N/A", "policy_start_date": "N/A", "policy_end_date": "N/A", "is_motor_policy": False, "rto_location": "", "vehicle_make_model": "", "fuel_type": "", "cubic_capacity": "", "mfg_or_reg_date": "", "source_file": file.filename, "parsing_error": str(e)})
            
    await manager.broadcast("POLICY|✅ Batch extraction complete!")
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
# PHASE 2: COMMISSION PROCESSING
# ==========================================
def safe_float(val):
    if pd.isna(val) or val is None or str(val).strip() == "": return 0.0
    val_str = str(val).strip()
    if val_str.lower() in ['nil', 'na', 'n/a', '-']: return 0.0
    clean_val = re.sub(r'[^\d.-]', '', val_str)
    try: return float(clean_val) if clean_val else 0.0
    except ValueError: return 0.0

def get_col_name(columns, aliases):
    normalized_file_cols = [re.sub(r'[^a-z0-9]', '', str(c).lower()) for c in columns]
    for alias in aliases:
        normalized_alias = re.sub(r'[^a-z0-9]', '', str(alias).lower())
        if not normalized_alias: continue
        if normalized_alias in normalized_file_cols:
            return columns[normalized_file_cols.index(normalized_alias)]
    return None

def read_file_to_dfs(filename, contents):
    dfs = []
    filename_lower = filename.lower()
    
    def safe_read_csv(content_bytes):
        try: return pd.read_csv(io.BytesIO(content_bytes), dtype=str)
        except Exception:
            try: return pd.read_csv(io.BytesIO(content_bytes), encoding='cp1252', dtype=str, on_bad_lines='skip')
            except Exception:
                try: return pd.read_csv(io.BytesIO(content_bytes), sep='\t', dtype=str)
                except Exception: return None

    if filename_lower.endswith('.csv'):
        res = safe_read_csv(contents)
        if res is not None: dfs.append(res)
    elif filename_lower.endswith('.pdf'):
        all_data = []
        try:
            with pdfplumber.open(io.BytesIO(contents)) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_data.extend(table)
            if len(all_data) > 1: dfs.append(pd.DataFrame(all_data[1:], columns=all_data[0]))
        except Exception: pass
    else:
        engine = 'pyxlsb' if filename_lower.endswith('.xlsb') else ('openpyxl' if filename_lower.endswith('.xlsx') else None)
        try:
            xls = pd.ExcelFile(io.BytesIO(contents), engine=engine)
            for sheet in xls.sheet_names:
                dfs.append(pd.read_excel(xls, sheet_name=sheet, dtype=str))
        except Exception:
            res = safe_read_csv(contents)
            if res is not None: dfs.append(res)
    return dfs

@app.post("/analyze-commission-files")
async def analyze_commission_files(files: List[UploadFile] = File(...)):
    await manager.broadcast("COMM_ANALYZE|Loading Master Mapping Database...")
    mappings = get_latest_rules()
    results = []
    
    for file in files:
        await manager.broadcast(f"COMM_ANALYZE|Scanning filename: {file.filename}...")
        fname_lower = file.filename.lower()
        detected = "Unknown"
        for insurer in mappings.keys():
            if insurer.lower() in fname_lower:
                detected = insurer
                break
        if detected == "Unknown":
            for insurer in mappings.keys():
                keywords = [w for w in insurer.lower().split() if len(w) > 3 and w not in ['insurance', 'general', 'life', 'health', 'company', 'ltd', 'limited']]
                if any(kw in fname_lower for kw in keywords):
                    detected = insurer
                    break
        results.append({"filename": file.filename, "detected_insurer": detected})
        
    await manager.broadcast("COMM_ANALYZE|✅ Analysis complete.")
    return {"success": True, "results": results}

@app.post("/process-commission-batch")
async def process_commission_batch(files: List[UploadFile] = File(...), insurers: str = Form(...)):
    try:
        await manager.broadcast("COMM_PROCESS|Initializing data extraction engine...")
        mappings = get_latest_rules()
        insurer_list = json.loads(insurers)
        standardized_data = []
        total = len(files)
        
        for i, file in enumerate(files):
            final_insurer = insurer_list[i]
            await manager.broadcast(f"COMM_PROCESS|[{i+1}/{total}] Processing {file.filename} as {final_insurer}...")
            
            mapping = None
            for key in mappings.keys():
                if key.strip().lower() == final_insurer.strip().lower():
                    mapping = mappings[key]
                    break
                    
            if not mapping:
                await manager.broadcast(f"COMM_PROCESS|⚠️ Skipping {file.filename} (No mapping found).")
                continue 
                
            contents = await file.read()
            dfs = read_file_to_dfs(file.filename, contents)

            for target_df in dfs:
                target_df.columns = target_df.columns.astype(str).str.strip()
                pol_col = get_col_name(target_df.columns, mapping['pol'])
                
                if not pol_col:
                    for idx, row in target_df.head(20).iterrows():
                        row_strs = [str(x).strip() for x in row.values]
                        found_pol = get_col_name(row_strs, mapping['pol'])
                        if found_pol:
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
                    await manager.broadcast(f"COMM_PROCESS|[{i+1}/{total}] Exact column match found. Extracting rows...")
                    cust_col = get_col_name(target_df.columns, mapping['cust'])
                    prod_col = get_col_name(target_df.columns, mapping['prod'])
                    prem_col = get_col_name(target_df.columns, mapping['prem'])
                    comm_col = get_col_name(target_df.columns, mapping['comm'])
                    date_col = get_col_name(target_df.columns, mapping['date'])
                    
                    for _, row in target_df.iterrows():
                        pol_val = str(row.get(pol_col, '')).strip()
                        if pd.isna(row.get(pol_col)) or pol_val == "" or pol_val.lower() == "nan" or "total" in pol_val.lower() or "grand" in pol_val.lower():
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
                    
        await manager.broadcast("COMM_PROCESS|✅ Data normalization complete! Building Excel array...")            
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
