import streamlit as st
import os
import io
import pandas as pd
import requests
import re
import unicodedata
import base64

# Page config
st.set_page_config(
    page_title="Excel to WhatsApp",
    page_icon="📊",
    layout="centered"
)

# Environment variables - WhatsApp number is FIXED, not user input
WASENDER_API_KEY = os.getenv("WASENDER_API_KEY")
WASENDER_SESSION_ID = os.getenv("WASENDER_SESSION_ID")
WA_TO = os.getenv("WA_TO")  # Fixed WhatsApp number

# Your existing column allow-list
WANTED_COLS = [
    "PRODUCTION","GOLD","COLOUR STONE","BLACK BEADS","DIAMOND",
    "NO","PRODUCT ID","PRODUCT","STYLE","QTY","G Qly","Gr. WT","Nt. WT",
    "ITEMCODE","STONE PCS","STONE WT","STONE RATE","STONE AMT",
    "BEADS PCS","BEADS WT","BEADS RATE","BEADS AMT",
    "DIA PCS","DIA WT","SR NO","SR.NO.","Sr.No","SR. NO",
    "LAB","REPORT","LOT NUMBER","GIVEN TO",
    "SHAPE","WT.","COL","CLA","CUT","POL","SYM","FLO",
    "STK","SIZE","MM","CRTS.","PCS.","COLOR","CLARITY",
    "CODE","JOB NO","ITEM","DESIGN NO.","METAL AND CLR.","GROSS WT.",
    "NET WT.","METAL AMT.","DIAMOND PCS","STUDDING TYPE",
    "STUDDING WT","QUALITY","DIAMOND TYPE",
    "SIZE (mm)","PIECES","CARAT","TYPE",
    "PARTICULAR","CTS","cts.","PURITY","TOTAL PCS","CARATS",
    "DESCRIPTION","COLOUR","PCS/CT","PCS PER CT",

    "Description of Goods", "HSN CODE", "PCS/CTS", "PCS",
    "STONE TYPE", "Stone ID", "Cert", "Ratio", "Table", "Depth",
]

def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = s.replace("\u00a0", " ")
    s = s.strip().lower()
    s = re.sub(r"[\s._\-]+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

def process_excel(file_bytes: bytes, wanted_norm: list[str], file_ext: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)

    # Read file depending on type
    if file_ext == "csv":
        df = pd.read_csv(bio, dtype=str, header=None)
    else:
        # xlsx / xls
        if file_ext == "xls":
            engine = "xlrd"
        else:
            engine = "openpyxl"
        df = pd.read_excel(bio, dtype=str, engine=engine, header=None)

    df = df.fillna("")
    # --- rest of your existing logic below stays same ---
    # Find header row...
    best_idx = -1
    best_score = -1
    for i in range(min(300, len(df))):
        vals = [str(v) for v in df.iloc[i].tolist()]
        score = sum(1 for v in vals if norm(v) in WANTED_NORM)
        if score > best_score:
            best_idx, best_score = i, score

    if best_idx < 0:
        raise ValueError("Could not find header row matching allow-list")

    header_vals = [str(v) for v in df.iloc[best_idx].tolist()]
    body = df.iloc[best_idx+1:].reset_index(drop=True)
    body.columns = header_vals
    nm = {c: norm(c) for c in body.columns}
    keep = [c for c in body.columns if nm[c] in WANTED_NORM]

    if not keep:
        raise ValueError("No matching columns found")

    return body[keep]

def send_to_whatsapp(file_bytes: bytes, filename: str, to: str) -> dict:
    # Encode to base64
    base64_data = base64.b64encode(file_bytes).decode('utf-8')
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    data_url = f"data:{mime_type};base64,{base64_data}"
    
    upload_url = "https://www.wasenderapi.com/api/upload"
    headers = {
        "Authorization": f"Bearer {WASENDER_API_KEY}",
        "Content-Type": "application/json"
    }
    
    upload_response = requests.post(
        upload_url,
        headers=headers,
        json={"base64": data_url},
        timeout=60
    )
    
    if upload_response.status_code >= 300:
        raise Exception(f"Upload failed: {upload_response.text}")
    
    temp_url = upload_response.json().get("publicUrl")
    
    send_url = "https://www.wasenderapi.com/api/send-message"
    send_response = requests.post(
        send_url,
        headers=headers,
        json={
            "to": to.replace("+", ""),
            "documentUrl": temp_url,
            "fileName": filename,
        },
        timeout=60
    )
    
    if send_response.status_code >= 300:
        raise Exception(f"Send failed: {send_response.text}")
    
    return send_response.json()

# UI
st.title("📊 Excel to WhatsApp")
st.markdown("Upload your Excel file (.xlsx, .xls, .csv) and send to WhatsApp instantly.")

# File upload
# File upload
uploaded_file = st.file_uploader("Select Excel File", type=["xlsx", "xls", "csv"])

# Default filename from uploaded file (without extension)
default_name = ""
if uploaded_file is not None and uploaded_file.name:
    default_name = uploaded_file.name.rsplit(".", 1)[0]

filename = st.text_input(
    "Filename (optional – defaults to uploaded name)",
    value=default_name,
    placeholder="Enter filename (or leave as is)"
)

# Extra columns input
extra_cols_input = st.text_input(
    "Additional columns to include (comma-separated)",
    placeholder="e.g., REMARKS, CATEGORY, DATE",
)

# Process user input into a clean list
extra_cols = []
if extra_cols_input.strip():
    extra_cols = [c.strip() for c in extra_cols_input.split(",") if c.strip()]

# Merge with your predefined allow-list
WANTED_COLS_FINAL = WANTED_COLS + extra_cols
WANTED_NORM = [norm(c) for c in WANTED_COLS_FINAL]

# Submit button
if st.button("📤 Upload & Send to WhatsApp", type="primary", use_container_width=True):
    if not uploaded_file:
        st.error("⚠️ Please select a file")
    else:
        try:
            with st.spinner("Processing Excel file..."):
                file_bytes = uploaded_file.read()
                file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()

                filtered_df = process_excel(file_bytes, WANTED_NORM, file_ext)

                output = io.BytesIO()
                filtered_df.to_excel(output, index=False, engine="openpyxl")
                processed_bytes = output.getvalue()

                # Use typed name if provided, otherwise original base name
                base_name = (filename or default_name or "report").strip()
                final_filename = base_name
                if not final_filename.lower().endswith(".xlsx"):
                    final_filename = f"{base_name}.xlsx"
                
            with st.spinner(f"Sending to WhatsApp ({WA_TO})..."):
                result = send_to_whatsapp(processed_bytes, final_filename, WA_TO)
            
            # Clean success message - NO JSON display
            st.success(f"✅ File sent successfully to WhatsApp ({WA_TO})!")
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

# Footer
st.markdown("---")
st.markdown("Built with Streamlit • Powered by WasenderAPI")
