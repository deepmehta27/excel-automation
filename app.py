import streamlit as st
import os
import io
import pandas as pd
import requests
import re
import unicodedata
import base64
import time

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
    "Cert.", "Cert. No.", "CertificateNo", "Diameter",
    "SETIAL", "WEIGHT", "MM SIZE", "PT",
    "Cert. No", "SIZE RANGE","POLISH","MEASUREMENT","RATIO",
    "CERT NUMBER","FL",
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
uploaded_files = st.file_uploader(
    "Select Excel Files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

rename_map = {}

if uploaded_files:
    st.write("### Rename files (optional)")
    for file in uploaded_files:
        default_base = file.name.rsplit(".", 1)[0]
        new_name = st.text_input(
            f"Rename for {file.name} (optional)",
            value=default_base,
            key=file.name  # unique key for each file
        )
        rename_map[file.name] = new_name

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

    if not uploaded_files:
        st.error("⚠️ Please select at least one file")
    else:
        try:
            for uploaded_file in uploaded_files:
                file_bytes = uploaded_file.read()
                file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()

                # Process
                filtered_df = process_excel(file_bytes, WANTED_NORM, file_ext)

                output = io.BytesIO()
                filtered_df.to_excel(output, index=False, engine="openpyxl")
                processed_bytes = output.getvalue()

                # Determine filename
                base_name = rename_map.get(uploaded_file.name, uploaded_file.name.rsplit(".", 1)[0])
                final_filename = base_name + ".xlsx"

                # Send
                with st.spinner(f"Sending {final_filename} to WhatsApp..."):
                    send_to_whatsapp(processed_bytes, final_filename, WA_TO)

                st.success(f"✅ Sent: {final_filename}")
                time.sleep(7)

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")


# Footer
st.markdown("---")
st.markdown("Built with Streamlit • Powered by WasenderAPI")
