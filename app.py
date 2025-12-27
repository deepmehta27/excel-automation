import streamlit as st
import os
import io
import pandas as pd
import requests
import re
import unicodedata
import base64
import time
import boto3
import json
from botocore.exceptions import ClientError

# =========================
# AWS DynamoDB Setup
# =========================
dynamodb = boto3.resource("dynamodb", region_name="eu-west-2")
table = dynamodb.Table("allowed_columns")

# =========================
# Page config
# =========================
st.set_page_config(
    page_title="Excel to WhatsApp",
    page_icon="📊",
    layout="centered"
)

# =========================
# Environment variables
# =========================
RAW_SENDERS = os.getenv("WASENDER_SENDERS", "[]")
WASENDER_SENDERS = json.loads(RAW_SENDERS)

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s or "")).strip()

    # Exclude columns starting with special chars
    if not re.match(r"^[A-Za-z0-9]", s):
        return ""

    s = s.replace("\u00a0", " ")
    s = s.lower()
    s = re.sub(r"[\s._\-]+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s


def load_allowed_columns():
    try:
        resp = table.get_item(Key={"type": "columns"})
        return resp.get("Item", {}).get("values", [])
    except ClientError as e:
        st.error(f"DynamoDB read failed: {e}")
        return []


def save_allowed_columns(cols: list[str]):
    table.put_item(
        Item={
            "type": "columns",
            "values": sorted(set(cols))
        }
    )

# =========================
# Excel Processing
# =========================
def process_excel(file_bytes: bytes, wanted_norm: list[str], file_ext: str) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)

    if file_ext == "csv":
        df = pd.read_csv(bio, dtype=str, header=None)
    else:
        engine = "xlrd" if file_ext == "xls" else "openpyxl"
        df = pd.read_excel(bio, dtype=str, engine=engine, header=None)

    df = df.fillna("")

    best_idx = -1
    best_score = -1

    for i in range(min(300, len(df))):
        vals = [str(v) for v in df.iloc[i].tolist()]
        score = sum(1 for v in vals if norm(v) in wanted_norm)
        if score > best_score:
            best_idx, best_score = i, score

    if best_idx < 0:
        raise ValueError("Could not find header row matching allow-list")

    header_vals = [str(v) for v in df.iloc[best_idx].tolist()]
    body = df.iloc[best_idx + 1 :].reset_index(drop=True)
    body.columns = header_vals

    nm = {c: norm(c) for c in body.columns}
    keep = [c for c in body.columns if nm[c] in wanted_norm]

    if not keep:
        raise ValueError("No matching columns found")

    return body[keep]

# =========================
# WhatsApp Sender
# =========================
def send_to_whatsapp(file_bytes: bytes, filename: str, sender: dict) -> dict:
    base64_data = base64.b64encode(file_bytes).decode("utf-8")
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    data_url = f"data:{mime_type};base64,{base64_data}"

    headers = {
        "Authorization": f"Bearer {sender['api_key']}",
        "Content-Type": "application/json",
    }

    upload_response = requests.post(
        "https://www.wasenderapi.com/api/upload",
        headers=headers,
        json={"base64": data_url},
        timeout=60,
    )

    if upload_response.status_code >= 300:
        raise Exception(f"Upload failed: {upload_response.text}")

    temp_url = upload_response.json().get("publicUrl")

    send_response = requests.post(
        "https://www.wasenderapi.com/api/send-message",
        headers=headers,
        json={
            "sessionId": sender["session_id"],
            "to": sender["wa_to"],
            "documentUrl": temp_url,
            "fileName": filename,
        },
        timeout=60,
    )

    if send_response.status_code >= 300:
        raise Exception(f"Send failed: {send_response.text}")

    return send_response.json()


# =========================
# UI
# =========================
st.title("📊 Excel to WhatsApp")
st.markdown("Upload Excel files and send filtered data to WhatsApp.")

# 🔧 INIT SESSION STATE (IMPORTANT)
if "column_added_msg" not in st.session_state:
    st.session_state.column_added_msg = None
    
# -------- Permanent Column Input --------
st.markdown("### ➕ Add Column (Permanent)")

new_col = st.text_input(
    "Enter column name to permanently allow",
    placeholder="e.g. REMARKS, CATEGORY, DATE"
)

if st.button("Add Column Permanently"):
    if not new_col.strip():
        st.warning("Column name cannot be empty")
    else:
        cols = load_allowed_columns()

        # ✅ SUPPORT COMMA-SEPARATED INPUT
        input_cols = [c.strip() for c in new_col.split(",") if c.strip()]
        existing_norms = [norm(c) for c in cols]

        added = []
        for col in input_cols:
            if norm(col) not in existing_norms:
                cols.append(col)
                added.append(col)

        if not added:
            st.warning("All columns already exist")
        else:
            save_allowed_columns(cols)
            st.session_state.column_added_msg = (
                f"✅ Added columns: {', '.join(added)}"
            )
            st.rerun()
            
if st.session_state.column_added_msg:
    st.success(st.session_state.column_added_msg)

# -------- Load Columns --------
DB_COLS = load_allowed_columns()
if not DB_COLS:
    st.warning("⚠️ No allowed columns found in database")

WANTED_NORM = [norm(c) for c in DB_COLS]

st.markdown("### 📱 Select WhatsApp Sender")

labels = [s["label"] for s in WASENDER_SENDERS]

selected_label = st.selectbox(
    "Send files from",
    labels
)

selected_sender = next(
    s for s in WASENDER_SENDERS if s["label"] == selected_label
)

# -------- File Upload --------
uploaded_files = st.file_uploader(
    "Select Excel Files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

rename_map = {}

if uploaded_files:
    st.write("### Rename files (optional)")
    for file in uploaded_files:
        rename_map[file.name] = st.text_input(
            f"Rename for {file.name}",
            value=file.name.rsplit(".", 1)[0],
            key=file.name
        )

# -------- Send Button --------
if st.button("📤 Upload & Send to WhatsApp", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("⚠️ Please select at least one file")
    else:
        try:
            for uploaded_file in uploaded_files:
                file_bytes = uploaded_file.read()
                file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()

                filtered_df = process_excel(file_bytes, WANTED_NORM, file_ext)

                output = io.BytesIO()
                filtered_df.to_excel(output, index=False, engine="openpyxl")
                processed_bytes = output.getvalue()

                base_name = rename_map.get(uploaded_file.name, uploaded_file.name)
                final_filename = base_name + ".xlsx"

                with st.spinner(f"Sending {final_filename}..."):
                    send_to_whatsapp(
                                processed_bytes,
                                final_filename,
                                selected_sender
                            )

                st.success(f"✅ Sent: {final_filename}")
                time.sleep(7)

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

st.markdown("---")
st.markdown("Built with Streamlit • Powered by WasenderAPI")
