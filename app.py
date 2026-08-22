import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
pd.options.mode.string_storage = "python"  
import base64
import gspread
import json
import jinja2
import re
import tempfile
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials as HumanCredentials
from google.oauth2.service_account import Credentials as BotCredentials
from googleapiclient.http import MediaFileUpload

import pdf_engine

# ==========================================
# 1. GLOBAL SETUP & CSS
# ==========================================
st.set_page_config(page_title="Rattan Freezone Logistics", page_icon="🚢", layout="wide")

def get_img_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as f: return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
    return None

left_logo_b64 = get_img_b64("logo_left.png")
right_logo_b64 = get_img_b64("logo_right.png")

left_img_tag = f'<img src="{left_logo_b64}" style="height: 65px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.2);">' if left_logo_b64 else ''
right_img_tag = f'<img src="{right_logo_b64}" style="height: 65px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.2);">' if right_logo_b64 else ''

st.markdown(f"""
<style>
    .stApp {{ background-color: #f8fafc; color: #1e293b; }}
    .custom-header {{
        background: linear-gradient(135deg, #e60000 0%, #8b0000 100%);
        color: white; padding: 20px 30px; border-radius: 12px;
        box-shadow: 0 8px 20px rgba(220, 38, 38, 0.25), inset 0 2px 10px rgba(255,255,255,0.1);
        display: flex; justify-content: space-between; align-items: center;
        position: relative; overflow: hidden; margin-bottom: 20px; margin-top: 10px;
    }}
    .header-center {{ display: flex; flex-direction: column; align-items: center; text-align: center; flex-grow: 1; }}
    .header-title {{ font-family: 'Arial', sans-serif; font-size: 28px; font-weight: 900; letter-spacing: 2px; margin: 0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }}
    .header-subtitle {{ color: #fecaca; font-size: 13px; margin: 0; margin-top: 4px; font-weight: 500; letter-spacing: 1px; display: flex; align-items: center; justify-content: center; gap: 10px; }}
    .header-badge {{ background-color: #ffffff; color: #dc2626; padding: 2px 8px; border-radius: 20px; font-weight: 800; font-size: 11px; box-shadow: 0 2px 5px rgba(0,0,0,0.15); }}
    [data-testid="stExpander"] {{
        background-color: #ffffff !important; border: 1px solid #e2e8f0; border-top: 4px solid #dc2626;
        border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.04); margin-bottom: 15px;
    }}
    [data-testid="stExpander"] summary p {{ font-weight: 700 !important; color: #1e293b !important; font-size: 1.02rem !important; }}
    [data-testid="stExpander"] p, [data-testid="stExpander"] h3, [data-testid="stExpander"] h4, [data-testid="stExpander"] h5 {{ color: #1e293b !important; }}
</style>

<div class="custom-header">
    <div>{left_img_tag}</div>
    <div class="header-center">
        <h1 class="header-title">RATTAN FREEZONE LOGISTICS</h1>
        <p class="header-subtitle">Pennywise Plaza | East Chaguanas <span class="header-badge">VAT REG# 202049</span></p>
    </div>
    <div>{right_img_tag}</div>
</div>
""", unsafe_allow_html=True)

for folder in ["uploaded_docs", "logos", "signatures", "watermarks", "templates", "generated_documents"]:
    if not os.path.exists(folder): os.makedirs(folder)

def to_decimal(val):
    try:
        if isinstance(val, (int, float)): return round(float(val), 2)
        clean_val = re.sub(r'[^\d.-]', '', str(val))
        return round(float(clean_val), 2)
    except: return 0.00

def safe_qty_parse(val):
    try:
        if isinstance(val, (int, float)): return int(val)
        val_str = str(val).replace(",", "").strip()
        if not val_str or val_str.lower() in ['nan', 'none', 'n/a']: return 0
        return int(float(val_str))
    except: return 0

# ==========================================
# 2. CONSTANTS & DATA SCHEMA
# ==========================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1Rcifpu4GRFAYFPQNBGrl96DpHXDiQM1JYys-Dhi0rrU/edit?usp=sharing"
ROOT_FOLDER_ID = "1GtZk2jfAHGqttyZVP9E8G4TA_MNrV9Pp"

ALL_COUNTRIES = ["", "USA", "China", "UK", "Canada", "Brazil", "Mexico", "Panama", "Japan", "Germany", "India", "France", "Italy", "South Korea", "Spain", "Australia", "Taiwan", "Netherlands", "Vietnam", "Malaysia", "Singapore", "South Africa", "UAE", "Saudi Arabia", "Switzerland", "Sweden", "Poland", "Belgium", "Thailand", "Indonesia", "Turkey", "Philippines", "Ireland", "Other"]

SYSTEM_DOCS = ["Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment", "Warehouse Delivery Note", "Finance Cost Statement"]
EXTERNAL_DOCS = ["Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", "Other Documents", "Miscellaneous Supporting Doc"]
ALL_DOCS = SYSTEM_DOCS + EXTERNAL_DOCS

LOG_COLUMNS = [
    "Row_UID", "Invoice No", "Client Name", "Container #", "Country of Origin", "ETA", 
    "Lodged Status", "Shipment Status", "NALDO", "Total Cartons", "B/L Number", "Freight", "Cargo Notes",
    "Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment", 
    "Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", 
    "Other Documents", "Miscellaneous Supporting Doc",
    "Subtotal (USD)", "Import Duties (TTD)", "Customs Deposit (TTD)", "Import VAT Paid (TTD)",
    "Additional Port Charges (TTD)", "Brokerage & Clearance Fees (TTD)", "Management Fees (TTD)",
    "Warehouse Delivery Note", "Finance Cost Statement", "Invoice Date"
]

# ==========================================
# 3. GOOGLE CLOUD & DRIVE SERVICE FUNCTIONS
# ==========================================
def get_gspread_client():
    creds_dict = json.loads(st.secrets["google_api"]["credentials"])
    creds = BotCredentials.from_service_account_info(
        creds_dict, 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]
    )
    return gspread.authorize(creds)

def get_drive_service():
    token_info = json.loads(st.secrets["google_drive_human"]["token"])
    creds = HumanCredentials.from_authorized_user_info(token_info)
    return build('drive', 'v3', credentials=creds)

@st.cache_data(ttl=60)
def load_log_data():
    try: 
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        rows = ws.get_all_values()
        if not rows or len(rows) < 2: return pd.DataFrame(columns=LOG_COLUMNS)
        headers = rows[0]
        data = rows[1:]
        df = pd.DataFrame(data, columns=headers, dtype=object)
        for col in LOG_COLUMNS:
            if col not in df.columns: df[col] = ""
        return df
    except Exception as e: 
        st.error(f"Failed to load data: {e}")
        return pd.DataFrame(columns=LOG_COLUMNS)

def save_log_data(df):
    try:
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        ws.clear()
        df = df.copy()
        for col in df.columns: df[col] = df[col].astype(str).replace(['nan', 'None'], '')
        for col in LOG_COLUMNS:
            if col not in df.columns: df[col] = ""
        df = df[LOG_COLUMNS]
        ws.update([df.columns.values.tolist()] + df.values.tolist())
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Failed to sync with Google Sheets: {e}")
        return False

def sync_local_pdf_to_google_drive(local_pdf_path, client_name, bl_number):
    if not os.path.exists(local_pdf_path):
        return None, "Local file not found on disk."
        
    try:
        drive = get_drive_service()
        file_name = os.path.basename(local_pdf_path)
        safe_client_name = str(client_name).replace("'", "\\'")
        safe_bl_number = str(bl_number).replace("'", "\\'")
        
        folders = drive.files().list(
            q=f"name='{safe_client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)"
        ).execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(
            body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}
        ).execute()['id']
        
        bl_folders = drive.files().list(
            q=f"name='{safe_bl_number}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)"
        ).execute().get('files', [])
        bl_folder_id = bl_folders[0]['id'] if bl_folders else drive.files().create(
            body={"name": str(bl_number), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}
        ).execute()['id']
        
        media = MediaFileUpload(local_pdf_path, mimetype='application/pdf', resumable=True)
        existing_files = drive.files().list(
            q=f"name='{file_name}' and '{bl_folder_id}' in parents and trashed=false",
            fields="files(id, webViewLink)"
        ).execute().get('files', [])
        
        if existing_files:
            file_id = existing_files[0]['id']
            final_file = drive.files().update(fileId=file_id, media_body=media, fields='id, webViewLink').execute()
        else:
            file_metadata = {'name': file_name, 'parents': [bl_folder_id]}
            final_file = drive.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
            
        return final_file.get('webViewLink'), None
    except Exception as e:
        return None, str(e)

def upload_system_pdf_to_drive(html_content, file_name, client_name, invoice_no):
    if not html_content: return "Pending Upload"
    try:
        from xhtml2pdf import pisa 
        drive = get_drive_service()
        safe_client_name = str(client_name).replace("'", "\\'")
        safe_invoice_no = str(invoice_no).replace("'", "\\'")
        
        folders = drive.files().list(q=f"name='{safe_client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        inv_folders = drive.files().list(q=f"name='{safe_invoice_no}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        inv_folder_id = inv_folders[0]['id'] if inv_folders else drive.files().create(body={"name": str(invoice_no), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
        with open(temp_pdf_path, "w+b") as result_file:
            pisa_status = pisa.CreatePDF(html_content, dest=result_file)
        if pisa_status.err:
            if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
            return "Upload Failed"
        
        pdf_media = MediaFileUpload(temp_pdf_path, mimetype='application/pdf', resumable=True)
        existing_files = drive.files().list(q=f"name='{file_name}' and '{inv_folder_id}' in parents and trashed=false", fields="files(id, webViewLink)").execute().get('files', [])
        if existing_files:
            file_id = existing_files[0]['id']
            final_pdf = drive.files().update(fileId=file_id, media_body=pdf_media, fields='id, webViewLink').execute()
        else:
            pdf_metadata = {'name': file_name, 'parents': [inv_folder_id]}
            final_pdf = drive.files().create(body=pdf_metadata, media_body=pdf_media, fields='id, webViewLink').execute()
        if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
        return final_pdf.get('webViewLink', 'Upload Failed')
    except Exception as e: return "Upload Failed"

def upload_physical_file_to_drive(uploaded_file, file_name, client_name, invoice_no):
    if not uploaded_file: return None
    try:
        drive = get_drive_service()
        safe_client_name = str(client_name).replace("'", "\\'")
        safe_invoice_no = str(invoice_no).replace("'", "\\'")
        
        folders = drive.files().list(q=f"name='{safe_client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        inv_folders = drive.files().list(q=f"name='{safe_invoice_no}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        inv_folder_id = inv_folders[0]['id'] if inv_folders else drive.files().create(body={"name": str(invoice_no), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        file_ext = os.path.splitext(uploaded_file.name)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name
            
        media = MediaFileUpload(temp_path, resumable=True)
        existing_files = drive.files().list(q=f"name='{file_name}' and '{inv_folder_id}' in parents and trashed=false", fields="files(id, webViewLink)").execute().get('files', [])
        if existing_files:
            file_id = existing_files[0]['id']
            file = drive.files().update(fileId=file_id, media_body=media, fields='id, webViewLink').execute()
        else:
            file_metadata = {'name': file_name, 'parents': [inv_folder_id]}
            file = drive.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        if os.path.exists(temp_path): os.remove(temp_path)
        return file.get('webViewLink')
    except Exception as e: return None

def get_eta_status(eta_date, shipment_status):
    if shipment_status == "Delivered": return "✅ DELIVERED", "#00b050"
    try:
        days_diff = (eta_date - datetime.now().date()).days
        if days_diff < 0: return "⚠️ OVERDUE", "#FF4500"
        if 0 <= days_diff <= 7: return "🔴 URGENT", "#FF0000"
        if 8 <= days_diff <= 14: return "🟡 APPROACHING", "#FFD700"
        return "🟢 IN TRANSIT", "#008000"
    except: return "TBD", "#808080"

def get_entity_profile(file_name, entity_name):
    profile = {"Name": entity_name, "Address": "Main Office Hub", "Template": "classic.html"}
    if os.path.exists(file_name) and os.path.getsize(file_name) > 0:
        df = pd.read_csv(file_name)
        match = df[df["Name"] == entity_name]
        if not match.empty:
            for col in df.columns: profile[col] = match.iloc[0][col]
    return profile

def get_supplier_mapping(supplier):
    if os.path.exists("supplier_mappings.csv") and os.path.getsize("supplier_mappings.csv") > 0:
        df = pd.read_csv("supplier_mappings.csv")
        match = df[df["Supplier"] == supplier]
        if not match.empty: return match.iloc[0]["DescCol"], match.iloc[0]["QtyCol"], match.iloc[0]["PriceCol"]
    return "-- Select --", "-- Select --", "-- Select --"

def save_supplier_mapping(supplier, desc, qty, price):
    df = pd.read_csv("supplier_mappings.csv") if os.path.exists("supplier_mappings.csv") else pd.DataFrame(columns=["Supplier", "DescCol", "QtyCol", "PriceCol"])
    df = df[df["Supplier"] != supplier]
    df = pd.concat([df, pd.DataFrame([{"Supplier": supplier, "DescCol": desc, "QtyCol": qty, "PriceCol": price}])], ignore_index=True)
    df.to_csv("supplier_mappings.csv", index=False)

def display_html_preview(raw_html):
    preview_html = f'<div style="background-color: white; padding: 40px; margin: 10px auto; border-radius: 5px; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); max-width: 900px; color: #333333;">{raw_html}</div>'
    components.html(preview_html, height=750, scrolling=True)

# ==========================================
# 4. NOMINATED AGENCY PORTAL (MODULE 3)
# ==========================================
def render_document_workspace(doc_title, doc_filename, bl_no, generate_callback, key_prefix, entity_client_name="Corinthian Pins Limited"):
    st.markdown(f"### {doc_title}")
    
    shipment_dir = os.path.join(pdf_engine.OUTPUT_DIR, str(bl_no).strip())
    file_path = os.path.join(shipment_dir, doc_filename)
    
    col_gen, col_status = st.columns([1, 2])
    with col_gen:
        if st.button(f"⚙️ Generate / Refresh {doc_title}", key=f"btn_gen_{key_prefix}", type="primary", use_container_width=True):
            filepath = generate_callback()
            st.success(f"Generated successfully: `{os.path.basename(filepath)}`")
            st.rerun()

    if os.path.exists(file_path):
        st.markdown("---")
        with open(file_path, "rb") as f:
            pdf_bytes = f.read()
            base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')

        c_dl, c_vault = st.columns(2)
        with c_dl:
            st.download_button(
                label=f"📥 Download {doc_title} (PDF)",
                data=pdf_bytes,
                file_name=os.path.basename(file_path),
                mime="application/pdf",
                key=f"dl_{key_prefix}",
                use_container_width=True
            )
        with c_vault:
            if st.button(f"☁️ Save to Vault / Google Drive", key=f"vault_{key_prefix}", use_container_width=True):
                with st.spinner("Synchronizing document to Google Drive Vault..."):
                    drive_url, err_msg = sync_local_pdf_to_google_drive(file_path, entity_client_name, bl_no)
                    if drive_url:
                        st.session_state[f"vault_link_{key_prefix}"] = drive_url
                        st.session_state[f"vault_status_{key_prefix}"] = True
                    else:
                        st.error(f"Vault Upload Failed: {err_msg}")

        if st.session_state.get(f"vault_status_{key_prefix}"):
            drive_link = st.session_state.get(f"vault_link_{key_prefix}", "#")
            st.success("✅ Saved and synchronized to Cloud Vault!")
            st.link_button("🔗 Open Document in Google Drive", url=drive_link, use_container_width=True)

        st.markdown("##### 👁️ Document Preview")
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf" style="border: 1px solid #cbd5e0; border-radius: 6px;"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
    else:
        st.info(f"ℹ️ {doc_title} has not been generated yet for B/L **{bl_no}**. Click the button above to generate.")

def render_nominated_agency_portal():
    st.subheader("🏛️ Nominated Agency Supply Chain Portal")
    st.caption("Corinthian Pins Limited | Trinidad Freight Solutions Limited | Rattans Freezone Limited")
    st.divider()

    active_shell_uid = st.session_state.get("active_shell_uid", "")

    if not active_shell_uid or active_shell_uid == "-- Choose Active Workspace --":
        st.warning("⚠️ No Active Workspace Shell selected. Please create or select an Active Workspace Shell from the top bar above to load shipment data from Google Sheets.")
        return

    df_current = load_log_data()
    match_row = df_current[df_current['Row_UID'].astype(str).str.strip() == active_shell_uid.strip()]

    if match_row.empty:
        st.error(f"Selected Workspace Shell ({active_shell_uid}) was not found in the Master Google Sheet.")
        return

    row_data = match_row.iloc[0].to_dict()

    # FORCE RELOAD SHEET VALUES INTO SESSION STATE WHEN WORKSPACE SWITCHES
    if st.session_state.get("nom_loaded_shell_uid") != active_shell_uid:
        st.session_state["nom_bl"] = str(row_data.get("B/L Number", "")).strip() or "BL-2026-001"
        st.session_state["nom_cntr"] = str(row_data.get("Container #", "")).strip() or "CNTR-40912"

        usd_val = to_decimal(row_data.get("Subtotal (USD)", 50000.00))
        st.session_state["nom_usd"] = usd_val if usd_val > 0 else 50000.00

        duty_val = to_decimal(row_data.get("Import Duties (TTD)", 12000.00))
        st.session_state["nom_dut"] = duty_val if duty_val > 0 else 12000.00

        vat_val = to_decimal(row_data.get("Import VAT Paid (TTD)", 11250.00))
        st.session_state["nom_vat"] = vat_val if vat_val > 0 else 11250.00

        dep_val = to_decimal(row_data.get("Customs Deposit (TTD)", 30000.00))
        st.session_state["nom_dep"] = dep_val if dep_val > 0 else 30000.00

        port_val = to_decimal(row_data.get("Additional Port Charges (TTD)", 4250.00))
        st.session_state["nom_port"] = port_val if port_val > 0 else 4250.00

        dem_val = to_decimal(row_data.get("Brokerage & Clearance Fees (TTD)", 2500.00))
        st.session_state["nom_dem"] = dem_val if dem_val > 0 else 2500.00

        fee_val = to_decimal(row_data.get("Management Fees (TTD)", 25000.00))
        st.session_state["nom_fee"] = fee_val if fee_val > 0 else 25000.00

        st.session_state["nom_loaded_shell_uid"] = active_shell_uid

    inv_no_disp = str(row_data.get('Invoice No', 'N/A')).strip() or 'N/A'
    client_disp = str(row_data.get('Client Name', 'N/A')).strip() or 'N/A'
    st.success(f"🔗 **Live Workspace Active:** `{active_shell_uid}` | **Invoice:** `{inv_no_disp}` | **Client:** `{client_disp}`")

    st.markdown("#### 1. Shipment & Port Outlay Details")
    col1, col2, col3 = st.columns(3)

    with col1:
        bl_no = st.text_input("Bill of Lading (B/L) No.", key="nom_bl")
        container_no = st.text_input("Container No.", key="nom_cntr")
        shipment_status = st.selectbox(
            "Operational Shipment Status",
            [
                "Pre-Clearance (Pending Deposit)",
                "Cleared & In Transit to Freezone",
                "Delivered (Draft / Reconciliation Pending)",
                "Delivered (Finalized & Reconciled)"
            ],
            key="nom_status"
        )

    with col2:
        usd_cargo_val = st.number_input("Foreign Cargo Valuation (USD)", step=1000.00, key="nom_usd")
        exchange_rate = st.number_input("Exchange Rate (TTD/USD)", value=6.80, step=0.01, key="nom_fx")
        ttd_cargo_val = usd_cargo_val * exchange_rate
        st.info(f"Converted TTD Cargo Value: **${ttd_cargo_val:,.2f} TTD**")

    with col3:
        bundled_service_fee = st.number_input("Corinthian Bundled Agency Fee (TTD)", step=500.00, key="nom_fee")
        contra_deposit_paid = st.number_input("Pre-Funded Advance Port Deposit (TTD)", step=1000.00, key="nom_dep")

    st.divider()

    st.markdown("#### 2. Itemized Statutory Port Outlays")
    outlay_col1, outlay_col2 = st.columns(2)

    with outlay_col1:
        customs_duty = st.number_input("Customs Import Duty (TTD)", step=500.00, key="nom_dut")
        import_vat = st.number_input("Customs Import VAT (12.5%) (TTD)", step=500.00, key="nom_vat")

    with outlay_col2:
        port_handling = st.number_input("Port Authority Handling & Storage (TTD)", step=250.00, key="nom_port")
        shipping_line_demurrage = st.number_input("Shipping Line Demurrage & Fees (TTD)", step=250.00, key="nom_dem")

    port_items = [
        {"desc": "Customs Import Duty", "amount": customs_duty},
        {"desc": "Customs Import VAT (12.5%)", "amount": import_vat},
        {"desc": "Port Authority Handling & Storage", "amount": port_handling},
        {"desc": "Shipping Line Local Demurrage & Terminal Charges", "amount": shipping_line_demurrage}
    ]

    total_port_outlays = sum(item["amount"] for item in port_items)
    service_vat = bundled_service_fee * 0.125
    gross_shipment_total = ttd_cargo_val + total_port_outlays + bundled_service_fee + service_vat
    net_contra_due = gross_shipment_total - ttd_cargo_val - contra_deposit_paid

    # Save Back to Google Sheet Trigger
    if st.button("💾 Save & Sync Outlays to Master Sheet", type="primary", use_container_width=True):
        with st.spinner("Syncing outlays to Google Sheet Master Ledger..."):
            df_update = load_log_data()
            matches = df_update.index[df_update['Row_UID'].astype(str).str.strip() == active_shell_uid.strip()].tolist()
            if matches:
                idx = matches[0]
                df_update.at[idx, "B/L Number"] = str(bl_no).strip()
                df_update.at[idx, "Container #"] = str(container_no).strip()
                df_update.at[idx, "Subtotal (USD)"] = f"{usd_cargo_val:,.2f}"
                df_update.at[idx, "Import Duties (TTD)"] = f"{customs_duty:,.2f}"
                df_update.at[idx, "Import VAT Paid (TTD)"] = f"{import_vat:,.2f}"
                df_update.at[idx, "Customs Deposit (TTD)"] = f"{contra_deposit_paid:,.2f}"
                df_update.at[idx, "Additional Port Charges (TTD)"] = f"{port_handling:,.2f}"
                df_update.at[idx, "Brokerage & Clearance Fees (TTD)"] = f"{shipping_line_demurrage:,.2f}"
                df_update.at[idx, "Management Fees (TTD)"] = f"{bundled_service_fee:,.2f}"
                if save_log_data(df_update):
                    # Force reload session_state on next run to reflect saved values
                    st.session_state["nom_loaded_shell_uid"] = ""
                    st.success("✅ Nominated Agency parameters successfully synced to Google Sheet Source of Truth!")
                    st.rerun()

    st.divider()

    st.markdown("#### 3. Document Generation & Action Triggers")
    stage1_tab, stage2_tab = st.tabs([
        "📄 Stage 1: Pre-Clearance & Operations", 
        "🚀 Stage 2: Post-Delivery Final Financial Discharge"
    ])

    with stage1_tab:
        subtab_port, subtab_service, subtab_tfs = st.tabs([
            "⚓ Advance Port Disbursement Request",
            "💼 Service Fee Disbursement Request",
            "🚚 Internal TFS Freight Invoice"
        ])
        
        with subtab_port:
            render_document_workspace(
                doc_title="Advance Port Disbursement Request",
                doc_filename=f"Disbursement_Request_Port_{bl_no}.pdf",
                bl_no=bl_no,
                generate_callback=lambda: pdf_engine.generate_port_disbursement_request(bl_no, container_no, port_items=port_items),
                key_prefix="stg1_port",
                entity_client_name="Corinthian Pins Limited"
            )
            
        with subtab_service:
            render_document_workspace(
                doc_title="Service Fee Disbursement Request",
                doc_filename=f"Disbursement_Request_Services_{bl_no}.pdf",
                bl_no=bl_no,
                generate_callback=lambda: pdf_engine.generate_service_disbursement_request(bl_no, container_no, bundled_service_fee=bundled_service_fee),
                key_prefix="stg1_service",
                entity_client_name="Corinthian Pins Limited"
            )
            
        with subtab_tfs:
            render_document_workspace(
                doc_title="Internal Upstream Freight Invoice (TFS)",
                doc_filename=f"TFS_Internal_Invoice_{bl_no}.pdf",
                bl_no=bl_no,
                generate_callback=lambda: pdf_engine.generate_internal_tfs_invoice(bl_no, container_no, tfs_base=5500.00),
                key_prefix="stg1_tfs",
                entity_client_name="Trinidad Freight Solutions Limited"
            )

    with stage2_tab:
        if shipment_status == "Delivered (Finalized & Reconciled)":
            st.markdown("##### Pre-Flight Settlement Audit Summary")
            summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)
            summary_col1.metric("Gross Shipment Total", f"${gross_shipment_total:,.2f} TTD")
            summary_col2.metric("Less Foreign Contra", f"-${ttd_cargo_val:,.2f} TTD")
            summary_col3.metric("Less Port Deposit", f"-${contra_deposit_paid:,.2f} TTD")
            summary_col4.metric("NET BALANCE PAYABLE", f"${net_contra_due:,.2f} TTD")
            
            audit_confirm = st.checkbox(f"I confirm container delivery and final port outlay reconciliation for B/L {bl_no}", key="audit_nom_chk")
            
            if audit_confirm:
                subtab_master, subtab_receipt = st.tabs([
                    "🧾 Master Agency Tax Invoice",
                    "✅ Official Payment Receipt & Account Discharge"
                ])
                
                with subtab_master:
                    render_document_workspace(
                        doc_title="Master Agency & Disbursement Tax Invoice",
                        doc_filename=f"Master_Tax_Invoice_{bl_no}.pdf",
                        bl_no=bl_no,
                        generate_callback=lambda: pdf_engine.generate_master_invoice(
                            bl_no, container_no, 
                            usd_cargo_val=usd_cargo_val, exchange_rate=exchange_rate, 
                            port_items=port_items, bundled_service_fee=bundled_service_fee, 
                            contra_deposit_paid=contra_deposit_paid
                        ),
                        key_prefix="stg2_master",
                        entity_client_name="Corinthian Pins Limited"
                    )
                    
                with subtab_receipt:
                    render_document_workspace(
                        doc_title="Official Payment Receipt & Account Discharge",
                        doc_filename=f"Official_Receipt_{bl_no}.pdf",
                        bl_no=bl_no,
                        generate_callback=lambda: pdf_engine.generate_official_receipt(bl_no, container_no, amount_paid=net_contra_due),
                        key_prefix="stg2_receipt",
                        entity_client_name="Corinthian Pins Limited"
                    )
            else:
                st.info("Check the confirmation box above to unlock final invoicing and discharge workspaces.")
        else:
            st.info("To generate Master Tax Invoices and Discharge Receipts, change Operational Status to **Delivered (Finalized & Reconciled)**.")

# ==========================================
# 5. CARGO TRACKER & DOCUMENT HUB VIEWS
# ==========================================
def render_master_log():
    st.subheader("📋 Cargo Tracker & Vault")
    df = load_log_data()

    if df.empty:
        st.info("No data found. Create a new shell to begin.")
    else:
        for idx, row in df.iterrows():
            row_uid = str(row.get('Row_UID', ''))
            if not row_uid.strip(): continue 
                
            inv_no = str(row.get('Invoice No', ''))
            display_inv = inv_no.strip() if inv_no.strip() else "[Blank Entry]"
            client_name = str(row.get('Client Name', 'Unknown Client'))
            ship_status = str(row.get("Shipment Status", "Active"))
            total_cartons = str(row.get("Total Cartons", "0"))
            cont_no = str(row.get("Container #", "")).strip()
            bl_no = str(row.get("B/L Number", "")).strip()
            origin_no = str(row.get("Country of Origin", "")).strip()
            lodged_val = str(row.get("Lodged Status", "")).strip()

            raw_eta = row.get("ETA")
            timestamp = pd.to_datetime(raw_eta, errors='coerce')
            current_date = timestamp.date() if not pd.isna(timestamp) else datetime.now().date()
            status_label, _ = get_eta_status(current_date, ship_status)
            naldo_val = str(row.get("NALDO", "No")).strip().upper()
            
            header_text = (f"📦 TOTAL CTNS: {total_cartons} | {status_label} | ETA: {current_date} | "
                           f"Container #: {cont_no} | B/L Number: {bl_no} | "
                           f"INV: {display_inv} | Origin: {origin_no} | Lodged: {lodged_val}")

            with st.expander(header_text):
                col1, col2, col3, col4, col5, col6 = st.columns(6)
                with col1: new_cont = st.text_input("Container #", value=str(row.get("Container #", "")), key=f"cont_{idx}")
                with col2: new_orig = st.selectbox("Country of Origin", ALL_COUNTRIES, index=ALL_COUNTRIES.index(row.get("Country of Origin", "")) if row.get("Country of Origin", "") in ALL_COUNTRIES else 0, key=f"orig_{idx}")
                with col3: new_eta = st.date_input("ETA (Container Arrival)", value=current_date, key=f"eta_{idx}")
                with col4: new_lodg = st.radio("Lodged", ["Yes", "No"], index=0 if row.get("Lodged Status") == "Yes" else 1, horizontal=True, key=f"lodged_{idx}")
                with col5: new_stat = st.selectbox("Shipment Status", ["Active", "Delivered"], index=0 if ship_status != "Delivered" else 1, key=f"stat_{idx}")
                with col6: new_naldo = st.radio("NALDO Code", ["Yes", "No"], index=0 if naldo_val == "YES" else 1, horizontal=True, key=f"naldo_{idx}")
                
                col7, col8, col9 = st.columns(3)
                with col7: new_bl = st.text_input("B/L Number", value=str(row.get("B/L Number", "")), key=f"bl_{idx}")
                with col8: new_freight = st.text_input("Freight (USD)", value=str(row.get("Freight", "")), key=f"fr_{idx}")
                with col9: new_subtotal = st.text_input("Subtotal (USD) [Auto]", value=str(row.get("Subtotal (USD)", "")), key=f"sub_{idx}")

                st.write("---")
                st.markdown("#### 💰 Post-Clearance TTD Expenses & Financial Reconciliation")
                c_f1, c_f2, c_f3 = st.columns(3)
                with c_f1: new_duties = st.text_input("Import Duties ($ TTD)", value=str(row.get("Import Duties (TTD)", "")), key=f"dut_{idx}")
                with c_f2: new_deposit = st.text_input("Customs Deposit ($ TTD)", value=str(row.get("Customs Deposit (TTD)", "")), key=f"dep_{idx}")
                with c_f3: new_vat = st.text_input("Import VAT Paid ($ TTD)", value=str(row.get("Import VAT Paid (TTD)", "")), key=f"vat_{idx}")
                
                c_f4, c_f5, c_f6 = st.columns(3)
                with c_f4: new_port = st.text_input("Additional Port Charges ($ TTD)", value=str(row.get("Additional Port Charges (TTD)", "")), key=f"port_{idx}")
                with c_f5: new_brokerage = st.text_input("Brokerage & Clearance Fees ($ TTD)", value=str(row.get("Brokerage & Clearance Fees (TTD)", "")), key=f"brok_{idx}")
                with c_f6: new_mgmt = st.text_input("Management Fees ($ TTD)", value=str(row.get("Management Fees (TTD)", "")), key=f"mgmt_{idx}")

                st.write("---")
                st.markdown("#### Document Vault (12-Slot Matrix)")
                
                grid = st.columns(6)
                upload_cache = {}

                for i, slot in enumerate(ALL_DOCS):
                    with grid[i % 6]:
                        st.markdown(f"**{slot}**")
                        file_link = str(row.get(slot, ""))
                        if file_link.startswith("http"):
                            clean_link = file_link
                            match = re.search(r'/d/([a-zA-Z0-9_-]+)', file_link)
                            if match:
                                file_id = match.group(1)
                                clean_link = f"https://drive.google.com/uc?export=download&id={file_id}"
                            st.link_button("📄 View Document", url=clean_link, key=f"view_{idx}_{i}", use_container_width=True)
                        else:
                            st.button("Pending Upload", disabled=True, key=f"pend_{idx}_{i}", use_container_width=True)
                        if slot in EXTERNAL_DOCS:
                            uploaded_file = st.file_uploader(f"Upload {slot}", key=f"up_{idx}_{i}", label_visibility="collapsed")
                            if uploaded_file: upload_cache[slot] = uploaded_file
                
                st.write("---")
                m_btn1, m_btn2 = st.columns(2)
                with m_btn1:
                    if st.button("💾 Save Shipment Updates", key=f"save_{idx}", type="primary", use_container_width=True):
                        with st.spinner("Processing updates..."):
                            df_update = load_log_data()
                            row_index = df_update.index[df_update['Row_UID'].astype(str).str.strip() == row_uid.strip()].tolist()[0]
                            df_update.at[row_index, "Container #"] = new_cont
                            df_update.at[row_index, "Country of Origin"] = new_orig
                            df_update.at[row_index, "ETA"] = str(new_eta)
                            df_update.at[row_index, "Lodged Status"] = new_lodg
                            df_update.at[row_index, "Shipment Status"] = new_stat
                            df_update.at[row_index, "NALDO"] = new_naldo
                            df_update.at[row_index, "B/L Number"] = str(new_bl).strip()
                            df_update.at[row_index, "Freight"] = str(new_freight).strip()
                            df_update.at[row_index, "Subtotal (USD)"] = str(new_subtotal).strip()
                            df_update.at[row_index, "Import Duties (TTD)"] = str(new_duties).strip()
                            df_update.at[row_index, "Customs Deposit (TTD)"] = str(new_deposit).strip()
                            df_update.at[row_index, "Import VAT Paid (TTD)"] = str(new_vat).strip()
                            df_update.at[row_index, "Additional Port Charges (TTD)"] = str(new_port).strip()
                            df_update.at[row_index, "Brokerage & Clearance Fees (TTD)"] = str(new_brokerage).strip()
                            df_update.at[row_index, "Management Fees (TTD)"] = str(new_mgmt).strip()
                            
                            for slot_name, up_file in upload_cache.items():
                                doc_filename = f"{inv_no if inv_no.strip() else row_uid}_{slot_name.replace(' ', '_')}.pdf"
                                new_link = upload_physical_file_to_drive(up_file, doc_filename, client_name, inv_no if inv_no.strip() else row_uid)
                                if new_link: df_update.at[row_index, slot_name] = new_link
                                
                            if save_log_data(df_update):
                                st.success("✅ Updates saved!")
                                st.rerun()

                with m_btn2:
                    if st.button("📄 Generate & Lock Post-Clearance Package", key=f"pkg_{idx}", use_container_width=True):
                        with st.spinner("Building Delivery Note & Finance Statement..."):
                            curr_date_str = datetime.now().strftime("%Y-%m-%d")
                            html_wh = pdf_engine.generate_warehouse_delivery_note_html(inv_no, new_cont, new_bl, total_cartons, curr_date_str) if hasattr(pdf_engine, 'generate_warehouse_delivery_note_html') else ""
                            wh_link = upload_system_pdf_to_drive(html_wh, f"{(inv_no if inv_no.strip() else row_uid)}_Warehouse_Delivery_Note.pdf", client_name, inv_no if inv_no.strip() else row_uid)
                            
                            sub_usd = to_decimal(new_subtotal)
                            fr_usd = to_decimal(new_freight)
                            dut_ttd = to_decimal(new_duties)
                            dep_ttd = to_decimal(new_deposit)
                            vat_ttd = to_decimal(new_vat)
                            port_ttd = to_decimal(new_port)
                            brok_ttd = to_decimal(new_brokerage)
                            mgmt_ttd = to_decimal(new_mgmt)
                            
                            html_fin = pdf_engine.generate_finance_cost_statement_html(inv_no, new_cont, new_bl, total_cartons, curr_date_str, sub_usd, fr_usd, dut_ttd, dep_ttd, vat_ttd, port_ttd, brok_ttd, mgmt_ttd) if hasattr(pdf_engine, 'generate_finance_cost_statement_html') else ""
                            fin_link = upload_system_pdf_to_drive(html_fin, f"{(inv_no if inv_no.strip() else row_uid)}_Finance_Cost_Statement.pdf", client_name, inv_no if inv_no.strip() else row_uid)
                            
                            df_update = load_log_data()
                            row_index = df_update.index[df_update['Row_UID'].astype(str).str.strip() == row_uid.strip()].tolist()[0]
                            df_update.at[row_index, "Warehouse Delivery Note"] = wh_link
                            df_update.at[row_index, "Finance Cost Statement"] = fin_link
                            
                            if save_log_data(df_update):
                                st.success("✅ Package linked and saved to Drive!")
                                st.rerun()

def render_admin_tracker():
    st.subheader("📦 Shipment Document Hub")
    active_shell_uid = st.session_state.get("active_shell_uid", "")
    if not active_shell_uid or active_shell_uid == "-- Choose Active Workspace --":
        st.warning("⚠️ Access Restriction: Please select an Active Workspace Shell from the top menu to enable data intake.")
        return

    df_current = load_log_data()
    match_row = df_current[df_current['Row_UID'].astype(str).str.strip() == active_shell_uid.strip()]
    row_data = match_row.iloc[0] if not match_row.empty else {}
    def get_val(key, default=""): return row_data.get(key, default)

    client_file = "clients.csv"
    supplier_file = "suppliers.csv"
    client_options = ["Select a Client..."] + sorted(pd.read_csv(client_file)["Name"].dropna().tolist()) if os.path.exists(client_file) and os.path.getsize(client_file) > 0 else ["Select a Client..."]
    supplier_options = ["Select a Supplier..."] + sorted(pd.read_csv(supplier_file)["Name"].dropna().tolist()) if os.path.exists(supplier_file) and os.path.getsize(supplier_file) > 0 else ["Select a Supplier..."]

    col1, col2 = st.columns([1, 1.3])
    with col1:
        st.markdown("#### Data Intake & Matrix Mapping")
        client_val = get_val("Client Name", "Select a Client...")
        client_idx = client_options.index(client_val) if client_val in client_options else 0
        client_name = st.selectbox("Client Workspace", client_options, index=client_idx)
        supplier_name = st.selectbox("Supplier Profile", supplier_options)
        
        uploaded_file = st.file_uploader("Drop Raw Vendor Spreadsheet (CSV or Excel)", type=["csv", "xlsx"])
        
    with col2:
        st.markdown("#### Target Workspace Selected")
        st.info(f"Active Workspace Shell: **{active_shell_uid}**")

# ==========================================
# 6. TOP NAVIGATION & WORKSPACE ROUTER
# ==========================================
if "active_module" not in st.session_state:
    st.session_state["active_module"] = "📋 Cargo Tracker & Vault"

col_nav1, col_nav2, col_nav3 = st.columns(3)
with col_nav1:
    if st.button("📋 Cargo Tracker & Vault", use_container_width=True): 
        st.session_state["active_module"] = "📋 Cargo Tracker & Vault"
with col_nav2:
    if st.button("📦 Shipment Document Hub", use_container_width=True): 
        st.session_state["active_module"] = "📦 Shipment Document Hub"
with col_nav3:
    if st.button("🏛️ Nominated Agency Portal", use_container_width=True): 
        st.session_state["active_module"] = "🏛️ Nominated Agency Portal"

st.write("---")

col_create, col_select = st.columns([1, 2])
with col_create:
    if st.button("➕ Create Empty Shipment Shell", type="primary", use_container_width=True):
        with st.spinner("Initializing Workspace Shell..."):
            df_current = load_log_data()
            new_uid = f"UID-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            blank_row = {col: "" for col in LOG_COLUMNS}
            blank_row["Row_UID"] = new_uid
            blank_row["Invoice No"] = ""
            blank_row["Total Cartons"] = 0
            blank_row["Shipment Status"] = "Active"
            blank_row["NALDO"] = "No"
            blank_row["Lodged Status"] = "No"
            blank_row["B/L Number"] = ""
            blank_row["Freight"] = ""
            blank_row["Cargo Notes"] = ""
            for doc_slot in ALL_DOCS: blank_row[doc_slot] = "Pending Upload"
            
            df_new = pd.concat([df_current, pd.DataFrame([blank_row])], ignore_index=True)
            if save_log_data(df_new):
                st.session_state["active_shell_uid"] = new_uid
                st.toast("Empty Workspace Shell successfully generated!", icon="✅")
                st.rerun()

with col_select:
    df_dropdown = load_log_data()
    dropdown_options = ["-- Choose Active Workspace --"]
    if not df_dropdown.empty:
        for _, r in df_dropdown.iterrows():
            r_uid = str(r.get("Row_UID", "")).strip()
            s_id = str(r.get("Invoice No", "")).strip()
            s_ctns = str(r.get("Total Cartons", "")).strip()
            s_client = str(r.get("Client Name", "")).strip()
            if not r_uid: continue
            display_name = s_id if s_id.strip() else "[Blank Entry]"
            label = f"[{r_uid}] INV: {display_name}"
            if s_client: label += f" | Client: {s_client}"
            if s_ctns and s_ctns != "0" and s_ctns != "": label += f" | Cartons: {s_ctns}"
            dropdown_options.append(label)

    current_target_uid = st.session_state.get("active_shell_uid", "")
    matching_indices = [i for i, opt in enumerate(dropdown_options) if f"[{current_target_uid}]" in opt]
    default_sel_idx = matching_indices[0] if matching_indices and current_target_uid else 0

    selected_option = st.selectbox("Select Target Workspace", dropdown_options, index=default_sel_idx, label_visibility="collapsed")
    if selected_option != "-- Choose Active Workspace --":
        match = re.search(r'\[(.*?)\]', selected_option)
        if match: 
            new_uid = match.group(1)
            if st.session_state.get("active_shell_uid") != new_uid:
                st.session_state["active_shell_uid"] = new_uid
                # Reset Nominated Agency loaded state so it picks up the newly selected shell's values
                st.session_state["nom_loaded_shell_uid"] = ""
    else: 
        st.session_state["active_shell_uid"] = ""

st.write("---")

if st.session_state["active_module"] == "📋 Cargo Tracker & Vault":
    render_master_log()
elif st.session_state["active_module"] == "📦 Shipment Document Hub":
    render_admin_tracker()
elif st.session_state["active_module"] == "🏛️ Nominated Agency Portal":
    render_nominated_agency_portal()
