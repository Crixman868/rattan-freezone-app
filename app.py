import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import base64
import gspread
import json
import jinja2
import re
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials as HumanCredentials
from google.oauth2.service_account import Credentials as BotCredentials
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
from weasyprint import HTML

st.set_page_config(page_title="Meridian Command Console", page_icon="📦", layout="wide")

# --- UTILITIES ---
def to_decimal(val):
    try: return Decimal(re.sub(r'[^\d.]', '', str(val))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except: return Decimal('0.00')

def safe_qty_parse(val):
    try:
        if isinstance(val, (int, float)): return int(val)
        val_str = str(val).replace(",", "").strip()
        if not val_str or val_str.lower() in ['nan', 'none', 'n/a']: return 0
        return int(float(val_str))
    except: return 0

# --- FOLDER SETUP ---
for folder in ["uploaded_docs", "logos", "signatures", "watermarks", "templates"]:
    if not os.path.exists(folder): os.makedirs(folder)

# --- GOOGLE SERVICES ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1wUBZSnB7cJ2T5_iY5_POpfsNmZn0INGj08EdcLc7TsQ/edit?usp=sharing"
ROOT_FOLDER_ID = "1CITSPAI-BoFeQQLLkmeoX2wkjunTbpGm"

def get_gspread_client():
    creds_dict = json.loads(st.secrets["google_api"]["credentials"])
    creds = BotCredentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"])
    return gspread.authorize(creds)

def get_drive_service():
    token_dict = json.loads(st.secrets["google_drive_human"]["token"])
    creds = HumanCredentials.from_authorized_user_info(token_dict)
    return build('drive', 'v3', credentials=creds)

# --- LOG HANDLING ---
LOG_COLUMNS = ["Row_UID", "Invoice No", "Client Name", "Container #", "Country of Origin", "ETA", "Lodged Status", "Shipment Status", "NALDO", "Total Cartons", "Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment", "Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", "Other Documents", "Miscellaneous Supporting Doc"]

def load_log_data():
    try: 
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        records = ws.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=LOG_COLUMNS)
        for col in df.columns: df[col] = df[col].astype(str).replace(['nan', 'None', '<NA>'], '')
        for col in LOG_COLUMNS:
            if col not in df.columns: df[col] = ""
        return df
    except Exception as e: 
        st.error(f"Failed to load: {e}")
        return pd.DataFrame(columns=LOG_COLUMNS)

def save_log_data(df):
    try:
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        ws.clear()
        df = df.copy()
        for col in df.columns: df[col] = df[col].astype(str).replace(['nan', 'None'], '')
        df = df[LOG_COLUMNS]
        ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"Sync Failed: {e}")
        return False

# --- UPLOADS ---
def upload_system_pdf_to_drive(html_content, file_name, client_name, invoice_no):
    if not html_content: return "Pending Upload"
    try:
        drive = get_drive_service()
        folders = drive.files().list(q=f"name='{client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        inv_folders = drive.files().list(q=f"name='{invoice_no}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        inv_folder_id = inv_folders[0]['id'] if inv_folders else drive.files().create(body={"name": str(invoice_no), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf: 
            HTML(string=html_content).write_pdf(temp_pdf.name)
            pdf_path = temp_pdf.name
        
        media = MediaFileUpload(pdf_path, mimetype='application/pdf', resumable=True)
        existing = drive.files().list(q=f"name='{file_name}' and '{inv_folder_id}' in parents and trashed=false", fields="files(id, webViewLink)").execute().get('files', [])
        if existing:
            file = drive.files().update(fileId=existing[0]['id'], media_body=media, fields='id, webViewLink').execute()
        else:
            file = drive.files().create(body={'name': file_name, 'parents': [inv_folder_id]}, media_body=media, fields='id, webViewLink').execute()
        os.remove(pdf_path)
        return file.get('webViewLink', 'Upload Failed')
    except Exception as e:
        st.error(f"Drive Error: {e}")
        return "Upload Failed"

def upload_physical_file_to_drive(uploaded_file, file_name, client_name, invoice_no):
    if not uploaded_file: return None
    try:
        drive = get_drive_service()
        folders = drive.files().list(q=f"name='{client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        inv_folders = drive.files().list(q=f"name='{invoice_no}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        inv_folder_id = inv_folders[0]['id'] if inv_folders else drive.files().create(body={"name": str(invoice_no), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name
        media = MediaFileUpload(temp_path, resumable=True)
        existing = drive.files().list(q=f"name='{file_name}' and '{inv_folder_id}' in parents and trashed=false", fields="files(id, webViewLink)").execute().get('files', [])
        if existing:
            file = drive.files().update(fileId=existing[0]['id'], media_body=media, fields='id, webViewLink').execute()
        else:
            file = drive.files().create(body={'name': file_name, 'parents': [inv_folder_id]}, media_body=media, fields='id, webViewLink').execute()
        os.remove(temp_path)
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Drive Upload Error: {e}")
        return None

# --- GENERATORS ---
def get_img_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as f: return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
    return None

def generate_caricom_printout(inv_num, date, client_name, client_addr, supplier_name, supplier_addr, additional_notes, subtotal, signatory_position, logo_path, sig_path, primary_hex, orientation):
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath="./templates"))
    template = env.get_template("caricom_template.html")
    return template.render(inv_num=inv_num, date=date, client_name=client_name, supplier_name=supplier_name, 
                           supplier_address=supplier_addr, additional_notes=additional_notes, subtotal=f"{subtotal:,.2f}",
                           logo_path=logo_path, sig_path=sig_path, signatory_position=signatory_position, 
                           primary_hex=primary_hex, orientation=orientation, client_address=client_addr)

def generate_html_document(title, inv_no, date, client, c_addr, supplier, s_profile, bl, total_ctns, df, total_val, freight=None, additional_notes="", payment_terms="", signatory_position="", is_packing=False, is_duties=False, duty_data=None):
    logo_path = get_img_b64(f"logos/{s_profile.get('Name', '')}_logo.png")
    sig_path = get_img_b64(f"signatures/{s_profile.get('Name', '')}_sig.png")
    
    if is_packing:
        table_rows = ""
        for idx, row in df.iterrows():
            qty = safe_qty_parse(row.get("QUANTITY", 0))
            table_rows += f'<tr><td style="padding:10px; border:1px solid #ccc;">{row.get("SPECIFICATION OF COMMODITIES","N/A")}</td><td style="padding:10px; border:1px solid #ccc; text-align:center;">{row.get("CTNS NOS","N/A")}</td><td style="padding:10px; border:1px solid #ccc; text-align:center;">{row.get("TOTAL CTNS",0)}</td><td style="padding:10px; border:1px solid #ccc; text-align:right;">{qty:,}</td></tr>'
        img_tag = f'<img src="{logo_path}" height="40">' if logo_path else ''
        sig_tag = f'<img src="{sig_path}" height="80">' if sig_path else ''
        return f'<html><body><table width="100%"><tr><td>{img_tag}</td><td align="right"><h2>{title}</h2></td></tr></table><table border="1" width="100%"><thead><tr bgcolor="#f7f7f7"><th>Description</th><th>Carton Nos</th><th>Total Ctns</th><th>Qty-Pcs</th></tr></thead><tbody>{table_rows}</tbody></table><div align="right">{sig_tag}<br><b>{signatory_position}</b></div></body></html>'
    elif is_duties:
        duty_data = duty_data or {}
        img_tag = f'<img src="{logo_path}" height="40">' if logo_path else ''
        return f'<html><body><table width="100%"><tr><td>{img_tag}</td><td align="right"><h2>{title}</h2></td></tr></table><p><b>Invoice:</b> {inv_no}</p><p>Total Customs Bill Due: ${duty_data.get("grand_total_ttd",0):,.2f} TTD</p></body></html>'
    else:
        env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath="./templates"))
        template = env.get_template(s_profile.get("Template", "classic.html")) if os.path.exists(f"./templates/{s_profile.get('Template', 'classic.html')}") else env.from_string("<h1>Invoice</h1>")
        items = []
        for idx, row in df.iterrows():
            items.append({"Description": str(row["Description"])[:250], "Qty": f"{safe_qty_parse(row.get('Qty', 0)):,}", "UnitPrice": f"{float(row.get('UnitPrice', 0)):.2f}", "Total": f"{float(row.get('Total Foreign (USD)', 0)):.2f}"})
        return template.render({"title": title, "inv_no": inv_no, "date": date, "client_name": client, "client_address": c_addr, "supplier_name": supplier, "supplier_address": s_profile.get("Address", "Main Office Hub"), "bl": bl, "total_ctns": total_ctns, "payment_terms": payment_terms, "additional_notes": additional_notes, "logo_path": logo_path, "sig_path": sig_path, "signatory_position": signatory_position, "subtotal": f"{total_val:,.2f}", "freight": (f"{freight:,.2f}" if freight else None), "grand_total": f"{(total_val + (freight or 0)):,.2f}", "items": items})

def display_html_preview(raw_html):
    components.html(f'<div style="background-color: white; padding: 40px;">{raw_html}</div>', height=750, scrolling=True)

# --- VIEWS ---
def render_master_log():
    st.title("🗄️ Master Log: Logistics Control Tower")
    df = load_log_data()
    if df.empty: st.info("No logs found.")
    for idx, row in df.iterrows():
        if st.expander(f"INV: {row.get('Invoice No')} | {row.get('Client Name')}"):
            st.write(row)
            if st.button("💾 Save", key=f"save_{idx}"): st.rerun()

def render_admin_tracker():
    st.title("📦 Command Console: Master Tracker")
    active_shell_uid = st.session_state.get("active_shell_uid", "")
    if not active_shell_uid: st.warning("Select Workspace."); return
    
    df = load_log_data()
    row = df[df['Row_UID'] == active_shell_uid].iloc[0] if not df[df['Row_UID'] == active_shell_uid].empty else {}
    def get_val(key, default=""): return row.get(key, default)
    
    col1, col2 = st.columns([1, 1.3])
    with col1:
        invoice_num = st.text_input("Invoice Number", value=get_val("Invoice No"))
        invoice_date = st.text_input("Invoice Date", value=get_val("ETA", datetime.now().strftime("%Y-%m-%d")))
        client_name = st.text_input("Client Name", value=get_val("Client Name"))
        supplier_name = st.selectbox("Supplier", ["Select Supplier..."]) # Add logic to load from csv
        
    with col2:
        t_inv, t_car, t_pck, t_dut = st.tabs(["📄 Invoice", "🌐 CARICOM", "📋 Packing Manifest", "🇹🇹 Customs Audit"])
        with t_car:
            orientation = st.radio("Orientation", ["portrait", "landscape"], index=1)
            if st.button("⚙️ Preview CARICOM"):
                s_profile = get_entity_profile("suppliers.csv", supplier_name)
                html = generate_caricom_printout(invoice_num, invoice_date, client_name, "Addr", supplier_name, "Addr", "Notes", 0, "Dir", None, None, "#000000", orientation)
                st.session_state["h_car"] = html
                display_html_preview(html)

def render_supplier_admin(): st.title("⚙️ Supplier Admin")
def render_client_admin(): st.title("👥 Client Admin")

# --- ROUTER ---
if "active_module" not in st.session_state: st.session_state["active_module"] = "📋 Master Log"
if st.sidebar.button("📋 Master Log"): st.session_state["active_module"] = "📋 Master Log"
if st.sidebar.button("📦 Master Tracker"): st.session_state["active_module"] = "📦 Master Tracker"
if st.sidebar.button("⚙️ Supplier Admin"): st.session_state["active_module"] = "⚙️ Supplier Admin"
if st.sidebar.button("👥 Client Admin"): st.session_state["active_module"] = "👥 Client Admin"

if st.session_state["active_module"] == "📋 Master Log": render_master_log()
elif st.session_state["active_module"] == "📦 Master Tracker": render_admin_tracker()
elif st.session_state["active_module"] == "⚙️ Supplier Admin": render_supplier_admin()
elif st.session_state["active_module"] == "👥 Client Admin": render_client_admin()