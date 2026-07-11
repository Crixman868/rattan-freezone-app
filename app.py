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
from googleapiclient.http import MediaFileUpload
from weasyprint import HTML

# ==========================================
# 1. GLOBAL SETUP & CSS
# ==========================================
st.set_page_config(page_title="Meridian Command Console", page_icon="📦", layout="wide")

COMPANY_LOGO_PATH = "company_logo.png"

def to_decimal(val):
    try:
        if isinstance(val, (int, float)):
            return Decimal(str(val)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        clean_val = re.sub(r'[^\d.]', '', str(val))
        return Decimal(clean_val).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except:
        return Decimal('0.00')

def safe_qty_parse(val):
    """FIX: Defensive parsing to prevent crashes."""
    try:
        if isinstance(val, (int, float)): return int(val)
        val_str = str(val).replace(",", "").strip()
        if not val_str or val_str.lower() in ['nan', 'none', 'n/a']: return 0
        return int(float(val_str))
    except:
        return 0

st.markdown("""
<style>
    .stApp {
        background-color: #ffffff;
        background-image: 
            linear-gradient(45deg, #f8f9fa 25%, transparent 25%, transparent 75%, #f8f9fa 75%, #f8f9fa), 
            linear-gradient(45deg, #f8f9fa 25%, transparent 25%, transparent 75%, #f8f9fa 75%, #f8f9fa);
        background-size: 20px 20px;
        background-position: 0 0, 10px 10px;
    }
    [data-testid="stExpander"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.04);
        margin-bottom: 10px;
    }
    [data-testid="stExpander"] summary p {
        font-weight: 600 !important;
        color: #1e293b !important;
        font-size: 1.05rem !important;
    }
    [data-testid="stExpander"] p, 
    [data-testid="stExpander"] h3, 
    [data-testid="stExpander"] h4, 
    [data-testid="stExpander"] h5 {
        color: #1e293b !important;
    }
</style>
""", unsafe_allow_html=True)

for folder in ["uploaded_docs", "logos", "signatures", "watermarks", "templates"]:
    if not os.path.exists(folder): os.makedirs(folder)

# ==========================================
# 2. CONSTANTS & DATA SCHEMA
# ==========================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1wUBZSnB7cJ2T5_iY5_POpfsNmZn0INGj08EdcLc7TsQ/edit?usp=sharing"
ROOT_FOLDER_ID = "1CITSPAI-BoFeQQLLkmeoX2wkjunTbpGm"

ALL_COUNTRIES = [
    "", "USA", "China", "UK", "Canada", "Brazil", "Mexico", "Panama", "Japan", "Germany", 
    "India", "France", "Italy", "South Korea", "Spain", "Australia", "Taiwan", 
    "Netherlands", "Vietnam", "Malaysia", "Singapore", "South Africa", "UAE", 
    "Saudi Arabia", "Switzerland", "Sweden", "Poland", "Belgium", "Thailand", 
    "Indonesia", "Turkey", "Philippines", "Ireland", "Other"
]

SYSTEM_DOCS = ["Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment"]
EXTERNAL_DOCS = ["Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", "Other Documents", "Miscellaneous Supporting Doc"]
ALL_DOCS = SYSTEM_DOCS + EXTERNAL_DOCS

LOG_COLUMNS = [
    "Row_UID", "Invoice No", "Client Name", "Container #", "Country of Origin", "ETA", 
    "Lodged Status", "Shipment Status", "NALDO", "Total Cartons", 
    "Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment", 
    "Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", 
    "Other Documents", "Miscellaneous Supporting Doc"
]

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def get_gspread_client():
    creds_dict = json.loads(st.secrets["google_api"]["credentials"])
    creds = BotCredentials.from_service_account_info(
        creds_dict, 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]
    )
    return gspread.authorize(creds)

def get_drive_service():
    token_dict = json.loads(st.secrets["google_drive_human"]["token"])
    creds = HumanCredentials.from_authorized_user_info(token_dict)
    return build('drive', 'v3', credentials=creds)

def load_log_data():
    try: 
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=LOG_COLUMNS)
        
        df = pd.DataFrame(records)
        for col in df.columns:
            df[col] = df[col].astype(str).replace(['nan', 'None', '<NA>'], '')
        
        for col in LOG_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df
    except Exception as e: 
        st.error(f"Failed to load data: {e}")
        return pd.DataFrame(columns=LOG_COLUMNS)

def save_log_data(df):
    try:
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        ws.clear()
        
        df = df.copy()
        for col in df.columns:
            df[col] = df[col].astype(str).replace(['nan', 'None'], '')
            
        for col in LOG_COLUMNS:
            if col not in df.columns: df[col] = ""
        df = df[LOG_COLUMNS]
        ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        st.error(f"Failed to sync with Google Sheets: {e}")
        return False

def upload_system_pdf_to_drive(html_content, file_name, client_name, invoice_no):
    if not html_content: return "Pending Upload"
    try:
        drive = get_drive_service()
        safe_client_name = str(client_name).replace("'", "\\'")
        safe_invoice_no = str(invoice_no).replace("'", "\\'")
        
        folders = drive.files().list(q=f"name='{safe_client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        inv_folders = drive.files().list(q=f"name='{safe_invoice_no}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute().get('files', [])
        inv_folder_id = inv_folders[0]['id'] if inv_folders else drive.files().create(body={"name": str(invoice_no), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}).execute()['id']
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
            
        HTML(string=html_content).write_pdf(temp_pdf_path)
        pdf_media = MediaFileUpload(temp_pdf_path, mimetype='application/pdf', resumable=True)
        
        existing_files = drive.files().list(q=f"name='{file_name}' and '{inv_folder_id}' in parents and trashed=false", fields="files(id, webViewLink)").execute().get('files', [])
        
        if existing_files:
            file_id = existing_files[0]['id']
            final_pdf = drive.files().update(fileId=file_id, media_body=pdf_media, fields='id, webViewLink').execute()
        else:
            pdf_metadata = {'name': file_name, 'parents': [inv_folder_id]}
            final_pdf = drive.files().create(body=pdf_metadata, media_body=pdf_media, fields='id, webViewLink').execute()
        
        os.remove(temp_pdf_path)
        return final_pdf.get('webViewLink', 'Upload Failed')
    except Exception as e:
        st.error(f"PDF Engine Error for {file_name}: {e}")
        return "Upload Failed"

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
        
        os.remove(temp_path)
        return file.get('webViewLink')
    except Exception as e:
        st.error(f"Drive Upload Error: {e}")
        return None

def get_eta_status(eta_date, shipment_status):
    if shipment_status == "Delivered":
        return "✅ DELIVERED", "#00b050"
    try:
        days_diff = (eta_date - datetime.now().date()).days
        if days_diff < 0: return "⚠️ OVERDUE", "#FF4500"
        if 0 <= days_diff <= 7: return "🔴 URGENT", "#FF0000"
        if 8 <= days_diff <= 14: return "🟡 APPROACHING", "#FFD700"
        return "🟢 IN TRANSIT", "#008000"
    except: return "TBD", "#808080"

def get_img_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as f: return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
    return None

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

# ==========================================
# 4. DOCUMENT GENERATORS
# ==========================================

# --- CARICOM MODULE ---
def generate_caricom_printout(inv_num, date, client_name, client_address, supplier_name, supplier_address, bl, total_ctns, subtotal, freight, grand_total, payment_terms, additional_notes, signatory_position, compliance_data, logo_path, sig_path, orientation, primary_hex):
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath="./templates"))
    template = env.get_template("caricom_template.html")
    return template.render(
        inv_num=inv_num, date=date, client_name=client_name, client_address=client_address,
        supplier_name=supplier_name, supplier_address=supplier_address, bl=bl, total_ctns=total_ctns,
        subtotal=subtotal, freight=freight, grand_total=grand_total, payment_terms=payment_terms,
        additional_notes=additional_notes, signatory_position=signatory_position,
        compliance_data=compliance_data, logo_path=logo_path, sig_path=sig_path,
        orientation=orientation, primary_hex=primary_hex
    )

def generate_html_document(title, inv_no, date, client, c_addr, supplier, s_profile, bl, total_ctns, df, total_val, freight=None, additional_notes="", payment_terms="", signatory_position="", is_packing=False, is_duties=False, duty_data=None):
    logo_path = get_img_b64(f"logos/{s_profile.get('Name', '')}_logo.png")
    sig_path = get_img_b64(f"signatures/{s_profile.get('Name', '')}_sig.png")

    if is_packing:
        table_rows = ""
        for idx, row in df.iterrows():
            # FIXED: Nuclear Defensive Parser Applied
            qty = safe_qty_parse(row.get("QUANTITY", 0))
            table_rows += f'<tr><td style="padding:10px; border:1px solid #ccc;">{row.get("SPECIFICATION OF COMMODITIES","N/A")}</td><td style="padding:10px; border:1px solid #ccc; text-align:center;">{row.get("CTNS NOS","N/A")}</td><td style="padding:10px; border:1px solid #ccc; text-align:center;">{row.get("TOTAL CTNS",0)}</td><td style="padding:10px; border:1px solid #ccc; text-align:right;">{qty:,}</td></tr>'
        img_tag = f'<img src="{logo_path}" height="50">' if logo_path else ''
        sig_tag = f'<img src="{sig_path}" height="80">' if sig_path else ''
        return f'<html><body><table width="100%"><tr><td>{img_tag}</td><td align="right"><h2>{title}</h2></td></tr></table><p><b>Exporter:</b> {supplier}<br><b>Consignee:</b> {client}<br>{c_addr}</p><table border="1" width="100%" cellspacing="0" cellpadding="5"><thead><tr bgcolor="#f7f7f7"><th>Description</th><th>Carton Nos</th><th>Total Ctns</th><th>Qty</th></tr></thead><tbody>{table_rows}</tbody></table><br><br><div align="right">{sig_tag}<br><b>{signatory_position}</b></div></body></html>'
    
    elif is_duties:
        duty_data = duty_data or {}
        img_tag = f'<img src="{logo_path}" height="50">' if logo_path else ''
        return f'<html><body><table width="100%"><tr><td>{img_tag}</td><td align="right"><h2>{title}</h2></td></tr></table><p><b>Invoice:</b> {inv_no}</p><p>Converted Base Value: ${duty_data.get("convert_to_ttd",0):,.2f} TTD</p><p>Customs Duty: ${duty_data.get("duty_owed",0):,.2f} TTD</p><p>VAT Owed: ${duty_data.get("vat_owed",0):,.2f} TTD</p><br><table border="1" width="100%" cellspacing="0" cellpadding="10"><tr><td bgcolor="#f9f9f9"><h3>Total Customs Bill Due: ${duty_data.get("grand_total_ttd",0):,.2f} TTD</h3></td></tr></table></body></html>'
    
    else:
        template_env = jinja2.Environment(loader=jinja2.FileSystemLoader(searchpath="./templates"))
        chosen_template = s_profile.get("Template", "classic.html")
        if not os.path.exists(f"./templates/{chosen_template}"): chosen_template = "classic.html"
        try: 
            template = template_env.get_template(chosen_template)
        except: 
            template = template_env.from_string("<h1>{{title}}</h1><p><b>Exporter:</b> {{supplier_name}}<br><b>Consignee:</b> {{client_name}}</p><table border='1' width='100%' cellspacing='0' cellpadding='5'><thead><tr bgcolor='#f2f2f2'><th>Description</th><th>Qty</th><th>Total</th></tr></thead><tbody>{% for item in items %}<tr><td>{{item.Description}}</td><td>{{item.Qty}}</td><td>{{item.Total}}</td></tr>{% endfor %}</tbody></table>")

        items = []
        for idx, row in df.iterrows():
            desc = str(row["Description"])[:250]
            qty = f"{safe_qty_parse(row.get('Qty', 0)):,}"
            try:
                price = f"{float(row.get('UnitPrice', 0)):.2f}" if pd.notna(row.get('UnitPrice')) else ""
            except ValueError:
                price = ""
            try:
                total = f"{float(row.get('Total Foreign (USD)', 0)):.2f}" if pd.notna(row.get('Total Foreign (USD)')) else ""
            except ValueError:
                total = ""
                
            items.append({"Description": desc, "Qty": qty, "UnitPrice": price, "Total": total})
            
        rendered_html = template.render({
            "title": title, "inv_no": inv_no, "date": date, "client_name": client, 
            "client_address": c_addr, "supplier_name": supplier, 
            "supplier_address": s_profile.get("Address", "Main Office Hub"), 
            "bl": bl, "total_ctns": total_ctns, "payment_terms": payment_terms, 
            "additional_notes": additional_notes