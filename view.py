import streamlit as st
import streamlit.components.v1 as components  # Make sure this is imported!
import pandas as pd
import os
import base64
import gspread
import json
import re
from datetime import datetime

# UPDATED TAB TITLE AND CUSTOM ICON
st.set_page_config(page_title="Rattan's Logistics", page_icon="logo_left.png", layout="wide")

def get_img_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as f: return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
    return None

left_logo_b64 = get_img_b64("logo_left.png")
right_logo_b64 = get_img_b64("logo_right.png")

# --- INJECT CUSTOM APP INSTALL ICON (LOGO LEFT ONLY) ---
if left_logo_b64:
    components.html(f"""
    <script>
        const doc = window.parent.document;
        
        // 1. Overwrite existing shortcut icons
        let icons = doc.querySelectorAll('link[rel="shortcut icon"], link[rel="icon"], link[rel="apple-touch-icon"]');
        icons.forEach(icon => icon.href = '{left_logo_b64}');
        
        // 2. Force inject an Apple Touch Icon if one doesn't exist
        let appleIcon = doc.querySelector('link[rel="apple-touch-icon"]');
        if (!appleIcon) {{
            appleIcon = doc.createElement('link');
            appleIcon.rel = 'apple-touch-icon';
            appleIcon.href = '{left_logo_b64}';
            doc.head.appendChild(appleIcon);
        }}
    </script>
    """, height=0, width=0)

# LOGOS SET TO 120PX, NO MARGINS SO THEY TOUCH THE TOP/BOTTOM OF THE BANNER
left_img_tag = f'<img src="{left_logo_b64}" style="height: 120px; border-radius: 12px; object-fit: contain; display: block;">' if left_logo_b64 else ''
right_img_tag = f'<img src="{right_logo_b64}" style="height: 120px; border-radius: 12px; object-fit: contain; display: block;">' if right_logo_b64 else ''

st.markdown(f"""
<style>
    .stApp {{ background-color: #f8fafc; color: #1e293b; }}
    .custom-header {{
        background: linear-gradient(135deg, #e60000 0%, #8b0000 100%);
        color: white; 
        padding: 0px 30px; /* REMOVED TOP AND BOTTOM PADDING TO HUG LOGOS */
        border-radius: 12px;
        box-shadow: 0 8px 20px rgba(220, 38, 38, 0.25), inset 0 2px 10px rgba(255,255,255,0.1);
        display: flex; justify-content: space-between; align-items: center;
        position: relative; overflow: hidden; margin-bottom: 20px; margin-top: 10px;
    }}
    .header-center {{ display: flex; flex-direction: column; align-items: center; text-align: center; flex-grow: 1; }}
    .header-title {{ 
        font-family: 'Arial', sans-serif; 
        font-size: 46px; /* MASSIVELY INCREASED FONT SIZE */
        font-weight: 900; 
        letter-spacing: 3px; 
        margin: 0; 
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3); 
    }}
    [data-testid="stExpander"] {{
        background-color: #ffffff !important; border: 1px solid #e2e8f0; border-top: 4px solid #dc2626;
        border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.04); margin-bottom: 15px;
    }}
    [data-testid="stExpander"] summary p {{ font-weight: 700 !important; color: #1e293b !important; font-size: 1.02rem !important; }}
</style>

<div class="custom-header">
    <div>{left_img_tag}</div>
    <div class="header-center">
        <h1 class="header-title">RATTAN FREEZONE</h1>
    </div>
    <div>{right_img_tag}</div>
</div>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1Rcifpu4GRFAYFPQNBGrl96DpHXDiQM1JYys-Dhi0rrU/edit?usp=sharing"

SYSTEM_DOCS = ["Commercial Invoice", "CARICOM Invoice", "Sequential Packing List", "Official Duties Assessment", "Warehouse Delivery Note", "Finance Cost Statement"]
EXTERNAL_DOCS = ["Bill of Lading Scan", "Original Invoice", "Original Packing List", "Tracker Document", "Other Documents", "Miscellaneous Supporting Doc"]
ALL_DOCS = SYSTEM_DOCS + EXTERNAL_DOCS

def get_gspread_client():
    from google.oauth2.service_account import Credentials as BotCredentials
    creds_dict = json.loads(st.secrets["google_api"]["credentials"])
    creds = BotCredentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"])
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_log_data():
    try: 
        ws = get_gspread_client().open_by_url(SHEET_URL).sheet1
        rows = ws.get_all_values()
        if not rows or len(rows) < 2: return pd.DataFrame()
        headers = rows[0]
        data = rows[1:]
        return pd.DataFrame(data, columns=headers, dtype=object)
    except Exception as e: 
        st.error("Failed to connect to Master Database.")
        return pd.DataFrame()

def get_eta_status(eta_date, shipment_status):
    if shipment_status == "Delivered": return "✅ DELIVERED"
    try:
        days_diff = (eta_date - datetime.now().date()).days
        if days_diff < 0: return "⚠️ OVERDUE"
        if 0 <= days_diff <= 7: return "🔴 URGENT"
        if 8 <= days_diff <= 14: return "🟡 APPROACHING"
        return "🟢 IN TRANSIT"
    except: return "TBD"

# UPDATED SUBHEADER
st.subheader("📋 Cargo Tracker & Vault")

df = load_log_data()

if df.empty:
    st.info("No active shipments to display.")
else:
    for idx, row in df.iterrows():
        row_uid = str(row.get('Row_UID', ''))
        if not row_uid.strip(): continue
        
        # Pull all core data
        inv_no = str(row.get('Invoice No', '')).strip()
        display_inv = inv_no if inv_no else "[Blank Entry]"
        client_name = str(row.get('Client Name', '')).strip()
        inv_date = str(row.get('Invoice Date', '')).strip()
        ship_status = str(row.get("Shipment Status", "Active"))
        total_cartons = str(row.get("Total Cartons", "0"))
        cont_no = str(row.get("Container #", "")).strip()
        bl_no = str(row.get("B/L Number", "")).strip()
        origin_no = str(row.get("Country of Origin", "")).strip()
        lodged_val = str(row.get("Lodged Status", "")).strip()
        naldo_val = str(row.get("NALDO", "")).strip()
        cargo_notes = str(row.get("Cargo Notes", "")).strip()

        # Pull financial data
        freight = str(row.get("Freight", "")).strip()
        subtotal = str(row.get("Subtotal (USD)", "")).strip()
        duties = str(row.get("Import Duties (TTD)", "")).strip()
        deposit = str(row.get("Customs Deposit (TTD)", "")).strip()
        vat = str(row.get("Import VAT Paid (TTD)", "")).strip()
        port = str(row.get("Additional Port Charges (TTD)", "")).strip()
        brokerage = str(row.get("Brokerage & Clearance Fees (TTD)", "")).strip()
        mgmt = str(row.get("Management Fees (TTD)", "")).strip()

        # ETA Logic
        raw_eta = row.get("ETA")
        timestamp = pd.to_datetime(raw_eta, errors='coerce')
        current_date = timestamp.date() if not pd.isna(timestamp) else datetime.now().date()
        status_label = get_eta_status(current_date, ship_status)
        
        header_text = (f"📦 TOTAL CTNS: {total_cartons} | {status_label} | ETA: {current_date} | "
                       f"Container #: {cont_no} | B/L Number: {bl_no} | "
                       f"INV: {display_inv} | Origin: {origin_no} | Lodged: {lodged_val}")

        with st.expander(header_text):
            st.markdown("#### 🚢 Essential Shipment Details")
            c1, c2, c3, c4 = st.columns(4)
            c1.caption("Client Name")
            c1.write(f"**{client_name or 'N/A'}**")
            c2.caption("Invoice Date")
            c2.write(f"**{inv_date or 'N/A'}**")
            c3.caption("Container Number")
            c3.write(f"**{cont_no or 'Pending'}**")
            c4.caption("B/L Number")
            c4.write(f"**{bl_no or 'Pending'}**")

            c5, c6, c7 = st.columns(3)
            c5.caption("Origin Country")
            c5.write(f"**{origin_no or 'N/A'}**")
            c6.caption("NALDO Code")
            c6.write(f"**{naldo_val or 'No'}**")
            c7.caption("Current Status")
            c7.write(f"**{ship_status}**")
            
            st.caption("Cargo Notes")
            st.write(f"*{cargo_notes or 'None'}*")

            st.write("---")
            st.markdown("#### 💰 Financial Overview (USD & Post-Clearance TTD)")
            f1, f2, f3, f4 = st.columns(4)
            f1.caption("Subtotal (USD)")
            f1.write(f"**${subtotal or '0.00'}**")
            f2.caption("Freight (USD)")
            f2.write(f"**${freight or '0.00'}**")
            f3.caption("Import Duties (TTD)")
            f3.write(f"**${duties or '0.00'}**")
            f4.caption("Customs Deposit (TTD)")
            f4.write(f"**${deposit or '0.00'}**")

            f5, f6, f7, f8 = st.columns(4)
            f5.caption("Import VAT Paid (TTD)")
            f5.write(f"**${vat or '0.00'}**")
            f6.caption("Add. Port Charges (TTD)")
            f6.write(f"**${port or '0.00'}**")
            f7.caption("Brokerage Fees (TTD)")
            f7.write(f"**${brokerage or '0.00'}**")
            f8.caption("Management Fees (TTD)")
            f8.write(f"**${mgmt or '0.00'}**")
            
            st.write("---")
            st.markdown("#### 📑 Secure Document Vault")
            grid = st.columns(6)
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
                        st.link_button("📄 Download", url=clean_link, use_container_width=True)
                    else:
                        st.markdown("🔒 *Pending*")