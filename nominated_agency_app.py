import os
import re
import json
import base64
import tempfile
import pandas as pd
import streamlit as st
import pdf_engine

# Google API & Drive Dependencies
import gspread
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials as HumanCredentials
from google.oauth2.service_account import Credentials as BotCredentials
from googleapiclient.http import MediaFileUpload

# Page Configuration
st.set_page_config(
    page_title="Nominated Agency Portal",
    page_icon="🚢",
    layout="wide"
)

st.title("🚢 Nominated Agency Supply Chain Portal")
st.caption("Corinthian Pins Limited | Trinidad Freight Solutions Limited | Rattans Freezone Limited")
st.divider()

# ---------------------------------------------------------
# 1. EXTRACTED GOOGLE CLOUD & DRIVE BRIDGE FUNCTIONS
# ---------------------------------------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1Rcifpu4GRFAYFPQNBGrl96DpHXDiQM1JYys-Dhi0rrU/edit?usp=sharing"
ROOT_FOLDER_ID = "1GtZk2jfAHGqttyZVP9E8G4TA_MNrV9Pp"

def get_gspread_client():
    """Authenticates gspread client using st.secrets['google_api']."""
    creds_dict = json.loads(st.secrets["google_api"]["credentials"])
    creds = BotCredentials.from_service_account_info(
        creds_dict, 
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]
    )
    return gspread.authorize(creds)

def get_drive_service():
    """Authenticates Drive v3 client using human OAuth token in st.secrets['google_drive_human']."""
    token_info = json.loads(st.secrets["google_drive_human"]["token"])
    creds = HumanCredentials.from_authorized_user_info(token_info)
    return build('drive', 'v3', credentials=creds)

def sync_local_pdf_to_google_drive(local_pdf_path, client_name, bl_number):
    """
    Uploads or updates a local PDF file in Google Drive inside:
    ROOT_FOLDER_ID -> Client Folder -> B/L Subfolder -> PDF File
    Returns the shareable webViewLink.
    """
    if not os.path.exists(local_pdf_path):
        return None, "Local file not found on disk."
        
    try:
        drive = get_drive_service()
        file_name = os.path.basename(local_pdf_path)
        safe_client_name = str(client_name).replace("'", "\\'")
        safe_bl_number = str(bl_number).replace("'", "\\'")
        
        # 1. Find or create Client Folder in Root
        folders = drive.files().list(
            q=f"name='{safe_client_name}' and '{ROOT_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)"
        ).execute().get('files', [])
        client_folder_id = folders[0]['id'] if folders else drive.files().create(
            body={"name": client_name, "parents": [ROOT_FOLDER_ID], "mimeType": "application/vnd.google-apps.folder"}
        ).execute()['id']
        
        # 2. Find or create B/L Subfolder in Client Folder
        bl_folders = drive.files().list(
            q=f"name='{safe_bl_number}' and '{client_folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)"
        ).execute().get('files', [])
        bl_folder_id = bl_folders[0]['id'] if bl_folders else drive.files().create(
            body={"name": str(bl_number), "parents": [client_folder_id], "mimeType": "application/vnd.google-apps.folder"}
        ).execute()['id']
        
        # 3. Upload or Update PDF File
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


# ---------------------------------------------------------
# 2. HELPER: TAB-BASED PDF WORKSPACE RENDERER
# ---------------------------------------------------------
def render_document_workspace(doc_title, doc_filename, bl_no, generate_callback, key_prefix, entity_client_name="Corinthian Pins Limited"):
    """
    Renders generation triggers, download buttons, Cloud Vault upload controls,
    and a full-width interactive PDF viewer inside a dedicated tab workspace.
    """
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
        
        # Read PDF bytes for preview & download
        with open(file_path, "rb") as f:
            pdf_bytes = f.read()
            base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')

        # Control Row: Download & Save to Vault
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

        # Full-Width Interactive PDF Preview
        st.markdown("##### 👁️ Document Preview")
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf" style="border: 1px solid #cbd5e0; border-radius: 6px;"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
    else:
        st.info(f"ℹ️ {doc_title} has not been generated yet for B/L **{bl_no}**. Click the button above to generate.")


# ---------------------------------------------------------
# 3. SHIPMENT ENTRY & AUDIT FORM
# ---------------------------------------------------------
st.header("1. Shipment & Port Outlay Details")

col1, col2, col3 = st.columns(3)

with col1:
    bl_no = st.text_input("Bill of Lading (B/L) No.", value="BL-2026-001")
    container_no = st.text_input("Container No.", value="CNTR-40912")
    shipment_status = st.selectbox(
        "Operational Shipment Status",
        [
            "Pre-Clearance (Pending Deposit)",
            "Cleared & In Transit to Freezone",
            "Delivered (Draft / Reconciliation Pending)",
            "Delivered (Finalized & Reconciled)"
        ]
    )

with col2:
    usd_cargo_val = st.number_input("Foreign Cargo Valuation (USD)", value=50000.00, step=1000.00)
    exchange_rate = st.number_input("Exchange Rate (TTD/USD)", value=6.80, step=0.01)
    ttd_cargo_val = usd_cargo_val * exchange_rate
    st.info(f"Converted TTD Cargo Value: **${ttd_cargo_val:,.2f} TTD**")

with col3:
    bundled_service_fee = st.number_input("Corinthian Bundled Agency Fee (TTD)", value=25000.00, step=500.00)
    contra_deposit_paid = st.number_input("Pre-Funded Advance Port Deposit (TTD)", value=30000.00, step=1000.00)

st.divider()

# ---------------------------------------------------------
# 4. ITEMIZED PORT OUTLAYS BREAKDOWN
# ---------------------------------------------------------
st.header("2. Itemized Statutory Port Outlays")
st.caption("Break down official customs and port charges incurred prior to release.")

outlay_col1, outlay_col2 = st.columns(2)

with outlay_col1:
    customs_duty = st.number_input("Customs Import Duty (TTD)", value=12000.00, step=500.00)
    import_vat = st.number_input("Customs Import VAT (12.5%) (TTD)", value=11250.00, step=500.00)

with outlay_col2:
    port_handling = st.number_input("Port Authority Handling & Storage (TTD)", value=4250.00, step=250.00)
    shipping_line_demurrage = st.number_input("Shipping Line Demurrage & Fees (TTD)", value=2500.00, step=250.00)

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

st.divider()

# ---------------------------------------------------------
# 5. DOCUMENT GENERATION & ACTION TRIGGERS (TAB SYSTEM)
# ---------------------------------------------------------
st.header("3. Document Generation & Action Triggers")

stage1_tab, stage2_tab = st.tabs([
    "📄 Stage 1: Pre-Clearance & Operations", 
    "🚀 Stage 2: Post-Delivery Final Financial Discharge"
])

# STAGE 1: PRE-CLEARANCE & OPERATIONS
with stage1_tab:
    st.caption("Manage pre-clearance cash requests and private upstream freight invoices.")
    
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
            generate_callback=lambda: pdf_engine.generate_port_disbursement_request(
                bl_no, container_no, port_items=port_items
            ),
            key_prefix="stg1_port",
            entity_client_name="Corinthian Pins Limited"
        )
        
    with subtab_service:
        render_document_workspace(
            doc_title="Service Fee Disbursement Request",
            doc_filename=f"Disbursement_Request_Services_{bl_no}.pdf",
            bl_no=bl_no,
            generate_callback=lambda: pdf_engine.generate_service_disbursement_request(
                bl_no, container_no, bundled_service_fee=bundled_service_fee
            ),
            key_prefix="stg1_service",
            entity_client_name="Corinthian Pins Limited"
        )
        
    with subtab_tfs:
        render_document_workspace(
            doc_title="Internal Upstream Freight Invoice (TFS)",
            doc_filename=f"TFS_Internal_Invoice_{bl_no}.pdf",
            bl_no=bl_no,
            generate_callback=lambda: pdf_engine.generate_internal_tfs_invoice(
                bl_no, container_no, tfs_base=5500.00
            ),
            key_prefix="stg1_tfs",
            entity_client_name="Trinidad Freight Solutions Limited"
        )

# STAGE 2: POST-DELIVERY FINAL FINANCIAL DISCHARGE
with stage2_tab:
    if shipment_status == "Delivered (Finalized & Reconciled)":
        st.markdown("### Pre-Flight Settlement Audit Summary")
        
        summary_col1, summary_col2, summary_col3, summary_col4 = st.columns(4)
        summary_col1.metric("Gross Shipment Total", f"${gross_shipment_total:,.2f} TTD")
        summary_col2.metric("Less Foreign Contra", f"-${ttd_cargo_val:,.2f} TTD")
        summary_col3.metric("Less Port Deposit", f"-${contra_deposit_paid:,.2f} TTD")
        summary_col4.metric("NET BALANCE PAYABLE", f"${net_contra_due:,.2f} TTD")
        
        st.warning("⚠️ Verify all figures above prior to compiling final legal tax invoices and payment discharge receipts.")
        
        audit_confirm = st.checkbox(f"I confirm container delivery and final port outlay reconciliation for B/L {bl_no}", key="audit_checkbox_stg2")
        
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
                        usd_cargo_val=usd_cargo_val, 
                        exchange_rate=exchange_rate, 
                        port_items=port_items, 
                        bundled_service_fee=bundled_service_fee, 
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
                    generate_callback=lambda: pdf_engine.generate_official_receipt(
                        bl_no, container_no, 
                        amount_paid=net_contra_due
                    ),
                    key_prefix="stg2_receipt",
                    entity_client_name="Corinthian Pins Limited"
                )
        else:
            st.info("Check the confirmation box above to unlock final invoicing and discharge workspaces.")
    else:
        st.info("To generate Master Tax Invoices and Discharge Receipts, change Operational Status to **Delivered (Finalized & Reconciled)** in Section 1.")