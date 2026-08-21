import os
import platform
import subprocess
import urllib.request
from datetime import datetime
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_documents")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
ENTITIES_CSV = os.path.join(BASE_DIR, "entities_config.csv")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Jinja Environment setup with Currency Formatting Filter
env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
env.filters['currency'] = lambda val: f"{float(val):,.2f}" if val is not None else "0.00"

def ensure_linux_wkhtmltopdf():
    """
    Downloads and extracts a standalone wkhtmltopdf binary on Linux environments 
    (Streamlit Cloud) when apt-get installation is unavailable.
    """
    local_bin = os.path.join(BASE_DIR, "bin", "wkhtmltopdf")
    if os.path.exists(local_bin):
        return local_bin

    bin_dir = os.path.join(BASE_DIR, "bin")
    os.makedirs(bin_dir, exist_ok=True)
    
    deb_url = "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox_0.12.6-1.buster_amd64.deb"
    deb_path = os.path.join(bin_dir, "wkhtmltox.deb")
    
    try:
        urllib.request.urlretrieve(deb_url, deb_path)
        subprocess.run(["ar", "x", deb_path], cwd=bin_dir, check=True)
        subprocess.run(["tar", "-xf", os.path.join(bin_dir, "data.tar.xz")], cwd=bin_dir, check=True)
        
        extracted_bin = os.path.join(bin_dir, "usr", "local", "bin", "wkhtmltopdf")
        if os.path.exists(extracted_bin):
            os.rename(extracted_bin, local_bin)
            os.chmod(local_bin, 0o755)
            return local_bin
    except Exception as e:
        print(f"Error auto-installing wkhtmltopdf binary: {e}")
        
    return None

def get_entity_info(company_name):
    """
    Fetches company details from entities_config.csv and converts logo, 
    stamp, and signature files to absolute file:/// URIs for wkhtmltopdf.
    """
    if not os.path.exists(ENTITIES_CSV):
        return {"Company_Name": company_name, "Registered_Address": "", "logo_path": "", "stamp_path": "", "sig_path": ""}
        
    df = pd.read_csv(ENTITIES_CSV)
    match = df[df["Company_Name"].str.strip().str.lower() == company_name.strip().lower()]
    
    if not match.empty:
        data = match.iloc[0].to_dict()
        for key in ["logo_file", "stamp_file", "sig_file"]:
            val = data.get(key)
            path_key = key.replace("_file", "_path")
            if pd.notna(val) and str(val).strip():
                full_path = os.path.join(ASSETS_DIR, str(val).strip()).replace("\\", "/")
                data[path_key] = f"file:///{full_path}"
            else:
                data[path_key] = ""
        return data
        
    return {
        "Company_Name": company_name,
        "Registered_Address": "",
        "logo_path": "",
        "stamp_path": "",
        "sig_path": ""
    }

def generate_pdf(template_name, context, output_filename, bl_no):
    """
    Renders Jinja2 HTML template and compiles PDF document via pdfkit.
    Detects Windows vs Linux environment dynamically for wkhtmltopdf binary path.
    """
    shipment_dir = os.path.join(OUTPUT_DIR, str(bl_no).strip())
    os.makedirs(shipment_dir, exist_ok=True)
    
    template = env.get_template(template_name)
    html_out = template.render(context)
    output_path = os.path.join(shipment_dir, output_filename)
    
    options = {
        'page-size': 'Letter',
        'margin-top': '0.5in',
        'margin-right': '0.5in',
        'margin-bottom': '0.5in',
        'margin-left': '0.5in',
        'encoding': "UTF-8",
        'enable-local-file-access': None
    }
    
    config = None
    if platform.system() == "Windows":
        path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        if os.path.exists(path_wkhtmltopdf):
            config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    elif platform.system() == "Linux":
        if os.path.exists('/usr/bin/wkhtmltopdf'):
            path_wkhtmltopdf = '/usr/bin/wkhtmltopdf'
        else:
            path_wkhtmltopdf = ensure_linux_wkhtmltopdf()
            
        if path_wkhtmltopdf and os.path.exists(path_wkhtmltopdf):
            config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    pdfkit.from_string(html_out, output_path, options=options, configuration=config)
    return output_path

# 1A. ADVANCE PORT DISBURSEMENT REQUEST
def generate_port_disbursement_request(bl_no, container_no, port_items=None, issue_date=None):
    corinthian = get_entity_info("Corinthian Pins Limited")
    rattans = get_entity_info("Rattans Freezone Limited")
    
    if port_items is None:
        port_items = [
            {"desc": "Customs Import Duty", "amount": 12000.00},
            {"desc": "Customs Import VAT (12.5%)", "amount": 11250.00},
            {"desc": "Port Authority Handling & Wharfage", "amount": 4250.00},
            {"desc": "Shipping Line Local Demurrage & Terminal Charges", "amount": 2500.00}
        ]
    
    subtotal = sum(item["amount"] for item in port_items)
    vat = 0.00
    total_due = subtotal + vat
    issue_date = issue_date or datetime.now().strftime("%d/%m/%Y")
    
    context = {
        "issuer": corinthian,
        "recipient": rattans,
        "bl_no": bl_no,
        "container_no": container_no,
        "issue_date": issue_date,
        "title": "ADVANCE PORT DISBURSEMENT REQUEST",
        "items": port_items,
        "subtotal": subtotal,
        "vat": vat,
        "total_due": total_due
    }
    return generate_pdf("disbursement_request.html", context, f"Disbursement_Request_Port_{bl_no}.pdf", bl_no)

# 1B. SERVICE FEE DISBURSEMENT REQUEST
def generate_service_disbursement_request(bl_no, container_no, bundled_service_fee=25000.00, issue_date=None):
    corinthian = get_entity_info("Corinthian Pins Limited")
    rattans = get_entity_info("Rattans Freezone Limited")
    
    items = [
        {"desc": "Customs Clearance Agency & Integrated Freight Management Services", "amount": bundled_service_fee}
    ]
    subtotal = bundled_service_fee
    vat = subtotal * 0.125
    total_due = subtotal + vat
    issue_date = issue_date or datetime.now().strftime("%d/%m/%Y")
    
    context = {
        "issuer": corinthian,
        "recipient": rattans,
        "bl_no": bl_no,
        "container_no": container_no,
        "issue_date": issue_date,
        "title": "SERVICE FEE DISBURSEMENT REQUEST",
        "items": items,
        "subtotal": subtotal,
        "vat": vat,
        "total_due": total_due
    }
    return generate_pdf("disbursement_request.html", context, f"Disbursement_Request_Services_{bl_no}.pdf", bl_no)

# 2. PRIVATE UPSTREAM FREIGHT INVOICE (Internal TFS -> Corinthian)
def generate_internal_tfs_invoice(bl_no, container_no, tfs_base=5500.00, issue_date=None):
    tfs = get_entity_info("Trinidad Freight Solutions Limited")
    corinthian = get_entity_info("Corinthian Pins Limited")
    vat_amount = tfs_base * 0.125
    issue_date = issue_date or datetime.now().strftime("%d/%m/%Y")
    
    context = {
        "issuer": tfs,
        "recipient": corinthian,
        "bl_no": bl_no,
        "container_no": container_no,
        "issue_date": issue_date,
        "freight_fee": tfs_base,
        "vat_amount": vat_amount,
        "total_due": tfs_base + vat_amount
    }
    return generate_pdf("tfs_internal_invoice.html", context, f"TFS_Internal_Invoice_{bl_no}.pdf", bl_no)

# 3. MASTER AGENCY & DISBURSEMENT TAX INVOICE
def generate_master_invoice(bl_no, container_no, usd_cargo_val=50000.00, exchange_rate=6.80, port_items=None, bundled_service_fee=25000.00, contra_deposit_paid=30000.00, issue_date=None):
    corinthian = get_entity_info("Corinthian Pins Limited")
    rattans = get_entity_info("Rattans Freezone Limited")
    
    ttd_cargo_val = usd_cargo_val * exchange_rate
    
    if port_items is None:
        port_items = [
            {"desc": "Customs Import Duty", "amount": 12000.00},
            {"desc": "Customs Import VAT (12.5%)", "amount": 11250.00},
            {"desc": "Port Authority Handling & Wharfage", "amount": 4250.00},
            {"desc": "Shipping Line Local Demurrage & Terminal Charges", "amount": 2500.00}
        ]
        
    total_port_outlays = sum(item["amount"] for item in port_items)
    service_vat = bundled_service_fee * 0.125
    
    gross_shipment_total = ttd_cargo_val + total_port_outlays + bundled_service_fee + service_vat
    net_contra_due = gross_shipment_total - ttd_cargo_val - contra_deposit_paid
    issue_date = issue_date or datetime.now().strftime("%d/%m/%Y")
    
    context = {
        "issuer": corinthian,
        "recipient": rattans,
        "bl_no": bl_no,
        "container_no": container_no,
        "issue_date": issue_date,
        "usd_cargo_val": usd_cargo_val,
        "exchange_rate": exchange_rate,
        "ttd_cargo_val": ttd_cargo_val,
        "port_items": port_items,
        "total_port_outlays": total_port_outlays,
        "bundled_service_fee": bundled_service_fee,
        "service_vat": service_vat,
        "gross_shipment_total": gross_shipment_total,
        "contra_deposit_paid": contra_deposit_paid,
        "net_contra_due": net_contra_due
    }
    return generate_pdf("master_agency_tax_invoice.html", context, f"Master_Tax_Invoice_{bl_no}.pdf", bl_no)

# 4. OFFICIAL PAYMENT RECEIPT & ACCOUNT DISCHARGE
def generate_official_receipt(bl_no, container_no, amount_paid=28125.00, issue_date=None):
    corinthian = get_entity_info("Corinthian Pins Limited")
    rattans = get_entity_info("Rattans Freezone Limited")
    issue_date = issue_date or datetime.now().strftime("%d/%m/%Y")
    
    context = {
        "issuer": corinthian,
        "recipient": rattans,
        "bl_no": bl_no,
        "container_no": container_no,
        "amount_paid": amount_paid,
        "issue_date": issue_date
    }
    return generate_pdf("official_receipt.html", context, f"Official_Receipt_{bl_no}.pdf", bl_no)