import os
import re
import pdfplumber
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from datetime import datetime
from num2words import num2words  # pip install num2words

# ================= CONFIG =================
PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"
OUTPUT_PREFIX = "Final_Output_MULTI_VENDOR_FINAL"
# =========================================

# -----------------------------
# TEXT EXTRACTION
# -----------------------------
def extract_text(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception:
        pass
    if len(text.strip()) < 200:
        images = convert_from_path(pdf_path)
        for img in images:
            text += pytesseract.image_to_string(img) + "\n"
    return text

# -----------------------------
# HELPERS
# -----------------------------
def extract_field(text, patterns):
    for pat in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            return m.group(1).strip()
    return None

def normalize_number(val):
    if not val:
        return None
    return re.sub(r"[₹,\s]", "", val)

def clean_address(val):
    if not val:
        return None
    val = re.sub(r"\*.*", "", val)  # remove footnotes
    return re.sub(r"\s{2,}", "\n", val).strip()

def clean_amount_in_words(val):
    if not val:
        return None
    val = val.split("\n")[0].strip()
    stop_words = ["For", "Authorized Signatory", "Whether", "*ASSPL"]
    for word in stop_words:
        if word in val:
            val = val.split(word)[0].strip()
    return val

def parse_sold_by_block(text):
    pat = r"Sold By\s*[:\-]?\s*(.*?)\s*(?=PAN\s*[:\-]?|GSTIN|Invoice\s*No|Order\s*Id|Billing\s*Address|Shipping\s*Address)"
    m = re.search(pat, text, re.I | re.S)
    if not m:
        return None, None, None
    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    seller_name = lines[0] if lines else None
    seller_address = "\n".join(lines[1:]) if len(lines) > 1 else None
    return block, seller_name, seller_address

def extract_seller_flexible(text, billing_address):
    seller_info, seller_name, seller_address = parse_sold_by_block(text)
    if seller_name:
        return seller_info, seller_name, seller_address
    if billing_address:
        parts = [p.strip() for p in billing_address.split("\n") if p.strip()]
        if parts:
            seller_name = parts[0]
            seller_address = "\n".join(parts[1:]) if len(parts) > 1 else None
            return parts[0], seller_name, seller_address
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines:
        if re.match(r"^[A-Z][A-Z\s,]{3,}", line):
            parts = [p.strip() for p in line.split(",") if p.strip()]
            seller_name = parts[0] if parts else line
            seller_address = "\n".join(parts[1:]) if len(parts) > 1 else None
            return line, seller_name, seller_address
    return None, None, None

def extract_state_codes(text):
    codes = re.findall(r"State\/UT\s*Code\s*[:\-]?\s*(\d{1,2})", text, re.I)
    if not codes:
        return None, None
    if len(codes) == 1:
        return codes[0], codes[0]
    return codes[0], codes[1]

def extract_tax(text, vendor="Amazon"):
    if vendor.lower() == "flipkart":
        m = re.search(r"IGST\s*[:\-]?\s*([\d,]+\.\d{2})", text, re.I)
        if m:
            return normalize_number(m.group(1))
    tax_lines = re.findall(r"(?:Total\s*Tax\s*Amount|Tax\s*Amount|Tax)[^\d]{0,5}[:\-]?\s*₹?\s*([\d,]+\.\d{2})", text, re.I)
    if tax_lines:
        return normalize_number(tax_lines[-1])
    return None

def extract_total_amount(text):
    patterns = [r"(Invoice\s*Value|TOTAL\s*Amount|Grand\s*Total)[^\d]{0,20}([\d,]+\.\d{2})"]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return normalize_number(m.group(2))
    nums = re.findall(r"\b\d+\.\d{2}\b", text)
    if nums:
        return nums[-1]
    return None

def detect_vendor(text):
    t = text.lower()
    if "amazon" in t:
        return "amazon"
    if "flipkart" in t or "shopler estore" in t:
        return "flipkart"
    if "swiggy" in t:
        return "swiggy"
    return "unknown"

# -----------------------------
# AMAZON EXTRACTION
# -----------------------------
def extract_amazon(text, row):
    row["invoice_number"] = extract_field(text, [r"Invoice\s*Number\s*[:\-]?\s*([A-Z0-9\-]+)"])
    row["invoice_date"] = extract_field(text, [r"Invoice\s*Date\s*[:\-]?\s*([\d./-]+)"])
    row["order_number"] = extract_field(text, [r"Order\s*Number\s*[:\-]?\s*([\d\-]+)"])
    row["order_date"] = row["invoice_date"]
    row["seller_pan"] = extract_field(text, [r"PAN\s*No\s*[:\-]?\s*([A-Z0-9]+)"])
    row["seller_gst"] = extract_field(text, [r"GST\s*Registration\s*No\s*[:\-]?\s*([A-Z0-9]+)"])
    row["billing_address"] = clean_address(extract_field(text, [r"Billing\s*Address\s*:\s*(.*?)(?=Shipping|Invoice|Order)"]))
    row["shipping_address"] = clean_address(extract_field(text, [r"Shipping\s*Address\s*:\s*(.*?)(?=Invoice|Order)"]))
    row["place_of_supply"] = extract_field(text, [r"Place\s*of\s*supply\s*[:\-]?\s*([A-Z ]+)"])
    taxes = re.findall(r"(CGST|SGST|IGST)[^\d]{0,10}([\d,]+\.\d{2})", text, re.I)
    if taxes:
        row["total_tax"] = round(sum(float(v.replace(",", "")) for _, v in taxes), 2)
    row["total_amount"] = extract_total_amount(text)
    amt_words = extract_field(text, [r"Amount\s*in\s*Words\s*:\s*(.*)"])
    row["amount_in_words"] = clean_amount_in_words(amt_words)

# -----------------------------
# FLIPKART EXTRACTION
# -----------------------------
def extract_flipkart(text, row):
    # same as your previous Flipkart code
    row["invoice_number"] = extract_field(text, [r"Invoice\s*No\s*[:\-]?\s*([A-Z0-9]+)"])
    row["order_number"] = extract_field(text, [r"Order\s*ID\s*[:\-]?\s*(OD\d+)"])
    row["invoice_date"] = extract_field(text, [r"Invoice\s*Date\s*[:\-]?\s*([\d]{2}[-/][\d]{2}[-/][\d]{4})"])
    row["order_date"] = row["invoice_date"]
    row["seller_name"] = extract_field(text, [r"Sold\s*By\s*[:\-]?\s*(.*?)(?=GSTIN|PAN|Invoice)"])
    row["seller_pan"] = extract_field(text, [r"PAN\s*[:\-]?\s*([A-Z0-9]+)"])
    row["seller_gst"] = extract_field(text, [r"GSTIN\s*[:\-]?\s*([A-Z0-9]+)"])
    row["billing_address"] = clean_address(extract_field(text, [r"Billing\s*Address\s*:\s*(.*?)(?=Shipping|Invoice)"]))
    row["shipping_address"] = clean_address(extract_field(text, [r"Shipping\s*Address\s*:\s*(.*?)(?=Invoice|Sold)"]))
    row["place_of_supply"] = extract_field(text, [r"Place\s*of\s*Supply\s*[:\-]?\s*([A-Z ]+)"])
    row["place_of_delivery"] = row["billing_address"]
    igst = extract_field(text, [r"IGST\s*[:\-]?\s*([\d,]+\.\d{2})"])
    cgst = extract_field(text, [r"CGST\s*[:\-]?\s*([\d,]+\.\d{2})"])
    sgst = extract_field(text, [r"SGST\s*[:\-]?\s*([\d,]+\.\d{2})"])
    tax = 0
    for v in [igst, cgst, sgst]:
        if v:
            tax += float(v.replace(",", ""))
    if tax:
        row["total_tax"] = round(tax, 2)
    row["total_amount"] = extract_total_amount(text)
    amt_words = extract_field(text, [r"Amount\s*in\s*Words\s*:\s*(.*)"])
    row["amount_in_words"] = clean_amount_in_words(amt_words)

# -----------------------------
# SWIGGY EXTRACTION
# -----------------------------
def extract_swiggy(text, pdf_name):
    rows = []
    blocks = re.split(r"\bTAX\s+INVOICE\b", text, flags=re.I)
    for b in blocks:
        inv = extract_field(b, [r"Invoice\s*Value\s*([\d]+)"])
        if not inv or inv in ("0", "00"):
            continue
        row = {c: None for c in TEMPLATE_COLUMNS}
        row["Field"] = pdf_name
        row["invoice_number"] = extract_field(b, [r"Invoice\s*No\s*[:\-]?\s*([A-Z0-9]+)"])
        row["order_number"] = extract_field(b, [r"Order\s*ID\s*[:\-]?\s*(\d+)"])
        row["invoice_date"] = extract_field(b, [r"Date\s*of\s*Invoice\s*[:\-]?\s*([\d\-]+)"])
        row["order_date"] = row["invoice_date"]
        row["seller_name"] = extract_field(b, [r"Seller\s*Name\s*[:\-]?\s*([A-Z][A-Z\s&.-]+)"])
        row["seller_gst"] = extract_field(b, [r"Seller\s*GSTIN\s*[:\-]?\s*([A-Z0-9]+)"])
        row["fssai_license"] = extract_field(b, [r"FSSAI\s*[:\-]?\s*(\d+)"])
        row["billing_address"] = clean_address(extract_field(b, [r"Customer\s*Address\s*:\s*(.*?)(?=Order\s*ID|Invoice\s*No|FSSAI)"]))
        row["place_of_delivery"] = row["billing_address"]
        cgst = extract_field(b, [r"Total\s*CGST\s*([\d]+\.\d{2})"])
        sgst = extract_field(b, [r"Total\s*SGST\s*([\d]+\.\d{2})"])
        if cgst and sgst:
            row["total_tax"] = round(float(cgst) + float(sgst), 2)
        row["total_amount"] = inv
        amt_words = extract_field(b, [r"Amount\s*in\s*words\s*[:\-]?\s*([A-Za-z\s]+?)(?=\n|Invoice|Whether|Discount|Disclaimer)"])
        row["amount_in_words"] = clean_amount_in_words(amt_words)
        rows.append(row)
    return rows

# -----------------------------
# MAIN LOOP
# -----------------------------
template_df = pd.read_excel(TEMPLATE_FILE)
TEMPLATE_COLUMNS = template_df.columns.tolist()
final_rows = []

for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print("\n" + "=" * 80)
    print(f"📄 Processing PDF: {pdf}")
    print("=" * 80)

    text = extract_text(os.path.join(PDF_FOLDER, pdf))
    vendor = detect_vendor(text)

    row = {c: None for c in TEMPLATE_COLUMNS}
    row["Field"] = pdf

    if vendor == "swiggy":
        rows = extract_swiggy(text, pdf)
        for r in rows:
            for k, v in r.items():
                print(f"{k:25} : {v}")
        final_rows.extend(rows)
        continue

    if vendor == "amazon":
        extract_amazon(text, row)
    elif vendor == "flipkart":
        extract_flipkart(text, row)

    seller_info, seller_name, seller_address = extract_seller_flexible(text, row.get("billing_address"))
    row["seller_info"] = seller_info
    row["seller_name"] = seller_name
    row["seller_address"] = seller_address

    billing_state_code, shipping_state_code = extract_state_codes(text)
    row["billing_state_code"] = billing_state_code
    row["shipping_state_code"] = shipping_state_code

    for k, v in row.items():
        print(f"{k:25} : {v}")

    final_rows.append(row)

# -----------------------------
# OUTPUT
# -----------------------------
df = pd.DataFrame(final_rows)
df = df.reindex(columns=TEMPLATE_COLUMNS)
output_file = f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)
print(f"\n✅ SUCCESS: {len(df)} invoices written to {output_file}")
