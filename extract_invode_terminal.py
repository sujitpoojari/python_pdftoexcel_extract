import os
import re
import pdfplumber
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from datetime import datetime

# ================= CONFIG =================
PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"
OUTPUT_PREFIX = "Final_Output_PRODUCTION_FINAL_V2"
# =========================================

# -----------------------------
# TEXT EXTRACTION (pdfplumber -> OCR)
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
    if len(text.strip()) < 200:  # OCR fallback
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
    return re.sub(r"\s+", " ", val).strip()

# -----------------------------
# SOLD BY BLOCK
# -----------------------------
def parse_sold_by_block(text):
    pat = (
        r"Sold By\s*[:\-]?\s*(.*?)\s*"
        r"(?=PAN\s*No|GST\s*Registration\s*No|Invoice\s*Number|Order\s*Number|Billing\s*Address|Shipping\s*Address)"
    )
    m = re.search(pat, text, re.I | re.S)
    if not m:
        return None, None, None
    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    seller_name = lines[0] if lines else None
    seller_address = " ".join(lines[1:]) if len(lines) > 1 else None
    return block, seller_name, seller_address

# -----------------------------
# FLEXIBLE SELLER EXTRACTION
# -----------------------------
def extract_seller_flexible(text, billing_address):
    # 1️⃣ Try Sold By block
    seller_info, seller_name, seller_address = parse_sold_by_block(text)
    if seller_name:
        return seller_info, seller_name, seller_address

    # 2️⃣ Fallback to first line of billing address
    if billing_address:
        parts = [p.strip() for p in billing_address.split(",") if p.strip()]
        if parts:
            seller_name = parts[0]
            seller_address = ", ".join(parts[1:]) if len(parts) > 1 else None
            return parts[0], seller_name, seller_address

    # 3️⃣ Fallback to first capitalized line in text
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for line in lines:
        if re.match(r"^[A-Z][A-Z\s,]{3,}", line):
            parts = [p.strip() for p in line.split(",") if p.strip()]
            seller_name = parts[0] if parts else line
            seller_address = ", ".join(parts[1:]) if len(parts) > 1 else None
            return line, seller_name, seller_address

    return None, None, None

# -----------------------------
# STATE / UT CODE
# -----------------------------
def extract_state_codes(text):
    codes = re.findall(r"State\/UT\s*Code\s*[:\-]?\s*(\d{1,2})", text, re.I)
    if not codes:
        return None, None
    if len(codes) == 1:
        return codes[0], codes[0]
    return codes[0], codes[1]

# -----------------------------
# TAX & TOTAL
# -----------------------------
def extract_tax(text):
    tax_values = re.findall(r"(IGST|CGST|SGST)[^\d]{0,10}([\d,]+\.\d{2})", text, re.I)
    if not tax_values:
        return None
    return round(sum(float(v.replace(",", "")) for _, v in tax_values), 2)

def extract_total_amount(text):
    patterns = [
        r"(Invoice\s*Value|TOTAL\s*Amount|Grand\s*Total)[^\d]{0,20}([\d,]+\.\d{2})"
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return normalize_number(m.group(2))
    # Amount-in-words fallback
    nums = re.findall(r"\b\d+\.\d{2}\b", text)
    if nums:
        return nums[-1]
    return None

# -----------------------------
# FIELD PATTERNS
# -----------------------------
FIELD_PATTERNS = {
    "invoice_number": [r"Invoice\s*Number\s*[:\-]?\s*([A-Z0-9\-]+)"],
    "invoice_date": [r"Invoice\s*Date\s*[:\-]?\s*([\d./-]+)"],
    "order_number": [r"Order\s*Number\s*[:\-]?\s*([\d\-]+)"],
    "order_date": [r"Order\s*Date\s*[:\-]?\s*([\d./-]+)"],
    "seller_pan": [r"PAN\s*No\s*[:\-]?\s*([A-Z0-9]+)"],
    "seller_gst": [r"GST\s*Registration\s*No\s*[:\-]?\s*([A-Z0-9]+)"],
    "billing_address": [
        r"Billing\s*Address\s*:\s*(.*?)\s*(?=Shipping\s*Address|Invoice\s*Number|Order\s*Number)"
    ],
    "shipping_address": [
        r"Shipping\s*Address\s*:\s*(.*?)\s*(?=Invoice\s*Number|Order\s*Number|Place\s*of)"
    ],
    "place_of_supply": [r"Place\s*of\s*supply\s*[:\-]?\s*([A-Z ]+)"],
    "amount_in_words": [r"Amount\s+in\s+Words\s*:\s*(.*?)\n"]
}

# -----------------------------
# MAIN
# -----------------------------
template_df = pd.read_excel(TEMPLATE_FILE)
template_columns = template_df.columns.tolist()

rows = []

for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print("\n" + "=" * 80)
    print(f"📄 Processing PDF: {pdf}")
    print("=" * 80)

    text = extract_text(os.path.join(PDF_FOLDER, pdf))

    row = {col: None for col in template_columns}
    row["Field"] = pdf  # Field column in template

    # Standard fields
    for field, patterns in FIELD_PATTERNS.items():
        val = extract_field(text, patterns)
        if field in ("billing_address", "shipping_address"):
            val = clean_address(val)
        row[field] = val

    # Seller info (enhanced fallback)
    seller_info, seller_name, seller_address = extract_seller_flexible(text, row.get("billing_address"))
    row["seller_info"] = seller_info
    row["seller_name"] = seller_name
    row["seller_address"] = seller_address

    # State codes
    billing_state_code, shipping_state_code = extract_state_codes(text)
    row["billing_state_code"] = billing_state_code
    row["shipping_state_code"] = shipping_state_code

    # Tax & Total
    row["total_tax"] = extract_tax(text)
    row["total_amount"] = extract_total_amount(text)

    # Terminal debug print
    for k, v in row.items():
        print(f"{k:25} : {v}")

    rows.append(row)

# -----------------------------
# OUTPUT
# -----------------------------
df = pd.DataFrame(rows)
df = df.reindex(columns=template_columns)

output_file = f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n✅ SUCCESS: {len(df)} invoices written to {output_file}")
