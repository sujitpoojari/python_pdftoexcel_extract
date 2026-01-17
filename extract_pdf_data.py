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
OUTPUT_PREFIX = "Final_Output_PRODUCTION"
# =========================================


# -------------------------------------------------
# TEXT EXTRACTION (pdfplumber → OCR fallback)
# -------------------------------------------------
def extract_text(pdf_path):
    text = ""

    # 1️⃣ Try pdfplumber (BEST for text PDFs)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception:
        pass

    # 2️⃣ OCR fallback if text too small
    if len(text.strip()) < 200:
        images = convert_from_path(pdf_path)
        for img in images:
            text += pytesseract.image_to_string(img) + "\n"

    return text


# -------------------------------------------------
# HELPERS
# -------------------------------------------------
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
    val = re.sub(r"\s+", " ", val)
    return val.strip()


# -------------------------------------------------
# SOLD BY BLOCK (VERY IMPORTANT)
# -------------------------------------------------
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


# -------------------------------------------------
# TAX + TOTAL (FINAL ROW ONLY)
# -------------------------------------------------
def extract_tax_and_total(text):
    tax = None
    total = None

    # TOTAL row with two numbers
    rows = re.findall(
        r"TOTAL\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)\s+₹?\s*([\d,]+(?:\.\d+)?)",
        text,
        re.I
    )
    if rows:
        tax, total = rows[-1]

    if not total:
        m = re.findall(r"Invoice\s*Value\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)", text, re.I)
        if m:
            total = m[-1]

    return normalize_number(tax), normalize_number(total)


# -------------------------------------------------
# FIELD PATTERNS
# -------------------------------------------------
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


# -------------------------------------------------
# MAIN
# -------------------------------------------------
template_df = pd.read_excel(TEMPLATE_FILE)
template_columns = template_df.columns.tolist()

rows = []

for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print(f"Processing: {pdf}")
    text = extract_text(os.path.join(PDF_FOLDER, pdf))

    if not text.strip():
        print(f"❌ Skipped empty: {pdf}")
        continue

    row = {col: None for col in template_columns}
    row["Source File"] = pdf

    # Sold By
    seller_info, seller_name, seller_address = parse_sold_by_block(text)
    row["seller_info"] = seller_info
    row["seller_name"] = seller_name
    row["seller_address"] = seller_address

    # Other fields
    for field, patterns in FIELD_PATTERNS.items():
        val = extract_field(text, patterns)
        if field in ("billing_address", "shipping_address"):
            val = clean_address(val)
        row[field] = val

    # Tax + Total
    tax, total = extract_tax_and_total(text)
    row["total_tax"] = tax
    row["total_amount"] = total

    rows.append(row)


# -------------------------------------------------
# OUTPUT
# -------------------------------------------------
df = pd.DataFrame(rows)
df = df.reindex(columns=template_columns)

output_file = f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n✅ SUCCESS: {len(df)} invoices written to {output_file}")
