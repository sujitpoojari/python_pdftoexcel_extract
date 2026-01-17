import os
import re
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from datetime import datetime

# ================= CONFIG =================
PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"
OUTPUT_PREFIX = "Final_Output_VERIFIED"
# =========================================

# ---------- TEXT HELPERS ----------
def clean_text(text):
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def find_first(patterns, text):
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return m.group(1)
    return None

# ---------- LOAD TEMPLATE ----------
template_df = pd.read_excel(TEMPLATE_FILE)
columns = template_df.columns.tolist()

rows = []

# ---------- PROCESS PDFs ----------
for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print(f"Processing: {pdf}")

    images = convert_from_path(os.path.join(PDF_FOLDER, pdf))
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img)

    text = clean_text(text)

    # ---------- INIT ROW ----------
    row = {col: None for col in columns}
    row["Field"] = pdf

    # ---------- ORDER NUMBER ----------
    row["order_number"] = find_first(
        [r"(\d{3}-\d{7}-\d{7})"], text
    )

    # ---------- INVOICE NUMBER ----------
    row["invoice_number"] = find_first(
        [
            r"Invoice\s*(?:Number|No|#)?\s*[:\-]?\s*([A-Z0-9\-]{8,})",
            r"Tax\s*Invoice\s*([A-Z0-9\-]{8,})"
        ],
        text
    )

    # ---------- INVOICE DATE ----------
    row["invoice_date"] = find_first(
        [r"(\d{2}[./-]\d{2}[./-]\d{4}|\d{4}[./-]\d{2}[./-]\d{2})"],
        text
    )

    # ---------- SELLER NAME ----------
    row["seller_name"] = find_first(
        [
            r"(?:Sold\s*By|Seller)\s*[:\-]?\s*(.*?)\s*(?:GST|PAN|Invoice|Address)"
        ],
        text
    )

    # ---------- GST / PAN ----------
    row["seller_pan"] = find_first(
        [r"(\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z]\w)"],
        text
    )

    # ---------- TOTAL AMOUNT ----------
    row["total_amount"] = find_first(
        [
            r"(?:Grand\s*Total|Invoice\s*Value|Amount\s*Payable)[^\d]{0,15}([\d,]+\.\d{2})"
        ],
        text
    )

    # ---------- TOTAL TAX ----------
    tax_matches = re.findall(
        r"(IGST|CGST|SGST)[^\d]{0,10}([\d,]+\.\d{2})",
        text,
        re.I
    )

    if tax_matches:
        row["total_tax"] = round(
            sum(float(t[1].replace(",", "")) for t in tax_matches), 2
        )

    rows.append(row)

# ---------- FINAL DATAFRAME ----------
df = pd.DataFrame(rows)

# Ensure correct column order
df = df[columns]

# ---------- WRITE OUTPUT ----------
output_file = f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output_file, index=False)

print(f"\n✅ SUCCESS: {len(df)} invoices written to {output_file}")
