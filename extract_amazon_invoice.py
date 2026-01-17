import os
import re
import pdfplumber
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from datetime import datetime

PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"
OUTPUT_PREFIX = "Final_Output_PRODUCTION_V5"


# ==========================================================
# TEXT EXTRACTION
# ==========================================================
def extract_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += "\n" + t

    if len(text.strip()) < 300:
        images = convert_from_path(pdf_path)
        for img in images:
            text += pytesseract.image_to_string(img) + "\n"

    return text


# ==========================================================
# COMMON HELPERS
# ==========================================================
def clean(val):
    return re.sub(r"\s+", " ", val).strip() if val else None


def amount(val):
    if not val:
        return None
    return float(re.sub(r"[₹,\s]", "", val))


# ==========================================================
# TAX & TOTAL
# ==========================================================
def extract_total_tax(text):
    cgst = re.findall(r"Total\s+CGST\s+([\d.]+)", text, re.I)
    sgst = re.findall(r"Total\s+SGST\s+([\d.]+)", text, re.I)
    igst = re.findall(r"IGST[^\d]{0,10}([\d.]+)", text, re.I)

    if cgst or sgst:
        return round(sum(map(float, cgst + sgst)), 2)
    if igst:
        return round(sum(map(float, igst)), 2)
    return None


def extract_total_amount(text):
    m = re.findall(r"Invoice\s*Value\s*([\d.]+)", text, re.I)
    return amount(m[-1]) if m else None


# ==========================================================
# AMAZON & FLIPKART (SINGLE INVOICE)
# ==========================================================
def extract_standard_invoice(text, source):
    row = {"Field": source}

    patterns = {
        "invoice_number": r"Invoice\s*(No|Number)[:\-]?\s*([A-Z0-9\-]+)",
        "order_number": r"(Order\s*(ID|Number))[:\-]?\s*([A-Z0-9\-]+)",
        "invoice_date": r"Date\s*of\s*Invoice[:\-]?\s*([\d\-]+)",
        "seller_pan": r"PAN[:\-]?\s*([A-Z0-9]+)",
        "seller_gst": r"GST(IN)?[:\-]?\s*([A-Z0-9]+)",
        "fssai_license": r"FSSAI[:\-]?\s*(\d+)",
        "amount_in_words": r"Amount\s*in\s*words[:\-]?\s*(.*?)\n",
        "place_of_supply": r"Place\s*of\s*Supply[:\-]?\s*([A-Za-z ]+)",
    }

    for k, p in patterns.items():
        m = re.search(p, text, re.I | re.S)
        row[k] = clean(m.group(m.lastindex)) if m else None

    row["total_tax"] = extract_total_tax(text)
    row["total_amount"] = extract_total_amount(text)
    return row


# ==========================================================
# 🔥 SWIGGY MULTI-INVOICE HANDLER (FINAL FIX)
# ==========================================================
def extract_swiggy_invoices(full_text, source):
    rows = []

    # Split using the ONLY reliable anchor
    blocks = re.split(r"Invoice\s*No\s*:", full_text, flags=re.I)

    for b in blocks:
        if "Invoice Value" not in b:
            continue  # ❌ skip blank / service invoice

        text = "Invoice No: " + b  # restore removed text

        row = {"Field": source}

        row["invoice_number"] = clean(
            re.search(r"Invoice\s*No[:\-]?\s*([A-Z0-9]+)", text, re.I).group(1)
        )
        row["order_number"] = clean(
            re.search(r"Order\s*ID[:\-]?\s*([0-9]+)", text, re.I).group(1)
        )
        row["invoice_date"] = clean(
            re.search(r"Date\s*of\s*Invoice[:\-]?\s*([\d\-]+)", text, re.I).group(1)
        )

        row["seller_name"] = clean(
            re.search(r"Seller\s*Name[:\-]?\s*(.*?)\n", text, re.I).group(1)
        )

        row["seller_gst"] = clean(
            re.search(r"Seller\s*GSTIN[:\-]?\s*([A-Z0-9]+)", text, re.I).group(1)
        )

        row["fssai_license"] = clean(
            re.search(r"FSSAI[:\-]?\s*(\d+)", text, re.I).group(1)
        )

        row["place_of_supply"] = clean(
            re.search(r"Place\s*of\s*Supply[:\-]?\s*([A-Za-z ]+)", text, re.I).group(1)
        )

        row["amount_in_words"] = clean(
            re.search(r"Amount\s*in\s*words[:\-]?\s*(.*?)\n", text, re.I).group(1)
        )

        row["total_tax"] = extract_total_tax(text)
        row["total_amount"] = extract_total_amount(text)

        rows.append(row)

    return rows


# ==========================================================
# MAIN
# ==========================================================
template_cols = pd.read_excel(TEMPLATE_FILE).columns.tolist()
final_rows = []

for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print(f"\n📄 Processing: {pdf}")
    text = extract_text(os.path.join(PDF_FOLDER, pdf))

    if "Swiggy" in text:
        final_rows.extend(extract_swiggy_invoices(text, pdf))
    else:
        final_rows.append(extract_standard_invoice(text, pdf))


df = pd.DataFrame(final_rows)

for c in template_cols:
    if c not in df:
        df[c] = None

df = df[template_cols]

outfile = f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(outfile, index=False)

print(f"\n✅ SUCCESS: {len(df)} invoices written to {outfile}")
