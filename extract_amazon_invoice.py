import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

# ------------------------------
PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"
# ------------------------------

rows = []

# ------------------------------
# Helpers
# ------------------------------

def extract_text(pdf_path):
    """Extract concatenated text from all pages."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

def extract_field(text, patterns):
    """Try multiple regex patterns; return first match group(1) or full match."""
    for pat in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            if m.lastindex and m.lastindex >= 1:
                return m.group(1).strip()
            else:
                return m.group(0).strip()
    return None

def normalize_number(s):
    """Strip ₹, commas, spaces; return plain numeric string or None."""
    if not s:
        return None
    s = re.sub(r"[₹,\s]", "", s)
    return s if s else None

def clean_address(raw_text):
    """Remove leaked seller info and normalize whitespace."""
    if not raw_text:
        return None
    cleaned = re.sub(r"Sold By.*", "", raw_text, flags=re.I)
    cleaned = re.sub(r"WISELIFE WELLNESS.*", "", cleaned, flags=re.I)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def parse_sold_by_block(text):
    """
    Parse the Sold By block into seller_info, seller_name, seller_address.
    """
    block_pat = r"Sold By\s*[:\-]?\s*(.*?)\s*(?=PAN\s*No|GST\s*Registration\s*No|Order\s*Number|Invoice\s*Number|Ship From|Shipping Address|Billing Address)"
    m = re.search(block_pat, text, flags=re.I | re.S)
    if not m:
        return None, None, None

    block = m.group(1).strip()
    lines = [ln.strip() for ln in re.split(r"\n|\r\n", block) if ln.strip()]

    if not lines:
        return re.sub(r"\s+", " ", block).strip(), None, None

    seller_name = lines[0]
    seller_address = " ".join(lines[1:]).strip() if len(lines) > 1 else None
    seller_info = re.sub(r"\s+", " ", block).strip()

    return seller_info, seller_name, seller_address

def extract_tax_and_total(text):
    """
    Extract FINAL tax amount and total amount.
    """
    tax_amount = None
    total_amount = None

    # TOTAL row with two amounts (tax + total)
    m_total_row = re.findall(
        r"TOTAL\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)\s+₹?\s*([\d,]+(?:\.\d+)?)",
        text,
        flags=re.I
    )
    if m_total_row:
        last_tax, last_total = m_total_row[-1]
        tax_amount = normalize_number(last_tax)
        total_amount = normalize_number(last_total)

    # Explicit "Total Tax Amount"
    if tax_amount is None:
        m_tax = re.findall(r"Total\s*Tax\s*Amount\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)", text, flags=re.I)
        if m_tax:
            tax_amount = normalize_number(m_tax[-1])

    # Explicit "TOTAL:" single amount
    if total_amount is None:
        m_total = re.findall(r"TOTAL\s*Amount\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)", text, flags=re.I)
        if m_total:
            total_amount = normalize_number(m_total[-1])

    # Fallback: "Invoice Value"
    if total_amount is None:
        m_invoice_val = re.findall(r"Invoice\s*Value\s*[:\-]?\s*₹?\s*([\d,]+(?:\.\d+)?)", text, flags=re.I)
        if m_invoice_val:
            total_amount = normalize_number(m_invoice_val[-1])

    return tax_amount, total_amount

# ------------------------------
# Field patterns (all with capturing groups)
# ------------------------------

field_patterns = {
    "invoice_type": [
        r"(Tax\s+Invoice\/Bill\s+of\s+Supply\/Cash\s+Memo)",
        r"(#\s*TAX\s+INVOICE)",
        r"(\bTAX\s+INVOICE\b)",
        r"(\bINV\b)"
    ],
    "invoice_number": [r"Invoice\s*Number\s*[:\-]?\s*([A-Z0-9\-]+)"],
    "invoice_date": [r"Invoice\s*Date\s*[:\-]?\s*([\d\.\/\-]+)"],
    "order_number": [r"Order\s*Number\s*[:\-]?\s*([\d\-]+)"],
    "order_date": [r"Order\s*Date\s*[:\-]?\s*([\d\.\/\-]+)"],
    "invoice_details": [r"Invoice\s*Details\s*[:\-]?\s*([A-Z0-9\-]+)"],
    "seller_pan": [r"PAN\s*No\s*[:\-]?\s*([A-Z0-9]+)"],
    "seller_gst": [r"GST\s*Registration\s*No\s*[:\-]?\s*([A-Z0-9]+)"],
    "fssai_license": [r"FSSAI\s*License\s*No\.?\s*([0-9]+)"],
    "billing_address": [
        r"Billing\s*Address\s*:\s*(.*?)\s*(?=Shipping\s*Address|State\/UT\s*Code|Invoice\s*Number|Order\s*Number)"
    ],
    "shipping_address": [
        r"Shipping\s*Address\s*:\s*(.*?)\s*(?=State\/UT\s*Code|Place\s*of\s*supply|Invoice\s*Number|Order\s*Number)"
    ],
    "billing_state_code": [r"Billing\s*Address.*?State\/UT\s*Code\s*[:\-]?\s*(\d+)"],
    "shipping_state_code": [r"Shipping\s*Address.*?State\/UT\s*Code\s*[:\-]?\s*(\d+)"],
    "place_of_supply": [r"Place\s*of\s*supply\s*[:\-]?\s*([A-Z ]+)"],
    "place_of_delivery": [r"Place\s*of\s*delivery\s*[:\-]?\s*([A-Z ]+)"],
    "reverse_charge": [r"Whether\s+tax\s+is\s+payable\s+under\s+reverse\s+charge\s*[-:]?\s*(\w+)"],
    "amount_in_words": [r"Amount\s+in\s+Words\s*:\s*(.*?)\n"]
}

# ------------------------------
# Process PDFs
# ------------------------------

if not os.path.exists(PDF_FOLDER):
    raise Exception(f"Folder not found: {PDF_FOLDER}")

pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
if not pdf_files:
    raise Exception("No PDF files found in invoices folder")

for pdf_file in pdf_files:
    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    print(f"\nProcessing: {pdf_file}")

    text = extract_text(pdf_path)
    if not text.strip():
        print(f"❌ Skipping empty PDF: {pdf_file}")
        continue

    row = {"Source File": pdf_file}

    # Parse Sold By block into seller_info, seller_name, seller_address
    seller_info, seller_name, seller_address = parse_sold_by_block(text)
    row["seller_info"] = seller_info
    row["seller_name"] = seller_name
    row["seller_address"] = seller_address

    # Extract other fields
    for field, patterns in field_patterns.items():
        value = extract_field(text, patterns)
        if field in ["billing_address", "shipping_address"]:
            value = clean_address(value)
        row[field] = value

    # Extract tax and total amounts correctly (final totals)
    tax_amount, total_amount = extract_tax_and_total(text)
    row["total_tax"] = tax_amount
    row["total_amount"] = total_amount

    rows.append(row)

# ------------------------------
# Verify extraction
# ------------------------------

if not rows:
    raise Exception("❌ No invoice data extracted. Ensure PDFs are text-based.")

df = pd.DataFrame(rows)

# ------------------------------
# Align with template columns
# ------------------------------

try:
    template_df = pd.read_excel(TEMPLATE_FILE)
    for col in template_df.columns:
        if col not in df.columns:
            df[col] = None
    df = df.reindex(columns=template_df.columns)
except FileNotFoundError:
    print("⚠️ Template not found. Using extracted columns only.")

# ------------------------------
# Write Excel with timestamp
# ------------------------------

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
final_output_file = f"Final_Output_{timestamp}.xlsx"
df.to_excel(final_output_file, index=False)
print(f"\n✅ SUCCESS! Data extracted into '{final_output_file}'")
