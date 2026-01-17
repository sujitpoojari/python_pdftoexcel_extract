from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import re
import os
from datetime import datetime

PDF_FOLDER = "invoices"
TEMPLATE_FILE = "Output Template.xlsx"

rows = []

def normalize(text):
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"\s+", " ", text)
    return text

for pdf in os.listdir(PDF_FOLDER):
    if not pdf.lower().endswith(".pdf"):
        continue

    print(f"\nProcessing OCR: {pdf}")

    images = convert_from_path(os.path.join(PDF_FOLDER, pdf))
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img)

    text = normalize(text)

    print("---- OCR TEXT SAMPLE ----")
    print(text[:1000])
    print("-------------------------")

    row = {"Source File": pdf}

    row["Invoice Number"] = re.search(r"[A-Z]{2,5}-\d{4,}", text)
    row["Order Number"] = re.search(r"\d{3}-\d{7}-\d{7}", text)
    row["GST Number"] = re.search(r"\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z]\w", text)
    row["Invoice Value"] = re.search(r"\b\d+\.\d{2}\b", text)

    for key in row:
        if hasattr(row[key], "group"):
            row[key] = row[key].group()
        elif row[key] is None:
            row[key] = None

    rows.append(row)

df = pd.DataFrame(rows)

try:
    template = pd.read_excel(TEMPLATE_FILE)
    df = df.reindex(columns=template.columns)
except:
    pass

output = f"Final_Output_OCR_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
df.to_excel(output, index=False)

print(f"\n✅ Output written: {output}")
