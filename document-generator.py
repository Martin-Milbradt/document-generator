from pathlib import Path

import glob
import os
import pandas as pd
import shutil
from docx2pdf import convert
from docxtpl import DocxTemplate

# pip install -r requirements.txt
base_dir = Path(__file__).parent
# Check for existence of sheet.xlsx and create a copy from sheet.example.xlsx if not found
excel_path = os.path.join(base_dir, "sheet.xlsx")
if not os.path.exists(excel_path):
    print("'sheet.xlsx' does not exist. Creating a copy from 'sheet.example.xlsx'.")
    src_path = os.path.join(base_dir, "sheet.example.xlsx")
    shutil.copy2(src_path, excel_path)

# Check for existence of template.docx and create a copy from template.example.docx if not found
template_path = os.path.join(base_dir, "template.docx")
if not os.path.exists(template_path):
    print(
        "'template.docx' does not exist. Created a copy from 'template.example.docx'."
    )
    src_path = os.path.join(base_dir, "template.example.docx")
    shutil.copy2(src_path, template_path)

word_dir = base_dir / "word"
pdf_dir = base_dir / "pdf"

# Create and clean output folders for the documents
word_dir.mkdir(exist_ok=True)
files = glob.glob(os.path.join(word_dir, "*"))
for f in files:
    os.remove(f)

pdf_dir.mkdir(exist_ok=True)
files = glob.glob(os.path.join(pdf_dir, "*"))
for f in files:
    os.remove(f)

# Convert Excel sheet to pandas dataframe
try:
    df = pd.read_excel(excel_path, sheet_name="Sheet1")
except Exception as e:
    raise OSError(f"Please check that '{excel_path}' exists and is not open: {e}")

doc = DocxTemplate(template_path)

# Iterate over each row in df and render word document
try:
    for record in df.to_dict(orient="records"):
        doc.render(record)
        output_path = word_dir / f"{record['Name']}.docx"
        doc.save(output_path)
except Exception as e:
    raise OSError(f"Please check that '{template_path}' exists and is not open: {e}")

# Convert word document to pdf
convert(word_dir, pdf_dir)
