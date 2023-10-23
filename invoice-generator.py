from pathlib import Path

import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate

# pip install -r requirements.txt
base_dir = Path(__file__).parent
word_template_path = base_dir / "template.docx"
excel_path = base_dir / "contracts.xlsx"
word_dir = base_dir / "invoices word"
pdf_dir = base_dir / "invoices pdf"

# Create output folder for the word documents
word_dir.mkdir(exist_ok=True)
pdf_dir.mkdir(exist_ok=True)

# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="List")

# Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = word_dir / f"{record['Name']}-contract.docx"
    doc.save(output_path)

# Convert word document to pdf
convert(word_dir, pdf_dir)
