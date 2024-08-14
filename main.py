import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8,txt=f"Invoice nr. {invoice_nr}", align="L", ln=1)
    pdf.cell(w=0, h=14, txt=f"Date: {date}",align="L", ln=1)

    for index, row in df.i
    pdf.output(f"PDFs/{filename}.pdf")
