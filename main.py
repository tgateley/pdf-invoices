import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()

    filename = Path(filepath).stem

    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8,txt=f"Invoice nr. {invoice_nr}", align="L", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=14, txt=f"Date: {date}", align="L", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family='Times', size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="Product ID", border=1)
    pdf.cell(w=50, h=8, txt="Product Name", border=1)
    pdf.cell(w=30, h=8, txt="Amount", border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total = df["total_price"].sum()
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=50, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=" ", border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=30, h=8, txt=f"The total amount due is {total} Euros", ln=1)
    pdf.cell(w=30, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
