import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    # Create the pdf
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Create the pdf name
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Create PDF header
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    # Create Date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Read in file and put values in respective cells
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Add headers to pdf file
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")

    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add Total Columns
    total_sum = df["total_price"].sum()

    pdf.set_font(family="Times", size=10, style="BIU")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=2)

    # Add Text stating the total price
    pdf.ln(2)
    pdf.set_font(family="Times", size=20, style="B")
    pdf.cell(w=30, h=20, txt="Summary Statement", ln=1)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=100, h=8, txt=f"The Total Price is {total_sum}", ln=1)

    pdf.ln(80)
    # Add company name
    pdf.image("Visionary.png", w=40)
    pdf.set_font(family="Times", size=23, style="BI")
    pdf.cell(w=100, h=8, txt="Visionary Technology inc.", ln =1)

    # Add company address
    pdf.set_font(family="Times", size=15, style="BI")
    pdf.cell(w=80, h=8, txt="007 N.E Dream St.", ln=1)
    pdf.cell(w=80, h=8, txt="Lorton, VA , 22079", ln=1)
    pdf.cell(w=80, h=8, txt="suite. 007", ln=1)

    pdf.output(f"pdf's/{filename}.pdf")
