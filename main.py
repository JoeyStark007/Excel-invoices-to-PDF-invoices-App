import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Read in file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

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
    pdf.cell(w=50, h=8, txt=f"Date: {date}")


    pdf.output(f"pdf's/{filename}.pdf")
