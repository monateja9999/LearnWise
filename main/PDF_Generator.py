import os
import shutil
import pandas as pd
from docx2pdf import convert as docx_to_pdf
import win32com.client
from fpdf import FPDF
from PyPDF2 import PdfReader
import json

class PDFTable(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.epw = self.w - 2 * self.l_margin  # Effective page width

    def table_from_dataframe(self, df):
        self.set_font("Arial", "", 10)
        line_height = 8
        col_width = self.epw / max(len(df.columns), 1)  # Avoid division by zero
        for col_name in df.columns:
            self.cell(col_width, line_height, str(col_name), border=1)
        self.ln(line_height)
        for _, row in df.iterrows():
            for item in row:
                self.cell(col_width, line_height, str(item), border=1)
            self.ln(line_height)

def csv_to_pdf(csv_path, pdf_path):
    df = pd.read_csv(csv_path)
    pdf = PDFTable()
    pdf.add_page()
    pdf.table_from_dataframe(df)
    pdf.output(pdf_path)

def xlsx_to_pdf(xlsx_path, pdf_path):
    xls = pd.ExcelFile(xlsx_path)
    pdf = PDFTable()
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, sheet_name, ln=True, align="C")
        pdf.table_from_dataframe(df)
    pdf.output(pdf_path)

def pptx_to_pdf_windows(pptx_path, pdf_path):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = True
    deck = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
    deck.SaveAs(pdf_path, 32)
    deck.Close()
    powerpoint.Quit()

def convert_course_files_to_pdf(folder_path, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        name, ext = os.path.splitext(filename)
        ext = ext.lower()
        pdf_output = os.path.join(output_folder, f"{name}.pdf")
        try:
            if ext == ".docx":
                docx_to_pdf(filepath, pdf_output)
            elif ext == ".csv":
                csv_to_pdf(filepath, pdf_output)
            elif ext == ".xlsx":
                xlsx_to_pdf(filepath, pdf_output)
            elif ext == ".pptx":
                pptx_to_pdf_windows(filepath, pdf_output)
            elif ext == ".pdf":
                shutil.copy(filepath, pdf_output)
            else:
                print(f"Skipped unsupported file: {filename}")
        except Exception as e:
            print(f"Error converting {filename}: {e}")

def pdf_to_json(pdf_path):
    result = []
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        for i, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            result.append({"page": i, "content": text.strip()})
    return result

def convert_pdfs_to_single_json(pdf_folder, json_output_file):
    combined_json = {}
    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            try:
                combined_json[filename] = pdf_to_json(pdf_path)
                print(f"Processed {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {e}")
    with open(json_output_file, "w", encoding="utf-8") as jf:
        json.dump(combined_json, jf, indent=2, ensure_ascii=False)
    print(f"All PDFs combined into JSON file: {json_output_file}")

# Usage
src_folder = r"C:\Users\monat\Desktop\LearnWise\\Course1"
pdf_folder = r"C:\Users\monat\Desktop\LearnWise\\Course1\PDFs"
json_output_file = r"C:\Users\monat\Desktop\LearnWise\\Course1\\combined_pdfs.json"

# Step 1: Convert source files to PDFs
convert_course_files_to_pdf(src_folder, pdf_folder)

# Step 2: Convert all PDFs into a single JSON file
convert_pdfs_to_single_json(pdf_folder, json_output_file)
