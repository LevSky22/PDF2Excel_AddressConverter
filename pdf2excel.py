import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog
import os
from datetime import datetime
import re

def get_pdf_path():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    return file_path

def extract_with_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_data = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_data.extend(table[1:])  # Skip the header row
        return pd.DataFrame(all_data, columns=['centris_no', 'municipality_borough', 'address', 'postal_code'])

def clean_city(city):
    # Remove everything from the first opening parenthesis onwards
    cleaned = re.sub(r'\s*\(.*$', '', city)
    # Remove any trailing whitespace or punctuation
    cleaned = re.sub(r'[\s\-,]+$', '', cleaned)
    return cleaned.strip()

def process_pdf(pdf_path):
    df = extract_with_pdfplumber(pdf_path)
    
    output_df = pd.DataFrame({
        'FNAM': 'À',
        'LNAM': "l'occupant",
        'ADD1': df['address'],
        'CITY': df['municipality_borough'].apply(clean_city),
        'PROV': 'QC',
        'PC': df['postal_code']
    })
    
    return output_df

def save_to_excel(df, pdf_path):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_simple_{current_time}.xlsx'
    df.to_excel(output_filename, index=False)
    print(f"Excel file '{output_filename}' has been created successfully.")

if __name__ == "__main__":
    pdf_path = get_pdf_path()
    if not pdf_path:
        print("No PDF file selected. Exiting.")
        exit(1)
    
    output_df = process_pdf(pdf_path)
    save_to_excel(output_df, pdf_path)