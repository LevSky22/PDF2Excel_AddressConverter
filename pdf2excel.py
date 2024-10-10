import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog
import os
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def get_pdf_paths():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Select PDF File(s)", filetypes=[("PDF Files", "*.pdf")])
    return file_paths

def extract_with_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_data = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                all_data.extend(table[1:])  # Skip the header row
        return pd.DataFrame(all_data, columns=['centris_no', 'municipality_borough', 'address', 'postal_code'])

def clean_text_and_extract_address(text):
    # Split at the first occurrence of a pattern that looks like the start of an address
    match = re.search(r'\b(?:\d+[a-z]?|[a-z]\d+)\b', text, re.IGNORECASE)
    
    if match:
        split_index = match.start()
        municipality = text[:split_index].strip()
        address = text[split_index:].strip()
        
        # Clean up municipality
        municipality = re.sub(r'\s*\(.*$', '', municipality)  # Remove everything from '(' onwards
        municipality = re.sub(r'[\s,\-/]+$', '', municipality)  # Remove trailing spaces, commas, hyphens, slashes
        
        return municipality, address
    else:
        return text.strip(), ''

def process_pdf(pdf_path):
    df = extract_with_pdfplumber(pdf_path)
    
    # Apply the new function to split municipality and address
    df[['CITY', 'ADD1']] = df['municipality_borough'].apply(lambda x: pd.Series(clean_text_and_extract_address(x)))
    
    output_df = pd.DataFrame({
        'FNAM': 'À',
        'LNAM': "l'occupant",
        'ADD1': df['ADD1'],
        'CITY': df['CITY'],
        'PROV': 'QC',
        'PC': df['postal_code']
    })
    
    return output_df

def save_to_excel(df, pdf_paths):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if len(pdf_paths) == 1:
        output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_paths[0]))[0]}_simple_{current_time}.xlsx'
    else:
        output_filename = f'output_excel/merged_pdfs_simple_{current_time}.xlsx'
    
    # Save DataFrame to Excel
    df.to_excel(output_filename, index=False)
    
    # Load the workbook and select the active sheet
    wb = load_workbook(output_filename)
    ws = wb.active
    
    # Auto-adjust columns
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save the workbook
    wb.save(output_filename)
    print(f"Excel file '{output_filename}' has been created successfully with auto-adjusted columns.")

def merge_dataframes(dfs):
    return pd.concat(dfs, ignore_index=True)

if __name__ == "__main__":
    pdf_paths = get_pdf_paths()
    if not pdf_paths:
        print("No PDF files selected. Exiting.")
        exit(1)
    
    all_dfs = []
    for pdf_path in pdf_paths:
        output_df = process_pdf(pdf_path)
        all_dfs.append(output_df)
    
    if len(all_dfs) > 1:
        merge = input("Multiple PDFs selected. Do you want to merge them into a single Excel file? (y/n): ").lower()
        if merge == 'y':
            final_df = merge_dataframes(all_dfs)
            save_to_excel(final_df, pdf_paths)
        else:
            for i, df in enumerate(all_dfs):
                save_to_excel(df, [pdf_paths[i]])
    else:
        save_to_excel(all_dfs[0], pdf_paths)