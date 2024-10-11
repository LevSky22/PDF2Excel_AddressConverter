import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog
import os
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import logging

def setup_logging():
    logs_dir = 'logs'
    os.makedirs(logs_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    log_file = os.path.join(logs_dir, f'conversion_log_{timestamp}.txt')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return log_file

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
                for row in table[1:]:  # Skip the header row
                    # Join any split cells and clean up whitespace
                    cleaned_row = [' '.join(str(cell).split()) for cell in row if cell]
                    if len(cleaned_row) == 4:
                        all_data.append(cleaned_row)
                    else:
                        logging.warning(f"Skipping malformed row: {cleaned_row}")
        return pd.DataFrame(all_data, columns=['centris_no', 'municipality_borough', 'address', 'postal_code'])

def clean_text(text):
    original = text
    # Remove everything from the first opening parenthesis onwards
    cleaned = re.sub(r'\s*\(.*$', '', text)
    # Remove apartment numbers (various formats)
    cleaned = re.sub(r',?\s*(?:app?t?|unit|suite|#)\s*\d+[a-z]?$', '', cleaned, flags=re.IGNORECASE)
    # Remove any trailing whitespace or punctuation, but preserve trailing E. or O.
    cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned)
    cleaned = cleaned.strip()
    
    if cleaned != original:
        logging.info(f"Cleaned text: '{original}' -> '{cleaned}'")
    return cleaned

def process_pdfs(pdf_paths, merge=False):
    all_dfs = []
    for pdf_path in pdf_paths:
        logging.info(f"Processing PDF: {pdf_path}")
        df = extract_with_pdfplumber(pdf_path)
        logging.info(f"Extracted {len(df)} rows from {pdf_path}")
        
        # Data validation and cleaning
        df['municipality_borough'] = df['municipality_borough'].apply(lambda x: x.split('(')[0].strip())
        
        # Clean up addresses
        df['address'] = df['address'].apply(lambda x: re.sub(r'^[a-zA-Z]', '', x).strip())  # Remove any leading single letter
        df['address'] = df.apply(lambda row: row['address'] if row['address'].strip() else f"{row['municipality_borough']} {row['address']}", axis=1)
        
        output_df = pd.DataFrame({
            'FNAM': 'À',
            'LNAM': "l'occupant",
            'ADD1': df['address'].apply(clean_text),
            'CITY': df['municipality_borough'],  # Don't apply clean_text to city names
            'PROV': 'QC',
            'PC': df['postal_code']
        })
        
        # Additional validation
        output_df['ADD1'] = output_df.apply(lambda row: row['ADD1'].replace(row['CITY'], '', 1).strip() if row['ADD1'].startswith(row['CITY']) else row['ADD1'], axis=1)
        
        logging.info(f"Processed {len(output_df)} rows for {pdf_path}")
        all_dfs.append(output_df)
    
    if merge:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        logging.info(f"Merged {len(merged_df)} total rows from all PDFs")
        return merged_df.sort_values('CITY')
    else:
        return all_dfs

def save_to_excel(dfs, pdf_paths, merge=False):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if merge:
        output_filename = f'output_excel/merged_output_{current_time}.xlsx'
        dfs.sort_values('CITY', inplace=True)  # Sort by CITY column
        dfs.to_excel(output_filename, index=False)
        auto_adjust_columns(output_filename)
        print(f"Merged Excel file '{output_filename}' has been created successfully.")
    else:
        for df, pdf_path in zip(dfs, pdf_paths):
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_{current_time}.xlsx'
            df.sort_values('CITY', inplace=True)  # Sort by CITY column
            df.to_excel(output_filename, index=False)
            auto_adjust_columns(output_filename)
            print(f"Excel file '{output_filename}' has been created successfully.")

def auto_adjust_columns(filename):
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    workbook = load_workbook(filename)
    worksheet = workbook.active

    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2  # Multiply by 1.2 for better fit
        worksheet.column_dimensions[column_letter].width = adjusted_width

    workbook.save(filename)

if __name__ == "__main__":
    log_file = setup_logging()
    logging.info("Starting PDF to Excel conversion process")
    
    pdf_paths = get_pdf_paths()
    if not pdf_paths:
        logging.error("No PDF files selected. Exiting.")
        exit(1)
    
    logging.info(f"Selected PDFs: {', '.join(pdf_paths)}")
    
    merge = False
    if len(pdf_paths) > 1:
        merge = input("Do you want to merge the PDFs into a single Excel file? (y/n): ").lower() == 'y'
    logging.info(f"Merge option: {merge}")
    
    output_dfs = process_pdfs(pdf_paths, merge)
    save_to_excel(output_dfs, pdf_paths, merge)
    
    logging.info(f"Conversion complete. Log file: {log_file}")