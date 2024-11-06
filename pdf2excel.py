import pdfplumber
import pandas as pd
from tkinter import Tk, filedialog
import os
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import logging
import tabula
import time

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

def process_pdfs(pdf_paths, merge=False, column_names=None, merge_names=False, 
                merged_name="Full Name", default_values=None, file_format='xlsx',
                output_dir=None, custom_filename=None):
    if column_names is None:
        column_names = {
            'First Name': 'First Name',
            'Last Name': 'Last Name',
            'Address': 'Address',
            'City': 'City',
            'Province': 'Province',
            'Postal Code': 'Postal Code'
        }
    
    if default_values is None:
        default_values = {}

    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    all_dfs = []
    for pdf_path in pdf_paths:
        logging.info(f"Processing PDF: {pdf_path}")
        df = extract_with_pdfplumber(pdf_path)
        logging.info(f"Extracted {len(df)} rows from {pdf_path}")
        
        # Data validation and cleaning
        df['municipality_borough'] = df['municipality_borough'].apply(lambda x: x.split('(')[0].strip())
        
        # Clean up addresses
        df['address'] = df['address'].apply(lambda x: re.sub(r'^[a-zA-Z]', '', x).strip())
        df['address'] = df.apply(lambda row: row['address'] if row['address'].strip() else f"{row['municipality_borough']} {row['address']}", axis=1)
        
        # Create the output DataFrame with the appropriate columns
        output_data = {}
        if merge_names:
            col_name = merged_name
            output_data[col_name] = default_values.get(col_name, "À l'occupant")
        else:
            for name_type in ['First Name', 'Last Name']:
                col_name = column_names[name_type]
                output_data[col_name] = default_values.get(col_name, 'À' if name_type == 'First Name' else "l'occupant")
        
        # Add other columns with their default values
        output_data.update({
            column_names['Address']: df['address'].apply(clean_text),
            column_names['City']: df['municipality_borough'],
            column_names['Province']: default_values.get(column_names['Province'], ''),
            column_names['Postal Code']: df['postal_code']
        })
        
        output_df = pd.DataFrame(output_data)
        
        # Update the column name in the validation code
        address_col = column_names['Address']
        city_col = column_names['City']
        output_df[address_col] = output_df.apply(
            lambda row: row[address_col].replace(row[city_col], '', 1).strip() 
            if row[address_col].startswith(row[city_col]) else row[address_col], 
            axis=1
        )
        
        # Sort the DataFrame by city
        output_df = output_df.sort_values(by=city_col)
        
        logging.info(f"Processed {len(output_df)} rows for {pdf_path}")
        all_dfs.append(output_df)
    
    return all_dfs

def save_to_excel(dfs, pdf_paths, merge=False):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if merge:
        output_filename = f'output_excel/merged_output_{current_time}.xlsx'
        dfs.sort_values('City', inplace=True)  # Sort by City column
        dfs.to_excel(output_filename, index=False)
        auto_adjust_columns(output_filename)
        print(f"Merged Excel file '{output_filename}' has been created successfully.")
    else:
        for df, pdf_path in zip(dfs, pdf_paths):
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_{current_time}.xlsx'
            df.sort_values('City', inplace=True)  # Sort by City column
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

def convert_pdf_to_excel(pdf_files, output_dir, merge_files=False, custom_filename=None, 
                        enable_logging=False, column_names=None, merge_names=False, 
                        merged_name="Full Name", default_values=None, file_format='xlsx'):
    if enable_logging:
        setup_logging()
    else:
        logging.disable(logging.CRITICAL)
    
    logging.info(f"Converting PDFs: {pdf_files}")
    
    pdf_paths = [pdf_files] if isinstance(pdf_files, str) else pdf_files
    total_files = len(pdf_paths)
    
    all_data = []
    for i, pdf_path in enumerate(pdf_paths):
        # Convert generator to list and get first item
        df = list(process_pdfs([pdf_path], merge=False, 
                         column_names=column_names, 
                         merge_names=merge_names, 
                         merged_name=merged_name,
                         default_values=default_values, 
                         file_format=file_format,
                         output_dir=output_dir,
                         custom_filename=custom_filename))[0]
        all_data.append(df)
        progress = int((i + 1) / total_files * 90)
        yield progress
    
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if merge_files:
        merged_df = pd.concat(all_data, ignore_index=True)
        city_column = column_names['City']
        merged_df = merged_df.sort_values(city_column)
        if custom_filename:
            output_filename = os.path.join(output_dir, f'{custom_filename}.{file_format}')
        else:
            output_filename = os.path.join(output_dir, f'merged_output_{current_time}.{file_format}')
        
        logging.info(f"Attempting to save merged file: {output_filename}")
        if file_format == 'xlsx':
            merged_df.to_excel(output_filename, index=False)
            auto_adjust_columns(output_filename)
        else:  # CSV format
            merged_df.to_csv(output_filename, index=False, encoding='utf-8-sig')  # Added encoding for proper UTF-8 handling
        
        yield output_filename
    else:
        last_file = None
        for i, df in enumerate(all_data):
            city_column = column_names['City']
            df = df.sort_values(city_column)
            if custom_filename:
                output_filename = os.path.join(output_dir, f'{custom_filename}_{i+1}.{file_format}')
            else:
                base_name = os.path.splitext(os.path.basename(pdf_paths[i]))[0]
                output_filename = os.path.join(output_dir, f'{base_name}_{current_time}.{file_format}')
            
            if file_format == 'xlsx':
                df.to_excel(output_filename, index=False)
                auto_adjust_columns(output_filename)
            else:  # CSV format
                df.to_csv(output_filename, index=False, encoding='utf-8-sig')  # Added encoding for proper UTF-8 handling
            
            last_file = output_filename
            logging.info(f"Created file: {output_filename}")
        yield last_file
    
    yield 100  # Final progress update
    logging.info("Conversion complete")

    if not enable_logging:
        logging.disable(logging.NOTSET)

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