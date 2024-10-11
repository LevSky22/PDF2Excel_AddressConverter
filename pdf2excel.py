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

def clean_text(text):
    # Remove everything from the first opening parenthesis onwards
    cleaned = re.sub(r'\s*\(.*$', '', text)
    # Remove any trailing whitespace or punctuation
    cleaned = re.sub(r'[\s\-,]+$', '', cleaned)
    return cleaned.strip()

def clean_and_separate_city_address(city, address):
    # Combine city and address
    combined = f"{city} {address}".strip()
    
    # Regular expression to find the first number in the string
    number_match = re.search(r'\d+', combined)
    
    if number_match:
        # Split at the first number
        split_index = number_match.start()
        city = combined[:split_index].strip()
        address = combined[split_index:].strip()
    else:
        # If no number found, keep original values
        city = city.strip()
        address = address.strip()
    
    # If either city or address is empty, use the original combined string
    if not city or not address:
        city = combined
        address = combined
    
    return clean_text(city), clean_text(address)

def process_pdfs(pdf_paths, merge=False):
    all_dfs = []
    for pdf_path in pdf_paths:
        df = extract_with_pdfplumber(pdf_path)
        
        # Apply the new function to clean and separate city and address
        cleaned_data = df.apply(lambda row: clean_and_separate_city_address(row['municipality_borough'], row['address']), axis=1)
        
        output_df = pd.DataFrame({
            'FNAM': 'À',
            'LNAM': "l'occupant",
            'ADD1': cleaned_data.apply(lambda x: x[1]),  # Always use the address part
            'CITY': cleaned_data.apply(lambda x: x[0]),
            'PROV': 'QC',
            'PC': df['postal_code']
        })
        all_dfs.append(output_df.sort_values('CITY'))
    
    if merge:
        merged_df = pd.concat(all_dfs, ignore_index=True)
        return merged_df.sort_values('CITY')
    else:
        return all_dfs

def save_to_excel(dfs, pdf_paths, merge=False):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if merge:
        output_filename = f'output_excel/merged_output_{current_time}.xlsx'
        dfs.to_excel(output_filename, index=False)
        auto_adjust_columns(output_filename)
        print(f"Merged Excel file '{output_filename}' has been created successfully.")
    else:
        for df, pdf_path in zip(dfs, pdf_paths):
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_{current_time}.xlsx'
            df.sort_values('CITY').to_excel(output_filename, index=False)  # Sort before saving
            auto_adjust_columns(output_filename)
            print(f"Excel file '{output_filename}' has been created successfully.")

def auto_adjust_columns(filename):
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
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    # Use Excel's built-in autofit
    for column in worksheet.columns:
        column_letter = get_column_letter(column[0].column)
        worksheet.column_dimensions[column_letter].auto_size = True

    workbook.save(filename)

if __name__ == "__main__":
    pdf_paths = get_pdf_paths()
    if not pdf_paths:
        print("No PDF files selected. Exiting.")
        exit(1)
    
    merge = False
    if len(pdf_paths) > 1:
        merge = input("Do you want to merge the PDFs into a single Excel file? (y/n): ").lower() == 'y'
    
    output_dfs = process_pdfs(pdf_paths, merge)
    save_to_excel(output_dfs, pdf_paths, merge)