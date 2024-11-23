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
from quebec_regions_mapping import get_shore_region

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

def extract_apartment(text):
    # Expanded regex pattern to handle more apartment formats
    apt_patterns = [
        # Handle any combination after apt. (letters, numbers, or both)
        r',?\s+((?:app(?:artement)?\.?|apt\.?)\s*[a-z0-9]+(?:-[a-z0-9]+)?)\s*$',
        
        # Handle single letters after apt.
        r',?\s+((?:app(?:artement)?\.?|apt\.?)\s*[a-z])\s*$',
        
        # Other unit indicators
        r',?\s+((?:unit|suite|#)\s*[a-z0-9]+(?:-[a-z0-9]+)?)\s*$',
        
        # Special unit types
        r',?\s+((?:app(?:artement)?\.?|apt\.?|unit|suite|#)\s*(?:ph|th|penthouse|townhouse)\s*\d*)\s*$',
        
        # Basement/ground floor
        r',?\s+((?:sous-sol|basement|rez-de-chaussée|ss|rdj|rdc)(?:\s*\d+)?)\s*$',
        
        # French logement
        r',?\s+((?:logement|log\.?)\s*[a-z0-9]+(?:-[a-z0-9]+)?)\s*$'
    ]
    
    # Convert text to lowercase for case-insensitive matching
    text_lower = text.lower() if text else ""
    
    # First check for explicit apartment indicators with any characters
    for pattern in apt_patterns:
        match = re.search(pattern, text_lower, flags=re.IGNORECASE)
        if match:
            # Use the original text case for the apartment number
            apartment = text[match.start(1):match.end(1)].strip()
            # Remove the apartment info and any trailing commas/spaces
            clean_address = re.sub(pattern, '', text, flags=re.IGNORECASE).strip().rstrip(',')
            return clean_address, apartment
    
    # Check for apartment indicators in middle of address
    middle_apt_patterns = [
        # Handle any combination after apt.
        r'((?:app(?:artement)?\.?|apt\.?)\s*[a-z0-9]+(?:-[a-z0-9]+)?)',
        # Other unit indicators
        r'((?:unit|suite|#)\s*[a-z0-9]+(?:-[a-z0-9]+)?)',
        # French variations
        r'((?:logement|log\.?)\s*[a-z0-9]+(?:-[a-z0-9]+)?)'
    ]
    
    for pattern in middle_apt_patterns:
        match = re.search(pattern, text_lower)
        if match:
            # Return the full matched apartment text
            apartment = text[match.start(1):match.end(1)].strip()
            clean_address = text[:match.start(0)].strip()
            return clean_address, apartment
    
    # Check for just characters after comma as last resort
    comma_apt_pattern = r',\s*((?:no\.?\s*)?[a-z0-9]+(?:-[a-z0-9]+)?)\s*$'
    match = re.search(comma_apt_pattern, text_lower)
    if match:
        # Add 'apt.' prefix to bare numbers/letters
        apartment = 'apt. ' + text[match.start(1):match.end(1)].strip()
        clean_address = re.sub(comma_apt_pattern, '', text).strip()
        return clean_address, apartment
    
    return text, None

def clean_text(text, extract_apt=False):
    if not text:
        return ("", None) if extract_apt else ""
        
    # Remove everything from the first opening parenthesis onwards
    cleaned = re.sub(r'\s*\(.*$', '', text)
    
    if extract_apt:
        # Extract apartment number and clean the address
        cleaned, apartment = extract_apartment(cleaned)
        # Clean up any remaining trailing punctuation
        cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned)
        cleaned = cleaned.strip()
        return cleaned, apartment
    else:
        # When not extracting apartments, preserve the original address format
        # Just clean up trailing punctuation without removing apartment info
        cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned)
        cleaned = cleaned.strip()
        return cleaned

def process_pdfs(pdf_paths, merge=False, column_names=None, merge_names=False, 
                merged_name="Full Name", default_values=None, file_format='xlsx',
                output_dir=None, custom_filename=None, merge_address=False,
                merged_address_name="Complete Address", address_separator=", ",
                province_default="QC", extract_apartment=False, apartment_column_name="Apartment",
                filter_apartments=False, include_apartment_column=True,
                include_phone=False, phone_default="", 
                include_date=False, date_value=None,
                filter_by_region=False, region_branch_ids=None):
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
        
        # Initialize output_data dictionary
        output_data = {}
        
        # Add branch ID based on region if filtering is enabled
        if filter_by_region and region_branch_ids:
            filtered_df = pd.DataFrame()
            
            for idx, row in df.iterrows():
                region = get_shore_region(row['municipality_borough'])
                logging.info(f"City: {row['municipality_borough']} -> Region: {region}")
                branch_id = region_branch_ids.get(f'flyer_{region}', region_branch_ids.get('flyer_unknown', 'unknown'))
                logging.info(f"Branch ID resolved to: {branch_id}")
                if branch_id != 'unknown':
                    row_df = pd.DataFrame([row])
                    row_df['Branch ID'] = branch_id
                    filtered_df = pd.concat([filtered_df, row_df])
                else:
                    logging.warning(f"Unknown region for city: {row['municipality_borough']}")
            
            if len(filtered_df) > 0:
                df = filtered_df
                output_data['Branch ID'] = df['Branch ID'].tolist()
            else:
                logging.error("No valid regions found in the data")
                return [pd.DataFrame()]
        
        # Handle name fields
        if merge_names:
            output_data[merged_name] = [default_values.get(merged_name, "À l'occupant")] * len(df)
        else:
            for name_type in ['First Name', 'Last Name']:
                col_name = column_names[name_type]
                output_data[col_name] = [default_values.get(col_name, 'À' if name_type == 'First Name' else "l'occupant")] * len(df)
        
        # Handle address fields
        if merge_address:
            merged_addresses = []
            apartments = []
            has_apartment = []
            
            for _, row in df.iterrows():
                if extract_apartment:
                    clean_addr, apt = clean_text(row['address'], extract_apt=True)
                    has_apartment.append(apt is not None)
                    apartments.append(apt)
                else:
                    clean_addr = clean_text(row['address'])
                
                address_parts = [
                    clean_addr,
                    row['municipality_borough'],
                    province_default,
                    row['postal_code']
                ]
                merged_address = address_separator.join(filter(None, address_parts))
                merged_addresses.append(merged_address)
            
            # Create DataFrame with all data
            output_data[merged_address_name] = merged_addresses
            if extract_apartment and include_apartment_column:
                output_data[apartment_column_name] = apartments
        
        else:
            if extract_apartment:
                # Extract apartments while cleaning addresses
                cleaned_addresses = []
                apartments = []
                has_apartment = []
                
                for addr in df['address']:
                    clean_addr, apt = clean_text(addr, extract_apt=True)
                    cleaned_addresses.append(clean_addr)
                    apartments.append(apt)
                    has_apartment.append(apt is not None)
                
                output_data.update({
                    column_names['Address']: cleaned_addresses,
                    column_names['City']: df['municipality_borough'],
                    column_names['Province']: default_values.get(column_names['Province'], ''),
                    column_names['Postal Code']: df['postal_code']
                })
                
                if include_apartment_column:
                    output_data[apartment_column_name] = apartments
            else:
                # Just pass through addresses without any apartment processing
                output_data.update({
                    column_names['Address']: df['address'].apply(lambda x: clean_text(x, extract_apt=False)),
                    column_names['City']: df['municipality_borough'],
                    column_names['Province']: default_values.get(column_names['Province'], ''),
                    column_names['Postal Code']: df['postal_code']
                })
        
        # Add phone number if enabled
        if include_phone:
            output_data[column_names.get('Phone', 'Phone')] = [phone_default] * len(df)
        
        # Add date if enabled
        if include_date:
            output_data[column_names.get('Date', 'Date')] = [date_value] * len(df)
        
        # Create output DataFrame with all data
        output_df = pd.DataFrame(output_data)
        
        # Sort the DataFrame
        if merge_address:
            sort_column = merged_address_name
        else:
            sort_column = column_names['City']
            # Clean up address if city is included
            address_col = column_names['Address']
            city_col = column_names['City']
            output_df[address_col] = output_df.apply(
                lambda row: row[address_col].replace(row[city_col], '', 1).strip() 
                if row[address_col].startswith(row[city_col]) else row[address_col], 
                axis=1
            )
        
        output_df = output_df.sort_values(by=sort_column).copy()
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

def auto_adjust_columns(filename, df=None):
    """Auto-adjust column widths for Excel files or format CSV content"""
    if filename.endswith('.xlsx'):
        # Excel file adjustment
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
    
    elif filename.endswith('.csv') and df is not None:
        # For CSV, format the DataFrame before saving
        formatted_df = df.copy()
        
        # Calculate max width for each column
        max_lengths = {}
        for column in formatted_df.columns:
            # Get max length including header
            max_lengths[column] = max(
                formatted_df[column].astype(str).str.len().max(),
                len(str(column))
            )
        
        # Format each column with consistent width
        for column in formatted_df.columns:
            width = max_lengths[column]
            formatted_df[column] = formatted_df[column].astype(str).str.ljust(width)
        
        return formatted_df

def convert_pdf_to_excel(pdf_files, output_dir, merge_files=False, custom_filename=None, 
                        enable_logging=False, column_names=None, merge_names=False, 
                        merged_name="Full Name", default_values=None, file_format='xlsx',
                        merge_address=False, merged_address_name="Complete Address",
                        address_separator=", ", province_default="QC",
                        extract_apartment=False, apartment_column_name="Apartment",
                        filter_apartments=False, include_apartment_column=True,
                        include_phone=False, phone_default="",
                        include_date=False, date_value=None,
                        filter_by_region=False, region_branch_ids=None):
    if enable_logging:
        setup_logging()
    else:
        logging.disable(logging.CRITICAL)
    
    logging.info(f"Converting PDFs: {pdf_files}")
    
    pdf_paths = [pdf_files] if isinstance(pdf_files, str) else pdf_files
    total_files = len(pdf_paths)
    
    all_data = []
    for i, pdf_path in enumerate(pdf_paths):
        dfs = process_pdfs([pdf_path], merge=False, 
                         column_names=column_names, 
                         merge_names=merge_names, 
                         merged_name=merged_name,
                         default_values=default_values, 
                         file_format=file_format,
                         output_dir=output_dir,
                         custom_filename=custom_filename,
                         merge_address=merge_address,
                         merged_address_name=merged_address_name,
                         address_separator=address_separator,
                         province_default=province_default,
                         extract_apartment=extract_apartment,
                         apartment_column_name=apartment_column_name,
                         filter_apartments=filter_apartments,
                         include_apartment_column=include_apartment_column,
                         include_phone=include_phone,
                         phone_default=phone_default,
                         include_date=include_date,
                         date_value=date_value,
                         filter_by_region=filter_by_region,
                         region_branch_ids=region_branch_ids)
        all_data.extend(dfs)
        progress = int((i + 1) / total_files * 90)
        yield progress
    
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if merge_files:
        merged_df = pd.concat(all_data, ignore_index=True)
        
        # Format date column if it exists
        if include_date and 'Date' in merged_df.columns:
            merged_df['Date'] = pd.to_datetime(merged_df['Date']).dt.strftime('%Y-%m-%d')
        
        # Remove apartment column if not included
        if extract_apartment and not include_apartment_column and apartment_column_name in merged_df.columns:
            merged_df = merged_df.drop(columns=[apartment_column_name])
        
        # Choose the correct sorting column
        if merge_address:
            sort_column = merged_address_name
        else:
            sort_column = column_names['City']
            
        merged_df = merged_df.sort_values(by=sort_column)
        
        if custom_filename:
            output_filename = os.path.join(output_dir, f'{custom_filename}.{file_format}')
        else:
            output_filename = os.path.join(output_dir, f'merged_output_{current_time}.{file_format}')
        
        logging.info(f"Attempting to save merged file: {output_filename}")
        if file_format == 'xlsx':
            merged_df.to_excel(output_filename, index=False)
            auto_adjust_columns(output_filename)
        else:  # CSV format
            formatted_df = auto_adjust_columns(output_filename, merged_df)
            formatted_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
        
        yield output_filename
    else:
        last_file = None
        for i, df in enumerate(all_data):
            # Format date column if it exists
            if include_date and 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
            
            if custom_filename:
                output_filename = os.path.join(output_dir, f'{custom_filename}_{i+1}.{file_format}')
            else:
                base_name = os.path.splitext(os.path.basename(pdf_paths[i]))[0]
                output_filename = os.path.join(output_dir, f'{base_name}_{current_time}.{file_format}')
            
            # Choose the correct sorting column based on merge_address setting
            if merge_address:
                sort_column = merged_address_name
            else:
                sort_column = column_names['City']
                
            df = df.sort_values(by=sort_column)
            
            if file_format == 'xlsx':
                df.to_excel(output_filename, index=False)
                auto_adjust_columns(output_filename)
            else:  # CSV format
                formatted_df = auto_adjust_columns(output_filename, df)
                formatted_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            
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