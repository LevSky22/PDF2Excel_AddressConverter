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

def extract_apartment(address):
    if not address:
        return ("", None)
    
    # Look for apartment indicator
    apt_index = address.lower().find('apt.')
    if apt_index == -1:
        return address, None
        
    # Split at the comma before apt.
    base_address = address[:apt_index].rstrip(' ,')
    
    # Find the postal code pattern after apt.
    postal_pattern = r'[A-Z][0-9][A-Z]\s*[0-9][A-Z][0-9]'
    postal_match = re.search(postal_pattern, address[apt_index:])
    
    if postal_match:
        # Extract everything from 'apt.' up to the postal code
        apartment = address[apt_index:apt_index + postal_match.start()].strip()
        return base_address, apartment
    else:
        # If no postal code found (shouldn't happen with our data), take rest of string
        apartment = address[apt_index:].strip()
        return base_address, apartment

def clean_text(text, extract_apt=False):
    if not text:
        return ("", None) if extract_apt else ""
    
    # First clean up any obvious typos where first letter is missing
    common_prefixes = {
        'ue ': 'Rue ',
        'v. ': 'Av. ',
        'h. ': 'Ch. ',
        'te ': 'Côte ',
        'l. ': 'Boul. '
    }
    
    for wrong, correct in common_prefixes.items():
        if text.startswith(wrong):
            text = correct + text[len(wrong):]
            break
    
    # Remove everything from the first opening parenthesis onwards
    cleaned = re.sub(r'\s*\(.*$', '', text)
    
    if extract_apt:
        # Extract apartment number and clean the address
        cleaned, apartment = extract_apartment(cleaned)
        # Clean up any remaining trailing punctuation
        # Modified to preserve single letters at start of street names
        cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned)
        cleaned = cleaned.strip()
        return cleaned, apartment
    else:
        # When not extracting apartments, preserve the original address format
        # Modified to preserve single letters at start of street names
        cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned)
        cleaned = cleaned.strip()
        return cleaned

def process_pdfs(pdf_paths, merge=False, column_names=None, merge_names=False, 
                merged_name="Full Name", default_values=None, file_format='xlsx',
                output_dir=None, custom_filename=None, merge_address=False,
                merged_address_name="Complete Address", address_separator=", ",
                province_default="QC", should_extract_apartment=False, apartment_column_name="Apartment",
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
        
        # Just clean up any extra whitespace
        df['address'] = df['address'].apply(lambda x: x.strip())
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
        
        # Handle name fields AFTER filtering the DataFrame
        if merge_address:
            merged_addresses = []
            apartments = []
            valid_indices = []
            
            for idx, row in df.iterrows():
                if should_extract_apartment:
                    logging.info(f"Processing address: {row['address']}")
                    clean_addr, apt = extract_apartment(row['address'])
                    logging.info(f"Extracted: Address='{clean_addr}', Apartment='{apt}'")
                    
                    # First check if we should filter out this address
                    if filter_apartments and apt is not None:
                        logging.info(f"Filtering out address with apartment: {row['address']} (apt: {apt})")
                        continue
                    
                    # If we get here, either filtering is off or this address has no apartment
                    if merge_address:
                        logging.info(f"Merging address components: {clean_addr}, {row['municipality_borough']}, {province_default}")
                    
                    address_parts = [
                        clean_addr,
                        row['municipality_borough'],
                        province_default,
                        row['postal_code']
                    ]
                    merged_address = address_separator.join(filter(None, address_parts))
                    
                    # Check for duplicates before adding
                    if merged_address not in merged_addresses:
                        merged_addresses.append(merged_address)
                        if include_apartment_column and not filter_apartments:
                            apartments.append(apt)
                        valid_indices.append(idx)
                    else:
                        logging.info(f"Skipping duplicate address: {merged_address}")
                
                else:
                    # Non-apartment extraction case
                    if filter_apartments:
                        _, apt = extract_apartment(row['address'])
                        if apt is not None:
                            logging.info(f"Filtering out address with apartment: {row['address']}")
                            continue
                    
                    clean_addr = clean_text(row['address'])
                    address_parts = [
                        clean_addr,
                        row['municipality_borough'],
                        province_default,
                        row['postal_code']
                    ]
                    merged_address = address_separator.join(filter(None, address_parts))
                    
                    # Check for duplicates before adding
                    if merged_address not in merged_addresses:
                        merged_addresses.append(merged_address)
                        valid_indices.append(idx)
                    else:
                        logging.info(f"Skipping duplicate address: {merged_address}")
            
            # Filter the DataFrame to only include valid rows
            if valid_indices:
                logging.info(f"Filtering DataFrame to {len(valid_indices)} valid entries")
                df = df.loc[valid_indices].copy()
                base_length = len(df)
                logging.info(f"Creating output data with base length: {base_length}")
                
                output_data = {}
                
                if filter_by_region and 'Branch ID' in df.columns:
                    logging.info("Adding Branch IDs to output")
                    output_data['Branch ID'] = df['Branch ID'].tolist()[:base_length]
                
                if merge_names:
                    logging.info(f"Adding merged names with default: {default_values.get(merged_name, 'À l\'occupant')}")
                    output_data[merged_name] = [default_values.get(merged_name, "À l'occupant")] * base_length
                else:
                    logging.info("Adding separate first and last names")
                    output_data[column_names['First Name']] = [default_values.get(column_names['First Name'], 'À')] * base_length
                    output_data[column_names['Last Name']] = [default_values.get(column_names['Last Name'], "l'occupant")] * base_length
                
                logging.info("Adding merged addresses")
                output_data[merged_address_name] = merged_addresses[:base_length]
                
                if include_apartment_column:
                    logging.info("Adding apartment numbers")
                    output_data[apartment_column_name] = apartments[:base_length]
                
                if include_phone:
                    logging.info(f"Adding phone numbers with default: {phone_default}")
                    output_data['Phone'] = [phone_default] * base_length
                
                if include_date:
                    logging.info(f"Adding dates with value: {date_value}")
                    output_data['Date'] = [date_value] * base_length

                # Verify array lengths before creating DataFrame
                lengths = {k: len(v) for k, v in output_data.items()}
                logging.info(f"Column lengths before DataFrame creation: {lengths}")
                if len(set(lengths.values())) > 1:
                    logging.error(f"Mismatched column lengths detected: {lengths}")
                    # Ensure all arrays match base_length
                    for key in output_data:
                        if len(output_data[key]) != base_length:
                            logging.warning(f"Adjusting length of {key} from {len(output_data[key])} to {base_length}")
                            if len(output_data[key]) > base_length:
                                output_data[key] = output_data[key][:base_length]
                            else:
                                output_data[key].extend([None] * (base_length - len(output_data[key])))

                # Create DataFrame after ensuring all arrays have same length
                logging.info("Creating final DataFrame")
                output_df = pd.DataFrame(output_data)
                logging.info(f"Final DataFrame shape: {output_df.shape}")
                all_dfs.append(output_df)
        
        else:
            if should_extract_apartment:
                cleaned_addresses = []
                apartments = []
                valid_indices = []
                
                for idx, addr in enumerate(df['address']):
                    clean_addr, apt = clean_text(addr, extract_apt=True)
                    
                    # First check if we should filter out this address
                    if filter_apartments and apt is not None:
                        logging.info(f"Filtering out address with apartment: {addr}")
                        continue
                    
                    # If we get here, either filtering is off or this address has no apartment
                    cleaned_addresses.append(clean_addr)
                    
                    # Only add apartment info if we're extracting but not filtering
                    if include_apartment_column and not filter_apartments:
                        apartments.append(apt)
                    valid_indices.append(idx)
            
                # Filter the DataFrame to only include valid rows
                if valid_indices:
                    df = df.iloc[valid_indices].copy()
                    # Ensure all arrays match the filtered DataFrame length
                    cleaned_addresses = cleaned_addresses[:len(df)]
                    if include_apartment_column:
                        apartments = apartments[:len(df)]
                
                output_data.update({
                    column_names['Address']: cleaned_addresses,
                    column_names['City']: df['municipality_borough'].tolist(),
                    column_names['Province']: [default_values.get(column_names['Province'], '')] * len(df),
                    column_names['Postal Code']: df['postal_code'].tolist()
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
        
        # Clean up address if city is included (do this before sorting)
        if not merge_address:
            address_col = column_names['Address']
            city_col = column_names['City']
            output_df[address_col] = output_df.apply(
                lambda row: row[address_col].replace(row[city_col], '', 1).strip() 
                if row[address_col].startswith(row[city_col]) else row[address_col], 
                axis=1
            )
        
        # Update the sorting logic
        if filter_by_region and 'Branch ID' in output_df.columns:
            sort_columns = ['Branch ID']
            if merge_address:
                if merged_address_name in output_df.columns:
                    sort_columns.append(merged_address_name)
            else:
                # Check if columns exist in the DataFrame
                city_col = column_names.get('City')
                addr_col = column_names.get('Address')
                
                if city_col and city_col in output_df.columns:
                    sort_columns.append(city_col)
                if addr_col and addr_col in output_df.columns:
                    sort_columns.append(addr_col)
            
            if len(sort_columns) > 0:
                output_df = output_df.sort_values(sort_columns)
                logging.info(f"Sorting by columns: {sort_columns}")
            else:
                logging.warning("No valid sort columns found")
        else:
            # Check which column to sort by
            sort_column = merged_address_name if merge_address else column_names.get('City')
            if sort_column and sort_column in output_df.columns:
                output_df = output_df.sort_values(by=sort_column)
                logging.info(f"Sorting by column: {sort_column}")
            else:
                logging.warning(f"Sort column '{sort_column}' not found in DataFrame. Skipping sort.")
        
        logging.info(f"Processed {len(output_df)} rows for {pdf_path}")
        all_dfs.append(output_df)
    
    return all_dfs

def save_to_excel(dfs, pdf_paths, merge=False):
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if merge:
        output_filename = f'output_excel/merged_output_{current_time}.xlsx'
        dfs.sort_values('City', inplace=True)
        dfs.to_excel(output_filename, index=False)
        auto_adjust_columns(output_filename)
        print(f"Merged Excel file '{output_filename}' has been created successfully.")
    else:
        for df, pdf_path in zip(dfs, pdf_paths):
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_{current_time}.xlsx'
            df.sort_values('City', inplace=True)
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
                        should_extract_apartment=False, apartment_column_name="Apartment",
                        filter_apartments=False, include_apartment_column=True,
                        include_phone=False, phone_default="",
                        include_date=False, date_value=None,
                        filter_by_region=False, region_branch_ids=None):
    logging.info(f"Converting PDFs: {pdf_files}")
    
    pdf_paths = [pdf_files] if isinstance(pdf_files, str) else pdf_files
    total_files = len(pdf_paths)
    
    # Create a set to track unique addresses across all files
    all_unique_addresses = set()
    all_data = []
    
    for i, pdf_path in enumerate(pdf_paths):
        dfs = process_pdfs(
            [pdf_path], 
            merge=False,
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
            should_extract_apartment=should_extract_apartment,
            apartment_column_name=apartment_column_name,
            filter_apartments=filter_apartments,
            include_apartment_column=include_apartment_column,
            include_phone=include_phone,
            phone_default=phone_default,
            include_date=include_date,
            date_value=date_value,
            filter_by_region=filter_by_region,
            region_branch_ids=region_branch_ids
        )
        
        if merge_files:
            # For each DataFrame in the results
            for df in dfs:
                if merged_address_name in df.columns:
                    # Get mask of unique addresses
                    unique_mask = ~df[merged_address_name].isin(all_unique_addresses)
                    # Update our set of unique addresses
                    all_unique_addresses.update(df[merged_address_name][unique_mask])
                    # Only keep rows with unique addresses
                    df_unique = df[unique_mask].copy()
                    if not df_unique.empty:
                        all_data.append(df_unique)
                else:
                    all_data.append(df)
        else:
            all_data.extend(dfs)
        
        progress = int((i + 1) / total_files * 90)
        yield progress
    
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if merge_files:
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            
            # Format date column if it exists
            if include_date and 'Date' in merged_df.columns:
                merged_df['Date'] = pd.to_datetime(merged_df['Date']).dt.strftime('%Y-%m-%d')
            
            # Remove apartment column if not included
            if should_extract_apartment and not include_apartment_column and apartment_column_name in merged_df.columns:
                merged_df = merged_df.drop(columns=[apartment_column_name])
            
            # Sort by Branch ID first if region filtering is enabled
            if filter_by_region and 'Branch ID' in merged_df.columns:
                sort_columns = ['Branch ID']
                if merge_address and merged_address_name in merged_df.columns:
                    sort_columns.append(merged_address_name)
                elif not merge_address:
                    city_col = column_names.get('City')
                    addr_col = column_names.get('Address')
                    if city_col and city_col in merged_df.columns:
                        sort_columns.append(city_col)
                    if addr_col and addr_col in merged_df.columns:
                        sort_columns.append(addr_col)
                
                if len(sort_columns) > 0:
                    merged_df = merged_df.sort_values(sort_columns)
                    logging.info(f"Sorting merged DataFrame by columns: {sort_columns}")
            else:
                # Original sorting logic
                sort_column = merged_address_name if merge_address else column_names.get('City')
                if sort_column and sort_column in merged_df.columns:
                    merged_df = merged_df.sort_values(by=sort_column)
                    logging.info(f"Sorting merged DataFrame by column: {sort_column}")
                else:
                    logging.warning(f"Sort column '{sort_column}' not found in merged DataFrame. Skipping sort.")
            
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
            
            # Sort by Branch ID first if region filtering is enabled
            if filter_by_region and 'Branch ID' in df.columns:
                if merge_address:
                    df = df.sort_values(['Branch ID', merged_address_name])
                else:
                    df = df.sort_values(['Branch ID', column_names['City'], column_names['Address']])
            else:
                sort_column = merged_address_name if merge_address else column_names['City']
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