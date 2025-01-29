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
from quebec_regions_mapping import get_shore_region, get_custom_sector
from unidecode import unidecode

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
    """Extracts rows from a PDF with columns [centris_no, municipality_borough, address, postal_code]."""
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
    """Extract apartment substring (e.g. 'Apt. 101') from an address. Returns (address_without_apt, apartment_text)."""
    if not address:
        return ("", None)
    
    apt_index = address.lower().find('apt.')
    if apt_index == -1:
        return address, None
        
    base_address = address[:apt_index].rstrip(' ,')
    
    # Find a postal code pattern after "apt." to see if we can isolate the apt portion
    postal_pattern = r'[A-Z][0-9][A-Z]\s*[0-9][A-Z][0-9]'
    postal_match = re.search(postal_pattern, address[apt_index:])
    
    if postal_match:
        # Extract everything from 'apt.' up to the postal code
        apartment = address[apt_index:apt_index + postal_match.start()].strip()
        return base_address, apartment
    else:
        # If no postal code found, take the rest
        apartment = address[apt_index:].strip()
        return base_address, apartment

def clean_text(text, extract_apt=False, remove_accents=False):
    """Clean text, optionally remove accents, optionally handle apt extraction (but typically we do that separately)."""
    if not text:
        return ("", None) if extract_apt else ""
    
    # Fix a few common prefix typos
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
        # Extract apartment number and clean address
        cleaned, apartment = extract_apartment(cleaned)
        cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned).strip()
        if remove_accents:
            cleaned = unidecode(cleaned)
            if apartment:
                apartment = unidecode(apartment)
        return cleaned, apartment
    
    # Not extracting apt
    cleaned = re.sub(r'[\s\-,]+(?<!E)(?<!O)\.?$', '', cleaned).strip()
    if remove_accents:
        cleaned = unidecode(cleaned)
    return cleaned

def add_name_columns_to_df(
    df,
    merge_names,
    merged_name,
    column_names,
    default_values,
    remove_accents
):
    """
    Adds name columns (merged or separate) to the DataFrame `df`.
    
    - If `merge_names` is True:
       * Creates one column (e.g., 'Full Name') and places it first.
       * If DF doesn't have 'First Name'/'Last Name', uses default occupant name.
    - Else:
       * Creates separate 'First Name' and 'Last Name' columns (or uses defaults),
         and places them first in order (First Name, Last Name).
    - Also removes accents if `remove_accents` is True.
    
    Returns a new DataFrame that includes existing df columns plus the name columns.
    """
    if df is None or df.empty:
        return df
    
    df = df.copy()  # work on a copy
    base_length = len(df)
    
    if merge_names:
        # Merge both into a single column
        if 'First Name' in df.columns and 'Last Name' in df.columns:
            merged_col = (df['First Name'].fillna('') + ' ' + df['Last Name'].fillna('')).str.strip()
            df[merged_name] = merged_col
        else:
            # If we don't have those columns, use default occupant
            default_full_name = default_values.get(merged_name, "À l'occupant")
            df[merged_name] = [default_full_name]*base_length
        
        # Remove accents from merged_name column if needed
        if remove_accents:
            df[merged_name] = df[merged_name].astype(str).apply(lambda x: unidecode(x))
        
        # (NEW) Reorder so that merged_name is the first column
        cols = df.columns.tolist()
        if merged_name in cols:
            cols.remove(merged_name)
            cols.insert(0, merged_name)
        df = df[cols]
    
    else:
        # Separate columns for First & Last
        fn_key = 'First Name'
        ln_key = 'Last Name'
        # The user might have given custom labels in column_names
        col_first = column_names.get(fn_key, fn_key)
        col_last = column_names.get(ln_key, ln_key)

        # If "First Name" isn't physically in df, create it with a default
        if fn_key not in df.columns:
            default_fn = default_values.get(col_first, "À l'occupant")
            df[col_first] = [default_fn]*base_length
        else:
            # rename in case user changed "First Name" => "Prénom"
            df.rename(columns={fn_key: col_first}, inplace=True)
            default_fn = default_values.get(col_first, "À l'occupant")
            df[col_first] = df[col_first].fillna(default_fn)
        
        # If "Last Name" isn't physically in df, create it with a default
        if ln_key not in df.columns:
            default_ln = default_values.get(col_last, "")
            df[col_last] = [default_ln]*base_length
        else:
            # rename in case user changed "Last Name" => "Nom"
            df.rename(columns={ln_key: col_last}, inplace=True)
            default_ln = default_values.get(col_last, "")
            df[col_last] = df[col_last].fillna(default_ln)

        # Remove accents if needed
        if remove_accents:
            df[col_first] = df[col_first].astype(str).apply(lambda x: unidecode(x))
            df[col_last] = df[col_last].astype(str).apply(lambda x: unidecode(x))
        
        # (NEW) Reorder so that First Name, Last Name are the first columns
        cols = df.columns.tolist()
        # move col_first to front if it exists
        if col_first in cols:
            cols.remove(col_first)
            cols.insert(0, col_first)
        # move col_last to second
        if col_last in cols:
            cols.remove(col_last)
            cols.insert(1, col_last)
        df = df[cols]
    
    return df

def process_pdfs(
    pdf_paths,
    merge=False,
    column_names=None,
    merge_names=False, 
    merged_name="Full Name",
    default_values=None,
    file_format='xlsx',
    output_dir=None,
    custom_filename=None,
    merge_address=False,
    merged_address_name="Complete Address",
    address_separator=", ",
    province_default="QC",
    should_extract_apartment=False, 
    apartment_column_name="Apartment",
    filter_apartments=False, 
    include_apartment_column=True, 
    include_phone=False, 
    phone_default="", 
    include_date=False, 
    date_value=None, 
    filter_by_region=False, 
    region_branch_ids=None, 
    use_custom_sectors=False, 
    remove_accents=False
):
    """
    Main logic that processes PDF(s) and returns a list of DataFrames.
    Each DataFrame corresponds to one PDF, unless `merge` is True (then it might behave differently).
    """
    logging.info(f"process_pdfs remove_accents setting: {remove_accents}")

    def clean_accents(text):
        if isinstance(text, str) and remove_accents:
            cleaned = unidecode(text)
            if cleaned != text:
                logging.info(f"Cleaned text: '{text}' -> '{cleaned}'")
            return cleaned
        return text

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
        
        # Clean up the city/municipality
        df['municipality_borough'] = df['municipality_borough'].apply(lambda x: x.split('(')[0].strip() if x else x)
        
        # Trim addresses
        df['address'] = df['address'].apply(lambda x: x.strip() if x else x)
        df['address'] = df.apply(
            lambda row: row['address'] if row['address'] and row['address'].strip() else f"{row['municipality_borough']} {row['address']}",
            axis=1
        )
        
        logging.info("PDF Processing Settings:")
        logging.info(f"use_custom_sectors = {use_custom_sectors}")
        logging.info(f"filter_by_region = {filter_by_region}")
        logging.info(f"region_branch_ids = {region_branch_ids}")

        # Possibly filter by region
        if filter_by_region or use_custom_sectors:
            filtered_df = pd.DataFrame(columns=df.columns)
            
            for idx, row in df.iterrows():
                city = row['municipality_borough']
                
                if use_custom_sectors:
                    # custom sector filtering
                    sector = get_custom_sector(city, row['postal_code'])
                    if sector and sector in region_branch_ids:
                        row_df = pd.DataFrame([row])
                        row_df['Branch ID'] = region_branch_ids[sector]
                        filtered_df = pd.concat([filtered_df, row_df], ignore_index=True)
                        logging.info(f"Added to custom sector {sector}: {city} ({row['postal_code']})")
                    else:
                        logging.info(f"Skipping city not in selected sectors: {city} ({row['postal_code']})")
                else:
                    # standard region filtering
                    region = get_shore_region(city)
                    branch_id = region_branch_ids.get(f'flyer_{region}', region_branch_ids.get('flyer_unknown', 'unknown'))
                    if branch_id != 'unknown':
                        row_df = pd.DataFrame([row])
                        row_df['Branch ID'] = branch_id
                        filtered_df = pd.concat([filtered_df, row_df], ignore_index=True)
            
            if len(filtered_df) > 0:
                df = filtered_df.copy()
            else:
                logging.error("No valid cities found after filtering.")
                # Return an empty DF so we don't break downstream steps
                all_dfs.append(pd.DataFrame())
                continue  # move on to next PDF

        output_df_final = None  # We'll store the final DF for this PDF

        # -----------------------------------------------------------------
        # (A) MERGE_ADDRESS = True
        # -----------------------------------------------------------------
        if merge_address:
            merged_addresses = []
            apartments = []
            branch_ids = []
            valid_indices = []

            for idx, row in df.iterrows():
                if should_extract_apartment:
                    clean_addr, apt = extract_apartment(row['address'])
                    
                    if filter_apartments and apt is not None:
                        # skipping addresses with apartments
                        logging.info(f"Filtering out address with apartment: {row['address']} (apt={apt})")
                        continue
                    
                    address_parts = [
                        clean_addr.strip() if clean_addr else "",
                        row['municipality_borough'].strip() if row['municipality_borough'] else "",
                        province_default,
                        row['postal_code']
                    ]
                    merged_address = address_separator.join(filter(None, address_parts))
                    
                    if merged_address not in merged_addresses:
                        merged_addresses.append(merged_address)
                        if 'Branch ID' in row:
                            branch_ids.append(row['Branch ID'])
                        if include_apartment_column and not filter_apartments:
                            apartments.append(apt)
                        valid_indices.append(idx)
                    else:
                        logging.info(f"Skipping duplicate merged address: {merged_address}")
                
                else:
                    # Non-apartment extraction
                    if filter_apartments:
                        # Even if we're not extracting apt, check for "apt."
                        _, apt = extract_apartment(row['address'])
                        if apt is not None:
                            logging.info(f"Filtering out address with apartment: {row['address']}")
                            continue
                    
                    # Just clean the address for final
                    plain_addr = clean_text(row['address'], extract_apt=False, remove_accents=remove_accents)
                    address_parts = [
                        plain_addr,
                        row['municipality_borough'],
                        province_default,
                        row['postal_code']
                    ]
                    merged_address = address_separator.join(filter(None, address_parts)).strip()
                    
                    if merged_address not in merged_addresses:
                        merged_addresses.append(merged_address)
                        if 'Branch ID' in row:
                            branch_ids.append(row['Branch ID'])
                        valid_indices.append(idx)
                    else:
                        logging.info(f"Skipping duplicate merged address: {merged_address}")

            if valid_indices:
                df_filtered = df.loc[valid_indices].copy()
                base_length = len(df_filtered)
                
                # Build the partial output_data
                output_data = {}

                # Insert region/branch if it existed
                if 'Branch ID' in df_filtered.columns:
                    output_data['Branch ID'] = df_filtered['Branch ID'].tolist()
                elif branch_ids:
                    output_data['Branch ID'] = branch_ids[:base_length]
                
                # Insert the merged address
                cleaned_merged_addresses = []
                for addr in merged_addresses[:base_length]:
                    if remove_accents:
                        cleaned_merged_addresses.append(unidecode(addr))
                    else:
                        cleaned_merged_addresses.append(addr)
                output_data[merged_address_name] = cleaned_merged_addresses

                partial_df = pd.DataFrame(output_data)
                output_df_final = partial_df
            else:
                # No valid addresses after filtering
                output_df_final = pd.DataFrame()
        
        # -----------------------------------------------------------------
        # (B) MERGE_ADDRESS = False
        # -----------------------------------------------------------------
        else:
            valid_indices = []
            cleaned_addresses = []
            apartments = []

            for idx, row in df.iterrows():
                if should_extract_apartment:
                    clean_addr, apt = clean_text(row['address'], extract_apt=True, remove_accents=remove_accents)
                    
                    if filter_apartments and apt is not None:
                        logging.info(f"Filtering out address with apartment: {row['address']}")
                        continue
                    
                    cleaned_addresses.append(clean_addr)
                    if include_apartment_column and not filter_apartments:
                        apartments.append(apt)
                    valid_indices.append(idx)
                
                else:
                    if filter_apartments:
                        # even if not extracting apt explicitly, check it
                        _, apt = extract_apartment(row['address'])
                        if apt is not None:
                            logging.info(f"Filtering out address with apartment: {row['address']}")
                            continue
                    # normal cleaning
                    plain_addr = clean_text(row['address'], extract_apt=False, remove_accents=remove_accents)
                    cleaned_addresses.append(plain_addr)
                    valid_indices.append(idx)
            
            df_filtered = df.iloc[valid_indices].copy()
            output_data = {}
            output_data[column_names['Address']] = cleaned_addresses
            output_data[column_names['City']] = df_filtered['municipality_borough'].tolist()
            output_data[column_names['Province']] = [
                default_values.get(column_names['Province'], province_default)
            ] * len(df_filtered)
            output_data[column_names['Postal Code']] = df_filtered['postal_code'].tolist()
            
            if 'Branch ID' in df_filtered.columns:
                output_data['Branch ID'] = df_filtered['Branch ID'].tolist()

            if include_apartment_column and not filter_apartments and should_extract_apartment:
                output_data[apartment_column_name] = apartments
            
            partial_df = pd.DataFrame(output_data)

            # Remove city duplication from address if it starts with city
            address_col = column_names['Address']
            city_col = column_names['City']
            if address_col in partial_df.columns and city_col in partial_df.columns:
                partial_df[address_col] = partial_df.apply(
                    lambda row: row[address_col].replace(row[city_col], '', 1).strip()
                    if row[address_col].startswith(row[city_col]) else row[address_col],
                    axis=1
                )
            
            output_df_final = partial_df

        # -----------------------------------------------------------------
        # COMMON STEP: Add NAME columns (merged or not) at front
        # -----------------------------------------------------------------
        if output_df_final is not None and not output_df_final.empty:
            output_df_final = add_name_columns_to_df(
                df=output_df_final,
                merge_names=merge_names,
                merged_name=merged_name,
                column_names=column_names,
                default_values=default_values,
                remove_accents=remove_accents
            )

            # Handle phone
            if include_phone:
                phone_col = column_names.get('Phone', 'Phone')
                output_df_final[phone_col] = [phone_default]*len(output_df_final)

            # Handle date
            if include_date:
                date_col = column_names.get('Date', 'Date')
                output_df_final[date_col] = [date_value]*len(output_df_final)

            # Final sorting
            if filter_by_region and 'Branch ID' in output_df_final.columns:
                sort_columns = ['Branch ID']
                if merge_address and merged_address_name in output_df_final.columns:
                    sort_columns.append(merged_address_name)
                else:
                    if column_names['City'] in output_df_final.columns:
                        sort_columns.append(column_names['City'])
                    if column_names['Address'] in output_df_final.columns:
                        sort_columns.append(column_names['Address'])
                output_df_final = output_df_final.sort_values(sort_columns)
            else:
                sort_column = merged_address_name if merge_address else column_names.get('City')
                if sort_column and sort_column in output_df_final.columns:
                    output_df_final = output_df_final.sort_values(by=sort_column)
        else:
            output_df_final = pd.DataFrame()

        logging.info(f"Processed {len(output_df_final)} rows for {pdf_path}")
        all_dfs.append(output_df_final)

    return all_dfs

def save_to_excel(dfs, pdf_paths, merge=False):
    """Provided utility to save output as Excel; not heavily changed."""
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    
    if merge:
        # Merged file logic
        output_filename = f'output_excel/merged_output_{current_time}.xlsx'
        merged_df = pd.concat(dfs, ignore_index=True)
        if 'City' in merged_df.columns:
            merged_df.sort_values('City', inplace=True)
        merged_df.to_excel(output_filename, index=False)
        auto_adjust_columns(output_filename)
        print(f"Merged Excel file '{output_filename}' has been created successfully.")
    else:
        # One file per DF
        for df, pdf_path in zip(dfs, pdf_paths):
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_path))[0]}_{current_time}.xlsx'
            if 'City' in df.columns:
                df.sort_values('City', inplace=True)
            df.to_excel(output_filename, index=False)
            auto_adjust_columns(output_filename)
            print(f"Excel file '{output_filename}' has been created successfully.")

def auto_adjust_columns(filename, df=None):
    """Auto-adjust column widths for Excel or format CSV content if needed."""
    if filename.endswith('.xlsx'):
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
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(filename)
    
    elif filename.endswith('.csv') and df is not None:
        # For CSV, we can left-pad or just return the same df
        formatted_df = df.copy()
        max_lengths = {}
        for column in formatted_df.columns:
            max_lengths[column] = max(
                formatted_df[column].astype(str).str.len().max(),
                len(str(column))
            )
        for column in formatted_df.columns:
            width = max_lengths[column]
            formatted_df[column] = formatted_df[column].astype(str).str.ljust(width)
        return formatted_df

def convert_pdf_to_excel(
    pdf_files,
    output_dir,
    merge_files=False,
    custom_filename=None, 
    enable_logging=False,
    column_names=None,
    merge_names=False, 
    merged_name="Full Name",
    default_values=None,
    file_format='xlsx',
    merge_address=False,
    merged_address_name="Complete Address",
    address_separator=", ",
    province_default="QC",
    should_extract_apartment=False,
    apartment_column_name="Apartment",
    filter_apartments=False,
    include_apartment_column=True,
    include_phone=False,
    phone_default="",
    include_date=False,
    date_value=None,
    filter_by_region=False,
    region_branch_ids=None,
    use_custom_sectors=False,
    remove_accents=False
):
    """
    High-level function that calls process_pdfs() and then writes outputs.
    Yields progress or a final filename.
    """
    logging.info(f"Starting conversion with remove_accents={remove_accents}")
    
    pdf_paths = [pdf_files] if isinstance(pdf_files, str) else pdf_files
    total_files = len(pdf_paths)
    
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
            region_branch_ids=region_branch_ids,
            use_custom_sectors=use_custom_sectors,
            remove_accents=remove_accents
        )
        
        # If merging all files into one, gather them
        if merge_files:
            for df in dfs:
                if merged_address_name in df.columns:
                    unique_mask = ~df[merged_address_name].isin(all_unique_addresses)
                    all_unique_addresses.update(df[merged_address_name][unique_mask])
                    df_unique = df[unique_mask].copy()
                    if not df_unique.empty:
                        all_data.append(df_unique)
                else:
                    # no merged_address_name column => just add
                    all_data.append(df)
        else:
            # If not merging, just store them individually
            all_data.extend(dfs)
        
        progress = int((i + 1) / total_files * 90)
        yield progress

    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    if merge_files:
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            if include_date and 'Date' in merged_df.columns:
                merged_df['Date'] = pd.to_datetime(merged_df['Date']).dt.strftime('%Y-%m-%d')
            
            if should_extract_apartment and not include_apartment_column and apartment_column_name in merged_df.columns:
                merged_df.drop(columns=[apartment_column_name], inplace=True, errors='ignore')

            # Sort if filter_by_region
            if filter_by_region and 'Branch ID' in merged_df.columns:
                sort_cols = ['Branch ID']
                if merge_address and merged_address_name in merged_df.columns:
                    sort_cols.append(merged_address_name)
                else:
                    city_col = column_names.get('City')
                    addr_col = column_names.get('Address')
                    if city_col and city_col in merged_df.columns:
                        sort_cols.append(city_col)
                    if addr_col and addr_col in merged_df.columns:
                        sort_cols.append(addr_col)
                merged_df = merged_df.sort_values(sort_cols)
            else:
                # normal sort
                sort_column = merged_address_name if merge_address else column_names.get('City')
                if sort_column and sort_column in merged_df.columns:
                    merged_df = merged_df.sort_values(by=sort_column)
            
            if custom_filename:
                output_filename = os.path.join(output_dir, f'{custom_filename}.{file_format}')
            else:
                output_filename = os.path.join(output_dir, f'merged_output_{current_time}.{file_format}')
            
            if file_format == 'xlsx':
                merged_df.to_excel(output_filename, index=False)
                auto_adjust_columns(output_filename)
            else:
                formatted_df = auto_adjust_columns(output_filename, merged_df)
                formatted_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            
            yield output_filename
    else:
        last_file = None
        for i, df in enumerate(all_data):
            if include_date and 'Date' in df.columns:
                df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
            
            if custom_filename:
                output_filename = os.path.join(output_dir, f'{custom_filename}_{i+1}.{file_format}')
            else:
                base_name = os.path.splitext(os.path.basename(pdf_paths[i]))[0]
                output_filename = os.path.join(output_dir, f'{base_name}_{current_time}.{file_format}')

            # sorting
            if filter_by_region and 'Branch ID' in df.columns:
                if merge_address and merged_address_name in df.columns:
                    df = df.sort_values(['Branch ID', merged_address_name])
                else:
                    df = df.sort_values(['Branch ID', column_names['City'], column_names['Address']])
            else:
                sort_col = merged_address_name if merge_address else column_names['City']
                if sort_col in df.columns:
                    df = df.sort_values(sort_col)
            
            if file_format == 'xlsx':
                df.to_excel(output_filename, index=False)
                auto_adjust_columns(output_filename)
            else:
                formatted_df = auto_adjust_columns(output_filename, df)
                formatted_df.to_csv(output_filename, index=False, encoding='utf-8-sig')
            
            last_file = output_filename
        
        yield last_file

    yield 100  # final progress
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
    
    output_dfs = process_pdfs(pdf_paths, merge=merge)
    save_to_excel(output_dfs, pdf_paths, merge=merge)
    
    logging.info(f"Conversion complete. Log file: {log_file}")