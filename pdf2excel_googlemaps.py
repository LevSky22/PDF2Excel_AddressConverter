import pdfplumber
import pandas as pd
import re
import requests
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from tkinter import Tk, filedialog
import os
import requests_cache
from dotenv import load_dotenv
from city_mappings import get_city_from_borough

# Enable in-memory caching for requests
requests_cache.install_cache('google_maps_cache', backend='memory', expire_after=86400)  # Cache expires after 1 day

# Google Maps API Key
load_dotenv()
API_KEY = os.getenv('GOOGLE_MAPS_API_KEY')

# Function to get file paths using a file dialog
def get_pdf_paths():
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window
    file_paths = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])
    return list(file_paths)

# Function to extract data using pdfplumber
def extract_with_pdfplumber(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    # Skip the first row (header) for each page
                    all_data.extend(table[1:])
            if all_data:
                combined_df = pd.DataFrame(all_data)
                return combined_df
    except Exception as e:
        print(f"PDFPlumber failed: {e}")
    return None

# Function to standardize city names
def standardize_city_name(city):
    if pd.isna(city):
        return city
    return get_city_from_borough(city)

# Function to parallelize PDF extraction
def parallel_pdf_extraction(paths):
    with ThreadPoolExecutor(max_workers=5) as executor:
        return list(executor.map(extract_with_pdfplumber, paths))

# Function to clean address
def clean_address(address):
    if not address:
        return ""
    # Remove trailing comma and any following whitespace
    address = re.sub(r',\s*$', '', address.strip())
    # Remove any other trailing punctuation
    address = re.sub(r'[.,;:]+$', '', address)
    # Remove 'Apt' or 'Apt.' if it's at the end of the address
    address = re.sub(r'\s+(?:Apt\.?|Apartment)?\s*$', '', address, flags=re.IGNORECASE)
    # Remove any leading or trailing whitespace
    return address.strip()

# Function to separate apartment number from address
def separate_apartment(address):
    if address is None:
        return None, None
    
    # Clean the address first
    address = clean_address(address)
    
    # Updated regex to capture more apartment formats, including "apt. 1" style
    apt_regex = re.compile(r'(.*?)\s*(?:,\s*)?((?:app?t?\.?\s*|unit\s*|suite\s*|#\s*)#?\s*\d+[a-zA-Z]?)\s*$', re.IGNORECASE)
    match = apt_regex.search(address)
    
    if match:
        main_address = match.group(1).strip()
        apt_number = match.group(2).strip()
        # Standardize apartment format, but keep original formatting if it's "apt. X"
        if not apt_number.lower().startswith('apt.'):
            apt_number = re.sub(r'^(#\s*|\s*#)', 'Apt. ', apt_number, flags=re.IGNORECASE)
        return main_address, apt_number
    
    return address, None

# Function to clean 'None' values
def clean_none(value):
    return '' if pd.isna(value) or value == 'None' else str(value)

# Main execution
if __name__ == "__main__":
    pdf_paths = get_pdf_paths()
    pdf_dataframes = parallel_pdf_extraction(pdf_paths)

    for i, df in enumerate(pdf_dataframes):
        if df is not None:
            df.dropna(how="all", inplace=True)

            expected_headers = ['centris_no', 'st', 'mun_bor', 'address', 'price', 'rent_price', 'pt', 'bt', 'rms', 'bdrm', 'bath_pr', 'f-s', 'p', 'g']
            if len(df.columns) >= len(expected_headers):
                extra_columns = [f'column_{i}' for i in range(len(expected_headers), len(df.columns))]
                df.columns = expected_headers + extra_columns
                print(f"Unexpected columns detected: {extra_columns}")
            else:
                df.columns = expected_headers[:len(df.columns)]

            df = df.reset_index(drop=True)

            mun_bor_column = next((col for col in df.columns if 'mun' in col or 'bor' in col), None)
            address_column = next((col for col in df.columns if 'address' in col), None)

            if mun_bor_column and address_column:
                df = df[[mun_bor_column, address_column]]
                df.columns = ['mun_bor', 'address']

                df['mun_bor'] = df['mun_bor'].apply(standardize_city_name)

                # Separate apartment numbers and clean addresses
                df['cleaned_address'], df['apartment'] = zip(*df['address'].apply(separate_apartment))

                # Google Maps API Geocoding Cache to store results and avoid redundant requests
                cache = {}

                def geocode_address(row, retries=3):
                    address = row.cleaned_address
                    city = row.mun_bor
                    full_address = f"{address}, {city}, QC, Canada"
                    encoded_address = urllib.parse.quote(full_address)
                    if full_address in cache:
                        return cache[full_address]
                    try:
                        for attempt in range(retries):
                            time.sleep(0.5)
                            response = requests.get(f"https://maps.googleapis.com/maps/api/geocode/json?address={encoded_address}&components=locality:{city}|country:CA&key={API_KEY}")
                            data = response.json()
                            if data['status'] == 'OK':
                                result = data['results'][0]
                                address_components = result['address_components']
                                postal_code = next((c['long_name'] for c in address_components if 'postal_code' in c['types']), None)
                                province = next((c['long_name'] for c in address_components if 'administrative_area_level_1' in c['types']), None)
                                country = next((c['short_name'] for c in address_components if 'country' in c['types']), None)
                                
                                # Retrieve full city name from geocoding result
                                full_city = next((c['long_name'] for c in address_components if 'locality' in c['types']), city)
                                
                                if country == 'CA':
                                    cache[full_address] = (postal_code, province, full_city)
                                    return postal_code, province, full_city

                            elif data['status'] in ['OVER_QUERY_LIMIT', 'UNKNOWN_ERROR']:
                                time.sleep(1)
                            else:
                                print(f"Failed to geocode address: {full_address}, Status: {data['status']}, Error: {data.get('error_message', 'N/A')}")
                                break
                    except Exception as e:
                        print(f"Geocoding error: {e}")
                    return None, None, None

                # Use ThreadPoolExecutor to speed up geocoding requests
                with ThreadPoolExecutor(max_workers=10) as executor:
                    geocode_results = list(executor.map(geocode_address, [row for row in df.itertuples(index=False)]))

                # Apply geocoding results to the DataFrame
                df[['postal_code', 'province', 'full_city']] = pd.DataFrame(geocode_results, index=df.index)

                # Use the full city name from geocoding if available, otherwise use the original city name
                df['city'] = df.apply(lambda row: row['full_city'] if pd.notna(row['full_city']) else row['mun_bor'], axis=1)

                # Adding headers for the output: FNAM, LNAM, ADD1, CITY, PROV, PC
                df['fnam'] = 'À'
                df['lnam'] = "l'occupant"
                df['add1'] = df.apply(lambda row: f"{row['cleaned_address']}, {row['apartment']}" if pd.notna(row['apartment']) else row['cleaned_address'], axis=1)
                df['prov'] = df['province']
                df['pc'] = df['postal_code']

                # Rearrange columns to match the required output format
                output_df = df[['fnam', 'lnam', 'add1', 'city', 'prov', 'pc']].copy()

                # Clean the 'add1' column one last time
                output_df.loc[:, 'add1'] = output_df['add1'].apply(clean_address)

                # Apply the cleaning function to all relevant columns
                columns_to_clean = ['add1', 'city', 'prov', 'pc']
                for col in columns_to_clean:
                    if col in output_df.columns:
                        output_df.loc[:, col] = output_df[col].apply(clean_none)
                    else:
                        print(f"Warning: Column '{col}' not found in DataFrame")

                # Remove rows where all address-related fields are empty
                output_df = output_df[~(output_df['add1'].isna() & output_df['city'].isna() & output_df['prov'].isna() & output_df['pc'].isna())]

                # Export the final DataFrame to an Excel file
                output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_paths[i]))[0]}_listings.xlsx'
                output_df.to_excel(output_filename, index=False, engine='openpyxl')

                print(f"Excel file '{output_filename}' has been created successfully.")
                print(f"Validated addresses: {output_df['add1'].notna().sum()}")
                print(f"Total rows in output: {len(output_df)}")
            else:
                print(f"Required columns 'Mun/Bor.' and 'Address' not found in the extracted data for {os.path.basename(pdf_paths[i])}.")
        else:
            print(f"No data extracted from the provided PDF file: {os.path.basename(pdf_paths[i])}.")