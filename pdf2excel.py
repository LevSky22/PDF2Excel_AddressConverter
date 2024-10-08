import pdfplumber
import pandas as pd
import re
import requests
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from tkinter import Tk, filedialog
import os
import requests_cache  # To use persistent caching
from dotenv import load_dotenv

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

# Get the file paths for input PDFs
pdf_paths = get_pdf_paths()

# Function to extract data using pdfplumber
def extract_with_pdfplumber(pdf_path):
    try:
        # Open the PDF with pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
            for page in pdf.pages:
                # Extract tables from each page using pdfplumber
                table = page.extract_table()
                if table:
                    all_data.extend(table)
            # Convert to DataFrame
            if all_data:
                combined_df = pd.DataFrame(all_data)
                return combined_df
    except Exception as e:
        # Print error message if extraction with pdfplumber fails
        print(f"PDFPlumber failed: {e}")
    return None

# Function to standardize city names
def standardize_city_name(city):
    if pd.isna(city):
        return city  # If the city is NaN, return it as is
    city_mappings = {
        "Ahuntsic": "Montreal",
        "Ahuntsic-Cartierville": "Montreal",
        "Anjou": "Montreal",
        "Baie d'Urfe": "Montreal",
        "Beaconsfield": "Montreal",
        "Cote-des-Neiges-Notre-Dame-de-Grace": "Montreal",
        "Côte-des-Neiges–Notre-Dame-de-Grâce": "Montreal",
        "CDN/NDG": "Montreal",
        "Lachine": "Montreal",
        "LaSalle": "Montreal",
        "L'Île-Bizard–Sainte-Geneviève": "Montreal",
        "L'Île Biz/Geneviève": "Montreal",
        "Le Sud Ouest": "Montreal",
        "Le Sud-Ouest": "Montreal",
        "Mercier–Hochelaga-Maisonneuve": "Montreal",
        "Montréal-Nord": "Montreal",
        "Outremont": "Montreal",
        "Pierrefonds-Roxboro": "Montreal",
        "Plateau-Mont-Royal": "Montreal",
        "Rivière-des-Prairies–Pointe-aux-Trembles": "Montreal",
        "RDP/PAT": "Montreal",
        "Rosemont–La Petite-Patrie": "Montreal",
        "Rosemont": "Montreal",
        "Saint-Laurent": "Montreal",
        "S Laurent": "Montreal",
        "S Léonard": "Montreal",
        "Saint-Léonard": "Montreal",
        "Sud-Ouest": "Montreal",
        "Verdun": "Montreal",
        "Verdun/Île Soeurs": "Montreal",
        "Ville Marie": "Montreal",
        "Villeray–Saint-Michel–Parc-Extension": "Montreal",
        "Villeray/S Michel": "Montreal",
        "Auteuil": "Laval",
        "Chomedey": "Laval",
        "Duvernay": "Laval",
        "Fabreville": "Laval",
        "Îles Laval": "Laval",
        "Laval des Rapides": "Laval",
        "Laval Ouest": "Laval",
        "Laval sur le Lac": "Laval",
        "Pont Viau": "Laval",
        "S François": "Laval",
        "S Martin": "Laval",
        "S Vincent de Paul": "Laval",
        "S Dorothée": "Laval",
        "S Rose": "Laval",
        "Rosemère": "Laval",
        "Vimont": "Laval"
    }
    for key, value in city_mappings.items():
        if city.startswith(key):
            return city.replace(key, value)
    return city

# Function to parallelize PDF extraction
def parallel_pdf_extraction(paths):
    with ThreadPoolExecutor(max_workers=5) as executor:
        return list(executor.map(extract_with_pdfplumber, paths))

# Extract data in parallel
pdf_dataframes = parallel_pdf_extraction(pdf_paths)

# Function to clean address for Google API use while retaining important parts
def clean_address_for_api(address):
    apt_regex = re.compile(r',?\s*apt\.?\s*\d+[a-zA-Z]?', re.IGNORECASE)
    cleaned_address = apt_regex.sub('', str(address)).strip()
    return cleaned_address

# Extract data from each provided PDF file separately
for i, df in enumerate(pdf_dataframes):
    if df is not None:
        # Drop rows and columns that are completely empty
        df.dropna(how="all", inplace=True)

        # Assuming the first row is likely not the header, assign expected headers explicitly
        expected_headers = ['centris_no', 'st', 'mun_bor', 'address', 'price', 'rent_price', 'pt', 'bt', 'rms', 'bdrm', 'bath_pr', 'f-s', 'p', 'g']
        if len(df.columns) >= len(expected_headers):
            extra_columns = [f'column_{i}' for i in range(len(expected_headers), len(df.columns))]
            df.columns = expected_headers + extra_columns
            print(f"Unexpected columns detected: {extra_columns}")
        else:
            df.columns = expected_headers[:len(df.columns)]

        # Drop the first row if it contains incorrect header information or is irrelevant
        if df.iloc[0].str.contains('^[A-Za-z]', regex=True).all():
            df = df.drop(0).reset_index(drop=True)

        # Drop columns with NaN headers to avoid type issues
        df = df.loc[:, ~df.columns.isna()]

        # Standardize column names to lowercase to avoid mismatches
        df.columns = df.columns.str.lower()

        # Print headers for debugging
        print(f"Extracted column headers for {os.path.basename(pdf_paths[i])}:", df.columns.tolist())

        # Attempt to find 'mun_bor' and 'address' columns dynamically
        mun_bor_column = next((col for col in df.columns if 'mun' in col or 'bor' in col), None)
        address_column = next((col for col in df.columns if 'address' in col), None)

        # Ensure required columns exist before proceeding
        if mun_bor_column and address_column:
            # Select relevant columns for processing
            df = df[[mun_bor_column, address_column]]  # Extract only city and address columns

            # Rename columns to standard names for further processing
            df.columns = ['mun_bor', 'address']

            # Standardize city names
            df['mun_bor'] = df['mun_bor'].apply(standardize_city_name)  # Vectorized city standardization

            # Clean addresses to prepare for Google Maps lookup
            df['cleaned_address'] = df['address'].apply(clean_address_for_api)

            # Google Maps API Geocoding Cache to store results and avoid redundant requests
            cache = {}

            def geocode_address(row, retries=3):
                address = row.cleaned_address
                city = row.mun_bor
                full_address = f"{address}, {city}, QC, Canada"  # Append province and country to help accuracy
                # Encode the full address for use in the URL
                encoded_address = urllib.parse.quote(full_address)
                # Check if the address is already cached
                if full_address in cache:
                    return cache[full_address]
                try:
                    # Retry mechanism for geocoding
                    for attempt in range(retries):
                        time.sleep(0.5)  # Keep delay minimal to prevent being rate-limited
                        response = requests.get(f"https://maps.googleapis.com/maps/api/geocode/json?address={encoded_address}&components=locality:{city}|country:CA&key={API_KEY}")
                        data = response.json()
                        if data['status'] == 'OK':
                            # Extract postal code, province, and validate country from the response
                            result = data['results'][0]
                            address_components = result['address_components']
                            postal_code = next((c['long_name'] for c in address_components if 'postal_code' in c['types']), None)
                            province = next((c['long_name'] for c in address_components if 'administrative_area_level_1' in c['types']), None)
                            country = next((c['short_name'] for c in address_components if 'country' in c['types']), None)
                            
                            # Ensure the result is within Canada
                            if country == 'CA':
                                # Cache the result for future use
                                cache[full_address] = (postal_code, province)
                                return postal_code, province

                        # Retry if there is an error or an empty response
                        elif data['status'] in ['OVER_QUERY_LIMIT', 'UNKNOWN_ERROR']:
                            time.sleep(1)  # Wait a bit before retrying
                        else:
                            print(f"Failed to geocode address: {full_address}, Status: {data['status']}, Error: {data.get('error_message', 'N/A')}")
                            break  # If it's not a retryable error, break the loop
                except Exception as e:
                    # Print an error message if the geocoding request fails
                    print(f"Geocoding error: {e}")
                return None, None

            # Use ThreadPoolExecutor to speed up geocoding requests
            with ThreadPoolExecutor(max_workers=10) as executor:
                geocode_results = list(executor.map(geocode_address, [row for row in df.itertuples(index=False)]))

            # Apply geocoding results to the DataFrame
            df[['postal_code', 'province']] = pd.DataFrame(geocode_results, index=df.index)

            # Adding headers for the output: FNAM, LNAM, ADD1, CITY, PROV, PC
            df['fnam'] = 'À'  # Set first name as 'À' for all records
            df['lnam'] = "l'occupant"  # Set last name as 'l'occupant' for all records
            df['add1'] = df['address']  # Full address field
            df['city'] = df['mun_bor']  # City field
            df['prov'] = df['province']  # Province field from geocoding
            df['pc'] = df['postal_code']  # Postal code from geocoding

            # Rearrange columns to match the required output format
            output_df = df[['fnam', 'lnam', 'add1', 'city', 'prov', 'pc']]

            # Export the final DataFrame to an Excel file
            output_filename = f'output_excel/{os.path.splitext(os.path.basename(pdf_paths[i]))[0]}_listings.xlsx'
            output_df.to_excel(output_filename, index=False, engine='openpyxl')

            # Print success message after file creation
            print(f"Excel file '{output_filename}' has been created successfully.")
        else:
            print(f"Required columns 'Mun/Bor.' and 'Address' not found in the extracted data for {os.path.basename(pdf_paths[i])}.")
    else:
        print(f"No data extracted from the provided PDF file: {os.path.basename(pdf_paths[i])}.")
