# PDF2Excel_AddressConverter


This project converts PDF listings of real estate properties from Centris into Excel files with standardized address information.

## Features

- Extracts data from multiple PDF files using pdfplumber
- Standardizes city names for Montreal and Laval areas
- Geocodes addresses using Google Maps API to obtain postal codes and provinces
- Outputs data in a standardized format (FNAM, LNAM, ADD1, CITY, PROV, PC)
- Uses multithreading for improved performance
- Implements caching to reduce API calls


## Requirements


- Python 3.x
- pdfplumber
- pandas
- requests
- requests_cache
- python-dotenv
- tkinter (usually comes with Python)


## Setup

1. Clone the repository

2. Install required packages: `pip install pdfplumber pandas requests requests_cache python-dotenv`

3. Create a `.env` file in the project root and add your Google Maps API key:
```

GOOGLE_MAPS_API_KEY=your_api_key_here

```  

### Obtaining a Google Maps API Key

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Geocoding API:
   - In the sidebar, click on "APIs & Services" > "Library"
   - Search for "Geocoding API" and click on it
   - Click "Enable"
4. Create credentials:
   - In the sidebar, click on "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "API Key"
5. Copy the generated API key and add it to your `.env` file

Note: Ensure you set up billing and quotas in the Google Cloud Console to use the API.

## Usage


1. Run the script: `python pdf2excel.py`
2. Select the PDF files you want to convert when prompted
3. The script will process the files and output Excel files in the `output_excel` directory

## Notes

- Ensure you have a valid Google Maps API key with Geocoding API enabled
- The script uses in-memory caching to reduce API calls. Adjust the cache expiration as needed
- City name standardization is currently set up for Montreal and Laval areas. Modify the `city_mappings` dictionary to add more mappings if needed

## License

Refer to the LICENSE file for more details.

### Made completely with ChatGPT + Canvas & Cursor, including this README file.