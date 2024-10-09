# PDF2Excel_AddressConverter


This project converts PDF listings of real estate properties from Centris into Excel files with standardized address information, primarily using PostGrid for high-quality address validation and geocoding.

## Features

- Extracts data from multiple PDF files using pdfplumber
- Standardizes city names for Montreal and Laval areas
- Validates and geocodes addresses using PostGrid API (SERP-certified by Canada Post)
- Implements a scoring system to evaluate address suggestion relevance
- Alternative option to use Google Maps API (not SERP-certified)
- Outputs data in a standardized format (FNAM, LNAM, ADD1, CITY, PROV, PC)
- Uses multithreading for improved performance
- Implements caching to reduce API calls
- Provides detailed logging for API interactions and address processing

## Why PostGrid?

This project primarily uses PostGrid for address validation and geocoding. PostGrid is recognized by Canada Post's Software Evaluation and Recognition Program (SERP), ensuring high-quality address validation for Canadian addresses. SERP-certified software must meet strict accuracy requirements, including:

- 98% accuracy in categorizing valid/invalid addresses
- 99% rejection rate for non-correctable addresses
- 99% correction rate for fixable addresses

By using PostGrid, we ensure our address data meets Canada Post's stringent standards, which is crucial for accurate real estate listings.

[Learn more about SERP-recognized software](https://www.canadapost-postescanada.ca/cpc/en/support/kb/business/customers-move/find-recognized-address-validation-service-providers)

While Google Maps API is also available as an alternative, it is not SERP-certified and may not provide the same level of accuracy for Canadian addresses.

## Scoring System

To ensure the highest quality of address validation, we've implemented a scoring system that evaluates the relevance of address suggestions returned by the PostGrid API. This system considers factors such as:

- Exact matches for street numbers
- Similarity of street names
- City name matches
- Apartment number accuracy

The scoring system helps in selecting the most appropriate suggestion when multiple options are available, improving the overall accuracy of the address validation process.

## Requirements


- Python 3.x
- pdfplumber
- pandas
- requests
- requests_cache
- python-dotenv
- tkinter (usually comes with Python)
- openpyxl

## Setup

1. Clone the repository

2. Install required packages: `pip install pdfplumber pandas requests requests_cache python-dotenv openpyxl`

3. Create a `.env` file in the project root and add your API key(s):
```

POSTGRID_API_KEY=your_postgrid_api_key_here
GOOGLE_MAPS_API_KEY=your_google_maps_api_key_here  # Optional, if using Google Maps

```

### Obtaining API Keys

#### PostGrid API Key (Recommended)

1. Sign up for a PostGrid account at [https://www.postgrid.com/](https://www.postgrid.com/)
2. Navigate to the API section in your dashboard
3. Generate a new API key
4. Copy the generated API key and add it to your `.env` file

Note: Ensure you set up billing and review the pricing for the PostGrid API usage.

#### Google Maps API Key (Alternative)

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Geocoding API
4. Create credentials (API Key)
5. Copy the generated API key and add it to your `.env` file

Note: While Google Maps API is available, it is not SERP-certified and may not provide the same level of accuracy for Canadian addresses.

## Usage


1. Run the desired script:
   - For PostGrid API (recommended): `python pdf2excel_postgrid.py`
   - For Google Maps API (alternative): `python pdf2excel.py`
2. Select the PDF files you want to convert when prompted
3. The script will process the files and output Excel files in the `output_excel` directory

## Notes

- Ensure you have valid API keys for the service you're using
- The scripts use in-memory caching to reduce API calls. Adjust the cache expiration as needed
- City name standardization is currently set up for Montreal and Laval areas. Modify the `city_mappings` dictionary to add more mappings if needed
- The scripts provide detailed logging in the `logs` directory for troubleshooting and monitoring API interactions
- The scoring system helps in selecting the most relevant address suggestions, but manual verification may still be necessary for complex cases

## License

Refer to the LICENSE file for more details.

#

*Made completely with ChatGPT + Canvas & Cursor, including this README file.*