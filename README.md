# PDF2Excel_AddressConverter

This project converts PDF listings of real estate properties from Centris into Excel files with standardized address information. It offers three main options for processing, with varying levels of address verification, as well as a user-friendly GUI application.

## Features

- Extracts data from multiple PDF files using pdfplumber
- Standardizes city names for Montreal and Laval areas
- Offers three processing options:
  1. PostGrid API (Recommended for highest accuracy with Canadian addresses):
     - Validates and geocodes addresses using PostGrid API (SERP-certified by Canada Post)
     - Implements a scoring system to evaluate address suggestion relevance
  2. Google Maps API (Good for international addresses):
     - Uses Google Maps API for address validation and geocoding
     - Retrieves full city names and postal codes from geocoding results
  3. Simple Processing (Fastest, relies on Centris data accuracy):
     - Extracts and cleans address data without external API calls
     - Assumes Centris data, including postal codes, is accurate
- Outputs data in a standardized format (FNAM, LNAM, ADD1, CITY, PROV, PC)
- Uses multithreading for improved performance
- Implements caching to reduce API calls (for API options)
- Provides detailed logging for API interactions and address processing (for API options)
- GUI application for easy use of the simple processing option

## Why PostGrid?

PostGrid is recognized by Canada Post's Software Evaluation and Recognition Program (SERP), ensuring high-quality address validation for Canadian addresses. SERP-certified software must meet strict accuracy requirements, including:

- 98% accuracy in categorizing valid/invalid addresses
- 99% rejection rate for non-correctable addresses
- 99% correction rate for fixable addresses

By using PostGrid, we ensure our address data meets Canada Post's stringent standards, which is crucial for accurate real estate listings.

[Learn more about SERP-recognized software](https://www.canadapost-postescanada.ca/cpc/en/support/kb/business/customers-move/find-recognized-address-validation-service-providers)

While Google Maps API is available as an alternative and provides good accuracy for international addresses, it is not SERP-certified and may not provide the same level of accuracy for Canadian addresses.

## Scoring System (PostGrid Option)

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
- PyQt5 (for GUI application)

## Setup

1. Clone the repository

2. Install required packages: 
   ```
   pip install pdfplumber pandas requests requests_cache python-dotenv openpyxl PyQt5
   ```

3. Create a `.env` file in the project root and add your API key(s):
   ```
   POSTGRID_API_KEY=your_postgrid_api_key_here
   GOOGLE_MAPS_API_KEY=your_google_maps_api_key_here
   ```

### Obtaining API Keys

#### PostGrid API Key (Recommended for Canadian addresses)

1. Sign up for a PostGrid account at [https://www.postgrid.com/](https://www.postgrid.com/)
2. Navigate to the API section in your dashboard
3. Generate a new API key
4. Copy the generated API key and add it to your `.env` file

Note: Ensure you set up billing and review the pricing for the PostGrid API usage.

#### Google Maps API Key

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Geocoding API
4. Create credentials (API Key)
5. Copy the generated API key and add it to your `.env` file

## Usage

### GUI Application

We now offer a graphical user interface (GUI) for easier use of the PDF to Excel converter.

#### Features

- File Import: Select multiple PDF files or use drag-and-drop functionality
- File Management: Add, remove, or mass delete files from the list
- Export Options: Choose destination folder for output Excel file(s)
- Conversion Settings: Merge PDFs or keep them separate, custom naming for output files
- Progress Tracking: Visual progress bar during conversion
- Error Handling: User-friendly error messages
- Internationalization: Support for French (default) and English languages
- About Section: Project information and link to GitHub repository
- Enhanced User Experience: Keyboard shortcuts and rubber band selection

#### Using the GUI Application

1. Launch the application by running:
   ```
   python pdf2excel_gui.py
   ```
2. Use the "Add Files" button or drag and drop PDF files into the list.
3. Select your desired output folder.
4. (Optional) Choose a custom filename for the output.
5. Click "Convert" to process the files.
6. The application will show progress and notify you when conversion is complete.

Note: The GUI application currently uses the simple processing method without API calls.

### Command-Line Usage

For more advanced options, including API-based processing, use the following command-line scripts:

1. For PostGrid API (recommended for highest accuracy with Canadian addresses): 
   ```
   python pdf2excel_postgrid.py
   ```
2. For Google Maps API (good for international addresses): 
   ```
   python pdf2excel_googlemaps.py
   ```
3. For simple processing (fastest, relies on Centris data): 
   ```
   python pdf2excel.py
   ```

Follow the prompts to select PDF file(s) and specify the output location.

## Executable Application

### Building the Executable

If you want to build the executable yourself:

1. Ensure you have PyInstaller installed:
   ```
   pip install pyinstaller
   ```

2. Use the provided spec file to build the executable:
   ```
   pyinstaller pdf2excel_gui.spec
   ```
2.1 In case the above doesn't work, try the following:
   ```
   python -m PyInstaller pdf2excel_gui.spec
   ```

   This will create an executable named `PDF2Excel_GUI_v1.0.0.exe` in the `dist` folder.

### Pre-built Executable

For convenience, we provide a pre-built executable for Windows users:

1. Go to the [Releases](https://github.com/YourUsername/PDF2Excel_AddressConverter/releases) page of this repository.
2. Download the latest `PDF2Excel_GUI_v1.0.0.exe` file.
3. Run the executable on your Windows machine.

Note: The executable version uses the simple processing method without API calls. For API-based processing, please use the command-line scripts as described in the Usage section above.

## Notes

- The simple processing option (`pdf2excel.py` and GUI) is fastest and doesn't require an API key, relying on the accuracy of Centris data
- The PostGrid option (`pdf2excel_postgrid.py`) provides additional verification for Canadian addresses but requires an API key and may incur costs
- The Google Maps option (`pdf2excel_googlemaps.py`) provides good accuracy for both Canadian and international addresses but requires an API key and may incur costs
- Both API options implement caching to reduce API calls and improve performance
- City name standardization is currently set up for Montreal and Laval areas. Modify the `city_mappings` dictionary in `city_mappings.py` to add more mappings if needed
- The API options provide more robust address parsing, including separation of apartment numbers
- The Google Maps option retrieves full city names and postal codes from the geocoding results
- All options output the data in the same standardized format (FNAM, LNAM, ADD1, CITY, PROV, PC)
- The scripts provide detailed logging in the `logs` directory for troubleshooting and monitoring API interactions (for API options)
- Manual verification may still be necessary for complex cases or addresses not found by the APIs
- For ultimate verification, users can cross-reference with Canada Post's database

## License

Refer to the LICENSE file for more details.

---

*Made with ChatGPT + Canvas & Cursor, including this README file.*