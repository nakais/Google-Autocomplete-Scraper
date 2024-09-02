# Google-Autocomplete-Scraper

# Google Autocomplete Scraper

This script automates the process of collecting Google autocomplete suggestions for a list of keywords stored in an Excel sheet. The script then identifies the longest and shortest suggestions for each keyword and updates the Excel sheet with this information.

## Features
- **Keyword Automation**: Searches Google for each keyword listed in an Excel file.
- **Autocomplete Suggestions**: Collects all autocomplete suggestions provided by Google.
- **Excel Integration**: Updates the Excel sheet with the longest and shortest autocomplete suggestions for each keyword.

## Prerequisites
- Python 3.x
- Selenium
- ChromeDriver
- openpyxl

## Setup
1. Clone the repository.
2. Install the Required Libraries using the following commands:
   ```bash
   pip install selenium openpyxl
    ```
3. Update the excel_file_path variable with the path to your Excel file.
4. Ensure ChromeDriver is installed and added to your PATH.

## Usage
1. Run the script:
   ```bash
   python automate_google_search.py
    ```
2. The Excel sheet will be automatically updated with the collected data.

## Excel sheet
   https://docs.google.com/spreadsheets/d/102r0zxPioHWHpgHKtoaklntOSmPtd-FYB8ffFghUvxI/edit?usp=sharing

##  Notes
- The script retries up to 3 times if it fails to collect suggestions.
- Ensure your Excel file has sheets named after the days of the week. 

   
