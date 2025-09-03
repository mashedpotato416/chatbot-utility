from gspread.utils import GridRangeType
from dotenv import load_dotenv
import os
import pandas as pd
import pyperclip
import math
import gspread
import csv

# Function for excel files
# Data will be in the format "[<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]]"
def read_excel_file(excel_filepath, sheet_name):

    # Read a specific sheet by name:
    df = pd.read_excel(excel_filepath, sheet_name)

    # Change NaN to Zero
    df_filled = df.fillna(0)

    return df_filled

# Function for google sheet
# Output will be in the format "[<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]]"
def read_googlesheet(wrsheet):

    # Load variables from .env file
    load_dotenv() 
    google_api_key = os.getenv("GOOGLESHEETS_API_KEY")
    google_link = os.getenv("GOOGLESHEETS_LINK")

    # Set gspread parameters
    gc = gspread.api_key(google_api_key)
    sh = gc.open_by_url(google_link)

    # Select worksheet
    worksheet = sh.worksheet(wrsheet)

    # Get all values from the worksheet
    data = worksheet.get_all_values()

    # Write the data to a CSV file
    with open('datagspread.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(data)

    # Read CSV File
    df = pd.read_csv('datagspread.csv', encoding='windows-1252')

    # Change NaN to Zero
    df_filled = df.fillna(0)

    return df_filled

# Function to extract keywords from data
def extract_keywords(excel_data, keyword_selection):

    # Initialize output variable
    combined_keywords = ''

    # Iterate thru the contents of the sheet Database
    for label,keywords in excel_data.items():
        # Select correct database or label
        if label == keyword_selection:
            # Iterate thru the keywords under the selected database or label
            for keyword in keywords:
                # Exclude zeroes
                if keyword != 0:
                    # Append to output variable
                    if combined_keywords == '':
                        combined_keywords = keyword
                    else:
                        combined_keywords += ', ' + keyword
            break

    # print(combined_keywords)
    pyperclip.copy(combined_keywords)
