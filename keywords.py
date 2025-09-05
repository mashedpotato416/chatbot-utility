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
def read_excel_file(excel_filepath, 
                        sheet_name):

    # Read a specific sheet by name:
    df = pd.read_excel(excel_filepath, 
                        sheet_name)

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
    gdata = worksheet.get_all_values()

    # Write the data to a CSV file
    with open('datagspread.csv', 
                    'w', 
                    newline='') as f:
        writer = csv.writer(f)
        writer.writerows(gdata)

    # Read CSV File
    df = pd.read_csv('datagspread.csv', 
                        encoding='windows-1252')

    # Change NaN to Zero
    df_filled = df.fillna(0)

    return df_filled

# Function to extract keywords from data
# Input 1 is a list with format "[<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]]"
# Input 2 is a string
def extract_keywords(excel_data, 
                        keyword_selection, 
                        minus_one=0):

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
                    new_keyword = keyword
                    # Check if prioritization needs to be reduced
                    if minus_one == 1:
                        minus_one_keyword_prio = int(keyword[-1]) - 1
                        prev_keyword_prio = keyword[-2:]
                        minus_one_keyword = keyword.replace(prev_keyword_prio, '~'+str(minus_one_keyword_prio))
                        new_keyword = minus_one_keyword
                    # Check if this is the first keyword
                    if combined_keywords == '':
                        combined_keywords = new_keyword
                    else:
                        combined_keywords += ', ' + new_keyword
            break

    # Copy output to clipboard
    pyperclip.copy(combined_keywords)
