from tkinter import *
from tkinter import ttk, messagebox
from keywords import read_excel_file, read_googlesheet, extract_keywords

# Successful Messagebox
def successful_popup():
    messagebox.showinfo("Keywords were copied to the clipboard")

# Wrapper function
def successful_extract(excel_data, keyword_selection):
    extract_keywords(excel_data, keyword_selection)
    messagebox.showinfo("Successful","Keywords were copied to the clipboard")

# Select columns with description in first row
# Input is a list with unnamed headers with format [<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]]
# Output is a list with named headers with format [<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]]
def get_data_columns(data_with_unnamed):
    output = []
    for column in data_with_unnamed:
        if 'Unnamed' not in column:
            output.append(column)
    return output

# Get number of rows
# 1st Input is a list without unnamed headers with format [<Desctiption>,[<keyword1>,<keyword2>,...,<keywordnN>]] 
# 2nd input is the desired number of column
def get_number_rows(list_label_and_keywords, num_column):
    count = len(list_label_and_keywords)
    return int(count/num_column)

# Get window size
# 1st input is the desired number of rows
# 2nd input is the desired number of columns
def get_window_size(num_rows, num_column):
    xsize = num_column*200
    ysize = num_rows*48
    return str(xsize) + 'x' + str(ysize)

# Add tkinter button
# 1st input is the tk object
# 2nd input is the desired text
# 3rd input is the desired command
# 4th input is the desired text padding
# 5th input is the desired column position
# 6th input is the desired row position
# 7th input is the desired button x and y padding
def add_button(tk_root, tk_text, btn_cmd, btn_width, txtypad, btn_col, btn_row, xypad):
    # Set button properties
    btn = Button(tk_root, text = tk_text, command = btn_cmd, anchor='w', justify="left", width=btn_width, pady=txtypad)

    # Set button position parameters
    btn.grid(column=btn_col, row=btn_row, pady=xypad, padx=xypad)  

def open_keywords_ui(data_source: str, sheet_name="EDatabase", path="test-file.xlsx"):
    
    # Initialize variables
    count_col = 0
    count_row = 0
    number_of_columns = 5
    descriptions = []

    if data_source == 'google':
        # Retrieve data from google sheets
        data = read_googlesheet(sheet_name)
    else:
        # Retrieve data from excel file
        data = read_excel_file(path, sheet_name)

    # Remove unnamed columns
    descriptions = get_data_columns(data)

    # Compute number of rows and columns
    number_of_rows = get_number_rows(descriptions,number_of_columns)

    # Compute/Define size of GUI
    button_width = 25
    button_padding = 5
    text_padding = 5
    window_size = get_window_size(number_of_rows, number_of_columns)

    # Create tkinter instance
    root = Tk()

    # Set the GUI Size         
    root.geometry(window_size)   

    # Add Keyword buttons
    for description in descriptions:

        btn_command = lambda description=description: successful_extract(data,description)
        add_button(root, description, btn_command, button_width, text_padding, count_col, count_row, button_padding)

        # Control number of buttons per column / Distributes buttons
        if count_row == number_of_rows:
            count_col += 1
            count_row = 0
        else:
            count_row += 1

    # Add Quit Button
    add_button(root, 'Quit', root.destroy, button_width, text_padding, count_col, number_of_rows, button_padding)

    root.mainloop() 
