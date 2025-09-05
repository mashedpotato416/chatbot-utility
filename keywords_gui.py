from tkinter import *
from tkinter import ttk, messagebox, font
from keywords import read_excel_file, read_googlesheet, extract_keywords
import tkinter as tk

# Global Variables
keyword_btns = []
data = []

# Desired number of columns for the app
number_of_columns = 7

# Desired number of rows for top section
rows_other = 3

# Desired number of rows for bottom section
rows_bottom = 1

# Successful Messagebox
def successful_popup(msg):
    messagebox.showinfo('Successful',msg)

# Wrapper function
def successful_extract(excel_data, keyword_selection, minus_one):
    extract_keywords(excel_data, keyword_selection, minus_one)
    successful_popup('Keywords were copied to the clipboard')

def open_keywords_ui(data_source: str, sheet_name='EDatabase', path='test-file.xlsx'):
    # Inner function to Retrieve data from source
    def get_data_from_source():
        if data_source == 'google':
            # Retrieve data from google sheets
            output = read_googlesheet(sheet_name)
        else:
            # Retrieve data from excel file
            output = read_excel_file(path, sheet_name)
        return output

    # Inner function to Select columns with description in first row
    # Input is a list with unnamed headers with format [<Header1>,<Header2>,<Header3>,...,<HeaderN>]
    # Output is a list with named headers with format [<Header1>,<Header2>,<Header3>,...,<HeaderN>]
    def get_data_columns(data_with_unnamed):
        output = []
        for column in data_with_unnamed:
            if 'Unnamed' not in column:
                output.append(column)
        return output

    # Inner function to Get number of rows
    # 1st Input is a list without unnamed headers with format [<Header1>,<Header2>,<Header3>,...,<HeaderN>]
    # 2nd input is the desired number of column
    def get_keywords_rows(list_label_and_keywords, num_column):
        count = len(list_label_and_keywords)
        return int(count/num_column)

    # Inner function to Get window size
    # 1st input is the desired number of rows
    # 2nd input is the desired number of columns
    def get_window_size(num_rows, num_column):
        xsize = num_column*195
        ysize = num_rows*50
        return str(xsize) + 'x' + str(ysize)

    # Inner function to Add tkinter checkbox
    def add_check_button(tk_root, tk_text, btn_var, btn_value_on, btn_value_off, btn_width, txtypad, btn_col, btn_row, xypad):
        # Set button properties
        btn = tk.Checkbutton(tk_root, text = tk_text, variable = btn_var, onvalue = btn_value_on, offvalue = btn_value_off, anchor='w', justify='left', width=btn_width, pady=txtypad)

        # Set button position parameters
        btn.grid(column=btn_col, row=btn_row, pady=xypad, padx=xypad)
    # Inner function to add tkinter button
    # 1st input is the tk object
    # 2nd input is the desired text
    # 3rd input is the desired command
    # 4th input is the desired text padding
    # 5th input is the desired column position
    # 6th input is the desired row position
    # 7th input is the desired button x and y padding
    def add_button(tk_root, tk_text, btn_cmd, btn_width, txtypad, btn_col, btn_row, xypad):
        # Set button properties
        btn = tk.Button(tk_root, text = tk_text, command = btn_cmd, anchor='w', justify='left', width=btn_width, pady=txtypad, font=font_btn, bg='#d3d4d5', fg='#000')

        # Set button position parameters
        btn.grid(column=btn_col, row=btn_row, pady=xypad, padx=xypad)

        global keyword_btns
        keyword_btns.append(btn)

    # Inner function for keyword button update
    def keyword_update():
        if min1.get() == 1:
            update_keyword_button_command(root, data, descriptions, 1)
        else:
            update_keyword_button_command(root, data, descriptions, 0)
    
    # Inner function for data refresh
    def data_refresh():
        global data
        # Update data
        data = get_data_from_source()
        # Update button
        keyword_update()
        successful_popup('Data has been updated')
    
    # Inner function update keyword button
    def update_keyword_button_command(tk_root, data, descriptions, min1):

        for btn in keyword_btns:
            for description in descriptions:
                if btn['text'] == description:
                    # Create a function for each button
                    btn_command = lambda description=description: successful_extract(data, description, min1)
                    # Update command
                    btn.config(command=btn_command)
                    # Update tk root
                    tk_root.update_idletasks()

    # Retrieve data from source
    global data
    data = get_data_from_source()

    # Declare variable as list
    descriptions = []

    # Retrieve headers (database names) and exclude unnamed ones
    descriptions = get_data_columns(data)

    # Compute number of rows and columns
    rows_keywords = get_keywords_rows(descriptions,number_of_columns)
    rows_keywords_start = rows_other
    rows_keywords_end = rows_keywords + rows_other
    rows_total = rows_keywords + rows_other + rows_bottom

    # Compute/Define size of UI
    button_width = 25
    button_padding = 5
    text_padding = 5
    window_size = get_window_size(rows_total, number_of_columns)

    # Create tkinter instance
    root = Tk()

    # Set the UI Size         
    root.geometry(window_size)   

    # Set the UI background
    # root.configure(bg='#f7f7f7')
 
    # Set UI font
    ui_font = 'Helvetica'
    font_custom = font.Font(family=ui_font, size=30, weight="bold")
    font_checkbtn = font.Font(family=ui_font, size=12)
    font_btn = font.Font(family=ui_font, size=9)
    font_btn2 = font.Font(family=ui_font, size=9, weight="bold")

    # Set tkinter title - window
    root.title('mp')

    # Set tkinter label - top label
    label_top = tk.Label(text='Utility App', font=font_custom)
    label_top.grid(column=0, row=0, pady=30, padx=5, columnspan=number_of_columns, rowspan=2, sticky=N)

    # Initialize count variables
    count_col = 0
    count_row = rows_keywords_start

    # Add checkbox Button
    # Set checkbox properties
    btn_txt = 'Reduce Priority by 1'
    btn_col = number_of_columns - 1
    min1 = tk.IntVar()
    btn = tk.Checkbutton(root, text=btn_txt, variable=min1, anchor='w', justify='center', width=15, pady=text_padding, command=keyword_update, font=font_checkbtn)

    # Set checkbox position parameters
    btn.grid(column=0, row=2, pady=5, padx=5, columnspan=number_of_columns)

    # Add Keyword buttons
    # Iterate thru headers
    for description in descriptions:

        # Temporary command
        btn_command = root.destroy

        # Add new buttons per database
        add_button(root, description, btn_command, button_width, text_padding, count_col, count_row, button_padding)

        # Control number of buttons per column / Distributes buttons
        if count_row == rows_keywords_end:
            # Next column
            count_col += 1
            # Reset row
            count_row = rows_keywords_start
        else:
            # Next row
            count_row += 1

    # Update button commands
    update_keyword_button_command(root, data, descriptions, min1)

    # Add Refresh Button
    # Set button properties
    btn = tk.Button(root, text = 'Refresh', command = data_refresh, anchor='w', justify='left', width=button_width, pady=text_padding, font=font_btn2, fg='#fff', bg='#198754')
    # Set button position parameters
    btn.grid(column=number_of_columns-2, row=rows_total, pady=button_padding, padx=button_padding)

    # Add Quit Button
    # Set button properties
    btn = tk.Button(root, text = 'Quit', command = root.destroy, anchor='w', justify='left', width=button_width, pady=text_padding, font=font_btn2, fg='#fff', bg='#dc3545')
    # Set button position parameters
    btn.grid(column=number_of_columns-1, row=rows_total, pady=button_padding, padx=button_padding)

    root.mainloop() 

