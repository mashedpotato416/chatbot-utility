from tkinter import *
from tkinter import ttk, messagebox, font
from keywords import read_excel_file, read_googlesheet, extract_keywords
from keywords_aivobot import command_popup, command_browser
import pyperclip
import time
import tkinter as tk
import asyncio

# Global Variables
keyword_btns = []
data = []

# Desired number of columns for the app
number_of_columns = 7

# Desired number of rows for top section
rows_other = 3

# Desired number of rows for bottom section
rows_bottom = 1

# Define Sizes
button_text_width = 20
button_padding = 5
button_paddingx = 5
button_paddingy = 5
grid_paddingx = 10
grid_paddingy = 5
col_multx = 180
col_multy = 50

def open_keywords_ui(data_source: str, 
                        sheet_name='EDatabase', 
                        path='test-file.xlsx'):

    # Inner function to Retrieve data from source
    def get_data_from_source():
        if data_source == 'google':
            # Retrieve data from google sheets
            output = read_googlesheet(sheet_name)
        else:
            # Retrieve data from excel file
            output = read_excel_file(path, 
                                        sheet_name)
        return output

    # Inner function to Select columns with description in first row
    def get_unnamed_headers():
        output = []
        for column in data:
            if 'Unnamed' not in column:
                output.append(column)
        return output

    # Inner function to Get number of rows
    def get_keywords_rows():
        count = len(descriptions)
        output = int(count/number_of_columns)
        return output

    # Inner function to Get window size
    def get_window_size():
        xsize = number_of_columns*col_multx
        ysize = rows_total*col_multy
        output = str(xsize) + 'x' + str(ysize)
        return output

    # Inner function to Add tkinter checkbox
    def add_check_button(tk_root, tk_text, btn_var, btn_value_on, btn_value_off, btn_width, txtypad, btn_col, btn_row, xypad):
        # Set button properties
        btn = tk.Checkbutton(tk_root, text = tk_text, variable = btn_var, onvalue = btn_value_on, offvalue = btn_value_off, anchor='w', justify='left', bd=btn_width, pady=txtypad)

        # Set button position parameters
        btn.grid(column=btn_col, row=btn_row, pady=xypad, padx=xypad)

    # Inner function to add tkinter button
    def add_button(tk_text, 
                    btn_cmd,
                    grid_col, 
                    grid_row,  
                    btn_width, 
                    btn_padx, 
                    btn_pady, 
                    btn_font, 
                    btn_bg, 
                    btn_fg, 
                    grid_padx, 
                    grid_pady, 
                    grid_sticky):

        # Set button properties
        btn = tk.Button(root, 
                        text = tk_text, 
                        command = btn_cmd, 
                        anchor = 'w', 
                        justify = 'left', 
                        width = btn_width, 
                        padx = btn_padx, 
                        pady = btn_pady, 
                        font = btn_font, 
                        bg = btn_bg, 
                        fg = btn_fg)

        # Set button position parameters
        btn.grid(column=grid_col, 
                    row=grid_row, 
                    padx=grid_padx, 
                    pady=grid_pady, 
                    sticky=grid_sticky)

        # Append to list of buttons
        global keyword_btns
        keyword_btns.append(btn)

    # Inner function for keyword button update
    def keyword_update():

        if min1.get() == 1:
            min1_value = 1
        else:
            min1_value = 0

        if paste_browser.get() == 1:
            paste_browser_value = 1
        else:
            paste_browser_value = 0

        update_keyword_button_command(min1_value, paste_browser_value)
        
    
    # Inner function for data refresh
    def data_refresh():
        global data
        # Update data
        data = get_data_from_source()
        # Update button
        keyword_update()
        successful_popup('Data has been updated')
    
    # Inner function update keyword button
    def update_keyword_button_command(minus_one, paste_browser):

        for btn in keyword_btns:
            for description in descriptions:
                if btn['text'] == description:
                    # Create a function for button
                    btn_command = lambda description=description: successful_extract(data, 
                                                                                        description, 
                                                                                        minus_one,
                                                                                        paste_browser)
                    # Update command
                    btn.config(command=btn_command)
                    # Update tk root
                    root.update_idletasks()

    # Successful Messagebox
    def successful_popup(msg):
        messagebox.showinfo('Successful',msg)

    # Successful Extract Prompt
    def successful_extract(excel_data, 
                            keyword_selection, 
                            minus_one,
                            paste_browser):

        # Copy keywords to clipboard
        combined_keywords = extract_keywords(excel_data, 
                                                keyword_selection, 
                                                minus_one)

        # Check if data will be copied directly to the browser
        if paste_browser == 1:
            # Minimize tkinter 
            root.iconify()
            time.sleep(0.5)
            # Find the correct database
            asyncio.run(command_browser(keyword_selection, combined_keywords, minus_one))
            # Successful
            successful_popup('Done')
        else:
            # Copy keywords to clipboard
            pyperclip.copy(combined_keywords)
            # Successful
            successful_popup('Keywords are now in your clipboard')


    # Retrieve data from source
    global data
    data = get_data_from_source()

    # Declare variable as list
    descriptions = []

    # Retrieve headers (database names) and exclude unnamed ones
    descriptions = get_unnamed_headers()

    # Compute number of rows and columns
    rows_keywords = get_keywords_rows()
    rows_keywords_start = rows_other
    rows_keywords_end = rows_keywords + rows_other
    rows_total = rows_keywords + rows_other + rows_bottom

    # Compute size of UI
    window_size = get_window_size()

    # Create tkinter instance
    root = Tk()

    # Set the UI Size         
    root.geometry(window_size)   
 
    # Set UI font
    ui_font = 'Helvetica'
    font_header = font.Font(family=ui_font, 
                                size=30, 
                                weight="bold")

    font_header2 = font.Font(family=ui_font, 
                                size=10, 
                                weight="bold")

    font_checkbtn = font.Font(family=ui_font, 
                                size=10)
                                
    font_btn = font.Font(family=ui_font, 
                            size=9)

    font_btn2 = font.Font(family=ui_font, 
                            size=9, 
                            weight="bold")
    
    # Set Button Color
    # Keyword Buttons
    btn_bg_color = '#d3d4d5'
    btn_fg_color = '#000'
    # Refresh Button
    btn2_bg_color = '#198754'
    btn2_fg_color = '#fff'
    # Quit Button
    btn3_bg_color = '#dc3545'
    btn3_fg_color = '#fff'

    # Set tkinter title - window
    root.title('mp')

    # Set tkinter label - top label
    label_top = tk.Label(text='Utility App', 
                            font=font_header)

    label_top.grid(column=0, 
                    row=0, 
                    pady=5, 
                    padx=5, 
                    columnspan=number_of_columns, 
                    rowspan=1, 
                    sticky=N)

    # Set tkinter label - reminders
    label_top = tk.Label(text='Reminder: Make sure knowledge base window is open in your browser', 
                            font=font_header2,
                            fg='#FF0000')

    label_top.grid(column=0, 
                    row=1, 
                    pady=5, 
                    padx=5, 
                    columnspan=number_of_columns, 
                    rowspan=1, 
                    sticky=N)

    # Initialize count variables
    count_col = 0
    count_row = rows_keywords_start

    # Add checkbox Button
    # Set checkbox properties
    btn_txt = 'Reduce Priority by 1'
    btn_col = 1
    min1 = tk.IntVar()
    btn = tk.Checkbutton(root, 
                            text=btn_txt, 
                            variable=min1, 
                            anchor='w', 
                            justify='center', 
                            width=15, 
                            pady=button_paddingy, 
                            command=keyword_update, 
                            font=font_checkbtn)

    # Set checkbox position parameters
    btn.grid(column=btn_col, 
                row=2, 
                pady=5, 
                padx=5, 
                columnspan=3)
    
    # Add checkbox Button
    # Set checkbox properties
    btn_txt2 = 'Paste on browser'
    btn_col2 = 3
    paste_browser = tk.IntVar()
    btn = tk.Checkbutton(root, 
                            text=btn_txt2, 
                            variable=paste_browser, 
                            anchor='w', 
                            justify='center', 
                            width=15, 
                            pady=button_paddingy, 
                            command=keyword_update, 
                            font=font_checkbtn)

    # Set checkbox position parameters
    btn.grid(column=btn_col2, 
                row=2, 
                pady=5, 
                padx=5, 
                columnspan=3)

    # Add Keyword buttons
    # Iterate thru headers
    for description in descriptions:

        # Temporary command
        btn_command = root.destroy

        # Add new buttons per keyword database
        add_button(tk_text = description, 
                    btn_cmd = btn_command, 
                    btn_width = button_text_width, 
                    btn_padx = button_paddingx, 
                    btn_pady = button_paddingy, 
                    btn_font = font_btn, 
                    btn_bg = btn_bg_color, 
                    btn_fg = btn_fg_color, 
                    grid_padx = grid_paddingx, 
                    grid_pady = grid_paddingy,
                    grid_col = count_col, 
                    grid_row = count_row,
                    grid_sticky = N)

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
    keyword_update()

    # Add Refresh Button

    add_button(tk_text = 'Refresh', 
                    btn_cmd = data_refresh, 
                    btn_width = button_text_width, 
                    btn_padx = button_paddingx, 
                    btn_pady = button_paddingy, 
                    btn_font = font_btn, 
                    btn_bg = btn2_bg_color, 
                    btn_fg = btn2_fg_color, 
                    grid_padx = grid_paddingx, 
                    grid_pady = grid_paddingy,
                    grid_col = number_of_columns-2, 
                    grid_row = rows_total,
                    grid_sticky = N)

    # Add Quit Button

    add_button(tk_text = 'Quit', 
                    btn_cmd = root.destroy, 
                    btn_width = button_text_width, 
                    btn_padx = button_paddingx, 
                    btn_pady = button_paddingy, 
                    btn_font = font_btn2, 
                    btn_bg = btn3_bg_color, 
                    btn_fg = btn3_fg_color, 
                    grid_padx = grid_paddingx, 
                    grid_pady = grid_paddingy,
                    grid_col = number_of_columns-1, 
                    grid_row = rows_total,
                    grid_sticky = N)

    root.mainloop() 

