import pyautogui, sys
import time
import pyperclip
import asyncio

img_conf = 0.8
img_conf2 = 0.9
loc_keyword_txtbox = './screenshots/screenshot1.png'
loc_delete_btn = './screenshots/screenshot2.png'
loc_save_btn = './screenshots/screenshot3.png'
loc_keyword_txt = './screenshots/screenshot4.png'
loc_desc_txt = './screenshots/screenshot5.png'
loc_google_search_btns = './screenshots/screenshot6.png'
loc_pencil_btns = './screenshots/screenshot7.png'
loc_header = './screenshots/screenshot8.png'
loc_scroll_main = './screenshots/screenshot9.png'

# Get center
def locate_center(x,y,w,h):
    loc_x = int(x + (w/2))
    loc_y = int(y + (h/2))
    output = [loc_x, loc_y]
    return output

# Retrieve index from text
def get_text_index(text):
    # Locate dot location
    loc_dot = text.find('.')
    # Get index using dot
    output = text[:loc_dot]
    return output

# Get image start location
async def img_location_start_and_end(path):
    loc = pyautogui.locateOnScreen(path, confidence=img_conf)
    loc_start = int(loc[0])
    loc_end = int(loc[0]) + int(loc[2])
    output = [loc_start, loc_end]
    return output

# Get image center location
async def img_location_center(path):
    loc = pyautogui.locateOnScreen(path, confidence=img_conf)
    output = locate_center(loc[0], 
                            loc[1], 
                            loc[2], 
                            loc[3])
    return output

# Get image 2/3 width location
async def img_location_75(path):
    loc = pyautogui.locateOnScreen(path, confidence=img_conf)
    loc_x = loc[0]+((loc[2]/4)*3)
    loc_y = loc[1]+(loc[3]/2)
    output = [loc_x, loc_y]
    return output

async def locate_and_click(path):
    pyautogui.moveTo(await img_location_center(path))
    pyautogui.click()

async def locate_and_click_75(path):
    pyautogui.moveTo(await img_location_75(path))
    pyautogui.click()

def scroll_down():
    pyautogui.scroll(-10000)

def scroll_up():
    pyautogui.scroll(10000)

async def command_popup():
    # Click on popup window
    await locate_and_click(loc_keyword_txt)
    # Scroll down
    scroll_down()
    # Click delete keywords button
    await locate_and_click(loc_delete_btn)
    # Click keywords text field
    await locate_and_click(loc_keyword_txtbox)
    # Paste keywords
    pyautogui.hotkey('ctrl', 'v')
    # Press enter
    pyautogui.press("enter")
    # Scroll up
    scroll_up()
    # Save changes
    await locate_and_click(loc_save_btn)

async def command_browser(keyword, combined_keywords):
    # Get keyword length
    len_keyword = len(keyword)
    # Click on browser
    await locate_and_click(loc_scroll_main)
    # Scroll up - This will make sure that loc_header will be searchable
    scroll_up()
    # Locate start and end of keyword database name
    loc_header_start_end = await img_location_start_and_end(loc_header)
    # Control + F
    pyautogui.hotkey('ctrl', 'f')
    # Type search keyword
    pyautogui.write(keyword[3:])
    # Close google search
    await locate_and_click_75(loc_google_search_btns)
    # Locate all edit buttons
    loc_pencil = pyautogui.locateAllOnScreen(loc_pencil_btns, confidence=img_conf2)

    for pos in loc_pencil:
        loc_center = locate_center(pos[0], pos[1], pos[2], pos[3])
        # Move to start of row
        pyautogui.moveTo(loc_header_start_end[0], loc_center[1])
        # Drag until end of text
        pyautogui.dragTo(loc_header_start_end[1], loc_center[1], duration=0.2)
        # Copy
        pyautogui.hotkey('ctrl', 'c')
        # Save to variable
        description_text = pyperclip.paste()
        # Check if text is matching with keyword
        if get_text_index(description_text) == get_text_index(keyword):
            # Click on edit button
            pyautogui.moveTo(loc_center)
            pyautogui.click()
            # Copy keywords to clipboard
            pyperclip.copy(combined_keywords)
            # Paste keywords
            await command_popup()
            break