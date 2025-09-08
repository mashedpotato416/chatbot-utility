import pyautogui, sys
import time

time_delay = 0.5
loc_keyword_txtbox = './screenshots/screenshot1.png'
loc_delete_btn = './screenshots/screenshot2.png'
loc_save_btn = './screenshots/screenshot3.png'
loc_keyword_txt = './screenshots/screenshot4.png'
loc_desc_txt = './screenshots/screenshot5.png'

def img_location_center(path):
    loc = pyautogui.locateOnScreen(path, confidence=0.8)
    loc_x = loc[0]+(loc[2]/2)
    loc_y = loc[1]+(loc[3]/2)
    output = [loc_x, loc_y]
    return output

def locate_and_click(path):
    pyautogui.moveTo(img_location_center(path))
    pyautogui.click()

def scroll_down():
    pyautogui.scroll(-10000)

def scroll_up():
    pyautogui.scroll(10000)

def command1():
    # Click on popup window
    locate_and_click(loc_keyword_txt)
    time.sleep(time_delay)
    # Scroll down
    scroll_down()
    time.sleep(time_delay)
    # Click delete keywords button
    locate_and_click(loc_delete_btn)
    time.sleep(time_delay)
    # Click keywords text field
    locate_and_click(loc_keyword_txtbox)
    time.sleep(time_delay)
    # Paste keywords
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(time_delay)
    # Press enter
    pyautogui.press("enter")
    time.sleep(time_delay)
    # Save changes
    locate_and_click(loc_save_btn)

