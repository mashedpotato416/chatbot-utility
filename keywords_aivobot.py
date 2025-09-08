import pyautogui, sys

loc_keyword_txtbox = './screenshots/screenshot1.png'
loc_delete_btn = './screenshots/screenshot2.png'
loc_save_btn = './screenshots/screenshot3.png'

def img_location_center(path):
    loc = pyautogui.locateOnScreen(path)
    loc_x = loc[0]+(loc[2]/2)
    loc_y = loc[1]+(loc[3]/2)
    output = [loc_x, loc_y]
    return output

def locate_and_click(path):
    pyautogui.moveTo(img_location_center(path))
    pyautogui.click()

def command1():
    # Click delete keywords button
    locate_and_click(loc_delete_btn)
    # Click keywords text field
    locate_and_click(loc_keyword_txtbox)
    # Paste keywords
    pyautogui.hotkey('ctrl', 'v')
    # Press enter
    pyautogui.press("enter")
    # Save changes
    locate_and_click(loc_save_btn)


