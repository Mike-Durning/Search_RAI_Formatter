import time
import pyautogui
import pygetwindow as gw

def hotkey_single(keys, delay=0.05):
    for key in keys:
        pyautogui.press(key)
        #print("Key", key, "Pressed!")
        time.sleep(delay)
    
def hotkey_double(keys, delay=0.05):
    pyautogui.hotkey(*keys)
    #print("Hotkey", keys, "Pressed!")
    time.sleep(delay)

def excel_macro(search_list_format_info):
    try:        
        todays_date = search_list_format_info["date"]
        search_list = search_list_format_info["format"]
        client_name = search_list_format_info["client_name"]
        
        filename_save = filename_save = client_name + " " + search_list + " " + todays_date  # noqa: E501
        
        time.sleep(0.15)
        
        target_window = gw.getWindowsWithTitle(filename_save)

        # Check if the window was found
        if target_window:
            # Bring the window to the foreground
            target_window[0].activate()
        else:
            print(f"Window with title '{filename_save}' not found.")
        
        # Select Total Column
        hotkey_single(['alt', 'j', 't', 't'])

        # Manuever to Total Row
        hotkey_double(['ctrl', 'down'])
        hotkey_single(['down'])
        
        pyautogui.typewrite("Total")
        
        hotkey_single(['right'])
        
        # Enter total function
        pyautogui.typewrite("=SUBTOTAL(103,[Patient Name])")
        hotkey_single(['enter'])
        hotkey_single(['up'])
        
        # Select total row
        hotkey_double(['shift', 'space'])
        
        # center total row
        hotkey_single(['alt', 'h', 'a', 'c'])
        
        # fill total row
        hotkey_single(['alt', 'h', 'h'])

        # Select Correct Fill Color
        for _ in range(3):
            pyautogui.press('down')
        pyautogui.press('enter')
        
        hotkey_single(['left'])
        
        hotkey_double(['ctrl', 'up'])

        # Select entire table
        hotkey_double(['ctrl', 'a'])
        
        hotkey_double(['ctrl', 'a'])

        # Select Autoresize
        hotkey_single(['alt', 'h', 'o', 'i'])
        
        hotkey_single(['left'])
        
        hotkey_double(['ctrl', 's'])
                
    except Exception as e:
        print(f"An error occurred: {e}")
