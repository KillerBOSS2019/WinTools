import win32gui
import os
import os.path
import mss
import mss.tools
import pyautogui


from io import BytesIO
import win32clipboard
from PIL import Image
import win32ui


##using pyautoGUI which is not a good method..
def screenshot2(window_title=None, clipboard=False, save_location=None):
    ##pyautogui wont work on more than one monitor.. have to use mss.tools
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            x, y, x1, y1 = win32gui.GetClientRect(hwnd)
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
            im = pyautogui.screenshot(region=(x, y, x1, y1))
            
            if clipboard == False:
                im.save(save_location)
                print("Saved to ", save_location)
                
            elif clipboard == True:
                #im.save("temp_window_snip.png")
                #file_to_bytes("temp_window_snip.png")
                copy_im_to_clipboard(im)
                print("Saved to Clipboard")
                pass
            return im
        else:
            print('Window not found!')
#im = screenshot2('Calculator', False, "C:/Users/dbcoo/Downloads/test_file.png")

# 'Taking Screenshot of Window'  
## this brings window to foreground first.. no good for me
def screenshot_old(window_title=None, clipboard=False, save_location=None):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            x, y, x1, y1 = win32gui.GetClientRect(hwnd)
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
   
            with mss.mss() as sct:
    # The screen part to capture 700 - 788
                monitor = {"top": y, "left": x, "width": y1, "height": x1}
                output = "sct-{top}x{left}_{width}x{height}.png".format(**monitor)
    # Grab the data
                sct_img = sct.grab(monitor)
                
    
    # Save picture to file or clipboard
            if clipboard == False:
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=save_location)
                print("Saved to ", save_location)
                
            elif clipboard == True:               
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                copy_im_to_clipboard(img)
                print("Saved to Clipboard")
        else:
            print('Window not found!')
