import os
from ctypes import windll
from io import BytesIO

import mss
import win32clipboard
import win32gui
import win32ui
from PIL import Image

from utils.clipboard import copy_im_to_clipboard

from .clipboard import send_to_clipboard


## Image to Bytes
def file_to_bytes(filepath):
    ## Take image into bytes and onto clipboard
    image = Image.open(filepath)
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()
    
    ### Sending to Clipboard
    send_to_clipboard(win32clipboard.CF_DIB, data)
    ### Deleting Temp File
    os.remove(filepath)
    print("Temp Image Deleted")

###screenshot window without bringing it to foreground 
def screenshot_window(capture_type, window_title=None, clipboard=False, save_location=None):
    hwnd = win32gui.FindWindow(None, window_title)
    try:
        left, top, right, bot = win32gui.GetClientRect(hwnd)
        #left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top
        
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        
        # Change the line below depending on whether you want the whole window
        # or just the client area as shown above. 
                              # 1, 2, 3 all give different results   ( 3 seems to work for everything)
        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), capture_type)
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        
        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        
        if result == 1:
            #PrintWindow Succeeded
            if clipboard == True:
                copy_im_to_clipboard(im)
                print("Copied to Clipboard")
            elif clipboard == False:
                im.save(save_location+".png")
                print("Saved to Folder")
    except Exception as e:
        print("error screenshot" + e )

def screenshot_monitor(monitor_number, filename="", clipboard = False):   
    monitor_number = int(monitor_number.split(":")[0])
    with mss.mss() as sct:
        try:
            mon = sct.monitors[monitor_number]  
            # Capturing Entire Monitor
            monitor = {
                "top": mon["top"],
                "left": mon["left"],
                "width": mon["width"],
                "height": mon["height"],
                "mon": monitor_number,
            }
            # Grab the Image
            sct_img = sct.grab(monitor)     
    
            if clipboard == True:
                """Having to Save to Temp File To Capture All Screens Successfully.. Investigate why"""
                if monitor_number==0:
                    print("capturing all, using temp file")  
                    mss.tools.to_png(sct_img.rgb, sct_img.size, output="temp.png")
                    file_to_bytes("temp.png") ## Saved to Clipboard
                    
                elif monitor_number != 0:
                    img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")  ##   Instead of making a temp file we get it direct from raw to clipboard
                    copy_im_to_clipboard(img)
                   # TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.winsettings.publicIP", getFrame_base64(img).decode())

            if clipboard == False:
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=filename + ".png")
                print("Image saved -> "+ filename+ ".png" )
                
        except IndexError:
            print("This Monitor does not exist")
