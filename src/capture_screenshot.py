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


### Clipboard Stuff
def copy_im_to_clipboard(image):
    bio = BytesIO()
    image.save(bio, 'BMP')
    data = bio.getvalue()[14:] # removing some headers
    bio.close()

    send_to_clipboard(win32clipboard.CF_DIB, data)


def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()


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
    #os.remove(filepath)
    #print("Temp Image Deleted")
### End of Clipboard Stuff  

def check_number_of_monitors():
    with mss.mss() as sct:
        monitor_count = (len(sct.monitors))
        return monitor_count

### Call back on_exist not being used            
def on_exists(fname):
    # type: (str) -> None
    if os.path.isfile(fname):
        newfile = fname + ".old"
        print("{} -> {}".format(fname, newfile))
        os.rename(fname, newfile)
        
def save_screenshot(sct_img_rgb, sct_img_size, filename):
    mss.tools.to_png(sct_img_rgb, sct_img_size, output=filename + ".png")
    ## Should we compress to JPG.. or save as JPG originally somehow?


###screenshot window without bringing it to foreground - https://stackoverflow.com/questions/19695214/python-screenshot-of-inactive-window-printwindow-win32gui/24352388#24352388
def screenshot_window(capture_type, window_title=None, clipboard=False, save_location=None):
    from ctypes import windll
    hwnd = win32gui.FindWindow(None, window_title)
    
    ##  Change the line below depending on whether you want the whole window
    ##  May be needed in future?
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
    # or just the client area. 
                          # 1, 2, 3 all give different results
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
            im.save(save_location)
            print("Saved to Folder")
            

#screenshot_window(capture_type=3, window_title="Calculator", clipboard= False, save_location="testing2.png")


def screenshot_monitor(monitor_number, filename=False, clipboard = False):    
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
                if monitor_number==0:
                    print("capturing all")  ## having to use temp file to capture all screens successfully?
                    mss.tools.to_png(sct_img.rgb, sct_img.size, output="temp.png")
                    file_to_bytes("temp.png") ## Saved to Clipboard
                    
                 ##   Instead of making a temp file we get it direct from raw to clipboard
                elif monitor_number != 0:
                    img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                    copy_im_to_clipboard(img)
                
                
            if clipboard == False:
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=filename + ".png")
                print("Image saved -> "+ filename+ ".png" )

        except IndexError:
            print("This Monitor does not exist")

screenshot_monitor(1, filename=f"C:\Users\dbcoo\Downloads\GGG.jpg", clipboard = False)


