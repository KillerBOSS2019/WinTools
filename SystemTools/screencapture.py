#screencapture

from util import ClipBoard, PLATFORM_SYSTEM

### Screenshot Monitor Imports ###
import mss.tools  # may need to find another module for this due to linux/macOS
from PIL import Image
from io import BytesIO


match PLATFORM_SYSTEM:
    case "Windows":
        import os
        import win32ui
        import win32gui
        import win32api
        import win32clipboard
        from util import windll
    case "Linux":
        pass
    case "Darwin":
        pass




"""
Screenshot is working, doesnt appear to be copying an image to clipbord anylonger?  
need to double check this is fact
Also need to find a module to replace mss.tools import which saves an RGB data to a file
"""


class ScreenShot:
    def screenshot_monitor(
            monitor_number,
            filename=None,
            clipboard=False):

        # Taking the User input of "Monitor:1 for example" and converting it to the correct monitor number
        monitor_number = int(monitor_number.split(":")[0])
        with mss.mss() as sct:
            try:
                mon = sct.monitors[monitor_number]
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
                    if monitor_number == 0:
                        image = Image.frombytes(
                            'RGB', (sct_img.width, sct_img.height), sct_img.rgb, 'raw', 'RGB', 0, 1)
                    else:
                        image = Image.frombytes(
                            "RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                    ClipBoard.copy_image_to_clipboard(image)
                else:
                    image = Image.frombytes(
                        'RGB', (sct_img.width, sct_img.height), sct_img.rgb, 'raw', 'RGB', 0, 1)
                    image.save(filename + ".png")
                    print("Image saved -> " + filename + ".png")

            except IndexError:

                print("[ERROR] This Monitor does not exist")

    def all_monitors_bytes_to_clipboard(self, filepath):
        """
        Converting an Image to Bytes to Copy to Clipboard
        """

        image = Image.open(filepath)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        if PLATFORM_SYSTEM == "Windows":
            ClipBoard.send_to_clipboard(win32clipboard.CF_DIB, data)

        os.remove(filepath)

    def get_monitors_Windows_OS():
        try:
           # import win32api
            # Code to retrieve monitor information
            objWMI = win32api.EnumDisplayMonitors()
        except Exception as e:
            print(f"An error occurred: {e}")
            return []

        monitor_list = []
        for idx, monitor in enumerate(objWMI):
            monitor_name = win32api.GetMonitorInfo(monitor[0])['Device']
            # monitor_manufacturer = win32api.GetMonitorInfo(monitor[0])['Manufacturer']
            monitor_list.append(f"{idx+1}: {monitor_name}")

        return monitor_list

    # screenshot window without bringing it to foreground
    def screenshot_window(capture_type=3, window_title=None, clipboard=False, save_location=None):
        hwnd = win32gui.FindWindow(None, window_title)
        try:
            left, top, right, bot = win32gui.GetClientRect(hwnd)
            w = right - left
            h = bot - top

            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
            saveDC.SelectObject(saveBitMap)

            # Change the line below depending on whether you want the whole window
            # or just the client area as shown above.
            # 1, 2, 3 all give different results   ( 3 seems to work for everything)
            result = windll.user32.PrintWindow(
                hwnd, saveDC.GetSafeHdc(), capture_type)
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
                # PrintWindow Succeeded
                if clipboard == True:
                    ClipBoard.copy_image_to_clipboard(im)
                    print("Copied to Clipboard")
                elif clipboard == False:
                    im.save(save_location+".png")
                    print("Saved to Folder")
        except Exception as e:
            print("error screenshot" + e)
