import base64
from io import BytesIO
from time import sleep

import mss
import pyautogui
from constants import TPClient
from PIL import Image
import os

cwd = os.path.join(os.getcwd())


def getFrame_base64(frame_image):
 #  # Get frame (only rgb - smaller size)
 #  frame_rgb     = mss.mss().grab(mss.mss().monitors[2]).rgb

 #  # Convert it from bytes to resize
 #  frame_image   = Image.frombytes("RGB", (1920, 1080), frame_rgb, "raw", "RGB")

    # TEMP SAVING IMAGE TO BUFFER THEN TO BASE 64
    buffer = BytesIO()
    frame_image.save(buffer, format='PNG')
    #### Save it to File #####
    b64_str = base64.standard_b64encode(buffer.getvalue())

    frame_image.close()
    return b64_str


def bgra_to_rgba(sct_img):
    img = Image.frombytes('RGB', sct_img.size, sct_img.rgb, 'raw')
    img = img.convert("RGBA")
    return img


def capture_around_mouse(height, width, livecap=False, overlay=False):
    m_position = pyautogui.position()
    if livecap:
        if overlay:
            monitor_number = 1
            with mss.mss() as sct:
                screenshot_size = [height, width]
                monitor = {
                    "top": m_position.y - screenshot_size[0] // 2,
                    "left": m_position.x - screenshot_size[1] // 2,
                    "width": screenshot_size[0],
                    "height": screenshot_size[1],
                    "mon": monitor_number,
                }
                sct_img = sct.grab(monitor)

              #  que = queue.Queue()
              #  thr = threading.Thread(target = lambda q, arg : q.put(bgra_to_rgba(sct_img)), args = (que, 2))
               # thr = threading.Thread(target = lambda q, arg : q.put(dosomething(arg)), args = (que, 2))
              #  thr.start()
              #  thr.join()
              #  while not que.empty():
              #      img=que.get()
              #      print("Whats this though?")

                finalresult = ""
                img = bgra_to_rgba(sct_img)
                try:
                    finalresult = Image.alpha_composite(img, overlay)
                except ValueError as err:
                    global cap_live
                    cap_live = False
                    print("Value Error", err)

                if finalresult:
                    return finalresult
                else:
                    return "NO RESULT"

        elif not overlay:
            monitor_number = 1
            with mss.mss() as sct:
                screenshot_size = [height, width]
                monitor = {
                    "top": m_position.y - screenshot_size[0] // 2,
                    "left": m_position.x - screenshot_size[1] // 2,
                    "width": screenshot_size[0],
                    "height": screenshot_size[1],
                    "mon": monitor_number,
                }
                sct_img = sct.grab(monitor)
                img = Image.frombytes(
                    "RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                return img

    #""" Else if not live cap"""
    else:
        print("PUSH TO HOLD CAPTURE ACTIVE")
        monitor_number = 1
        with mss.mss() as sct:
            screenshot_size = [height, width]
            monitor = {
                "top": m_position.y - screenshot_size[0] // 2,
                "left": m_position.x - screenshot_size[1] // 2,
                "width": screenshot_size[0],
                "height": screenshot_size[1],
                "mon": monitor_number,
            }
            sct_img = sct.grab(monitor)
            # Instead of making a temp file we get it direct from raw to clipboard
            img = Image.frombytes("RGB", sct_img.size,
                                  sct_img.bgra, "raw", "BGRX")
            return img


def turn_on_cap(height=None, width=None, overlayimage=None):
    global cap_live
    while cap_live:
        img = capture_around_mouse(
            height, width, livecap=True, overlay=overlayimage)
        if img != "NO RESULT":
            TPClient.stateUpdate(
                stateId="KillerBOSS.TP.Plugins.winsettings.active.mouseCapture", stateValue=getFrame_base64(img).decode())
            sleep(0.10)
        elif img == "NO RESULT":
            print("NO RESULT.. STOPPED")
            cap_live = False
            break
        if not cap_live:
            print("Live Cap Stopped")
            cap_live = False
            break


def get_cursor_choices():
    path = cwd+"\mouse_overlays"
    dirs = os.listdir(path)
    cursor_list = []
    for item in dirs:
      #  if item.endswith(".png"):
        #    cursor_list.append(item.split(".")[0])
        cursor_list.append(item)
    return cursor_list
