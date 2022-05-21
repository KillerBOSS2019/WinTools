import ctypes

import numpy as np
import pyautogui
import win32gui
from PIL import Image


def get_windows():
    results = []
    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetWindowText(hwnd):
    
                results.append(win32gui.GetWindowText(hwnd))
                
    win32gui.EnumWindows(winEnumHandler, None)
    return results

def winextra(action):
    
  # if action == "Keep Active, Minimize All (Toggle)":
  #     pyautogui.hotkey('win', 'home')
  #    
  # if action == "Minimize All (Toggle)":
  #     pyautogui.hotkey('win', 'd')
#         
  #  if action =="Clipboard History":
  #      pyautogui.hotkey('win', 'v')
        
    if action == "Emoji":
        pyautogui.hotkey('win', '.')
        
    if action == "Keyboard":
        pyautogui.hotkey('win', 'ctrl', "o")

def get_key_state(key):
    hllDll = ctypes.WinDLL ("User32.dll")
    if key == "NUM LOCK":
        return hllDll.GetKeyState(0x90)
    if key == "CAPS LOCK":
        return hllDll.GetKeyState(0x14)

def move_win_button(direction):
    check = get_key_state('NUM LOCK')
    if not check:
        pyautogui.hotkey('win', 'shift', direction)
    elif check:
        pyautogui.press('numlock')
        pyautogui.hotkey('win', 'shift', direction)
        pyautogui.press('numlock')

def resize_image(im, height, width):
    if im.size[0] == height and im.size[1] == width: 
        return im
    else:
        print("We are actually resizing")
        size = (height,width)
        im.load()
        bands = list(im.split())
        a = np.asarray(bands[-1])
        a.flags.writeable = True
        a[a != 0] = 1
        bands[-1] = Image.fromarray(a)
        bands = [b.resize(size, Image.LINEAR) for b in bands]
        a = np.asarray(bands[-1])
        a.flags.writeable = True
        a[a != 0] = 255
        bands[-1] = Image.fromarray(a)
        im = Image.merge('RGBA', bands)
        return im
