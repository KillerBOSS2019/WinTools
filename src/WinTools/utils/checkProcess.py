import os

import win32con
import win32gui


def check_process(process_name, shortcut ="", focus=True, focus_type="Restore"):
    exist = False
    processes = []
    win32gui.EnumWindows(lambda x, _: processes.append(x), None)
    
    for hwnd in processes:
        window_name = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        if process_name.lower() in window_name.lower():
            exist = True
            print(window_name, class_name, hwnd)
            
            
            if focus:
                print("attempting to focus window")
                #SW_SHOWMAXIMIZED,  SW_RESTORE, SW_SHOWNOACTIVATE, SW_SHOWNORMAL, Minimize 
                #### How to SHOW but not bring to front? certain times windows wont capture cause they arent in some sort of focus...
                if focus_type == "Normal":                 ### difference between SHOWNORMAL and NORMAL  ??  or SHOW_OPENWINDOW ?
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL) 
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Maximized":
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)  
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Restore":
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  
                    win32gui.SetForegroundWindow(hwnd)
                if focus_type == "Minimized":
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  
                    win32gui.SetForegroundWindow(hwnd)
                break
    if not exist:
        try:
            print("load via shortcut")
            os.system('"' + shortcut + '"')
        except Exception as e:
            print("error loading shortcut " + e)
