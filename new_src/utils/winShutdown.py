import pyautogui
import os

def win_shutdown(time, cancel=False):
    ## should we create a shutdown timer/countdown ??  we can warn the user 5 minutes before shutting down so they can cancel
    
    ## IF Blank then we show the GUI
    if time == "":
        time_set= pyautogui.prompt(text='💻 How many MINUTES do you want to wait before shutting down?', title='Shutdown PC?', default='')
        if time_set == "0":
            os.system(f"shutdown -a")
            pyautogui.alert("❗ ABORTED SYSTEM SHUTDOWN ❗")
            time="DONE"
        if type(time) == int:
            if int(time_set) > 0:
                print(f"Shutdown in {time_set} minutes")
                time_set = int(time_set) * 60
                os.system(f"shutdown -s -t {time}")
                time="DONE"
                
    if time == "NOW":
        os.system(f"shutdown -s -t 0")
        
    if type(time) == int:
    
        if int(time) >0:
            try:
                print("Shutdown soon?")
                time = int(time) * 60 
                os.system(f"shutdown -s -t {time}")
            except:
                pass
        elif int(time) == None or 0:
            os.system(f"shutdown -a")
            pyautogui.alert("❗ ABORTED SYSTEM SHUTDOWN ❗")