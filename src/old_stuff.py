#### MAKE ACTION IN PLUGIN TO CHECK IF PROCESS RUNNING, FOCUS PROCESS, ELSE IF NOT RUNNING START DESIRED SHORTCUT SET.
    

###  import time
###  
###  import win32gui
###  results = []
###  def winEnumHandler( hwnd, ctx ):
###      if win32gui.IsWindowVisible( hwnd ):
###         if win32gui.GetWindowText( hwnd ):
###             #print(win32gui.GetWindowText( hwnd ))
###             ctx.append(win32gui.GetWindowText( hwnd ))
###          
###  win32gui.EnumWindows(winEnumHandler, results )
###  
###  old_results = results
###  running = True
###  while running:
###      if old_results != results:
###          print("Previous Count:", len(old_results), "New Count:", len(results))
###          old_results = results
###      else:
###          time.sleep(2)
###          results = []
###          win32gui.EnumWindows(winEnumHandler, results )
###          #print("its the same")
###  
###  #print(results)




###    def check_process(process_name, focus=False):
###        import psutil
###        # Iterate over all running process
###        exists=False
###        for proc in psutil.process_iter():
###            try:
###                # Get process name & pid from process object.
###                processName = proc.name()
###                processID = proc.pid
###                print(processName , ' ::: ', processID)
###                if process_name.lower() in processName.lower():
###                    print("Its RUNNING!")
###                    exists=True
###                    ### bring it to focus now??
###                    import win32gui
###                    hwnd = win32gui.FindWindow(None, "Calculator")
###                    if hwnd:
###                        win32gui.SetForegroundWindow(hwnd)
###                    break
###            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
###                pass
###    
###        if exists == False:
###            import os
###            print("opening up shortcut now...")
###            path = r"C:\\Users\\dbcoo\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Discord Inc\\Discord.lnk"
###            os.system('"' + path + '"')
###            
###    #check_process("discord")
###    
###    
###    import win32gui
###    ## it will focus discord, but not calculator??
###    #hwnd = win32gui.FindWindow(None, "! xXKiller_BOSSXx - Discord")
###    hwnd = win32gui.FindWindow(None, "Calculator")
###    if hwnd:
###        win32gui.SetForegroundWindow(hwnd)
###        print("yup")















        
    #print("OPENING")
    #os.startfile (r"C:/Users/dbcoo/Desktop/MIDItoOBS.lnk")
    #import subprocess
    #subprocess.run(["C:\Users\dbcoo\AppData\Local\Discord\Update.exe --processStart Discord.exe"])
   
   ### this works for Origin Shortcut, but not discord?
   # import subprocess
  #  subprocess.Popen(r'C:\Users\dbcoo\Desktop\The Shortcuts\Origin.lnk', shell=True)
    
##import subprocess
##subprocess.Popen(r'C:\Program Files (x86)\Origin\Origin.exe')
##import os
##os.startfile (r"C:\Program Files (x86)\Origin\Origin.exe")
##print("fk")

from subprocess import call
#call(["C:\Program Files (x86)\Origin\Origin.exe"])
## works but loads stays running in idle window??
#call([r"C:\\Users\\dbcoo\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Discord Inc\\Discord.lnk"])

#import winapps 
#for app in winapps.list_installed(): 
#   print(app)
 
 
#import os
#os.system(Start-Process -FilePath  "C:\Users\dbcoo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Discord Inc\Discord.lnk")
#os.system(r"C:\\Users\\dbcoo\AppData\\Local\\Discord\\Update.exe") # To open any program by their name recognized by windows

# OR

#os.startfile("path to application or any file") # Open any program, text or office document


import os
path = 'C:\\Users\\dbcoo\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Discord Inc\\Discord.lnk'
path2 = r'C:\Users\dbcoo\AppData\Local\Discord\Update.exe --processStart Discord.exe'
path3 = 'C:\\Users\\dbcoo\\Desktop\\The Shortcuts\\Origin.lnk'
obs = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OBS Studio\OBS Studio (64bit).lnk"
#print(path)

#os.system('"' + path + '"')



### way to get running procceses/windows
###    import win32gui
###    def window_enum_handler(hwnd, resultList):
###        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
###            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))
###    
###    def get_app_list(handles=[]):
###        mlst=[]
###        win32gui.EnumWindows(window_enum_handler, handles)
###        for handle in handles:
###            mlst.append(handle)
###        return mlst
###    
###    appwindows = get_app_list()
###    for i in appwindows:
###        print (i)


from pywinauto.findwindows    import find_window
#from pywinauto.win32functions


#win32gui.SetForegroundWindow(find_window(title='taskeng.exe'))
import win32con
##   HWND = win32gui.FindWindow(None, "Calculator")
##   #win32gui.ShowWindow(hwnd,SW_SHOW)
##   
##   win32gui.ShowWindow(HWND, win32con.SW_RESTORE)
##   win32gui.SetWindowPos(HWND,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
##   win32gui.SetWindowPos(HWND,win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)  
##   win32gui.SetWindowPos(HWND,win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_SHOWWINDOW + win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)


#import win32gui
#hwnd = win32gui.FindWindowEx(0,0,0, "obs")
#win32gui.SetForegroundWindow(hwnd)









###  def get_process(process_name):
###      import win32gui
###      exists = False
###      def windowEnumerationHandler(hwnd, top_windows):
###          top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
###      if __name__ == "__main__":
###          results = []
###          top_windows = []
###          win32gui.EnumWindows(windowEnumerationHandler, top_windows)
###          for i in top_windows:
###              if process_name in i[1]:
###                  print (i[0])
###                  print(i)
###                  #win32gui.ShowWindow(i[0],5)  ### original
###                  
###                  win32gui.ShowWindow(i[0], win32con.SW_SHOWNORMAL)  ##updated n working
###                  win32gui.SetForegroundWindow(i[0])
###                  exists = True
###                  break
###          if exists == False:
###              print("Nope")

#get_process("Origin")



import win32gui
##  hwnd = win32gui.FindWindow(None, "Calculator")
##  if hwnd:
##      win32gui.SetForegroundWindow(hwnd)
##      win32gui.ShowWindow(hwnd, 0)
##      print(hwnd)



###Discord
#hwnd = 133888

### Calculator
##   hwnd = 1248156

### OBS 
##   hwnd = 528046
##   
##   ##   print(hwnd)
##   import win32gui as w
##   import win32con as wc
##   w.SetForegroundWindow(hwnd)
##   w.ShowWindow(hwnd, wc.SW_SHOWNORMAL)
##   print("HI")






#

####     Pile of Ducks — Today at 12:32 PM
####    I would guess the issue here is in the way you get the hwnd
#   processes = []
#   w.EnumWindows(lambda x, _: processes.append(x), None)
####       This will list all the processes into the processes list (the list contains hwnd s) 
#   window_name = w.GetWindowText(hwnd)
#   class_name = w.GetClassName(hwnd)
####        then you can loop that list and get the window and class names by doing this
####    then comparing for example class name to a target class name you can determine if its the right hwnd






def get_process(process_name):
    import win32gui
    
    exists = False
    def windowEnumerationHandler(hwnd, top_windows):
        top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
    if __name__ == "__main__":
        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        for i in top_windows:
            if process_name.lower() in i[1].lower():
                print(i[0])
                #win32gui.ShowWindow(i[0],5)  ### original
                win32gui.ShowWindow(i[0], win32con.SW_SHOWNORMAL)  ##updated n working
                win32gui.SetForegroundWindow(i[0])
                exists = True
                break
        if exists == False:
            print("load via shortcut")
            shortcut = "C:\\Program Files (x86)\\Origin\\Origin.exe"
            os.system('"' + shortcut + '"')
            print("Nope")

#get_process("origin")

#### SAVE THIS


### Looping thru running proccess' and getting Name + Class - how to get ID?
def check_proc(process_name, shortcut ="", focus=True):
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
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  ##updated n working
                win32gui.SetForegroundWindow(hwnd)
               #### break so it doesnt keep on triggering random process'
                break
                #print(f"{window_name:20.20} {class_name:20.20}")
    if not exist:
        print("load via shortcut")
        os.system('"' + shortcut + '"')
    
    
#check_proc("Discord", shortcut=r"C:\Users\dbcoo\AppData\Local\Discord\Update.exe --processStart Discord.exe", focus=True)
check_proc("Discord", focus=True)
















#### THE OLD CHECK PROCESS, REPLACED WITH ONE ABOVE
import psutil
def check_process(process_name, shortcut, focus=False):
    # Iterate over all running process
    exists=False
    for proc in psutil.process_iter():
        try:
            # Get process name & pid from process object.
            processName = proc.name()
            processID = proc.pid
           # print(processName , ' ::: ', processID)
            if process_name.lower() in processName.lower():
                print("Its RUNNING!")
                exists=True
                hwnd = win32gui.FindWindow(None, process_name)
                print(processName, processID, hwnd)
                if hwnd:
                    if focus:
                        print("attempting to focus")
                        #SW_SHOWMAXIMIZED,  SW_RESTORE, SW_SHOWNOACTIVATE, SW_SHOWNORMAL    - give options to show normal or restored or max'd ? 
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  ##updated n working
                        win32gui.SetForegroundWindow(hwnd)
                        break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    if exists == False:
        print("opening up shortcut now...")
        os.system('"' + shortcut + '"')
