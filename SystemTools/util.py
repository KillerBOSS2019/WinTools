import subprocess
from ast import literal_eval
from io import BytesIO
from TPPEntry import PLATFORM_SYSTEM
import pyautogui
import os

match PLATFORM_SYSTEM:
    case "Windows":
        from win32com.client import GetObject  # Used to Get Display Name / Details
        from ctypes import windll
        import win32clipboard
        import win32gui
    case "Linux":
        pass
    case "Darwin":
        pass


def runWindowsCMD(command):
    """ 
    Running a windows command using system level encoding
    """
    systemencoding = windll.kernel32.GetConsoleOutputCP()
    output = subprocess.run(command, stdout=subprocess.PIPE, shell=True)
    result = str(output.stdout.decode("cp{}".format(str(systemencoding))))
    return result


class SystemPrograms:
    def __init__(self):
        self.programs = self.getSystemApp()
        self.steam = self.programs["SteamApps"]
        self.microsoft = self.programs['Microsoft']
        self.other = self.programs["Other"]

    def getSystemApp(self):
        SteamsApps = {}
        OtherApps = {}
        Microsoft = {}

        programs = runWindowsCMD(
            ["powershell", "get-StartApps | ConvertTo-Json"])
        programs = list(literal_eval(programs))

        for program in programs:
            if not program['Name'].lower() in ["readme", "documentation"]:
                if "steam://rungameid/" in program['AppID']:
                    SteamsApps[program['Name']] = program['AppID']
                elif "Microsoft." in program['AppID'] and ".AutoGenerated." not in program['AppID']:
                    Microsoft[program['Name']] = program['AppID']
                else:
                    OtherApps[program['Name']] = program['AppID']
        return {"SteamApps": SteamsApps, "Microsoft": Microsoft, "Other": OtherApps}

    def start(self, appName, apptype):
        if self.programs[apptype].get(appName, False):
            command = "explorer shell:appsfolder\\" + \
                self.programs[apptype][appName]
            runWindowsCMD(command)


class Get_Windows:
    def get_windows_Windows_OS():
        results = []

        def winEnumHandler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                if win32gui.GetWindowText(hwnd):

                    results.append(win32gui.GetWindowText(hwnd))

        win32gui.EnumWindows(winEnumHandler, None)
        return results

    def get_windows_Linux():
        """ 
        This Includes the Current Active Window 
        - Docs -> https://lazka.github.io/pgi-docs/Wnck-3.0/classes/Window.html#Wnck.Window.get_class_group_name
        """
        import gi
        # It must be set to require 3.0 bfore we import Wnck
        gi.require_version("Wnck", "3.0")
        from gi.repository import Wnck

        scr = Wnck.Screen.get_default()

        # Force Update must be done
        scr.force_update()

   #     all_windows = scr.get_windows()
        # all_windows_list = [x for x in scr.get_windows()]

        window_name_list = []
        for x in scr.get_windows():
            window_name_list.append(x.get_name())

        """ Current Active Window Details"""
        ACTIVE_WINDOW_NAME = scr.get_active_window().get_name()
        ACTIVE_WINDOW_XID = scr.get_active_window().get_xid()
        ACTIVE_WINDOW_PID = scr.get_active_window().get_pid()

        return window_name_list


class ClipBoard:
    def copy_image_to_clipboard(image):
        if PLATFORM_SYSTEM == "Windows":
            bio = BytesIO()
            image.save(bio, 'BMP')
            data = bio.getvalue()[14:]  # removing some headers
            bio.close()
            ClipBoard.send_to_clipboard(win32clipboard.CF_DIB, data)

        # This currently works with Fedora 36 and saves image to clipboard
        if PLATFORM_SYSTEM == "Linux":
            memory = BytesIO()
            image.save(memory, format="png")
            output = subprocess.Popen(("xclip", "-selection", "clipboard", "-t", "image/png", "-i"),
                                      stdin=subprocess.PIPE)
            # write image to stdin
            output.stdin.write(memory.getvalue())
            output.stdin.close()

            # BACKUP OPTIONS
            # Option # 1
            # os.system(f"xclip -selection clipboard -t image/png -i {path + '/image.png'}")
            # os.system("xclip -selection clipboard -t image/png -i temp_file.png")

            # Option #2  - https://stackoverflow.com/questions/56618983/how-do-i-copy-a-pil-picture-to-clipboard
            # might be able to use module called klemboard ??

        # This needs tested/worked on...
        if PLATFORM_SYSTEM == "Darwin":
            # Option #1
            # os.system(f"pbcopy < {path + '/image.png'}")

            # Option #2

            subprocess.run(
                ["osascript", "-e", 'set the clipboard to (read (POSIX file "image.jpg") as JPEG picture)'])

    def send_to_clipboard(clip_type, data):
        if PLATFORM_SYSTEM == "Windows":
            if clip_type == "text":
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(data)
                win32clipboard.CloseClipboard()
            else:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(clip_type, data)
                win32clipboard.CloseClipboard()

def screenshot_current_linux():
    from gi.repository import Gdk, GdkPixbuf

    w = Gdk.get_default_root_window()
    sz = w.get_geometry()
    # print "The size of the window is %d x %d" % sz
    pb = GdkPixbuf.Pixbuf.new(GdkPixbuf.Colorspace.RGB, False, 8, sz[2], sz[3])
    pb = pb.get_from_drawable(w, w.get_colormap(), 0, 0, 0, 0, sz[2], sz[3])
    if (pb != None):
        pb.save("screenshot.png", "png")
        print("Screenshot saved to screenshot.png.")
    else:
        print("Unable to get the screenshot.")





def execute_cmd_command(command):
    """Executes a Command Prompt command."""
    os.system(f"cmd.exe /c {command}")

def execute_powershell_command(command):
    """Executes a PowerShell command."""
    os.system(f"powershell.exe -Command {command}")

## Example usage: Execute a Command Prompt command
#execute_cmd_command("dir")




class SystemState:
    def sleep(self):
        """Puts the system into sleep mode."""
        if PLATFORM_SYSTEM == "Windows":
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif PLATFORM_SYSTEM == "Linux":
            os.system("systemctl suspend")

    def logout(self):
        """Logs out the current user."""
        if PLATFORM_SYSTEM == "Windows":
            os.system("shutdown -l")
        elif PLATFORM_SYSTEM == "Linux":
            os.system("gnome-session-quit --logout")

    def restart(self):
        """Restarts the system."""
        if PLATFORM_SYSTEM == "Windows":
            os.system("shutdown -r")
        elif PLATFORM_SYSTEM == "Linux":
            os.system("systemctl reboot")

    def shutdown(self, delay=None, cancel=False):
        """Shuts down the system."""
        if PLATFORM_SYSTEM == "Windows":
            self.win_shutdown(delay, cancel=False)

        elif PLATFORM_SYSTEM == "Linux":
            if delay is not None:
                os.system(f"shutdown -P {delay}")
            else:
                os.system("shutdown -P now")

    def lock(self):
        """Locks the user session."""
        if PLATFORM_SYSTEM == "Windows":
            os.system("rundll32.exe user32.dll,LockWorkStation")
        elif PLATFORM_SYSTEM == "Linux":
            os.system("gnome-screensaver-command -l")

    def win_shutdown(self, time, cancel=False):
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

# ok = ScreenShot()
#
# ok.screenshot_monitor(monitor_number="0", filename="test", clipboard=True)
# import wx  ## apart of pyGUI ??  can this copy to clipboard on linux too ??  https://www.programcreek.com/python/?code=miloharper%2Fneural-network-animation%2Fneural-network-animation-master%2Fmatplotlib%2Fbackends%2Fbackend_wx.py
"""
def Copy_to_Clipboard(self, event=None):
    "copy bitmap of canvas to system clipboard"
    bmp_obj = wx.BitmapDataObject()
    bmp_obj.SetBitmap(self.bitmap)

    if not wx.TheClipboard.IsOpened():
        open_success = wx.TheClipboard.Open()
        if open_success:
            wx.TheClipboard.SetData(bmp_obj)
            wx.TheClipboard.Close()
            wx.TheClipboard.Flush()
"""


"""
Fedora Example 
{'arch': ('64bit', 'ELF'),
 'machine': 'x86_64',
 'platform full': 'Linux-5.19.4-200.fc36.x86_64-x86_64-with-glibc2.35',
 'platform mac': ('', ('', '', ''), ''),
 'release': '5.19.4-200.fc36.x86_64',
 'system': 'Linux',
 'version': '#1 SMP PREEMPT_DYNAMIC Thu Aug 25 17:42:04 UTC 2022'}
"""

""" 
According to my research this is the most reliable way and better than system.platform 
### MAC Examples     - os.name = 'posix'      /     platform.system = 'Darwin'       /   platform.release = '8.11.0'
### Linux Examples   - os.name = 'posix'     /      platform.system = 'Linux'       /    platform.release = '3.19.0-23-generic'
### Windows Examples - os.name = 'nt'       /       platform.system = 'Windows'    /     platform.release = '10'

                                          EXAMPLES                                      """
# PLATFORM_SYSTEM = platform.system()  # Windows / Darwin / Linux
