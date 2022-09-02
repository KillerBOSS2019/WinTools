import os
import json
import platform
import subprocess
import sounddevice as sd
from ast import literal_eval
from functools import reduce

### Screenshot Monitor Imports ###
import mss.tools # may need to find another module for this due to linux/macOS
from PIL import Image
from io import BytesIO

### Port Audio Trouble Shooting
from ctypes.util import find_library
# print(find_library('portaudio'))


""" 
According to my research this is the most reliable way and better than system.platform 
### MAC Examples     - os.name = 'posix'      /     platform.system = 'Darwin'       /   platform.release = '8.11.0'
### Linux Examples   - os.name = 'posix'     /      platform.system = 'Linux'       /    platform.release = '3.19.0-23-generic'
### Windows Examples - os.name = 'nt'       /       platform.system = 'Windows'    /     platform.release = '10'

                                          EXAMPLES                                      """
PLATFORM_SYSTEM = platform.system()       ### Windows / Darwin / Linux
PLATFORM_RELEASE = platform.release()     ### Windows 10 / 11 , or Linux 2.6.22 etc..
PLATFORM_MACHINE = platform.machine()     ### AMD64 / *Intel ??
PLATFORM_ARCH = platform.architecture()   ### 32bit / 64bit 
PLATFORM_PLATFORM = platform.platform()   ### Windows-10-10.0.19043-SP0
PLATFORM_VERSION = platform.version()     ### 10.0.19043 for windows 10
PLATFORM_MAC = platform.mac_ver()         ### This may or may not be useful im not sure.. its here for now - it just comes empty tuples if used on windows

PLATFORM_NODE = platform.node()           ### Computer Network name  - more than likely will never use this

OS_DEETS = {
    "version": PLATFORM_VERSION,
    "arch": PLATFORM_ARCH,
    "machine": PLATFORM_MACHINE,
    "system": PLATFORM_SYSTEM,
    "release": PLATFORM_RELEASE,
    "platform full": PLATFORM_PLATFORM,
    "platform mac": PLATFORM_MAC
    }


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




if PLATFORM_SYSTEM == "Windows":
    """ 
    Utilized for setting Windows Specific Imports and Variables based on OS details
    """
    from ctypes import windll
    import audio2numpy as a2n
    import win32clipboard
    import pyttsx3



if PLATFORM_SYSTEM == "Linux":
    """ 
    Utilized for setting Linux Specific Imports and Variables based on OS details
    """
    linux_display_server = subprocess.check_output("loginctl show-session $(awk '/tty/ {print $1}' <(loginctl)) -p Type | awk -F= '{print $2}'", shell=True).decode().split("\n")[0]

    if linux_display_server == "wayland":
        print("Its wayland bruh")

    if linux_display_server == "x11":
        print("Its x11 time!")



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
        self.other = self.programs["OtherApps"]

    def getSystemApp(self):
        SteamsApps = {}
        OtherApps = {}
        Microsoft = {}

        programs = runWindowsCMD(["powershell", "get-StartApps | ConvertTo-Json"])
        programs = list(literal_eval(programs))

        for program in programs:
            if not program['Name'].lower() in ["readme", "documentation"]:
                if "steam://rungameid/" in program['AppID']:
                    SteamsApps[program['Name']] = program['AppID']
                elif "Microsoft." in program['AppID'] and ".AutoGenerated." not in program['AppID']:
                    Microsoft[program['Name']] = program['AppID']
                else:
                    OtherApps[program['Name']] = program['AppID']
        return {"SteamApps": SteamsApps, "Microsoft": Microsoft, "OtherApps": OtherApps}

    def start(self, appName, apptype):
        if self.programs[apptype].get(appName, False):
            runWindowsCMD("explorer shell:appsfolder\\" + self.programs[apptype][appName])




class Powerplan:
    def __init__(self):
        self.currentPPlan = {}
        self.powerplans = {}
        self.GetPowerplan()

    def GetPowerplan(self):
        powerplanResult = runWindowsCMD("powercfg -List")

        for powerplan in powerplanResult.split("\n"):
            if ":" in powerplan:
                parsedData = powerplan.split(":")[1].split() # from "Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced) *" to ["381b4222-f694-41f0-9685-ff5bb260df2e", "(Balanced) *"]
                the_data = (parsedData[0])
                plan_name = (" ".join(parsedData[1:]))

                if "*" in plan_name:
                    plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                    self.powerplans[plan_name]=the_data

                    self.currentPPlan[plan_name]=the_data
                else:
                    plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                    self.powerplans[plan_name]=the_data
        return self.powerplans
    def changeTo(self, pplanName):
        if self.powerplans.get(pplanName, False):
            runWindowsCMD(f"powercfg.exe /S {self.powerplans[pplanName]}")
            return True
        return False



def jsonPathfinder(data, path):
    pathlist = []
    print(data)
    data = json.loads(data)
    for path in path.split("."):
        try:
            pathlist.append(int(path))
        except ValueError:
            pathlist.append(path)
    return reduce(lambda a, b: a[b], pathlist, data)


class TTS:
    def getAllVoices():
       engine = pyttsx3.init()
       return engine.getProperty('voices')

    def getAllOutput_TTS2():
       audio_dict = {}
       for outputDevice in sd.query_hostapis(0)['devices']:
           if sd.query_devices(outputDevice)['max_output_channels'] > 0:
               audio_dict[sd.query_devices(device=outputDevice)['name']] = sd.query_devices(device=outputDevice)['default_samplerate']
       return audio_dict

    def TextToSpeech(message, voicesChoics, volume=100, rate=100, output="Default"):
       engine = pyttsx3.init()
       voices = engine.getProperty('voices')
       engine.setProperty("volume", volume/100)
       engine.setProperty("rate", rate)
       for voice in voices:
           if voice.name == voicesChoics:
               engine.setProperty('voice', voice.id)
               print("Using", voice.name, "voices", voice)

       if output == "Default":
           engine.say(message)
       else:
           appdata = os.getenv('APPDATA')
           engine.save_to_file(message, rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
       engine.runAndWait()
       engine.stop()

       if output != "Default":
           device = TTS.getAllOutput_TTS2()  ### can make this list returned a global + save that will pull it once and thats it?  instead of every time this action is called..which could be troublesome..
           sd.default.samplerate = device[output]
           sd.default.device = output +", MME"
           x,sr=a2n.audio_from_file(rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
           sd.play(x, sr, blocking=True)
        
        



"""
Screenshot is working, doesnt appear to be copying an image to clipbord anylonger?  
need to double check this is fact
Also need to find a module to replace mss.tools import which saves an RGB data to a file
"""
class ScreenShot:
    def screenshot_monitor(self, 
                           monitor_number,
                           filename=None,
                           clipboard = False):
        
        ### Taking the User input of "Monitor:1 for example" and converting it to the correct monitor number
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
                        # Monitor 0 is ALL Monitors Combined, we need to save this to temp file and then to clipboard
                        
                        image = Image.frombytes('RGB', (sct_img.width, sct_img.height), sct_img.rgb, 'raw', 'RGB', 0, 1)
                        image.save("temp.png")
                     #   mss.tools.to_png(sct_img.rgb, sct_img.size, output="temp.png")
                        
                        # Converting to Bytes then off to Clipboard
                        self.all_monitors_bytes_to_clipboard("temp.png")

                    elif monitor_number != 0:
                        # Instead of making a temp file we get it direct from raw to clipboard
                        img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                        
                        ClipBoard.copy_image_to_clipboard(img)
                       # TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.winsettings.publicIP", getFrame_base64(img).decode())

                if clipboard == False:
                    image = Image.frombytes('RGB', (sct_img.width, sct_img.height), sct_img.rgb, 'raw', 'RGB', 0, 1)
                    image.save(filename + ".png")
                   # mss.tools.to_png(sct_img.rgb, sct_img.size, output=filename + ".png")
                    print("Image saved -> "+ filename+ ".png" )

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

        if PLATFORM_SYSTEM=="Windows":
            ClipBoard.send_to_clipboard(win32clipboard.CF_DIB, data)
        
        os.remove(filepath)





class ClipBoard:
    
    def copy_image_to_clipboard(image):
        if PLATFORM_SYSTEM == "Windows":
            bio = BytesIO()
            image.save(bio, 'BMP')
            data = bio.getvalue()[14:] # removing some headers
            bio.close()
            ClipBoard.send_to_clipboard(win32clipboard.CF_DIB, data)
            
        
        
        ## these may not work just found some random details online need to test
        if PLATFORM_SYSTEM == "Linux":
            
            ### Option # 1
           # os.system(f"xclip -selection clipboard -t image/png -i {path + '/image.png'}")
           # os.system("xclip -selection clipboard -t image/png -i temp_file.png")
           
           
           
           ### Option #2  - https://stackoverflow.com/questions/56618983/how-do-i-copy-a-pil-picture-to-clipboard
           ## might be able to use module called klemboard ??
            memory = BytesIO()
            image.save(memory, format="png")

            output = subprocess.Popen(("xclip", "-selection", "clipboard", "-t", "image/png", "-i"), 
                                      stdin=subprocess.PIPE)
            # write image to stdin
            output.stdin.write(memory.getvalue())
            output.stdin.close()


        ## This needs tested/worked on...
        if PLATFORM_SYSTEM == "Darwin":
            # Option #1
           # os.system(f"pbcopy < {path + '/image.png'}")
            
            # Option #2
            import subprocess
            subprocess.run(["osascript", "-e", 'set the clipboard to (read (POSIX file "image.jpg") as JPEG picture)']) 
        
        
        
        
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
            
