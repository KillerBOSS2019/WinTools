##p_imports 

import platform
import subprocess
import os
import platform
## moved to tts.py  - import sounddevice as sd
from ast import literal_eval


import sounddevice as sd

## used in screencapture.py
import mss.tools  # may need to find another module for this due to linux/macOS
from PIL import Image
from io import BytesIO
## used in screencapture.py

PLATFORM_SYSTEM = platform.system()  # Windows / Darwin / Linux


## This is not being utilized YET.. but it will be when we use linux for sure
OS_INFO = {
    "version": platform.version(),          # 10.0.19043 for windows 10
    "arch": platform.architecture(),        # 32bit / 64bit
    "machine": platform.machine(),          # AMD64 / *Intel ??
    "system": platform.system(),            # Windows / Darwin / Linux
    "release": platform.release(),          # Windows-10-10.0.19043-SP0
    "platform full": platform.platform(),
    "platform mac": platform.mac_ver(),
    "platform node": platform.node()
}


if PLATFORM_SYSTEM == "Windows":
    import win32clipboard
    import win32gui  # used to capture window
  # not needed ?? import win32process  # used to capture window
  # not needed ?? import win32ui  # used to capture window
    from win32com.client import GetObject  # Used to Get Display Name / Details

    ## for TTS.py
    import pyttsx3
    import audio2numpy as a2n
    import comtypes

    ## for screencapture.py
    import win32clipboard, os, win32gui, win32api, win32ui
    from ctypes import windll

    





if PLATFORM_SYSTEM == "Linux":
    """ 
    Utilized for setting Linux Specific Imports and Variables based on OS details
    """
    linux_display_server = subprocess.check_output(
        "loginctl show-session $(awk '/tty/ {print $1}' <(loginctl)) -p Type | awk -F= '{print $2}'", shell=True).decode().split("\n")[0]

    if linux_display_server == "wayland":
        print("Its wayland bruh")

    if linux_display_server == "x11":
        print("Its x11 time!")
