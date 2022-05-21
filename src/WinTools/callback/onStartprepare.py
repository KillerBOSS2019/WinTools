import pyttsx3
import sounddevice as sd


def getAllVoices():
    engine = pyttsx3.init()
    return engine.getProperty("voices")

def getAllOutput_TTS2():
    audio_dict = {}
    for outputDevice in sd.query_hostapis(0)['devices']:
        if sd.query_devices(outputDevice)['max_output_channels'] > 0:
            audio_dict[sd.query_devices(device=outputDevice)['name']] = sd.query_devices(device=outputDevice)['default_samplerate']
    return audio_dict

import os


def activate_windows_setting(choice=False):
    
    settings ={
        "SYSTEM:  Display": 'ms-settings:display',
        "SYSTEM:  Advanced Display": 'ms-settings:display-advanced',
        "SYSTEM:  Night Light": 'ms-settings:nightlight',
        "SYSTEM:  Sound": 'ms-settings:sound',
        "SYSTEM:  Manage Sound Devices": 'ms-settings:sound-devices',
        "SYSTEM:  Manage App/Device Volume": 'ms-settings:apps-volume',
        "SYSTEM:  App Volume & Device Preferences": 'ms-settings:apps-volume',
        "SYSTEM:  Notifcations & Actions": 'ms-settings:notifications',
        "SYSTEM:  Power & Sleep": 'ms-settings:powersleep',
        "SYSTEM:  Battery": 'ms-settings:batterysaver',
        "SYSTEM:  Battery Usage Details": 'ms-settings:batterysaver-usagedetails',
        "SYSTEM:  Default Save Locations": 'ms-settings:savelocations',
        "SYSTEM:  Multi-Tasking": 'ms-settings:multitasking',
        "SYSTEM:  Sign-in Options": 'ms-settings:signinoptions',     ## this was 'ACCOUNTS'
        "SYSTEM:  Date & Time": 'ms-settings:dateandtime',             ## this was 'TIME & LANGUAGE*
        "SYSTEM:  Time Region": 'ms-settings:regionformatting',        ## this was 'TIME & LANGUAGE*
        "SYSTEM:  Settings Home Page": 'ms-settings:',
        "NETWORK:  Ethernet": 'ms-settings:network-ethernet',
        "NETWORK:  Wi-Fi": 'ms-settings:network-wifi',
        "PERSONALIZATION:  Background": 'ms-settings:personalization-background',
        "PERSONALIZATION:  Colors": 'ms-settings:personalization-colors',
        "PERSONALIZATION:  Lock Screen": 'ms-settings:lockscreen',
        "PERSONALIZATION:  Themes": 'ms-settings:themes',
        "PERSONALIZATION:  Start Folders": 'ms-settings:personalization-start-places',
        "APPS:  Apps & Features": 'ms-settings:appsfeatures',
        "APPS:  Manage Startup Apps": 'ms-settings:startupapps',
        "APPS:  Manage Default Apps": 'ms-settings:defaultapps',
        "APPS:  Manage Optional Features": 'ms-settings:optionalfeatures',
        "GAMING:  Game Bar": 'ms-settings:gaming-gamebar',
        "GAMING:  Game DVR": 'ms-settings:gaming-gamedvr',
        "GAMING:  Game Mode": 'ms-settings:gaming-gamemode',
        "GAMING:  XBOX Networking": 'ms-settings:gaming-xboxnetworking',
        "PRIVACY:  Activity History": 'ms-settings:privacy-activityhistory',
        "PRIVACY:  Webcam": 'ms-settings:privacy-webcam',
        "PRIVACY:  Microphone": 'ms-settings:privacy-microphone',
        "PRIVACY:  Background Apps": 'ms-settings:privacy-backgroundapps',
        r"UPDATE & SECURITY:  Windows Update": 'ms-settings:windowsupdate',
        r"UPDATE & SECURITY:  Windows Recovery": 'ms-settings:recovery',
        r"UPDATE & SECURITY:  Update history": 'ms-settings:windowsupdate-history',
        r"UPDATE & SECURITY:  Restart Options": 'ms-settings:windowsupdate-restartoptions',
        r"UPDATE & SECURITY:  Delivery Optimization": 'ms-settings:delivery-optimization',
        r"UPDATE & SECURITY:  Windows Security": 'ms-settings:windowsdefender',
        r"UPDATE & SECURITY:  Windows Defender": 'windowsdefender:',
        r"UPDATE & SECURITY:  For Developers": 'ms-settings:developers',
    }
    if not choice:
        settings_list =[]
        for thing in settings:
            settings_list.append(thing)
        return settings_list
    else:
        os.system(f'explorer "{settings[choice]}"')
