from ctypes import POINTER, cast

import pythoncom
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class AudioController(object):
    def __init__(self, process_name):
        self.process_name = process_name
        self.volume = self.process_volume()

    def process_volume(self):
        pythoncom.CoInitialize()  # 3rd coinitilize...
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # print('Volume:', interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        print(self.process_name)
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # only set volume in the range 0.0 to 1.0
                self.volume = min(1.0, max(0.0, decibels))
                interface.SetMasterVolume(self.volume, None)

    def decrease_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # 0.0 is the min value, reduce by decibels
                self.volume = max(0.0, self.volume-decibels)
                interface.SetMasterVolume(self.volume, None)

    def increase_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:

                # 1.0 is the max value, raise by decibels
                self.volume = min(1.0, self.volume+decibels)
                interface.SetMasterVolume(self.volume, None)


def muteAndUnMute(process, value):
    if value == "Mute":
        value = 1
    elif value == "Unmute":
        value = 0
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == process:
            volume.SetMute(value, None)


def volumeChanger(process, action, value):
    if action == "Set":
        AudioController(str(process)).set_volume((int(value)*0.01))
    elif action == "Increase":
        AudioController(str(process)).increase_volume((int(value)*0.01))

    elif action == "Decrease":
        AudioController(str(process)).decrease_volume((int(value)*0.01))


def setMasterVolume(Vol):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    scalarVolume = int(Vol) / 100
    volume.SetMasterVolumeLevelScalar(scalarVolume, None)


def getMasterVolume() -> int:
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    return int(round(volume.GetMasterVolumeLevelScalar() * 100))
