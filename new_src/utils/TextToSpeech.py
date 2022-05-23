import os

import audio2numpy as a2n
import pyttsx3
import sounddevice as sd
from .util import getAllOutput_TTS2, getAllVoices


def TextToSpeech(message, voicesChoics, volume=100, rate=100, output="Default"):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty("volume", volume/100)
    engine.setProperty("rate", rate)
    for voice in getAllVoices():
        if voice.name == voicesChoics:
            engine.setProperty('voice', voice.id)
            print("Using", voice.name, "voices", voice)

    if output =="Default":
        try:
            engine.say(message)
            engine.runAndWait()
            engine.stop()
        except Exception as e:
            print("TTS Error", e)
    else:
        try:
            appdata = os.getenv('APPDATA')
            engine.save_to_file(message, rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
            engine.runAndWait()
            engine.stop()
            device = getAllOutput_TTS2()  ### can make this list returned a global + save that will pull it once and thats it?  instead of every time this action is called..which could be troublesome..
            sd.default.samplerate = device[output]
            sd.default.device = output +", MME"
            x,sr=a2n.audio_from_file(rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
            sd.play(x, sr, blocking=True)
        except Exception as e:
            print("EXCEPTION ERROR(Line:880)", e)