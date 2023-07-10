#tts.py
from util import PLATFORM_SYSTEM
import os
import sounddevice as sd


match PLATFORM_SYSTEM:
    case "Windows":
        import pyttsx3
        import audio2numpy as a2n
        import comtypes
    case "Linux":
        pass
    case "Darwin":
        pass




class TTS:
    def getAllVoices():
        comtypes.CoInitialize()
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        comtypes.CoUninitialize()
        return voices

    def getAllOutput_TTS2():
        audio_dict = {}
        for outputDevice in sd.query_hostapis(0)['devices']:
            if sd.query_devices(outputDevice)['max_output_channels'] > 0:
                audio_dict[sd.query_devices(device=outputDevice)['name']] = sd.query_devices(
                    device=outputDevice)['default_samplerate']
        return audio_dict

    def TextToSpeech(message, voicesChoics, volume=100, rate=100, output="Default"):
        sd.stop()
        sd.default.reset()
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        engine.setProperty("volume", volume/100)
        engine.setProperty("rate", rate)
        for voice in voices:
            if voice.name == voicesChoics:
                engine.setProperty('voice', voice.id)
                print("Using", voice.name, "voices", voice)


        appdata = os.getenv('APPDATA')
        engine.save_to_file(
            message, rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
        engine.runAndWait()
        engine.stop()

        if output != "Default":
            # can make this list returned a global + save that will pull it once and thats it?  instead of every time this action is called..which could be troublesome..
            device = TTS.getAllOutput_TTS2()
            sd.default.samplerate = device[output]
            sd.default.device = output + ", MME"

        x, sr = a2n.audio_from_file(
            rf"{appdata}/TouchPortal/Plugins/WinTools/speech.wav")
        sd.play(x, sr, blocking=False)
