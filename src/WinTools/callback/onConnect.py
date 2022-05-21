from pycaw.magic import MagicManager
from audio.WinAudioCallback import WinAudioCallBack
from constants import (TPClient, Tools)
from .onStartprepare import (activate_windows_setting, getAllVoices, getAllOutput_TTS2)
from utils.powerplan import get_powerplans
import pythoncom


def onConnect(data):
    """ Check for Update """
    try:
        github_check = Tools.updateCheck("KillerBOSS2019", "WinTools")
        plugin_version = str(data['pluginVersion'])
        plugin_version = plugin_version[:1] + "." + plugin_version[1:]
        if github_check[1:4] != plugin_version[0:3]:
            TPClient.showNotification(
                    notificationId="KillerBOSS.TP.Plugins.Update_Check",
                    title=f"WinTools v{github_check[1:4]} is available",
                    msg="A new Wintools Version is available and ready to Download. This may include Bug Fixes and or New Features",
                    options= [{
                    "id":"Download Update",
                    "title":"Click here to Update"
                    }])
    except:
        print("Something went wrong checking update")
        pass
    
    pplans = get_powerplans()
    pplanList = [item for item in pplans[0]]
    currentPPlan = pplans[1]
    TPClient.stateUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_current", currentPPlan)
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.powerplan_choice", pplanList)

    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.choice", activate_windows_setting())

    pythoncom.CoInitialize()
    ttsVoices = [voice.name for voice in getAllVoices()]
    ttsDevices = getAllOutput_TTS2()

    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.output", [ttsOutput for ttsOutput in ttsDevices])
    TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.voices", ttsVoices)

    """  Start audio listener """
    audioExemptionList = data['settings'][6]['Audio State Exemption List']
    try:
        MagicManager.magic_session(WinAudioCallBack, audioExemptionList)
        #MagicManager.add_magic_app
    except NotImplementedError as err:
        print(f"--------- Magic already in session!! ---------\n------{err}------")
    