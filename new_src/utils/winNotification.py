import os

from winotify import PY_EXE, Notifier, Registry, audio

app_id = "WinTools-Notify"
app_path = os.path.abspath(__file__)
r = Registry(app_id, PY_EXE, app_path, force_override=True)
notifier = Notifier(r)

@notifier.register_callback
def clear_notify():
    notifier.clear()


def win_toast(atitle="", amsg="", buttonText = "", buttonlink = "", sound = "", aduration="short", icon=""):
    ### setting the base notification stuff
    if not os.path.exists(rf"{icon}") or icon == "": icon = os.path.join(os.getcwd(),"src\icon.png")

    toast = notifier.create_notification(
                         title=atitle,
                         msg=amsg,
                         icon=icon,
                         duration = aduration.lower(),
                         )
    
    ### we can allow multiple links
    if buttonText != "" and buttonlink != "":
        toast.add_actions(label=buttonText, 
                        launch=buttonlink)
      # toast.add_actions("Open Github", "https://github.com/versa-syahptr/winotify")
      # toast.add_actions("Quit app", "Kewl")
      # toast.add_actions("spam", "sweet")

    audioDic = {
        "Default": audio.Default,
        "IM": audio.IM,
        "Mail": audio.Mail,
        "Reminder": audio.Mail,
        "SMS": audio.SMS,
        "LoopingAlarm1": audio.LoopingAlarm,
        "LoopingAlarm2": audio.LoopingAlarm2,
        "LoopingAlarm3": audio.LoopingAlarm3,
        "LoopingAlarm4": audio.LoopingAlarm4,
        "LoopingAlarm5": audio.LoopingAlarm6,
        "LoopingAlarm6": audio.LoopingAlarm8,
        "LoopingAlarm7": audio.LoopingAlarm9,
        "LoopingAlarm8": audio.LoopingAlarm10,
        "LoopingCall1": audio.LoopingCall,
        "LoopingCall2": audio.LoopingCall,
        "LoopingCall3": audio.LoopingCall,
        "LoopingCall4": audio.LoopingCall,
        "LoopingCall5": audio.LoopingCall,
        "LoopingCall6": audio.LoopingCall,
        "LoopingCall7": audio.LoopingCall,
        "LoopingCall8": audio.LoopingCall,
        "LoopingCall9": audio.LoopingCall,
        "LoopingCall10": audio.LoopingCall,
        "Silent": audio.Silent,
    }

    toast.set_audio(audioDic[sound], loop=False)
         
    toast.build()
    toast.show()