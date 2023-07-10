## Windows Notification

from winotify import Notification, audio, Registry, PY_EXE, Notifier
import os


class WinToast:
    ## make an init
    def __init__(self):
        self.app_id = "WinTools-Notify"
        self.app_path = os.path.abspath(__file__)
        r = Registry(self.app_id, PY_EXE, self.app_path, force_override=True)
        self.notifier = Notifier(r)

        @self.notifier.register_callback
        def clear_notify():
            self.notifier.clear()


    def show_notification(self,title="", msg="", buttonText = "", buttonlink = "", sound = "Default", duration="short", icon=""):
        """ 
        Preparing the notification
        """
        if not os.path.exists(icon) or icon == "":
            icon = os.path.join(os.getcwd(), "src", "wintools_icon.png")

        toast = self.notifier.create_notification(
                             title=title,
                             msg=msg,
                             icon=icon,
                             duration = duration.lower(),
                             )

        ### we can allow multiple links
        if buttonText != "" and buttonlink != "":
            toast.add_actions(label=buttonText, 
                            launch=buttonlink)

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
            "LoopingCall2": audio.LoopingCall2,
            "LoopingCall3": audio.LoopingCall3,
            "LoopingCall4": audio.LoopingCall4,
            "LoopingCall5": audio.LoopingCall5,
            "LoopingCall6": audio.LoopingCall6,
            "LoopingCall7": audio.LoopingCall7,
            "LoopingCall8": audio.LoopingCall8,
            "LoopingCall9": audio.LoopingCall9,
            "LoopingCall10": audio.LoopingCall10,
            "Silent": audio.Silent,
        }
      #  print("building..")
        toast.set_audio(audioDic[sound], loop=False)

        toast.build()
        toast.show()



#WinToast(title="Test", msg="This is a test", sound="LoopingAlarm1", buttonText="Open Github", buttonlink="http://www.github.com")


