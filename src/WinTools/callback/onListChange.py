from utils.virtualDesktop import vd_check
from utils.mouseCapture import get_cursor_choices
from constants import TPClient

def onListChange(data):
    if data['actionId'] == 'KillerBOSS.TP.Plugins.ChangeAudioOutput':
        pass # Need work here
        #try:
        #    updateDeviceOutput(data['value'])
        #except KeyError:
        #    pass
    if data['actionId'] == 'KillerBOSS.TP.Plugins.virtualdesktop.actions.move_window':
        vd_check()
        
    if data['actionId'] == 'KillerBOSS.TP.Plugins.winsettings.active.mouseCapture':
        thecursors = get_cursor_choices()
        thecursors.append("None")
        thecursors.append("LOAD MORE")
        if oldcursors != thecursors:
            print("not the same mmmk")
            TPClient.choiceUpdate("KillerBOSS.TP.Plugins.winsettings.active.mouseCapture.overlay_file", thecursors)
            oldcursors = thecursors