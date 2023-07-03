import TouchPortalAPI as TP
from TouchPortalAPI.logger import Logger
# from argparse import ArgumentParser
# from threading import Thread
# import sys
# from time import sleep, time
# import pyperclip
# import pygetwindow as gw
# import pyautogui
# from pyvda import AppView, VirtualDesktop, get_virtual_desktops
# from screeninfo import get_monitors


## Local Imports
from p_imports import *
from TPPEntry import PLUGIN_ID, TP_PLUGIN_STATES, TP_PLUGIN_ACTIONS, TP_PLUGIN_INFO, TP_PLUGIN_CONNECTORS, PLUGIN_NAME
from util import SystemPrograms, Get_Windows, PLATFORM_SYSTEM
from util import win32gui, os
from powerplan import Powerplan
from tts import TTS
from screencapture import ScreenShot
import Macro





# Create the Touch Portal API client.
try:
    TPClient = TP.Client(
        pluginId=PLUGIN_ID,  # required ID of this plugin
        sleepPeriod=0.05,    # allow more time than default for other processes
        autoClose=True,      # automatically disconnect when TP sends "closePlugin" message
        checkPluginId=True,  # validate destination of messages sent to this plugin
        maxWorkers=4,        # run up to 4 event handler threads
        # do not spam TP with state updates on every page change
        updateStatesOnBroadcast=False,
    )
except Exception as e:
    sys.exit(f"Could not create TP Client, exiting. Error was:\n{repr(e)}")
# TPClient: TP.Client = None  # instance of the TouchPortalAPI Client, created in main()


g_log = Logger(name=PLUGIN_ID)




def updateStates():
    g_log.debug("Running update state")

    update_time_short = 0.5              # Setting short update time to every 0.5 seconds
    update_time_long = time() + 5  # Setting long update time to every 60 seconds

    macroStateId = TP_PLUGIN_STATES["macro state"]['id']
    macroPlayProfile = TP_PLUGIN_ACTIONS["macroPlayer"]['data']["macro profile"]['id']


    while TPClient.isConnected():
        sleep(update_time_short)
        
        if Macro.States.macroRecordThread != None and Macro.States.macroRecordThread.is_alive():
            TPClient.stateUpdate(macroStateId, "RECORDING")
            Macro.States.macro_recordState = True
        else:
            TPClient.stateUpdate(macroStateId, "NOT RECORDING")
            Macro.States.macro_recordState = False

        if Macro.States.macroPlayThread != None and Macro.States.macroPlayThread.is_alive():
            TPClient.stateUpdate(
                TP_PLUGIN_STATES["macro play state"]['id'], "PLAYING")
            Macro.States.macro_playState = True
        else:
            TPClient.stateUpdate(
                TP_PLUGIN_STATES["macro play state"]['id'], "NOT PLAYING")
            Macro.States.macro_playState = False

        #NOTE: CHECK ME -  We should not be checking/opening a json file every 0.5 seconds, we should only do this when the macro profile changes or on startup?

        # Update macro profile
        newProfileList = list(Macro.getMacroProfile().keys())
        # default true cuz it does not exist in dict so need to update it.
        if TPClient.choiceUpdateList.get(macroPlayProfile, True) != newProfileList:
            TPClient.choiceUpdate(macroPlayProfile, newProfileList)


        ### Check for Current Active Windows
        windows_active = get_current_windows()

        ## Updating only if the list has changed
        if (windows_active != TPClient.choiceUpdateList.get(PLUGIN_ID + ".screencapturewindow.window_name")):
            TPClient.choiceUpdate(
                PLUGIN_ID + ".screencapturewindow.window_name", windows_active)
            TPClient.stateUpdate(
                PLUGIN_ID + ".Windows.activeCOUNT", str(len(windows_active)))
            

        ## Updating only every 60 seconds below
        if update_time_long < time():
            if PLATFORM_SYSTEM == "Windows":
                ### Getting TTS Output Devices and Updating Choices
                voices = [voice.name for voice in TTS.getAllVoices()]
                tts_outputs = list(TTS.getAllOutput_TTS2().keys())
                if TPClient.choiceUpdateList.get(TP_PLUGIN_ACTIONS["TTS"]["data"]["output"]['id']) != tts_outputs:
                    TPClient.choiceUpdate(PLUGIN_ID + "act.TSS.output", tts_outputs)
                if TPClient.choiceUpdateList.get(TP_PLUGIN_ACTIONS["TTS"]["data"]["voices"]['id']) != voices:
                    TPClient.choiceUpdate(PLUGIN_ID + "act.TSS.voices", voices)

                ### Getting Virtual Desktop info
                number_of_active_desktops = len(get_virtual_desktops())
                currentVdNum = VirtualDesktop.current().number        
                if (vd_list := [str(x + 1) for x in range(number_of_active_desktops)]) and TPClient.choiceUpdateList.get(PLUGIN_ID + ".act.vd_appchanger.vd_index") != vd_list:
                    vd_list.extend(["next", "previous"])
                    TPClient.choiceUpdate(PLUGIN_ID + ".act.vd_appchanger.vd_index", vd_list)
                    TPClient.choiceUpdate(PLUGIN_ID + ".act.vd_switcher.vd_index", vd_list)
                    TPClient.choiceUpdate(PLUGIN_ID + ".act.vd_rename.current_vds", vd_list)
                    TPClient.choiceUpdate(PLUGIN_ID + ".act.vd_remove.vd_name", vd_list)
                    TPClient.stateUpdate(TP_PLUGIN_STATES["num VD"]["id"], str(number_of_active_desktops))
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVD"]["id"], str(currentVdNum))


            #### Universal Stuff
            # Check for Monitor info & Details
            list_monitor = check_number_of_monitors()
            if (list_monitor != TPClient.choiceUpdateList.get(PLUGIN_ID + ".screencapturedisplay.monitors_choice")):
                TPClient.choiceUpdate(
                    PLUGIN_ID + ".screencapturedisplay.monitors_choice", list_monitor)
                TPClient.choiceUpdate(
                    PLUGIN_ID + ".winsettings.monchoice", list_monitor)
                TPClient.choiceUpdate(
                    PLUGIN_ID + ".winsettings.primary_monitor_choice", list_monitor)

            ### Resetting the next update time    
            update_time_long = time() + 60  # Set the next update time to 60 seconds later
            
    g_log.debug("UpdateState func exited")


updateStateThread = Thread(target=updateStates)


def window_callback(hwnd, app_name):
    if win32gui.IsWindowVisible(hwnd) and app_name in win32gui.GetWindowText(hwnd):
        return hwnd
    return False


# TP Client event handler callbacks

# Initial connection handler
@TPClient.on(TP.TYPES.onConnect)
def onConnect(data):
    g_log.info(
        f"Connected to TP v{data.get('tpVersionString', '?')}, plugin v{data.get('pluginVersion', '?')}.")
    g_log.debug(f"Connection: {data}")

    if PLATFORM_SYSTEM == "Windows":
        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS['Powerplan']['data']['powerplanChoices']['id'], list(
            pplan.powerplans.keys()))

  #  TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["macroPlayer"]['data']["macro profile"]['id'], list(Macro.getMacroProfile().keys()))
    TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["Keyboard presser"]["data"]["keys"]["id"], sorted(
        pyautogui.KEYBOARD_KEYS[4:]))

    updateStateThread.start()



# Settings handler
@TPClient.on(TP.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    g_log.debug(f"Settings: {data}")


# Action handler
@TPClient.on(TP.TYPES.onAction)
def onAction(data):
    g_log.info(f"Action: {data}")
    if not (action_data := data.get('data')) or not (aid := data.get('actionId')):
        return

## NOTE: Clipboard Actions - may not be needed as TP has a built in clipboard in the new version
    if aid == TP_PLUGIN_ACTIONS['Clipboard']['id']:
        pyperclip.copy(action_data[0]['value'])


## NOTE: Keyboard & Mouse Actions
    if aid == TP_PLUGIN_ACTIONS['Hold Mouse button']['id']:
        try:
            if action_data[0]['value'].lower() == "hold":
                pyautogui.mouseDown(button=action_data[1]['value'].lower())
            elif action_data[0]['value'].lower() == "release":
                pyautogui.mouseUp(button=action_data[1]['value'].lower())
        except pyautogui.PyAutoGUIException:  # Shouldn't happen but who knows...
            g_log.error("Hold or Release mouse have invaild mouse button")

    # all data send from TP value is in string eg 0 is "0" which is True
    if aid == TP_PLUGIN_ACTIONS['Mouse click']['id']:
        pyautogui.click(button=action_data[0]['value'].lower(), clicks=int(
            action_data[1]['value']), interval=int(action_data[1]['value']))

    if aid == TP_PLUGIN_ACTIONS['move Mouse']['id']:
        for adata in [1, 2]:
            action_data[adata]['value'] = int(action_data[adata]['value'])

        try:
            action_data[3]['value'] = float(action_data[3]['value'])
        except:
            action_data[3]['value'] = 0.1

        if action_data[0]['value'].lower() == "moveto":
            pyautogui.moveTo(x=action_data[1]['value'], y=action_data[2]['value'], duration=float(
                action_data[3]['value']))
        elif action_data[0]['value'].lower() == "move":
            pyautogui.move(
                xOffset=action_data[1]['value'], yOffset=action_data[2]['value'], duration=action_data[3]['value'])

    if aid == TP_PLUGIN_ACTIONS['Mouse scrolling']['id']:
        mouseScroll(action_data[0]['value'], action_data[1]['value'])

    if aid == TP_PLUGIN_ACTIONS['Drag mouse']['id']:
        for adata in [1, 2]:
            action_data[adata]['value'] = int(action_data[adata]['value'])
        try:
            action_data[3]['value'] = float(action_data[3]['value'])
        except:
            action_data[3]['value'] = 0.1

        if action_data[0]['value'].lower() == "dragto":
            pyautogui.dragTo(x=action_data[1]['value'], y=action_data[2]['value'],
                             duration=action_data[3]['value'], button=action_data[4]['value'].lower())
        elif action_data[0]['value'].lower() == "drag":
            pyautogui.drag(xOffset=action_data[1]['value'], yOffset=action_data[2]['value'],
                           duration=action_data[3]['value'], button=action_data[4]['value'].lower())

## NOTE: Virtual Desktop Actions

    if aid == TP_PLUGIN_ACTIONS["VD create"]["id"]:
        try:
            created_vd = VirtualDesktop.create()
            if action_data[0]["value"] != "":
                created_vd.rename(action_data[0]["value"])
        except Exception as e:
            g_log.error("VD create failed", e)

    if aid == TP_PLUGIN_ACTIONS["VD rename"]["id"]:
        try:
            VirtualDesktop(int(action_data[0]["value"])).rename(
                action_data[1]["value"])
        except Exception as e:
            g_log.error("VD rename failed", e)

    if aid == TP_PLUGIN_ACTIONS["VD remove"]["id"]:
        try:
            VirtualDesktop(int(action_data[0]["value"])).remove()
        except Exception as e:
            g_log.error("VD remove failed", e)

    if aid == TP_PLUGIN_ACTIONS["VD switcher"]["id"]:
        try:
            if action_data[0]["value"].lower() == "next":
                current_vd = VirtualDesktop.current()
                VirtualDesktop(int(current_vd.number)+1).go()
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVD"]["id"], str(current_vd.number))
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVDName"]["id"], str(current_vd.name))

            if action_data[0]["value"].lower() == "previous":
                current_vd = VirtualDesktop.current()
                VirtualDesktop(int(current_vd.number)-1).go()
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVD"]["id"], str(current_vd.number))
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVDName"]["id"], str(current_vd.name))
            else:
                VirtualDesktop(int(action_data[0]['value'])).go()
                current_vd = VirtualDesktop.current()
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVD"]["id"], str(current_vd.number))
                TPClient.stateUpdate(TP_PLUGIN_STATES["CurrentVDName"]["id"], str(current_vd.name))

        except Exception as e:
            g_log.error("VD switcher failed", e)
        

    if aid == TP_PLUGIN_ACTIONS["VD app changer"]:
        if (action_data[0]["value"].lower() == "current"):
            try:
                AppView.current().move(VirtualDesktop(
                    int(action_data[1]["value"])))
            except Exception as e:
                g_log.error("VD app changer failed", e)

        else:
            try:
                app_hwnd = win32gui.EnumWindows(
                    window_callback, action_data[0]["value"])
                if (app_hwnd):
                    AppView(app_hwnd).move(VirtualDesktop(
                        int(action_data[1]["value"])))
            except Exception as e:
                g_log.error("VD app changer failed", e)
            
    if aid == TP_PLUGIN_ACTIONS["VD app pin"]:
        if action_data[1]["value"].lower() == "current":
            try:
                app_to_pin = AppView.current()
            except:
                app_to_pin = None
        else:
            app_hwnd = win32gui.EnumWindows(
                window_callback, action_data[1]["value"])
            if app_hwnd:
                app_to_pin = AppView(app_hwnd)

        if app_to_pin:
            pin_option = action_data[0]["value"]

            if app_to_pin.is_pinned():
                pin_option = "Unpin"
            else:
                pin_option = "pin"

            if pin_option == "Pin":
                app_to_pin.pin()
            elif pin_option == "Unpin":
                app_to_pin.unpin()

    if PLATFORM_SYSTEM == "Windows":
        if aid == TP_PLUGIN_ACTIONS['App launcher']['id']:
            sysProgram.start(action_data[1]['value'], action_data[0]['value'])

        if aid == TP_PLUGIN_ACTIONS['Powerplan']['id']:
            pplan.changeTo(action_data[0]['value'])

        if aid == TP_PLUGIN_ACTIONS["TTS"]['id']:
            Thread(target=TTS.TextToSpeech, args=(
                action_data[1]['value'],
                action_data[2]['value'],
                int(action_data[2]['value']),
                int(action_data[3]['value']),
                action_data[4]['value']
            )).start()

## NOTE: Macro Recorder / Player
    if aid == TP_PLUGIN_ACTIONS["MacroRecorder"]['id']:
        ## We need to make this non blocking to other things, also using ESCAPE to end the macro recording seems wrong..
        ## We could show a taskbar icon that shows the current state of the maco recorder.. problem is it may not be visible without clicking for it
        if not Macro.States.macro_recordState:
            Macro.States.macroRecordThread = Thread(
                target=Macro.record, args=(action_data[1]['value'], action_data[0]['value']))
            Macro.States.macroRecordThread.start()

    if aid == TP_PLUGIN_ACTIONS["macroPlayer"]['id']:
        if not Macro.States.macro_playState and action_data[0]['value'] in Macro.getMacroProfile():
            Macro.States.macroPlayThread = Thread(
                target=Macro.play, args=(action_data[0]['value'],))
            Macro.States.macroPlayThread.start()

                # depreceated if aid == TP_PLUGIN_ACTIONS["json Parser"]["id"]:
                # depreceated     result = jsonPathfinder(
                # depreceated         action_data[0]["value"], action_data[1]["value"])
                # depreceated     TPClient.createState(PLUGIN_ID + ".state.jsonresult." + action_data[2]['value'],
                # depreceated                          action_data[2]['value'], result, "Json parser result")

## NOTE: Keyboard Writer
    if aid == TP_PLUGIN_ACTIONS["Keyboard writer"]["id"]:
        interval = 0
        try:
            interval = float(action_data[1]['value'])
        except ValueError:
            pass
        pyautogui.write(action_data[0]['value'], interval)

    if aid == TP_PLUGIN_ACTIONS["Keyboard presser"]["id"]:
        key = action_data[1]['value']
        if action_data[0]["value"] == "Press key":
            pyautogui.press(key)
        elif action_data[0]["value"] == "Release key":
            pyautogui.keyUp(key)
        elif action_data[0]["value"] == "Hold key":
            pyautogui.keyDown(key)

## NOTE: Screen / Display Capture
    if aid == TP_PLUGIN_ACTIONS["Screen Capture Display"]["id"]:
        if data['data'][1]['value'] == "Clipboard":
            try:
                ScreenShot.screenshot_monitor(
                    monitor_number=data['data'][0]['value'], clipboard=True)
            except Exception as e:
                print(e)

        elif data['data'][1]['value'] == "File":
            try:
                afile_name = (data['data'][2]['value']) + \
                    "/" + (data['data'][3]['value'])
                ScreenShot.screenshot_monitor(
                    monitor_number=data['data'][0]['value'], filename=afile_name, clipboard=False)
            except Exception as e:
                print(e)

    if aid == TP_PLUGIN_ACTIONS["Screen Capture Window"]["id"]:
        if data['data'][0]['value']:
            if data['data'][4]['value'] == "Clipboard":

                if PLATFORM_SYSTEM == "Windows":
                    ScreenShot.screenshot_window(capture_type=int(
                        data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=True)

                if PLATFORM_SYSTEM == "Linux":
                    ScreenShot.screenshot_window_linux(
                        window_name=data['data'][0]['value'], clipboard=True)

            if data['data'][4]['value'] == "File":
                afile_name = data['data'][2]['value'] + \
                    "/" + data['data'][3]['value']

                if PLATFORM_SYSTEM == "Windows":
                    ScreenShot.screenshot_window(capture_type=int(
                        data['data'][1]['value']), window_title=data['data'][0]['value'], clipboard=False, save_location=afile_name)

                if PLATFORM_SYSTEM == "Linux":
                    ScreenShot.screenshot_window_linux(
                        window_name=data['data'][0]['value'], file_name=afile_name)

    if aid == TP_PLUGIN_ACTIONS["Screenshot Window Current"]["id"]:
        current_window_title = gw.getActiveWindowTitle()
        if current_window_title:
            ScreenShot.screenshot_window(3, current_window_title, data['data'][0]['value'] == "Clipboard", os.path.join(
                data['data'][1]['value'], data['data'][2]['value']))

    if aid == TP_PLUGIN_ACTIONS["Screen Capture Window WildCard"]["id"]:
        ScreenShot.screenshot_window(window_title=data['data'][0]['value'], clipboard=data['data'][3]['value']
                                     == "Clipboard", save_location=os.path.join(data['data'][1]['value'], data['data'][2]['value']))


def get_current_windows():
    windows_active = []

    if PLATFORM_SYSTEM == "Windows":
        windows_active = Get_Windows.get_windows_Windows_OS()

    if PLATFORM_SYSTEM == "Linux":
        windows_active = Get_Windows.get_windows_Linux()

    if PLATFORM_SYSTEM == "Darwin":
        pass

    return windows_active


def check_number_of_monitors():
    if PLATFORM_SYSTEM == "Windows":
        list_monitor_full = ScreenShot.get_monitors_Windows_OS()
        list_monitor_full.insert(0, "0: ALL MONITORS")

    elif PLATFORM_SYSTEM == "Linux" or PLATFORM_SYSTEM == "Darwin":
        monitors = get_monitors()
        list_monitor_full = []

        count = 1
        for x in monitors:
            list_monitor_full.append(str(count) + ": " + x.name)
            count += 1
        list_monitor_full.insert(0, "0: ALL MONITORS")

    return list_monitor_full


def mouseScroll(mousescroll, speed):
    speed = int(speed)
    if mousescroll in ["DOWN", "LEFT"]:  # Set direction
        speed = speed * -1

   # if reverse:
   #     speed = speed * -1

    if mousescroll in ["UP", "DOWN"]:
        mousescroll = win32con.MOUSEEVENTF_WHEEL if PLATFORM_SYSTEM == "Windows" else "scroll"
    elif mousescroll in ["LEFT", "RIGHT"]:
        mousescroll = win32con.MOUSEEVENTF_HWHEEL if PLATFORM_SYSTEM == "Windows" else "hscroll"

    if PLATFORM_SYSTEM == "Windows":
        win32api.mouse_event(mousescroll, 0, 0, speed)
    else:
        if mousescroll == "hscroll":
            pyautogui.hscroll(speed)
        elif mousescroll == "scroll":
            pyautogui.scroll(speed)


# Connector handler
@TPClient.on(TP.TYPES.onConnectorChange)
def onConnector(data):
    g_log.debug("onConnector", data)
    if data['connectorId'] == TP_PLUGIN_CONNECTORS['MouseSliderCon']['id']:
        if data['value'] != 51:
            mouseScroll("UP" if data['value'] > 51 else "DOWN" if data['data'][0]['value'] == "Up/Down" else "RIGHT" if data['value'] > 51 else "LEFT",
                        data['value'] * 10,
                        False if data['data'][2] == "False" else True)
            print(data)


# on hold handler
@TPClient.on(TP.TYPES.onHold_down)
def onHold(data):
    while True:
        sleep(0.01)
        if TPClient.isActionBeingHeld(TP_PLUGIN_ACTIONS['Mouse scrolling']['id']):
            scrollSpeed = abs(int(data['data'][1]['value']) * 1)
            mouseScroll(data['data'][0]['value'], scrollSpeed)
        elif TPClient.isActionBeingHeld(TP_PLUGIN_ACTIONS['Mouse click']['id']):
            pyautogui.click(button=data['data'][0]['value'],
                            clicks=int(data['data'][1]['value']),
                            interval=float(data['data'][2]['value']))
        else:
            break


# Action data select event
@TPClient.on(TP.TYPES.onListChange)
def onListChange(data):
    if PLATFORM_SYSTEM == "Windows" and data['listId'] == TP_PLUGIN_ACTIONS['App launcher']['data']['appType']['id']:
       # if data['listId'] == TP_PLUGIN_ACTIONS['App launcher']['data']['appType']['id']:
        if data['value'] == "Steam":
            program = sysProgram.steam
        elif data['value'] == "Microsoft":
            program = sysProgram.microsoft
        else:
            program = sysProgram.other

        TPClient.choiceUpdate(
            TP_PLUGIN_ACTIONS['App launcher']['data']['appChoices']['id'], list(program.keys()))


# Shutdown handler
@TPClient.on(TP.TYPES.onShutdown)
def onShutdown(data):
    g_log.info('Received shutdown event from TP Client.')
    # We do not need to disconnect manually because we used `autoClose = True`


# Error handler
@TPClient.on(TP.TYPES.onError)
def onError(exc):
   # g_log.error(f'Error in TP Client event handler: {repr(exc)}')

   # This on error consistantly stops me from finding an error because it gives no reference to where it  came from
    pass


# main
def main():
    global TPClient, g_log
    ret = 0

    # default log file destination
    logFile = f"./{PLUGIN_ID}.log"
    # default log stream destination
    logStream = sys.stdout

    parser = ArgumentParser(fromfile_prefix_chars='@')
    parser.add_argument("-d", action='store_true',
                        help="Use debug logging.")
    parser.add_argument("-w", action='store_true',
                        help="Only log warnings and errors.")
    parser.add_argument("-q", action='store_true',
                        help="Disable all logging (quiet).")
    parser.add_argument("-l", metavar="<logfile>",
                        help=f"Log file name (default is '{logFile}'). Use 'none' to disable file logging.")
    parser.add_argument("-s", metavar="<stream>",
                        help="Log to output stream: 'stdout' (default), 'stderr', or 'none'.")

    # his processes the actual command line and populates the `opts` dict.
    opts = parser.parse_args()
    del parser

    # trim option string (they may contain spaces if read from config file)
    opts.l = opts.l.strip() if opts.l else 'none'
    opts.s = opts.s.strip().lower() if opts.s else 'stdout'
    print(opts)

    # Set minimum logging level based on passed arguments
    logLevel = "INFO"
    if opts.q:
        logLevel = None
    elif opts.d:
        logLevel = "DEBUG"
    elif opts.w:
        logLevel = "WARNING"

    # set log file if -l argument was passed
    if opts.l:
        logFile = None if opts.l.lower() == "none" else opts.l
    # set console logging if -s argument was passed
    if opts.s:
        if opts.s == "stderr":
            logStream = sys.stderr
        elif opts.s == "stdout":
            logStream = sys.stdout
        else:
            logStream = None

    # Configure the Client logging based on command line arguments.
    # Since the Client uses the "root" logger by default,
    # this also sets all default logging options for any added child loggers, such as our g_log instance we created earlier.
    TPClient.setLogFile(logFile)
    TPClient.setLogStream(logStream)
    TPClient.setLogLevel(logLevel)

    # ready to go
    g_log.info(
        f"Starting {TP_PLUGIN_INFO['name']} v{TP_PLUGIN_INFO['version']} on {PLUGIN_NAME}.")

    try:
        TPClient.connect()
        g_log.info('TP Client closed.')
    except KeyboardInterrupt:
        g_log.warning("Caught keyboard interrupt, exiting.")
    except Exception:
        from traceback import format_exc
        g_log.error(f"Exception in TP Client:\n{format_exc()}")
        ret = -1
    finally:
        TPClient.disconnect()

    del TPClient

    g_log.info(f"{TP_PLUGIN_INFO['name']} stopped.")
    return ret


if __name__ == "__main__":
    if PLATFORM_SYSTEM == "Windows":
        sysProgram = SystemPrograms()
        pplan = Powerplan()

    sys.exit(main())
