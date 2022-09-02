import sys
from time import sleep
import TouchPortalAPI as TP
from TouchPortalAPI.logger import Logger
from argparse import ArgumentParser
from threading import Thread
import pyperclip
import pyautogui

from TPPEntry import *
from util import SystemPrograms, Powerplan, TextToSpeech, getAllOutput_TTS2, getAllVoices
import Macro




if PLATFORM_SYSTEM == "Windows":
    import win32api
    import win32con



# Create the Touch Portal API client.
try:
    TPClient = TP.Client(
        pluginId = PLUGIN_ID,  # required ID of this plugin
        sleepPeriod = 0.05,    # allow more time than default for other processes
        autoClose = True,      # automatically disconnect when TP sends "closePlugin" message
        checkPluginId = True,  # validate destination of messages sent to this plugin
        maxWorkers = 4,        # run up to 4 event handler threads
        updateStatesOnBroadcast = False,  # do not spam TP with state updates on every page change
    )
except Exception as e:
    sys.exit(f"Could not create TP Client, exiting. Error was:\n{repr(e)}")
# TPClient: TP.Client = None  # instance of the TouchPortalAPI Client, created in main()

# Crate the (optional) global logger, an instance of `TouchPortalAPI::Logger` helper class.
# Logging configuration is set up in main().
g_log = Logger(name = PLUGIN_ID)

# macro_recordState = False
# macroRecordThread = None

# macro_playState = False
# macroPlayThread = None

def checkAllDataValue(data):
    return all([True if x['value'] else False for x in data])

if PLATFORM_SYSTEM == "Windows":
    sysProgram = SystemPrograms()
    pplan = Powerplan()


## Update states

def updateStates():
    g_log.debug("Running update state")


    macroStateId = TP_PLUGIN_STATES["macro state"]['id']
    macroPlayProfile = TP_PLUGIN_ACTIONS["macroPlayer"]['data']["macro profile"]['id']

    while TPClient.isConnected():
        sleep(0.1)
        if Macro.States.macroRecordThread != None and Macro.States.macroRecordThread.is_alive():
            TPClient.stateUpdate(macroStateId, "RECORDING")
            Macro.States.macro_recordState = True
        else:
            TPClient.stateUpdate(macroStateId, "NOT RECORDING")
            Macro.States.macro_recordState = False

        if Macro.States.macroPlayThread != None and Macro.States.macroPlayThread.is_alive():
            TPClient.stateUpdate(TP_PLUGIN_STATES["macro play state"]['id'], "PLAYING")
            Macro.States.macro_playState = True
        else:
            TPClient.stateUpdate(TP_PLUGIN_STATES["macro play state"]['id'], "NOT PLAYING")
            Macro.States.macro_playState = False



        if PLATFORM_SYSTEM == "Windows":
            """Getting TTS Output Devices and Updating Choices"""
            voices = [voice.name for voice in getAllVoices()]
            tts_outputs = list(getAllOutput_TTS2().keys())

            if TPClient.choiceUpdateList.get(TP_PLUGIN_ACTIONS["TTS"]["data"]["output"]) != tts_outputs:
                TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.output", tts_outputs)
            if TPClient.choiceUpdateList.get(TP_PLUGIN_ACTIONS["TTS"]["data"]["voices"]) != voices:
                TPClient.choiceUpdate("KillerBOSS.TP.Plugins.TextToSpeech.voices", voices)

        # Update macro profile
        newProfileList = list(Macro.getMacroProfile().keys())
        if macroPlayProfile in TPClient.choiceUpdateList and TPClient.choiceUpdateList[macroPlayProfile] != newProfileList:
            TPClient.choiceUpdate(macroPlayProfile, newProfileList)
    
    g_log.debug("UpdateState func exited")

updateStateThread = Thread(target=updateStates)


## TP Client event handler callbacks

# Initial connection handler
@TPClient.on(TP.TYPES.onConnect)
def onConnect(data):
    g_log.info(f"Connected to TP v{data.get('tpVersionString', '?')}, plugin v{data.get('pluginVersion', '?')}.")
    g_log.debug(f"Connection: {data}")

    if PLATFORM_SYSTEM == "Windows":
        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS['Powerplan']['data']['powerplanChoices']['id'], list(pplan.powerplans.keys()))

    TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["macroPlayer"]['data']["macro profile"]['id'], list(Macro.getMacroProfile().keys()))
    TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["Keyboard presser"]["data"]["keys"]["id"], pyautogui.KEYBOARD_KEYS[4:])

    updateStateThread.start()

# Settings handler
@TPClient.on(TP.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    g_log.debug(f"Settings: {data}")

# Action handler
@TPClient.on(TP.TYPES.onAction)
def onAction(data):
    g_log.debug(f"Action: {data}")
    # check that `data` and `actionId` members exist and save them for later use
    if not (action_data := data.get('data')) or not (aid := data.get('actionId')):
        return
    
    if aid == TP_PLUGIN_ACTIONS['Clipboard']['id']:
        pyperclip.copy(action_data[0]['value'])
    if aid == TP_PLUGIN_ACTIONS['Hold Mouse button']['id'] and checkAllDataValue(action_data):
        try:
            if action_data[0]['value'].lower() == "hold":
                pyautogui.mouseDown(button=action_data[1]['value'].lower())
            elif action_data[0]['value'].lower() == "release":
                pyautogui.mouseUp(button=action_data[1]['value'].lower())
        except pyautogui.PyAutoGUIException: # Shouldn't happen but who knows...
            g_log.error("Hold or Release mouse have invaild mouse button")

    # all data send from TP value is in string eg 0 is "0" which is True
    if aid == TP_PLUGIN_ACTIONS['Mouse click']['id'] and checkAllDataValue(action_data):
        pyautogui.click(button=action_data[0]['value'].lower(), clicks=int(action_data[1]['value']), interval=int(action_data[1]['value']))
    
    if aid == TP_PLUGIN_ACTIONS['move Mouse']['id'] and checkAllDataValue(action_data):
        for adata in [1,2]:
            action_data[adata]['value'] = int(action_data[adata]['value'])

        try:
            action_data[3]['value'] = float(action_data[3]['value'])
        except:
            action_data[3]['value'] = 0.1
            
        if action_data[0]['value'].lower() == "moveto":
            pyautogui.moveTo(x=action_data[1]['value'], y=action_data[2]['value'], duration=float(action_data[3]['value']))
        elif action_data[0]['value'].lower() == "move":
            pyautogui.move(xOffset=action_data[1]['value'], yOffset=action_data[2]['value'], duration=action_data[3]['value'])
    
    if aid == TP_PLUGIN_ACTIONS['Drag mouse']['id'] and checkAllDataValue(action_data):
        for adata in [1,2]:
            action_data[adata]['value'] = int(action_data[adata]['value'])
        try:
            action_data[3]['value'] = float(action_data[3]['value'])
        except:
            action_data[3]['value'] = 0.1

        if action_data[0]['value'].lower() == "dragto":
            pyautogui.dragTo(x=action_data[1]['value'], y=action_data[2]['value'], duration=action_data[3]['value'], button=action_data[4]['value'].lower())
        elif action_data[0]['value'].lower() == "drag":
            pyautogui.drag(xOffset=action_data[1]['value'], yOffset=action_data[2]['value'], duration=action_data[3]['value'], button=action_data[4]['value'].lower())

    if PLATFORM_SYSTEM == "Windows":
        if aid == TP_PLUGIN_ACTIONS['App launcher']['id'] and checkAllDataValue(action_data):
            sysProgram.start(action_data[1]['value'], action_data[0]['value'])


        if aid == TP_PLUGIN_ACTIONS['Powerplan']['id']:
            pplan.changeTo(action_data[0]['value'])


        if aid == TP_PLUGIN_ACTIONS["TTS"]['id']:
            Thread(target=TextToSpeech, args=(
                action_data[0]['value'],
                action_data[1]['value'],
                action_data[2]['value'],
                action_data[3]['value'],
                action_data[4]['value']
            )).start()


    if aid == TP_PLUGIN_ACTIONS["MacroRecorder"]['id'] and checkAllDataValue(action_data):
        if not Macro.States.macro_recordState:
            Macro.States.macroRecordThread = Thread(target=Macro.record, args=(action_data[1]['value'],))
            Macro.States.macroRecordThread.start()


    if aid == TP_PLUGIN_ACTIONS["macroPlayer"]['id'] and checkAllDataValue(action_data):
        if not Macro.States.macro_playState and action_data[0]['value'] in Macro.getMacroProfile():
            Macro.States.macroPlayThread = Thread(target=Macro.play, args=(action_data[0]['value'],))
            Macro.States.macroPlayThread.start()




    if aid == TP_PLUGIN_ACTIONS["json Parser"]["id"]:
        result = jsonPathfinder(action_data[0]["value"], action_data[1]["value"])
        TPClient.createState(PLUGIN_ID + ".state.jsonresult." + action_data[2]['value'],
        action_data[2]['value'], result, "Json parser result")

    

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





def mouseScroll(mousescroll, speed, reverse=False):
    if mousescroll in ["DOWN", "LEFT"]: # Set direction
        speed = speed * -1

    if reverse: speed = speed * -1

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
    if data['listId'] == TP_PLUGIN_ACTIONS['App launcher']['data']['appType']['id']:
        if data['value'] == "Steam":
            program = sysProgram.steam
        elif data['value'] == "Microsoft":
            program = sysProgram.microsoft
        else:
            program = sysProgram.other

        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS['App launcher']['data']['appChoices']['id'], list(program.keys()))



# Shutdown handler
@TPClient.on(TP.TYPES.onShutdown)
def onShutdown(data):
    g_log.info('Received shutdown event from TP Client.')
    # We do not need to disconnect manually because we used `autoClose = True`



# Error handler
@TPClient.on(TP.TYPES.onError)
def onError(exc):
    g_log.error(f'Error in TP Client event handler: {repr(exc)}')



## main
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
    if opts.q: logLevel = None
    elif opts.d: logLevel = "DEBUG"
    elif opts.w: logLevel = "WARNING"

    # set log file if -l argument was passed
    if opts.l:
        logFile = None if opts.l.lower() == "none" else opts.l
    # set console logging if -s argument was passed
    if opts.s:
        if opts.s == "stderr": logStream = sys.stderr
        elif opts.s == "stdout": logStream = sys.stdout
        else: logStream = None

    # Configure the Client logging based on command line arguments.
    # Since the Client uses the "root" logger by default,
    # this also sets all default logging options for any added child loggers, such as our g_log instance we created earlier.
    TPClient.setLogFile(logFile)
    TPClient.setLogStream(logStream)
    TPClient.setLogLevel(logLevel)

    # ready to go
    g_log.info(f"Starting {TP_PLUGIN_INFO['name']} v{TP_PLUGIN_INFO['version']} on {PLATFORM_SYSTEM}.")

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
    sys.exit(main())