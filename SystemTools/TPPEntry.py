import platform
import os

__version__ = "3.2"

#PLUGIN_ID = "com.github.KillerBOSS2019.TouchPortal.plugin.WinTool"
PLUGIN_ID = "plugin.SystemTools"

PLATFORM_SYSTEM = platform.system()

if PLATFORM_SYSTEM == "Windows":
    plugin_name = "Windows"
    appdata = os.getenv("LOCALAPPDATA")
elif PLATFORM_SYSTEM == "Darwin":
    plugin_name = "MacOS"
    appdata = "./Document/TouchPortal/plugins/SystemTools"
elif PLATFORM_SYSTEM == "Linux":
    plugin_name = "Linux"
    appdata = os.getenv("HOME") + "/.config/TouchPortal/plugins/SystemTools"

PLUGIN_NAME = plugin_name + " Tools"

TP_PLUGIN_INFO = {
    'sdk': 6,
    'version': int(float(__version__) * 100),
    'name': PLUGIN_NAME,
    'id': PLUGIN_ID,
    "plugin_start_cmd": "%TP_PLUGIN_FOLDER%SystemTools\\SystemTools.exe",
    'configuration': {
        'colorDark': "#25274c",
        'colorLight': "#707ab5"
    },
    "doc": {
        "repository": "KillerBOSS2019:WinTools",
        "Install": "1. Download latest version of plugin for your system.\n2. Import downloaded tpp by click the gear button next to email/notification icon.\n3. If this is first plugin, you will need to restart TouchPortal for it to work."
    }
}





TP_PLUGIN_CATEGORIES = {
    "main": {
        'id': PLUGIN_ID + ".main",
        'name' : plugin_name + " Utility",
        'imagepath' : "icon-24.png"
    },
    "mouse": {
        "id": PLUGIN_ID + ".AdvancedMouse",
        "name": plugin_name + " Mouse",
        "imagepath": "icon-24.png"
    },
    "keyboard": {
        "id": PLUGIN_ID + ".advancedKeyboard",
        "name": plugin_name + " Keyboard",
        "imagepath": "icon-24.png"
    }
}

# Setting(s) for this plugin. These could be either for users to
# set, or to persist data between plugin runs (as read-only settings).
TP_PLUGIN_SETTINGS = {}



TP_PLUGIN_CONNECTORS = {
    # Can't make this work as right now.
    # "MouseSliderCon": {
    #     "category": "main",
    #     "id": PLUGIN_ID + ".conor.Mouseslider",
    #     "name": "Slider Mouse scroll (HScroll/Scroll)",
    #     "format": "$[1]Scroll mouse at scroll speed $[2]x and Reverse$[3]",
    #     "label": "Mouse slider connector",
    #     "data": {
    #         "scrollChoice": {
    #             "id": PLUGIN_ID + ".conor.Mouseslider.scrollChoice",
    #             "type": "choice",
    #             "label": "scroll choice",
    #             "default": "Forward/Backward",
    #             "valueChoices": ["Forward/Backward", "Up/Down"]
    #         },
    #         'scrollSpeed': {
    #             'id': PLUGIN_ID + ".conor.Mouseslider.scrollChoice.speed",
    #             'type': "number",
    #             'label': "scroll speed",
    #             'default': 20,
    #             "minValue": 1,
    #             "maxValue": 120
    #         },
    #         "scrollReverse": {
    #             "id": PLUGIN_ID + ".conor.Mouseslider.reverse",
    #             "type": "choice",
    #             "label": "reverse",
    #             "default": "False",
    #             "valueChoices": ["True", "False"]
    #         },
    #     }
    # }
}




# Action(s) which this plugin supports.
TP_PLUGIN_ACTIONS = {
    'Hold Mouse button': {
        'category': "mouse",
        'id': PLUGIN_ID + ".act.holdMouse",
        'name': "Mouse Hold/Release mouse button",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        'format': "$[1]mouse button$[2]",
        'data': {
            'mouseState': {
                'id': PLUGIN_ID + ".act.holdMouse.state",
                'type': "choice",
                'label': "mouse state",
                'default': "Hold",
                'valueChoices': [
                    'Hold',
                    'Release'
                ]
            },
            'Buttonchoices': {
                'id': PLUGIN_ID + ".act.holdMouse.buttonChoices",
                'type': "choice",
                'label': "button choices",
                'default': "Left",
                'valueChoices': [
                    'Left',
                    'Middle',
                    'Right'
                ]
            },
        }
    },

    'Mouse click': {
        'category': "mouse",
        'id': PLUGIN_ID + ".act.mouseclick",
        'name': "Mouse Click",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        'format': "$[1]Click$[2]Times with interval$[3]",
        'data': {
            'mouseButton': {
                'id': PLUGIN_ID + ".act.mouseclick.buttonoptions",
                'type': "choice",
                'label': "mouse state",
                'default': "Left",
                'valueChoices': [
                    'Right',
                    'Middle',
                    'Left',
                ]
            },
            'clickTimes': {
                'id': PLUGIN_ID + ".act.mouseclick.clickTimes",
                'type': "number",
                'label': "click times",
                'allowDecimals': False,
                'default': 3,
                'minValue': 1
            },
            'clickInterval': {
                'id': PLUGIN_ID + ".act.mouseclick.clickInterval",
                'type': "number",
                'label': "click interval",
                'allowDecimals': True,
                'default': 1,
                'minValue': 0
            },
        }
    },

    'Mouse scrolling': {
        'category': "mouse",
        'id': PLUGIN_ID + ".act.mousescroll",
        'name': "Mouse scroll",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        'format': "Scroll Mouse$[1]by$[2]ticks",
        'data': {
            'mouseButton': {
                'id': PLUGIN_ID + ".act.mousescroll.scrolloptions",
                'type': "choice",
                'label': "mouse state",
                'default': "UP",
                'valueChoices': [
                    'UP',
                    'DOWN',
                    'LEFT',
                    'RIGHT',
                ]
            },
            'scroll ticks': {
                'id': PLUGIN_ID + ".act.mousescroll.ticks",
                'type': "number",
                'label': "scroll ticks",
                'allowDecimals': False,
                'default': 120,
                'minValue': 1
            }
        }
    },

    'move Mouse': {
        'category': "mouse",
        'id': PLUGIN_ID + ".act.mouseMovefunc",
        'name': "Move mouse",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "Mouse $[1]X$[2]Y$[3]in duration$[4]",
        'data': {
            'mouseState': {
                'id': PLUGIN_ID + ".act.mouseMovefunc.state",
                'type': "choice",
                'label': "mouse state",
                'default': "MoveTo",
                'valueChoices': [
                    'MoveTo',
                    'Move'
                ]
            },
            'moveXpos': {
                'id': PLUGIN_ID + ".act.mouseMovefunc.moveXPos",
                'type': "text",
                'label': "mouse move x pos",
                'default': "10",
            },
            'moveYpos': {
                'id': PLUGIN_ID + ".act.mouseMovefunc.moveYPos",
                'type': "text",
                'label': "mouse move Y pos",
                'default': "10"
            },
            'move duration': {
                'id': PLUGIN_ID + ".act.mouseMovefunc.duration",
                'type': "text",
                'label': "move duration",
                'default': "0.1",
            }
        }
    },

    'Drag mouse': {
        'category': "mouse",
        'id': PLUGIN_ID + ".act.DragmouseMovefunc",
        'name': "Drag Mouse",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "$[1]X$[2]Y$[3]in duration$[4]with button $[5]",
        'data': {
            'mouseState': {
                'id': PLUGIN_ID + ".act.DragmouseMovefunc.state",
                'type': "choice",
                'label': "mouse state",
                'default': "DragTo",
                'valueChoices': [
                    'DragTo',
                    'Drag'
                ]
            },
            'moveXpos': {
                'id': PLUGIN_ID + ".act.DragmouseMovefunc.moveXPos",
                'type': "text",
                'label': "mouse move x pos",
                'default': "10",
            },
            'moveYpos': {
                'id': PLUGIN_ID + ".act.DragmouseMovefunc.moveYPos",
                'type': "text",
                'label': "mouse move Y pos",
                'default': "10"
            },
            'move duration': {
                'id': PLUGIN_ID + ".act.DragmouseMovefunc.duration",
                'type': "number",
                'label': "move duration",
                'default': "0.1",
            },
            'mouseButton': {
                'id': PLUGIN_ID + ".act.DragmouseMovefunc.button",
                'type': "choice",
                'label': "mouse button",
                'default': "Left",
                'valueChoices': [
                    'Left',
                    'Middle',
                    'Right'
                ]
            },
        }
    },

    "Clipboard": {
        'category': "main",
        'id': PLUGIN_ID + ".act.clipboard",
        'name': "Save Text to clipboard",
        'prefix': TP_PLUGIN_CATEGORIES['main']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "Save$[1]to clipboard",
        'data': {
            'clipboardtext': {
                'id': PLUGIN_ID + ".act.clipboard.text",
                'type': "text",
                'label': "save clipboard text",
                'default': ""
            }
        }
    },

    "MacroRecorder": {
        "category": "keyboard",
        "id": PLUGIN_ID + "act.macroRecorder",
        "name": "Macro recorder",
        "prefix": TP_PLUGIN_CATEGORIES['keyboard']['name'],
        "type": "communicate",
        "tryInline": True,
        "format": "$[1] Macro Recording and save to $[2]",
        "data": {
            "microState": {
                "id": PLUGIN_ID + ".act.macroRecorder.state",
                "type": "choice",
                "label": "macro state",
                "default": "Toggle",
                "valueChoices": [
                    "Start",
                    "Stop",
                    "Toggle"
                ]
            },
            "microProfile": {
                "id": PLUGIN_ID + ".act.macroProfile",
                "type": "text",
                "label": "macro Profile",
                "default": "Main",
            }
        }
    },

    "macroPlayer": {
        "category": "keyboard",
        "id": PLUGIN_ID + "act.macroPlayer",
        "name": "Micro Player",
        "prefix": TP_PLUGIN_CATEGORIES['keyboard']['name'],
        "type": "communicate",
        "tryInline": True,
        "format": "Using macro profile$[1] and play$[2]",
        "data": {
            "macro profile": {
                "id": PLUGIN_ID + ".act.macroPlayer.profile",
                "type": "choice",
                "label": "macro profile",
                "default": "",
                "valueChoices": []
            },
            "macroPlayOpt": {
                "id": PLUGIN_ID + ".act.macroPlayer.playopt",
                "type": "choice",
                "label": "play opt",
                "default": "All",
                "valueChoices": [
                    "Keyboard",
                    "Mouse",
                    "All"
                ]
            }
        }
    },
    
    "Keyboard writer": {
        'category': "keyboard",
        'id': PLUGIN_ID + ".act.keyboardwriter",
        'name': "Write Text",
        'prefix': TP_PLUGIN_CATEGORIES['keyboard']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "Write$[1]with interval$[2]for every character",
        'data': {
            'text': {
                'id': PLUGIN_ID + ".act.keyboardwriter.text",
                'type': "text",
                'label': "text to write",
                'default': ""
            },
            'delay': {
                'id': PLUGIN_ID + ".act.keyboardwriter.delay",
                'type': "text",
                'label': "delay",
                'default': ""
            }
        }
    },
    "Keyboard presser": {
        'category': "keyboard",
        'id': PLUGIN_ID + ".act.keyboardpresser",
        'name': "Key Control",
        'prefix': TP_PLUGIN_CATEGORIES['keyboard']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "$[1]$[2]",
        'data': {
            "press options": {
                "id": PLUGIN_ID + ".act.keyboardpresser.options",
                "type": "choice",
                "label": "press options",
                "default": "Press key",
                "valueChoices": [
                    "Hold key",
                    "Release key",
                    "Press key"
                ]
            },
            'keys': {
                'id': PLUGIN_ID + ".act.keyboardpresser.keys",
                'type': "choice",
                'label': "key choices",
                'default': " ",
                "valueChoices": []
            }
        }
    },
    "json Parser": {
        'category': "main",
        'id': PLUGIN_ID + ".act.jsonparser",
        'name': "Json Parser",
        'prefix': TP_PLUGIN_CATEGORIES['main']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "With json$[1]to get$[2]and save to$[3]",
        'data': {
            "json data": {
                "id": PLUGIN_ID + ".act.jsonparser.jsondata",
                "type": "text",
                "label": "json data",
                "default": "",
            },
            'json path': {
                'id': PLUGIN_ID + ".act.keyboardpresser.jsonpath",
                'type': "text",
                'label': "json path",
                'default': "",
            },
            'save result': {
                'id': PLUGIN_ID + ".act.keyboardpresser.result",
                'type': "text",
                'label': "save result",
                'default': "",
            }
        }
    },
    
    "Screen Capture Display": {
        'category': "main",
        "id": PLUGIN_ID + ".screencapture.full.file",
        "name": "CAPTURE:  Display to File / Clipboard",
        "prefix": "plugin",
        "type": "communicate",
        "tryInline": True,
        "description": "Capture Display to Clipboard OR File",
        "format": "Display # $[monitors_choice] to $[file_clipboard_choice] to $[path] and $[name]",
        "data": [
          {
            "id": PLUGIN_ID + ".screencapture.monitors_choice",
            "type": "choice",
            "label": "choice",
            "default": "",
            "valueChoices": []
          },
          {
            "id": PLUGIN_ID + ".screencapture.file_clipboard_choice",
            "type": "choice",
            "label": "choice",
            "default": "Pick One",
            "valueChoices": [
                "Clipboard",
                "File"
            ]
          },
          {
            "id": PLUGIN_ID + ".screencapture.file.path",
            "type": "folder",
            "label": "folder"
          },
          {
            "id": PLUGIN_ID + ".screencapture.file.name",
            "type": "text",
            "label": "text",
            "default": ""
          }
        ]
      },

    "Screen Capture Window":{
        'category': "main",
        "id": PLUGIN_ID + ".screencapture.window.file",
        "name": "CAPTURE:  Window to File / Clipboard",
        "prefix": "plugin",
        "type": "communicate",
        "tryInline": True,
        "description": "Capture Window to Clipboard OR File",
        "format": "Window:$[window_name] Save:$[window_active_capture]  Directory->$[filepath] and file name ->$[filename] ",
        "data": [
          {
            "id": PLUGIN_ID + ".screencapture.window_name",
            "type": "choice",
            "label": "choice",
            "default": "",
            "valueChoices": []
          },
          {
            "id": PLUGIN_ID + ".screencapture.window_capture_type",
            "type": "choice",
            "label": "choice",
            "default": "3",
            "valueChoices": [
              "0",
              "1",
              "2",
              "3"
            ]
          },
          {
            "id": PLUGIN_ID + ".screencapture.filepath",
            "type": "folder",
            "label": "folder",
            "default": ""
          },
          {
            "id": PLUGIN_ID + ".screencapture.filename",
            "type": "text",
            "label": "text",
            "default": ""
          }, 
          {
            "id": PLUGIN_ID + ".screencapture.window_active_capture",
            "type": "choice",
            "label": "choice",
            "default": "Pick One",
            "valueChoices": [
                "Clipboard",
                "File"
            ]
          }
        ]
      },
    "Screen Capture Window WildCard": {
            'category': "main",
            "id": PLUGIN_ID + ".screencapture.window.file.wildcard",
            "name": "CAPTURE:  Window to File / Clipboard (*)",
            "prefix": "plugin",
            "type": "communicate",
            "tryInline": True,
            "description": "Capture Window by Name to Clipboard OR File ",
            "format": "Window Name*:$[window_name] Save:$[clipboard_file_choice]  Directory->$[path] and file name ->$[name] ",
            "data": [
            {
              "id": PLUGIN_ID + ".screencapture.window_name",
              "type": "text",
              "label": "text",
              "default": ""
            },
            {
              "id": PLUGIN_ID + ".screencapture.window_capture_type",
              "type": "choice",
              "label": "choice",
              "default": "3",
              "valueChoices": [
                "0",
                "1",
                "2",
                "3"
              ]
            },
            {
              "id": PLUGIN_ID + ".screencapture.file.path",
              "type": "folder",
              "label": "folder",
              "default": ""
            },
            {
              "id": PLUGIN_ID + ".screencapture.file.name",
              "type": "text",
              "label": "text",
              "default": ""
                },
            {
              "id": PLUGIN_ID + ".screencapture.clipboard_file_choice",
              "type": "choice",
              "label": "choice",
              "default": "Pick One",
              "valueChoices": [
                  "Clipboard",
                  "File"
              ]
            }
              ]
            },

    "Screenshot Window Current":{
            'category': "main",
            "id": PLUGIN_ID + ".window.current",
            "name": "CAPTURE:  Current Active Window to File / Clipboard",
            "prefix": "plugin",
            "type": "communicate",
            "tryInline": True,
            "description": "Capture CURRENT Active Window to Clipboard OR File ",
            "format": "Save to $[clipboard_file_choice]   IF file -> $[filepath] and $[filename].png",
            "data": [
          {
            "id": PLUGIN_ID + ".screencapture.window_current.clipboard_file_choice",
            "type": "choice",
            "label": "choice",
            "default": "Pick One",
            "valueChoices": [
                "Clipboard",
                "File"
            ]
          },
          {
            "id": PLUGIN_ID + ".screencapture.window_current.filepath",
            "type": "folder",
            "label": "folder",
            "default": ""
          },
          {
            "id": PLUGIN_ID + ".screencapture.window_current.filename",
            "type": "text",
            "label": "text",
            "default": ""
          }
        ]
        }
}


### Adding Windows Specific Actions
if plugin_name == "Windows": 

    TP_PLUGIN_ACTIONS["App launcher"] = {
        'category': "main",
        'id': PLUGIN_ID + ".act.advancedLauncher",
        'name': "Advanced Program launcher",
        'prefix': TP_PLUGIN_CATEGORIES['main']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "Launch$[1]$[2]",
        'data': {
            'appType': {
                'id': PLUGIN_ID + ".act.advancedLauncher.AppType",
                'type': "choice",
                'label': "App Type",
                'default': "Pick One",
                'valueChoices': [
                    'Steam',
                    'Microsoft',
                    'Other'
                ]
            },
            'appChoices': {
                'id': PLUGIN_ID + ".act.advancedLauncher.appChoices",
                'type': "choice",
                'label': "AppChoices",
                'default': "",
                'valueChoices': []
            }
        }
    }
    TP_PLUGIN_ACTIONS["Powerplan"] = {
        'category': "main",
        'id': PLUGIN_ID + ".act.Powerplan",
        'name': "Change Powerplan",
        'prefix': TP_PLUGIN_CATEGORIES['main']['name'],
        'type': "communicate",
        'tryInline': True,
        'format': "Change current powerplan to$[1]",
        'data': {
            'powerplanChoices': {
                'id': PLUGIN_ID + ".act.Powerplan.choices",
                'type': "choice",
                'label': "pplan choices",
                'default': "Balanced",
                'valueChoices': []
            }
        }
    }
    TP_PLUGIN_ACTIONS["TTS"] = {
        "category": "main",
        "id": PLUGIN_ID + ".act.TTS",
        "name": "Text to speech",
        "prefix": TP_PLUGIN_CATEGORIES['main']['name'],
        "type": "communicate",
        "description": "Play Text to Speech thru a specific Audio Output",
        "tryInline": True,
        "format": "Text$[1] Voice$[2] Volume$[3] Speech Rate$[4] Audio Output$[5]",
        "data": {
            'voices': {
                "id": PLUGIN_ID + "act.TSS.voices",
                "type": "choice",
                "label": "choice",
                "default": "",
                "valueChoices": []
            },
            'text': {
                "id": PLUGIN_ID + "act.TSS.text",
                "type": "text",
                "label": "text",
                "default": "Hello!"
            },
            'volume': {
                "id": PLUGIN_ID + "act.TSS.volume",
                "type": "number",
                "label": "tts volume",
                "allowDecimals": False,
                "default": "100", 
                "minValue":0,
                "maxValue":100
            },
            'rate': {
                "id": PLUGIN_ID + "act.TSS.rate",
                "type": "number",
                "label": "tts rate",
                "allowDecimals": False,
                "default": "175",
                "minValue":25,
                "maxValue":600
            },  
            'output': {
                "id": PLUGIN_ID + "act.TSS.output",
                "type": "choice",
                "label": "tts output choice",
                "default": "",
                "valueChoices":[]
            }
        }
    }


# Plugin static state(s). These are listed in the entry.tp file,
# vs. dynamic states which would be created/removed at runtime.
TP_PLUGIN_STATES = {
    'macro state': {
        'category': "keyboard",
        'id': PLUGIN_ID + ".state.macroState",
        'type': "text",
        'desc': "WinTool Keyboard show current macro state",
        'default': "NOT RECORDING"
    },
    'macro play state': {
        'category': "keyboard",
        'id': PLUGIN_ID + ".state.macroPlaystate",
        'type': "text",
        'desc': "WinTool Keyboard show macro play state",
        'default': "NOT PLAYING"
    },
}

# Plugin Event(s).
TP_PLUGIN_EVENTS = {}