from sys import platform

# Version string of this plugin (in Python style).
__version__ = "3.1"

# The unique plugin ID string is used in multiple places.
# It also forms the base for all other ID strings (for states, actions, etc).
PLUGIN_ID = "com.github.KillerBOSS2019.TouchPortal.plugin.WinTool"

## Start Python SDK declarations
# These will be used to generate the entry.tp file,
# and of course can also be used within this plugin's code.
# These could also live in a separate .py file which is then imported
# into your plugin's code, and be used directly to generate the entry.tp JSON.
#
# Some entries have default values (like "type" for a Setting),
# which are commented below and could technically be excluded from this code.
#
# Note that you may add any arbitrary keys/data to these dictionaries
# w/out breaking the generation routine. Only known TP SDK attributes
# (targeting the specified SDK version) will be used in the final entry.tp JSON.
##

if platform == "win32":
    plugin_name = "Windows"
elif platform == "darwin":
    plugin_name = "MacOS"
else:
    plugin_name = "Linux"

# Basic plugin metadata
TP_PLUGIN_INFO = {
    'sdk': 6,
    'version': int(float(__version__) * 100),  # TP only recognizes integer version numbers
    'name': plugin_name + " Tools",
    'id': PLUGIN_ID,
    # Startup command, with default logging options read from configuration file (see main() for details)
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

# This example only uses one Category for actions/etc., but multiple categories are supported also.
TP_PLUGIN_CATEGORIES = {
    "main": {
        'id': PLUGIN_ID + ".main",
        'name' : "WinTool Utility",
        'imagepath' : "icon-24.png"
    },
    "mouse": {
        "id": PLUGIN_ID + ".AdvancedMouse",
        "name": "WinTool Mouse",
        "imagepath": "icon-24.png"
    },
    "keyboard": {
        "id": PLUGIN_ID + ".advancedKeyboard",
        "name": "WinTool keyboard",
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
        # 'category' is optional, if omitted then this action will be added to all, or the only, category(ies)
        'category': "mouse",
        'id': PLUGIN_ID + ".act.holdMouse",
        'name': "Mouse Hold/Release mouse button",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        # 'format' tokens like $[1] will be replaced in the generated JSON with the corresponding data id wrapped with "{$...$}".
        # Numeric token values correspond to the order in which the data items are listed here, while text tokens correspond
        # to the last part of a dotted data ID (the part after the last period; letters, numbers, and underscore allowed).
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
        # 'category' is optional, if omitted then this action will be added to all, or the only, category(ies)
        'category': "mouse",
        'id': PLUGIN_ID + ".act.mouseclick",
        'name': "Mouse Click",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        # 'format' tokens like $[1] will be replaced in the generated JSON with the corresponding data id wrapped with "{$...$}".
        # Numeric token values correspond to the order in which the data items are listed here, while text tokens correspond
        # to the last part of a dotted data ID (the part after the last period; letters, numbers, and underscore allowed).
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
        # 'category' is optional, if omitted then this action will be added to all, or the only, category(ies)
        'category': "mouse",
        'id': PLUGIN_ID + ".act.mousescroll",
        'name': "Mouse scroll",
        'prefix': TP_PLUGIN_CATEGORIES['mouse']['name'],
        'type': "communicate",
        'tryInline': True,
        'hasHoldFunctionality': True,
        # 'format' tokens like $[1] will be replaced in the generated JSON with the corresponding data id wrapped with "{$...$}".
        # Numeric token values correspond to the order in which the data items are listed here, while text tokens correspond
        # to the last part of a dotted data ID (the part after the last period; letters, numbers, and underscore allowed).
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
        # 'category' is optional, if omitted then this action will be added to all, or the only, category(ies)
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
        # 'category' is optional, if omitted then this action will be added to all, or the only, category(ies)
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
    }
}

if platform == "win32": # add windows specific stuff
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
        "prefix": TP_PLUGIN_CATEGORIES['main']['name'],
        "type": "communicate",
        "description": "Play Text to Speech thru a specific Audio Output",
        "tryInline": True,
        "format": "Text[1] Voice[2] Volume[3] Speech Rate[4] Audio Output[5]",
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