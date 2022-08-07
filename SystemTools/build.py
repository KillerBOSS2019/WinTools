from TouchPortalAPI.tppbuild import *
from TPPEntry import __version__

PLUGIN_MAIN = "main.py"

PLUGIN_EXE_NAME = "SystemTools"

PLUGIN_EXE_ICON = ""

PLUGIN_ENTRY = "TPPEntry.py"

PLUGIN_ENTRY_INDENT = -1

PLUGIN_ROOT = "SystemTools"

PLUGIN_ICON = ""

OUTPUT_PATH = "./"

PLUGIN_VERSION = str(__version__)

ADDITIONAL_FILES = [
    "record.json"
]

ADDITIONAL_PYINSTALLER_ARGS = [
    "--log-level=WARN"
]

if __name__ == "__main__":
    runBuild()