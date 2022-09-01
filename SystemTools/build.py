from TouchPortalAPI.tppbuild import *
from TPPEntry import __version__, plugin_name
import subprocess

if plugin_name == "Linux":
    p = subprocess.Popen("sudo apt-get install python3-tk libportaudio2", stdout=subprocess.PIPE, shell=True)
    p = p.communicate()
    print(p)

    

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

ADDITIONAL_TPPSDK_ARGS = []

ADDITIONAL_PYINSTALLER_ARGS = [
    "--log-level=WARN"
]

if __name__ == "__main__":
    runBuild()