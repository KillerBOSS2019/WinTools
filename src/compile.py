import os
import json
import shutil
import zipfile

with open("./src/entry.tp") as entry:
    entry = json.loads(entry.read())

startcmd = entry['plugin_start_cmd'].split("%TP_PLUGIN_FOLDER%")[1].split("\\")
filedirectory = startcmd[0]
fileName = startcmd[1]

os.makedirs(filedirectory)

for file in os.listdir('./src'):
    if file not in ["compile.py", "utils", "requirements.txt", "build", "dist", "Main.py"]:
        print("copying", file)
        shutil.copy("./src/"+file, filedirectory)
#shutil.move("/dist/WinTools.exe", filedirectory)

shutil.make_archive("WinTools.tpp", 'zip', "WinTools/"+filedirectory)