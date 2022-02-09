import os
import json
import shutil

with open("entry.tp") as entry:
    entry = json.loads(entry.read())

startcmd = entry['plugin_start_cmd'].split("%TP_PLUGIN_FOLDER%")[1].split("\\")
filedirectory = startcmd[0]
fileName = startcmd[1]
    
if os.path.exists(filedirectory):
    os.remove(os.path.join(os.getcwd(), "WinTools"))
else:
    os.makedirs("temp/"+filedirectory)


for file in os.listdir("."):
    if file not in ["compile.py", "utils", "requirements.txt", "build", "dist", "Main.py", "main.spec", "__pycache__", "temp"]:
        print("copying", file)
        shutil.copy(os.path.join(os.getcwd(), file), os.path.join("temp", filedirectory))
        
os.rename("dist\Main.exe", "dist\WinTools.exe")
shutil.copy(os.path.join(os.getcwd(), r"dist\WinTools.exe"), "temp/"+filedirectory)

shutil.make_archive(base_name="WinTools", format='zip', root_dir="temp", base_dir="WinTools") 
os.rename("WinTools.zip", "WinTools.tpp")
