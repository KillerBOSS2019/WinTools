import subprocess
from ctypes import windll


def out(command):
    systemencoding = windll.kernel32.GetConsoleOutputCP()
    systemencoding= f"cp{systemencoding}"
    output = subprocess.run(command, stdout=subprocess.PIPE, shell=True)
    result = str(output.stdout.decode(systemencoding))
    return result

def get_powerplans():
    pplans={}
    current = None
    for powerplan in out("powercfg -List").split("\n"):
        if ":" in powerplan:
            ParsedData = powerplan.split(":")[1].split()
            the_data = (ParsedData[0])
            plan_name = (" ".join(ParsedData[1:]))

            if "*" in plan_name:
                plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                pplans[plan_name]=the_data
                current = plan_name
            else:
                plan_name = plan_name[plan_name.find("(") + 1: plan_name.find(")")]
                pplans[plan_name]=the_data
    return (pplans, current)

def change_pplan(choice):
    the_thing=(get_powerplans())
    out(f"powercfg.exe /S {the_thing[choice]}")
