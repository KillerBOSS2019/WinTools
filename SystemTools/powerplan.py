##powerplan.py
from util import runWindowsCMD




class Powerplan:
    def __init__(self):
        self.currentPPlan = {}
        self.powerplans = {}
        self.GetPowerplan()

    def GetPowerplan(self):
        powerplanResult = runWindowsCMD("powercfg -List")

        for powerplan in powerplanResult.split("\n"):
            if ":" in powerplan:
                # from "Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced) *" to ["381b4222-f694-41f0-9685-ff5bb260df2e", "(Balanced) *"]
                parsedData = powerplan.split(":")[1].split()
                the_data = (parsedData[0])
                plan_name = (" ".join(parsedData[1:]))

                if "*" in plan_name:
                    plan_name = plan_name[plan_name.find(
                        "(") + 1: plan_name.find(")")]
                    self.powerplans[plan_name] = the_data

                    self.currentPPlan[plan_name] = the_data
                else:
                    plan_name = plan_name[plan_name.find(
                        "(") + 1: plan_name.find(")")]
                    self.powerplans[plan_name] = the_data
        return self.powerplans

    def changeTo(self, pplanName):
        if self.powerplans.get(pplanName, False):
            runWindowsCMD(f"powercfg.exe /S {self.powerplans[pplanName]}")
            return True
        return False