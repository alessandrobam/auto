import subprocess
import time
import pyautogui
import utils
import pars
import excel
import pyperclip
import os

class pyWebLocation:
        url = ""

class pyIPPF (pyWebLocation):
        url = "https://w3-03.ibm.com/services/ippf/protected/LA/taskManagementAllEntities/p/1548161440075/1913339847/mainaction.wss"
        # regionEmailHeader = (479,272,1427,259)
        

        def __init__(self):
            pass

        def start(self):
            subprocess.Popen(pars.firefoxApp  + " " + self.url)
            utils.login()

            
        
        def saveLaborHours(self):
            saveDir = r"C:\Users\ALESSANDROAlves\Box\Plan & Build\Governance\Companion Agreements\Brazil\04 - Actuals - Labor" 
            fileName = os.path.join(saveDir , utils.getStampedStr("{} - IPPF Actuals Extract.csv",utils.YYmMM_DD))
            utils.automate("IPPF Actual Hours Report",1,10)
            if os.path.isfile(fileName ): 
                os.remove(fileName) #deleting existing file
            utils.automate("IPPF Actual Hours Report",11,11)
            pyautogui.typewrite(fileName)
            time.sleep(2)
            pyautogui.press("enter")

            # Refresh PBReports
            excel.openPBR(pars.pbReportsLocal, utils.image(r"excel\files\PB Reports Brazil.xlsm.PNG"))
            utils.automate("Refresh PowerQueries",10,30)

            # Enviando pro Rene






        
def getActualsHours():
    ippf = pyIPPF() 
    ippf.start()
    ippf.saveLaborHours()
    