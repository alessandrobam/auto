import subprocess
import time
import pyautogui
import utils
import pars
import excel

site = r"https://lmt.w3ibm.mybluemix.net"
dumpDir = r"C:\Users\ALESSANDROAlves\Box\Plan & Build\Governance\Companion Agreements\Brazil\04 - Actuals - Labor"

def start(wait):
    subprocess.Popen(pars.chromeApp  + " " +  site)
    time.sleep(wait) 

def login(wait):

    pyautogui.typewrite("abarbosa@br.ibm.com")
    pyautogui.press("tab")
    pyautogui.typewrite(pars.keypass)
    pyautogui.press("enter")
    time.sleep(wait) 
    

def runAndSaveReport():
    pyautogui.click(1028,210, interval=6)    #Click Montlhy
    pyautogui.click(1842,191, interval=2)    #Click Calendar
    Calendar_Region=(1673, 277, 180, 158)
    fileName = utils.expandPath("lmt_feb.png","img")
    x  = pyautogui.locateOnScreen(fileName,region=Calendar_Region, confidence=0.8)
    print(x)


    
def getLaborHourReport():
    start(8)
    login(10)
    runAndSaveReport()


getLaborHourReport()

# import time
# import pyautogui


# pyautogui.click(pyautogui.center(pyautogui.locateOnScreen(r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\MyPythonAutomation\imgPatterns\lmt_feb.png",region=(1673, 277, 180, 158) )))dir
