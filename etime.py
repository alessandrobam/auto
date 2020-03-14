import subprocess
import time
import pyautogui
import utils
import pars
import excel

site = r"https://w3-bz.ieb.ibm.com/hr/etime "
browser = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe "
# dumpDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\IBM Brazil\RH\People Management Misc\PB Reports CPG\ETIME Standby"

def start():
    subprocess.Popen(browser  + site)
    time.sleep(8) 

def login():
    # if pyautogui.locateOnScreen (r"C:\Temp\Etime_login_Box.PNG") is not None:
    pyautogui.typewrite("abarbosa@br.ibm.com")
    pyautogui.press("tab")
    pyautogui.typewrite(pars.keypass)
    pyautogui.press("enter")
    time.sleep(6) 
    pyautogui.click(971,111, interval=1)   #Enable flash
    pyautogui.click(309,163, interval=17)  #Run Flash

def runAndSaveReport():
    pyautogui.click(1721,485, interval=15) #Select Reports
    pyautogui.click(76,377, interval=1)    #Select favorite
    pyautogui.click(105,399, interval=4)   #Select report
    pyautogui.click(69,313, interval=12)   #Run report
    pyautogui.click(181,287, interval=5)   #Refresh 
    pyautogui.click(129,375,  clicks=3)    #get the report
    time.sleep(3)
    fileName = dumpDir + "\\" + utils.getPrefixStr() + "ETime Requests"
    utils.deleteIfExists(fileName + ".xls")
    pyautogui.typewrite(dumpDir + "\\" + utils.getPrefixStr() + "ETime Requests")
    # pyautogui.press("enter")

def getETimeRequests():
    # print (dumpDir)
    start()
    login()
    runAndSaveReport()
    excel.openPBR(pars.pbReportsCPG),



# getETimeRequests()

