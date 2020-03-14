# import pars
import subprocess
import time
import pyautogui
import utils
import pars
import pyperclip
import urllib.request

site = r"https://performancemanager4.successfactors.com/sf/start?company=C0000160864P&_s.crb=AjHVmFY2iiToHshZ199pT7%252bt7lc%253d"



def start(wait):
    subprocess.Popen(pars.chromeApp  + " " +  site)
    time.sleep(wait) 

def openAll():
    emps = pars.my_employees.split(",")
    start(6)
    utils.loginSSO(20,"checkpoint_logged.PNG")
    pyautogui.click(98,771, interval=3)  #Face click
    pyautogui.click(1356,240, interval=2) #Take action
    pyautogui.click(1370,346, interval=2) 
    pyautogui.click(625,60, interval=3)
    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey("ctrl","c")
    emp_checkpoint = pyperclip.paste()
    pos = emp_checkpoint.find("userid=")
    for x in emps:
        url = emp_checkpoint[0:pos+7] + x  + emp_checkpoint[pos+7+9:]
        subprocess.Popen(pars.chromeApp  + " " +  url)

def openByName(name):
#     name = "SERGIO SERRANO FILHO"
    start(6)
    utils.loginSSO(20,"checkpoint_logged.PNG")
    pyautogui.click(98,771, interval=1)  
    pyautogui.hotkey("ctrl","f")
    pyautogui.typewrite(name)
    pyautogui.press("esc")
    pyautogui.press("enter")    
    time.sleep(1)
    pyautogui.click(1356,240, interval=1) #Take action
    pyautogui.click(1370,346) #Goal to plan
    
