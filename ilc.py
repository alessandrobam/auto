import pyautogui
import pars
import subprocess
import time
import os


def start():
    os.chdir(r"C:\Program Files\IBM\BMS\ILC")
    subprocess.Popen(pars.ilcApp)
    time.sleep(5) 

def login():
    pyautogui.typewrite(pars.keypass)
    pyautogui.press("enter")
    time.sleep(8) 

def claimForVendor(forSerial):
    pyautogui.doubleClick(823,402)
    pyautogui.press("delete")
    pyautogui.typewrite(forSerial)
    pyautogui.press("enter",interval=3)
    pyautogui.press("tab", presses=2)
    pyautogui.press("delete")

def submit():
    pyautogui.click(490, 335)
    time.sleep(1)
    pyautogui.click(556, 335)

def avoidMissingClaim():
    start()
    login()
    # submit()
