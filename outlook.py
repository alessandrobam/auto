import subprocess
import time
import pyautogui
import utils
import pars
import excel
import pyperclip
import os


outlook = "https://outlook.office.com/owa/"
verseCal ="https://mail.notes.na.collabserv.com/verse#/calendar"

def start(wait):
    subprocess.Popen(pars.firefoxApp  + " " + outlook)
    time.sleep(wait) 

def login():
    pyautogui.click (937,569)
    time.sleep(7)
    pyautogui.typewrite("MARIAAX2")
    pyautogui.press("tab")
    pyautogui.typewrite(pars.keypass2)
    pyautogui.press("enter")

def getMeThere():
    start(4)
    login()