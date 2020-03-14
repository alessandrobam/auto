import subprocess
import time
import pyautogui
import utils
import pars
import excel
import pyperclip
import os

def start(wait):
    subprocess.Popen(pars.notesApp)
    time.sleep(5)
    pyautogui.typewrite(pars.keypass3)
    pyautogui.press("enter")

def getMeThere():
    start (1)
