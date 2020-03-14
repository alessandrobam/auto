import utils
import pyautogui
import time
import pars
import subprocess,os

saveDir = r"C:\Users\ALESSANDROAlves\Box\Plan & Build\Delivery\ITSM Extracts"



moreButton_x = 1772
downloadButton_offset_y = 57
downloadButton_x = 541

# region = (356, 622, 825, 679)
# region = (358, 456, 974, 1039)

def start():
    subprocess.Popen(pars.firefoxApp  + " " + pars.url_AbbotReporting)
    

def downloadFile(searchImg, Destination):

    location = utils.locateOnScreen(searchImg)
    pyautogui.click(moreButton_x, location.y, interval=1)
    pyautogui.click(downloadButton_x, location.y +downloadButton_offset_y , interval=3)

    pyautogui.hotkey("alt","s")
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("home")
    pyautogui.typewrite(os.path.join(saveDir , (utils.getStampedStr("{} - ", utils.YYmMM_DD))))
    time.sleep(2)
    pyautogui.press("enter")

def downloadFilesKTLO():
    start()
    utils.waitUntil("BaseReport.PNG")
    downloadFile("BaseReport.PNG", saveDir)
    downloadFile("Closed Tickets.PNG", saveDir)
