import subprocess
import pyautogui
from utils import *
import pars
import excel
import pyperclip
import os
import time

verseURL = "https://mail.notes.na.collabserv.com/verse?"
verseCalURL ="https://mail.notes.na.collabserv.com/verse#/calendar"

def start():
    # try:
        # utils.bring("E-mail - Mozilla Firefox")
    # except:
    subprocess.Popen(pars.firefoxApp  + " " + verseURL)
    
    utils.waitUntil(utils.img(r"verse\compose_button.PNG"), utils.img(r"sso\SSO_LogingScreen.PNG"))
    if utils.isItOn(utils.img(r"sso\SSO_LogingScreen.PNG")):
        utils.login()    

def newEmail(subject, distro,distroCC=""):
    # if utils.isItOn(utils.img(r"verse\CloseCurrentEdit.PNG")):
    #    utils.click(utils.waitUntil(utils.img(r"verse\CloseCurrentEdit.PNG")))
    #    time.sleep(1.5)
    
    utils.click(utils.waitUntil(utils.img(r"verse\compose_button.PNG")))
    utils.waitUntil(utils.img(r"verse\new_email_is_ready.PNG"))
    for i in distro:
        pyautogui.typewrite(i)
        pyautogui.press(";")
    pyautogui.hotkey("tab")
    if len(distroCC)  > 0:
        time.sleep(3)
        pyautogui.press("press")
        pyautogui.typewrite(distroCC)
        print(distroCC)
        pyautogui.hotkey("tab")
    pyautogui.hotkey("tab")
    pyautogui.typewrite(subject)
    pyautogui.hotkey("tab")
    
def newBodyLine(text):
    pyautogui.typewrite(text)
    pyautogui.press("enter")

def attachFile(fileName):
    if fileName:
        location = locateOnScreenWithinRegion(r"verse\attach_icon.PNG",(215,273, 800, 800))
        pyautogui.click(location)
        waitUntil(img("verse\open_button.PNG"))
        pyautogui.typewrite(fileName)
        time.sleep(2)
        pyautogui.press("enter")

def sendVerseEmail(subject, distro, body, attachment, pasteClipboard):
    start()
    newEmail(subject,distro)
    newBodyLine (body)
    if pasteClipboard:
        newBodyLine("")
        pyautogui.hotkey("ctrl","v")
    attachFile(attachment)

def saveAttachments(path):
    # this automation will donwload attachments from a in-screen email. If the list of attachements is big
    # this 
    lastFoundX = -1
    currentFoundX = 0

    region = (0,0,pyautogui.size().width, pyautogui.size().height)
    emailDate = getEmailDate()

    print(pyautogui.size().width)
    
    utils.bringFirefoxToFront()    
    
    while True:
        try:
            location = utils.locateOnScreenWithinRegion("left_corner_first_attached_file.PNG",region)
            pyautogui.moveTo(location.x-33, location.y)
            print("Downloading Attachement.....")
            region = (location.x+10,0,pyautogui.size().width, pyautogui.size().height)
            time.sleep(1)
            pyautogui.click()
            
            utils.automate("Firefox - Save File",9,9)
            pyautogui.hotkey("alt","s")            
            utils.automate("Firefox - Save File",10,100)

            pyautogui.hotkey("ctrl","c")
            fileName = os.path.join(path[0:path.find('"')], emailDate + pyperclip.paste()  )
            
            pyautogui.typewrite(fileName)
            # time.sleep(1)
            # pyautogui.press("enter")
            # time.sleep(1)
        except:  
            print("No attachement found")
            break
    

def savePSFile():
        utils.bringFirefoxToFront()
        time.sleep(1)
        #Preparing file Name
        pyautogui.doubleClick(799,349)
        pyautogui.hotkey("ctrl","c")
        subjectLine = pyperclip.paste()

        options = {"ME" : pars.ps_me_report,
                   "Weekly_ME_Extract" : pars.ps_me_report,
                   "Weekly_PD_Extract" : pars.ps_pd_report,
                   "Weekly_PR_Extract" : pars.ps_pr_report,
                   "Weekly_Status_Report_Submission" : pars.ps_subm_report
                  }

        fileName = os.path.join(options[subjectLine],  utils.getStampedStr("{} - " + subjectLine + ".xlsx",utils.YYmMM_DD))
        print (fileName)

        # Downloading the file
        location = utils.locateOnScreen("left_corner_first_attached_file.PNG")
        print(location)
        pyautogui.moveTo(location.x-33, location.y)
        time.sleep(1)
        pyautogui.click()
        time.sleep(2)
        pyautogui.hotkey("alt","s")
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.typewrite(fileName)

def saveCurrentEmailAsPDF(path):
    
        utils.bringFirefoxToFront()       
        fileName = utils.addPath("verse_envelop and 3 dots.png","imgPattern")
        emailDate = getEmailDate()
        utils.automate("Save Verse Email to PDF",20,50)
        time.sleep(4)
        fileName = os.path.join(path[0:path.find('"')], emailDate + "EMAIL_"  )
        print(fileName)

        pyautogui.typewrite(fileName)


def getEmailDate():
    utils.bringFirefoxToFront()    
    clicked  = utils.clickImg("verse\showMoreLabel.PNG",30,0)
    pyautogui.moveTo(clicked.x+38, clicked.y-30)
    pyautogui.drag(-150,0,duration=0.25, button="left")
    pyautogui.hotkey("ctrl","c")
    return convertDate(pyperclip.paste())
    

def convertDate(txtDate):
    txtsplit = txtDate.split()
    options = {
               "dez" : "12",
               "nov" : "11",
               "out" : "10",
               "set" : "09",
               "ago" : "08",
               "jun" : "07",
               "jul" : "06",
               "mai" : "05",
               "abr" : "04",
               "mar" : "03",
               "fev" : "02",
               "jan" : "01"
               }
    retorno =  time.strftime("%Ym") + options[txtsplit[3]] + "_" + txtsplit[1] + " - "    
    return retorno