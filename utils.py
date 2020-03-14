import datetime, calendar
import os
import pyautogui
from pathlib import Path
import time
import pyscreeze
import ntpath
import pars
import pywinauto
import sys
import utils
import ctypes
import subprocess

YYmMM = "%Ym%m"
YYmMM_DD = "%Ym%m_%d"
MonthDay = "%b %d"
MB_OK = 0
MB_OKCANCEL = 1
MB_YESNOCANCEL = 3
MB_YESNO = 4

IDOK = 0
IDCANCEL = 2
IDABORT = 3
IDYES = 6
IDNO = 7

def LastFriday():
    lastFriday = datetime.date.today() 
    oneday = datetime.timedelta(days=1)
    while lastFriday.weekday() != calendar.FRIDAY:
        lastFriday -= oneday
    return lastFriday

def today():
   return datetime.date.today() 

def prevMonth():
   return datetime.date.today() -  datetime.timedelta(days=30)


def getPrefixStr():
    return   datetime.date.today().strftime("%Ym%m" + " - ")

def deleteIfExists(fileName):
    if os.path.isfile(fileName ): 
        os.remove(fileName) #deleting existing file
 
def getStampedStr(fileName, stampFormat, lastFriday=False):
    
    if lastFriday:
            dateStr = LastFriday().strftime(stampFormat)
    else:
            dateStr = today().strftime(stampFormat)
    return (fileName.format(dateStr))

def expandPath(fileName, type):
        if type=="img":
                return os.path.join(pars.imgPatterns,fileName)
        else:
                return fileName


def img(fileName):
        print(os.path.join(pars.imgPatterns,fileName))
        return os.path.join(pars.imgPatterns,fileName)
        

def getEmailTemplate(fileName):
        fileName = os.path.join(pars.emailTemplates,fileName)
        with open(fileName, 'r') as file:
                return file.read()

def locateOnScreenUntilOFF(fileName):
    repeat = True
    while repeat == True:
       print ("Repeating")
       if locateOnScreen(fileName) == False:
          repeat = False
          print ("Not found. Continue with the automation")
       else:
           if msgbox("Automation Failed","Do you want it to try to continue?",4) == IDNO:
               sys.exit("Automation Aborted")
           else:
               repeat = True
    
def waitUntil(img1, img2="" , attempts = 50):
    repeat = True
    attempt = 0
    while repeat == True and attempt <= attempts:
       response = locateOnScreen(img1, "Search No " + str(attempt))
       response2 = locateOnScreen(img2, "Search No " + str(attempt))
       if response == False and response2 == False:
          repeat = True
          attempt = attempt + 1
       else:
          repeat = False 
    if response:
        return response
    else:
        return response2



def locateOnScreen(fileName, msg = ""):
        if msg !="":
            print(msg + " : " + fileName)

        fileName = os.path.join(pars.imgPatterns,fileName)
        # print (fileName)
        start_time = time.time()
        try:
            im2 = pyautogui.center(pyautogui.locateOnScreen(fileName))
            elapsed_time = time.time() - start_time
            print ("image recognition is complete: " + str(elapsed_time))
            return im2
        except:
            return False
            
            


def locateOnScreenWithinRegion(fileName, Region):
        print("searching image" + fileName)
        # fileName = os.path.join(pars.imgPatterns,fileName)
        fileName = addPath(fileName, "imgPattern")
        start_time = time.time()
        # print(Region)
        im2 = pyautogui.center(pyautogui.locateOnScreen(fileName, region=Region))
        elapsed_time = time.time() - start_time
        print ("image recognition is complete: " + str(elapsed_time))
        # pyautogui.click(im2)     
        return im2

def addPath(fileName, type):
        if type == "imgPattern":
                return os.path.join(pars.imgPatterns,fileName)




def relativePath(fileName):
        dir =  Path.cwd()
        return os.path.join(dir, fileName)
                
def moveToArch(fileName, archiveFolder):
        # 'os.rename()
        # archiveFolder = "Archive"
        # fileName = r"C:\Users\ALESSANDROAlves\Box\Abbott Latam\Abbott Latam - Workspace\BPCS ITSM Categorization\2019m05_28 - Uncategorized BPCS tickets.xlsx"
        baseName =  ntpath.basename(fileName)                   
        filePath = ntpath.dirname(fileName)   
        ArchivePath = os.path.join(filePath, archiveFolder)
        ArchivedfileName = os.path.join(ArchivePath,baseName)
        # print(fileName)
        # print(ArchivedfileName)
        # deleteIfExists(ArchivedfileName)
        os.rename(fileName,ArchivedfileName)

def moveFilesToArchive(fileName):
        mypath = ntpath.dirname(fileName)
        onlyfiles = [f for f in os.listdir(mypath) if os.path.isfile(os.path.join(mypath, f))]
        print (onlyfiles)
        for x in onlyfiles:
                moveToArch(os.path.join(mypath,x),"Archive")

def msgbox(title, msg, style):
    ##  Styles:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | No 
    ##  6 : Cancel | Try Again | Continue"
    return ctypes.windll.user32.MessageBoxW(0, msg, title, style + 4096)

def alert(msg):
    ##  Styles:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | No 
    ##  6 : Cancel | Try Again | Continue"
    return ctypes.windll.user32.MessageBoxW(0, msg, "alert", 0 + 4096)


def loginSSO(wait, isLoggedImage):
    isLogged = True
    try:
         utils.locateOnScreen(isLoggedImage)
    except:
        isLogged = False
    if not(isLogged):
        pyautogui.typewrite("abarbosa@br.ibm.com")
        pyautogui.press("tab")
        pyautogui.typewrite(pars.keypass)
        pyautogui.press("enter")
        time.sleep(wait) 
    print(isLogged)
 




def bringFirefoxToFront():
        app = pywinauto.application.Application().connect(title_re="E-mail - Mozilla Firefox")
        mozilla = app.window(best_match="E-mail - Mozilla FirefoxMozillaWindowClass")
        mozilla.SetFocus()


def extractFileName(fullName):
    head, tail = ntpath.split(fullName)
    return tail or ntpath.basename(head)

def extractPath(fullname):
    return '\\'.join(fullname.split('\\')[0:-1])


def bringWindowToFront(appName):
        app = pywinauto.application.Application().connect(title_re=appName, visible_only=True)
        mozilla = app.window(best_match=appName)
        mozilla.SetFocus()

def bring(windows_caption):
       utils.bringWindowToFront( windows_caption)

def takeScreenshot(ul_x, ul_y, br_x, br_y):
    snipApp = r'"C:\Windows\System32\SnippingTool.exe"'
    utils.subprocess.Popen(snipApp)
    time.sleep(1)
    pyautogui.hotkey("alt","n")
    time.sleep(1)
    pyautogui.moveTo(ul_x,ul_y)
    pyautogui.dragTo(br_x, br_y)  # drag mouse to XY
    pyautogui.hotkey("ctrl","c")
    pyautogui.hotkey("alt","f4")

def openURL(url):
   subprocess.Popen(pars.firefoxApp  + " " + url)
   time.sleep(4) 

def clickImg(fileName,x_offset=0, y_offset=0, region=0,  wait=0):
    if region != 0:
        target = utils.locateOnScreenWithinRegion(utils.expandPath(fileName,"img"),region)
    else:
        target = utils.locateOnScreen(utils.expandPath(fileName,"img"))

    print (target)
    pyautogui.click(target.x + x_offset, target.y +  y_offset)
    
    print (target.x+x_offset)
    print (target.y+ y_offset)
    return target

def getTimeStampStr(dateFormat, TodayOrFriday ):
    if TodayOrFriday == "today":
        fileStamp  = getStampedStr("{}",dateFormat,lastFriday=False)
    else:
        if TodayOrFriday == "last_friday":
            fileStamp  = getStampedStr("{}",dateFormat,lastFriday=True)
        else:
            fileStamp  = ""
    return fileStamp


def click(box,x_offset=0, y_offset=0, wait=0):
    time.sleep(wait)
    pyautogui.click(box.x + x_offset, box.y +  y_offset)
    print(box)


def findImage(fileName,x_offset, y_offset, region=0):
    if region != 0:
        target = utils.locateOnScreenWithinRegion(utils.expandPath(fileName,"img"),region)
    else:
        target = utils.locateOnScreen(utils.expandPath(fileName,"img"))
    return target

def isItOn(fileName):
    return utils.locateOnScreen(fileName) != False
        

def isSSORequired():
    utils.waitUntil(utils.expandPath("firefox\home_button.PNG","img"))
    r = utils.waitUntil(utils.expandPath("sso\SSO_LogingScreen.PNG","img"), attempts=5)
    return r != False

def image(file):
    return utils.expandPath("sso\SSO_LogingScreen.PNG","img")

def login():
    if isSSORequired():
        pyautogui.typewrite(pars.keypass)
        pyautogui.press("enter")
    else:
        print("login is not required")


def getArgument(file, type):
    
    def getLimit(file, start):
        le_dot =  file.find(".",3) 
        le_us =  file.find("_",start)  
        le = le_us if le_us !=-1 else le_dot
        return le
        
    us = file.find("_" + type.upper() ,3)
    
    if us != -1:
       le = getLimit (file, us+2 ) 
       return int(file[us+2:le])
    else:
        return 0    


def isWait(file):
    return file.upper().find("_WAIT") > -1

def isOptional(file):
    return file.upper().find("_OPTIONAL") > -1


def getIndex(file):
    return int(file[0:2])
    

def automate(process_name, first_step=1, last_step=100):
    processDir = r'C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\Processes'
    autoDir = os.path.join(processDir, process_name)
    files = []

    for r, d, f in os.walk(os.path.join(processDir, process_name)):
        for file in f:
            files.append(file)

    for f in files:
        i = getIndex(f)
        # print(first_step)
        # print(last_step)
        if i >= first_step and i <= last_step:
            
            
            x_offset = getArgument(f,"x")
            y_offset = getArgument(f,"y")
            sleep = getArgument(f,"s")

            # print("index: " + str(getIndex(f))  + " X_offset: " + str(x_offset) + " Y_offset: " + str(y_offset) + " sleep: " + str(sleep)  + " File Name: " + f)
            
            if not isOptional(f):
              r = utils.waitUntil(os.path.join(autoDir,f))
            else:
              r = utils.waitUntil(os.path.join(autoDir,f),attempts=5)
              print("optional: " + f)
            
            if r != False:
                if not isWait(f):
                   utils.click(r, x_offset, y_offset )
                else:
                   time.sleep(sleep)
                   print( r)
            else:
               if not isOptional(f):
                  msgbox("Error", "Automation failed. File not found: " + f,0)
                  sys.exit()
            

