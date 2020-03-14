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



class pyVerse (pyWebLocation):
        url = "https://mail.notes.na.collabserv.com/verse?"
        regionEmailHeader = (479,272,1427,259)
        

        def __init__(self):
            pass

        def start(self):
            subprocess.Popen(pars.firefoxApp  + " " + self.url)
            utils.waitUntil(utils.img("verse\compose_button.PNG"))
            time.sleep(1)
            
    
        
        def clickCompose(self, wait=1):
            utils.clickImg("verse\compose_button.PNG", x_offset=0,y_offset=0)
            time.sleep(wait)
        
        def clickCcCco(self, wait=1):
            utils.clickImg("verse\cc_button.PNG", x_offset=0,y_offset=0, region=self.regionEmailHeader)
            time.sleep(wait)

        
        def newEmail (self, subject, toEmails, ccEmails = ""):
            self.clickCompose(2)
            self.writeEmailList(toEmails)
            time.sleep(1)
            self.focusSubject(0)
            
            pyautogui.typewrite(subject)
            if ccEmails!="":
                self.clickCcCco()
                self.writeEmailList(ccEmails)
            time.sleep(1)    
            self.focusBody()


        def focusTo(self, wait=1):
            utils.clickImg(r"verse\to_label.PNG", x_offset=0,y_offset=0,region=self.regionEmailHeader)
            time.sleep(wait)
        #     pass
        
        def focusSubject(self, wait=1):
            utils.clickImg(r"verse\attach_icon.PNG", x_offset=0,y_offset=-40,region=self.regionEmailHeader)
            time.sleep(wait)
        
        
        def focusBody(self, wait=1):
            self.focusSubject()
            pyautogui.press("tab")
            


        def writeEmailList(self,distro):
            for i in distro:
                pyautogui.typewrite(i)
                pyautogui.press(";")
    
        def attachFile(self,fileName):
            if fileName:
                utils.clickImg(r"verse\attach_icon.PNG", x_offset=0,y_offset=0, region=self.regionEmailHeader)
                pyautogui.typewrite(fileName)
                time.sleep(2)
                pyautogui.press("enter")



            
                
                

        
                







verse = "https://mail.notes.na.collabserv.com/verse?"
verseCal ="https://mail.notes.na.collabserv.com/verse#/calendar"

def start():
    print("about to initiate a browser")
    subprocess.Popen(pars.firefoxApp  + " " + verse)
    utils.waitUntil(utils.img("verse\compose_button.PNG"))
    
    

def startCal():
    subprocess.Popen(pars.firefoxApp  + " " + verseCal)
    time.sleep(4) 


def newEmail(subject, distro, wait, distroCC=""):
    for i in distro:
        pyautogui.typewrite(i)
        pyautogui.press(";"
        )
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
    time.sleep(1)

def newBodyLine(text):
    pyautogui.typewrite(text)
    pyautogui.press("enter")

def attachFile(fileName):
    if fileName:
        location = utils.locateOnScreenWithinRegion(r"verse\attach_icon.PNG",(215,273, 800, 800))
        print(location)
        pyautogui.click(location)
        time.sleep(2)
        pyautogui.typewrite(fileName)
        time.sleep(2)
        pyautogui.press("enter")

def send():
    pyautogui.click(1855, 920) #Send the email



def newMeeting(meetingSubject, participants, day, timeofDay, duration, templateFile, webex):
        startCal()
        pyautogui.click(120,160) #clicking on Novo
        time.sleep(5)
        pyautogui.typewrite (meetingSubject)
        pyautogui.press("tab",presses=2)
        pyautogui.typewrite(participants)
        pyautogui.press("tab",presses=3)
        pyautogui.typewrite(day)
        pyautogui.press("tab",presses=1)
        pyautogui.typewrite(timeofDay)
        pyautogui.press("tab",presses=1)
        if duration != 60:
                steps = int((60 - duration)/15)
                print (steps)
                pyautogui.hotkey("alt","down")
                pyautogui.press("up",presses=steps)
                pyautogui.press("enter",presses=1)
        pyautogui.press("tab",presses=4)
        if webex=="Y":
            pyautogui.typewrite(pars.url_MyWebex)
        else:
            pyautogui.typewrite("Hortolandia - Room To Be Defined")

        pyautogui.click(190,655) #Activivating body
        time.sleep(1)
        a = utils.getEmailTemplate(templateFile)
        pyautogui.typewrite(a)


def createCalendarInvite():
        fileName = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\MyPythonAutomation\reference\meetings.xlsx"
        book = excel.open_workbook(fileName)
        for x in range(book.sheet_by_index(0).nrows):
                if x>0:
                        _day = excel.getXLDRValueAsText (book, x, 0)
                        _time = excel.getXLDRValueAsText (book, x,  1)
                        _duration = excel.getXLDRValueAsText (book, x,  2)
                        _subject = excel.getXLDRValueAsText (book, x,  3)
                        _tag = excel.getXLDRValueAsText (book,x,  4)
                        _participant = excel.getXLDRValueAsText (book,x, 5)
                        _webex = excel.getXLDRValueAsText (book,x, 6)
                        _template = excel.getXLDRValueAsText (book,x, 7)

                    
                        newMeeting(  meetingSubject = _subject.format(_tag),
                                participants = _participant,
                                day = _day.strftime(r"%d/%m/%Y") ,
                                timeofDay = _time.strftime(r"%H:%M"),
                                duration = int(_duration), templateFile= _template, webex= _webex)



def sendDirectMail():
        fileName = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\MyPythonAutomation\reference\directmail.xlsx"
        book = excel.open_workbook(fileName)
        # start(10)
        for x in range(book.sheet_by_index(0).nrows):
                if x>0:
                        _name = excel.getXLDRValueAsText (book, x,  0)
                        _email = excel.getXLDRValueAsText (book, x,  1)
                        _subject = excel.getXLDRValueAsText (book, x,  2)
                        _body = excel.getXLDRValueAsText (book, x,  3)
                        print(_body)
                        start(10)
                        newEmail(_subject,[_email],3)
                        newBodyLine(_body.format(_name))
                        time.sleep(5)

def sendDirectMailAmex():
    fileName = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\MyPythonAutomation\reference\directmail.xlsx"
    book = excel.open_workbook(fileName)
    
    verse = pyVerse()
   
    for x in range(book.sheet_by_index(0).nrows):
        if x>0:
            _name = excel.getXLDRValueAsText (book, x,  0)
            _email = excel.getXLDRValueAsText (book, x,  1)
            _cc = excel.getXLDRValueAsText (book, x,  2)
            _subject = excel.getXLDRValueAsText (book, x,  3)
            _body = excel.getXLDRValueAsText (book, x,  4)
        
            # verse.start(5)
            verse.newEmail(_subject,[_email],["cbakour@us.ibm.com","camolina@br.ibm.com"])
            

            newBodyLine(_name + ",")
            pyautogui.hotkey("ctrl","v")
            
            pyautogui.click(713,517)
            # time.sleep(3)

            pyautogui.typewrite(_body)
            pyautogui.press("enter")

            # time.sleep(5)


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
	

    
def getEmailDate():
    utils.bringFirefoxToFront("Mozilla Firefox")    
    clicked  = utils.clickImg("verse\showMoreLabel.PNG",30,0)
    pyautogui.moveTo(clicked.x+38, clicked.y-30)
    pyautogui.drag(-150,0,duration=0.25, button="left")
    pyautogui.hotkey("ctrl","c")
    return convertDate(pyperclip.paste())
    
    


    

def saveAttachments(path):
    # this automation will donwload attachments from a in-screen email. If the list of attachements is big
    # this 
    lastFoundX = -1
    currentFoundX = 0

    region = (0,0,pyautogui.size().width, pyautogui.size().height)
    emailDate = getEmailDate()

    print(pyautogui.size().width)
    
    utils.bringFirefoxToFront("Mozilla Firefox")    
    
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
        utils.bringFirefoxToFront("Mozilla Firefox")
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
    
        utils.bringFirefoxToFront("Mozilla Firefox")       
        fileName = utils.addPath("verse_envelop and 3 dots.png","imgPattern")
        emailDate = getEmailDate()
        utils.automate("Save Verse Email to PDF",1,50)
        time.sleep(4)
        fileName = os.path.join(path[0:path.find('"')], emailDate + "EMAIL_"  )
        print(fileName)

        pyautogui.typewrite(fileName)

# verse = pyVerse()
# verse.start(5)      
# verse.newEmail("Testes", pars.distro02, pars.distroADAM)
# verse.clickCcCco()
# verse.focusBody()

# verse.focusTo()
# verse.focusSubject()
# verse.writeEmailList(pars.distro01)
# verse.attachFile(r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\MyPythonAutomation\imgPatterns\verse\attach_icon.PNG")
# saveAttachments()
# saveCurrentEmailAsPDF(r"c:\temp")
# getEmailDate()
# convertDate()

