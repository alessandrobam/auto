import pyautogui
import time
import auto
import subprocess
import excel, verse, utils
import sys
import pars
import os
import pyperclip

RPT_CNB_BODY = "Non Billable Hours  Variance - By Resource"
RPT_CNB_ATTA = "Non Billable Hours  Variance - By Skill"
RPT_CWA_BODY = "ILC - Claim Waiting - Consolidated"
RPT_CWA_ATTA = "ILC - Claim Waiting"
RPT_LPR_BOTH = "All PRs"
RPT_MRO_BOTH = "Roster Status"
RPT_ITMS_CAT = "ITSM - Uncategorized Tickets - Latam"


def saveFile(reportNameFile, reportNamePaste, fileName):
    return excel.savePBRReport (reportNameFile, reportNamePaste, fileName )
       
def sendVerseEmail(subject, distro, body, attachment, wait, pasteClipboard):
    verse.start()
    verse.newEmail(subject,distro)
    time.sleep(2)
    verse.newBodyLine (body)
    if pasteClipboard:
        verse.newBodyLine("")
        pyautogui.hotkey("ctrl","v")
    time.sleep(2)
    verse.attachFile(attachment)
    
    # verse.send()

def sendCNBReport():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Staffing\Non Billable Hours\Weekly Update to Leadership (email)" 
    fileName = os.path.join(saveDir , utils.getStampedStr("CNB Weekly Report for {}.xlsx","%b %d", lastFriday=True))
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=True)
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName=RPT_CNB_ATTA)
    excel.saveReportToFile(fileName, excel.AutofitColumns)
    excel.runReportByName(reportName="Non Billable Hours Variance")
    pyautogui.press("enter")
    # time.sleep(10)
    excel.copyReportToClipBoard()
    subject = "Weekly CNB Variance Report - WE " + dateStr
    body = utils.getEmailTemplate("CNB.txt").format(dateStr)
    sendVerseEmail(subject, pars.distro01, body, fileName,11,True)


def sendIPPFActuals():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\AUTO - Weekly IPPF vs Actuals Email" 
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - IPPF vs Actuals Cost.xlsx",utils.YYmMM_DD, lastFriday=True))
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=True)
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName="IPPF Vs Actual Cost report")
    excel.saveReportToFile(fileName, excel.AutofitColumns)
    subject = "IPPF Cost vs Actual  - WE " + dateStr
    body = utils.getEmailTemplate("IPPF vs Cost Email.txt").format(dateStr)
    sendVerseEmail(subject, ["vinay.dhande@ibm.com"], body, fileName,10,True)



def sendCPHoursCostRevProfit():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\AUTO - Weekly CP - Hours, Cost, Rev, Profit report" 
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - CP - Hours, Cost, Rev, Profit.xlsx",utils.YYmMM_DD, lastFriday=True))
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=True)
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName="CP - Hours, Cost, Rev, Profit")
    excel.saveReportToFile(fileName, excel.AutofitColumns)
    subject = "CP - Hours, Cost, Rev, Profit (Confirmed)  - WE " + dateStr
    body = utils.getEmailTemplate("IPPF vs Cost Email.txt").format(dateStr)
    sendVerseEmail(subject, ["vinay.dhande@ibm.com"], body, fileName,10,True)


def send2019DMCApprovedPRs():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Client Support\Pipeline Management\Weekly PR Reports" 
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - 2019 DMC Approved PRs.xlsx",utils.YYmMM_DD))
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=False)
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName="All 2019 DMC Approved PRs")
    excel.saveReportToFile(fileName, excel.AutofitColumns)
    # excel.copyReportToClipBoard()
    subject = "2019 DMC Approved PRs - " + dateStr
    body = utils.getEmailTemplate("Weekly 2019 DMC PRs.txt").format(dateStr)
    sendVerseEmail(subject, ["vinay.dhande@ibm.com"], body, fileName,11,True)



def updateRosterWithCGPRates():
    formulaCol = "" 
    ratesRange = ""
  
    def formatCGPReport():
        nonlocal formulaCol
        nonlocal ratesRange
        
        # excel.AutofitColumns()
        
        limitCol = excel.getlimitY("A2")
        excel.right()
        excel.enterFormula("New Rate")
        formulaCol = excel.getCurrentCol()       
        
        limitRow = int(excel.getlimitX("A2",1))-1
        rowRange = "A{}:{}{}".format(limitRow,limitCol,limitRow) 
        excel.gotoCell(limitCol+str(limitRow))
        excel.right()
        excel.enterFormula(r'=LOOKUP(2,1/({}<>""),{})'.format(rowRange,rowRange))   
        excel.copy()
        excel.up()
        excel.endSelect("up")
        pyautogui.hotkey("shift","down")
        excel.paste()
        excel.gotoCell("a2")
        ratesRange = "$A$3:${}${}".format(excel.nextColLetter(limitCol),limitRow)

    def applyRatesToRoster():
        nonlocal formulaCol
        nonlocal ratesRange
        excel.open(pars.pbRoster, r"excel\files\Roster.xlsx.PNG")
        time.sleep(2)
        

        excel.gotoCell("A2000")
        excel.end( "up",0,0)
        limitRow = excel.getCurrentRow()
        rateCol = 9
        excel.gotoCell("i2")
        excel.addCol()
        
        print ("formulaCol: " + formulaCol)
        
        lookupFormula = "=IFERROR(VLOOKUP(A{},{}{},{},FALSE),J{})".format(limitRow,formulaFileName,ratesRange,int(excel.colNameToNum(formulaCol)),limitRow) 
        print(lookupFormula)
        excel.gotoCell("i" + limitRow)

        # fill up
        excel.enterFormula(lookupFormula)
        excel.copy()
        excel.up()
        excel.endSelect("up")
        excel.paste()



        excel.addCol()
        excel.right()
        excel.end("down")
        excel.left()
        excel.enterFormula("=IF(J{}>0,J{},K{})".format(limitRow,limitRow,limitRow))

        # fill up
        excel.copy()
        excel.up()
        excel.endSelect("up")
        excel.paste()

        excel.gotoCell("i2")
        pyautogui.hotkey("ctrl", "shift","down")
        excel.copy()


        excel.gotoCell("K2")
        excel.pasteValues()
        excel.left()
        excel.deleteCol()
        excel.left()
        excel.deleteCol()
        
        
        # excel.addCol()
        # lookupFormula = "=IFERROR(VLOOKUP(A{},{}{},{},FALSE),J{})".format(limitRow,formulaFileName,ratesRange,int(excel.colNameToNum(formulaCol)),limitRow) 



        # excel.deleteCol()
    
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\AUTO - Update Roster Rates Based on Latest CGP" 
    fileName = utils.getStampedStr("{} - CGP Rates by Emp.xlsx",utils.YYmMM_DD)
    fullFileName = os.path.join(saveDir , fileName)
    formulaFileName =  os.path.join("' " + saveDir , "["+ fileName + "]Sheet1'!")
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=False)
    excel.openPBR(pars.pbReportsABT )
    excel.runReportByName(reportName="CGP - Cost Rates by Employee")
    excel.saveReportToFile(fullFileName, formatCGPReport)
    
    applyRatesToRoster()
    
    
def sendClaimWatingReport():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Staffing\Non Billable Hours\Weekly Claim Waiting Variance Report - WE Month Day" 
    
    fileName = os.path.join(saveDir , utils.getStampedStr("Claim Waiting week of {}.xlsx","%b %d", lastFriday=True))
    dateStr = utils.getStampedStr("{}","%b %d",lastFriday=True)
    excel.openPBR(pars.pbReportsABT, "")

    excel.runReportByName(reportName=RPT_CWA_ATTA)
    excel.saveReportToFile(fileName, excel.AutofitColumns)
    excel.runReportByName(reportName=RPT_CWA_BODY)
    excel.copyReportToClipBoard()
    subject = "Weekly Claim Waiting Report - WE " + dateStr
    body = utils.getEmailTemplate("ClaimWaiting.txt").format(dateStr)
    sendVerseEmail(subject, pars.distro01, body, fileName,11,True)

def sendPRReports():
    def filteringLatin():
        excel.gotoCell("E10")
        time.sleep(5)
        pyautogui.hotkey("alt","down")

        pyautogui.press("down")
        pyautogui.press("right")
        pyautogui.press("right")
        pyautogui.press("t")
        pyautogui.press("enter")
        time.sleep(2)
        


    #  print(utils.msgbox("teste","teste",4))

    # sys.exit("")
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Client Support\Pipeline Management\Weekly PR Reports"   
    fileName = os.path.join(saveDir , utils.getStampedStr("Latin America PRs as of {}.xlsx","%b %d"))
    dateStr = utils.getStampedStr("{}","%b %d")
    
    excel.openPBR(pars.pbReportsABT,"")
    excel.runReportByName(reportName=RPT_LPR_BOTH)
    
    # filteringLatin()
    
    excel.saveReportToFile(fileName,excel.AutofitColumns)
    excel.copyReportToClipBoard()
    
    subject = "Latin America PR Report as of " + dateStr
    body = utils.getEmailTemplate("LatamPRs.txt")
    sendVerseEmail(subject, pars.distro02, body, fileName,11, True)


def sendLocalActuals():
    auto.sendPBReportsByEmail(pbrfile = pars.pbReportsLocal,
                     reportName = "ILC - Actual Hours",
                     emailSubject = "Actuals for Brazil local projects", 
                     distro = ["rrolivei@br.ibm.com","abarbosa@br.ibm.com"],
                     body = "Rene, segue as horas trabalhadas nos projetos locais. Por favor atualize o arquivo workbook. This is an automated email.",
                     saveExcelReportTo = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Finance\IPPF\AUTO - Publish the Weekly Actual Reports for Local Projects\Actuals for Local Projects.xlsx",
                     timeStampFileFormat = "last_friday",
                     IncludeReportInBodyReport = True)



def sendWrongContractFlag():
    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\2019m07 - Fixed Projects - Daily Report on Wrong Contract Type"   
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - Possibly Wrong Contract Type.xlsx",utils.YYmMM_DD))
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName="FP Projects Possibly with Incorrect Contract Type on CP")
    excel.saveReportToFile(fileName,excel.AutofitColumns)
    excel.copyReportToClipBoard()
    subject = "Daily Report - Possibly Wrong Contract Type in CP"
    body = utils.getEmailTemplate("Wrong Contract.txt")
    sendVerseEmail(subject, pars.distroTest, body, fileName,11, True)


def updateSLATrackerFile():
    excel.openPBR(pars.pbReportsJNJ)
    excel.runReportByName("SLA - Tracker")
    excel.copyTable("E14",ul_offsetX=1, lr_offsetX=-1)
    excel.openPBR(pars.slaTracker)
    excel.gotoCell("D5")
    excel.pasteValues()
    excel.openPBR(pars.pbReportsJNJ)
    excel.filterPivot("E9","All")
    excel.copyTable("E14",ul_offsetX=1, lr_offsetX=-1)
    excel.openPBR(pars.slaTracker)
    excel.gotoCell("H5")
    excel.pasteValues()
    excel.gotoCell("A17")


def pasteSLATrackerScreenShotToTheEmail():
    utils.takeScreenshot(41,236, 1119, 457)
    utils.bringFirefoxToFront()
    time.sleep(1)
    pyautogui.press("pageup")
    pyautogui.press("down",4)
    pyautogui.hotkey("ctrl","v")
    pyautogui.press("enter",2)
    pyautogui.typewrite("List of breached tickets:")
    pyautogui.press("delete",2)

def sendJNJBreachedTickets():
    saveDir = r"C:\Users\ALESSANDROAlves\Box\Ticket Data J&J Latam\Automation\00 - SLA Management\Daily Breached Tickets Report"   
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - Breached tickets in Current Tableau Month.xlsx",utils.YYmMM_DD))
    dateStr = utils.getStampedStr("{}","%b %d")
    
    excel.openPBR(pars.pbReportsJNJ,"")
    excel.runReportByName(reportName="Breached Incidents  - Current Month")
    excel.saveReportToFile(fileName,excel.AutofitColumns)
    excel.copyReportToClipBoard()
    subject = "JnJ Breached Tickets for the Current Tableau Month as of " + dateStr
    body = utils.getEmailTemplate("JNJ Breached Tickets.txt")
    sendVerseEmail(subject, ["lguerra@br.ibm.com","gomezhum@mx1.ibm.com","jrodrigo@mx1.ibm.com","jjrocha@mx1.ibm.com","ritwghos@in.ibm.com","abarbosa@br.ibm.com","adricff@br.ibm.com","acuriel@mx1.ibm.com"], body, fileName,20, True)
    # updateSLATrackerFile()
    # pasteSLATrackerScreenShotToTheEmail()

    



    
def sendMissingRoster():
    def none():
        excel.AutofitColumns()
        excel.setColumnSize("E", 30)
        excel.setColumnSize("F", 30)
        excel.setColumnSize("G", 30)
    
    def file_massage():
        # Highlighting the columns
        excel.gotoCell("H13")
        pyautogui.keyDown("shift")
        pyautogui.press("right")
        pyautogui.press("right")
        pyautogui.keyUp("shift")

        pyautogui.keyDown("alt")
        pyautogui.press("h")
        pyautogui.press("h")
        pyautogui.keyUp("alt")

        
        pyautogui.press("right",presses=5,interval=1)
        pyautogui.press("enter")

    saveDir = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\Weekly Missing Roster Report"   
    fileName = os.path.join(saveDir , utils.getStampedStr("Roster Status as of {}.xlsx","%b %d"))
    dateStr = utils.getStampedStr("{}","%b %d")
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName=RPT_MRO_BOTH)
    file_massage()
    excel.saveReportToFile(fileName, none)
    excel.copyReportToClipBoard()   
    subject = "Roster Status as of " + dateStr
    body = utils.getEmailTemplate("Missing Roster.txt")
    sendVerseEmail(subject, pars.distro01, body, fileName, 11, True)
    
    


def sendBPCSCategorizationReport():
    template =r'"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\CIC Brazil\Automation\F0001 - Create a automatic process for Ticket Categorization in BPCS\UncatTemplate.xlsx"'
    saveDir = r"C:\Users\ALESSANDROAlves\Box\Abbott Latam\Abbott Latam - Workspace\BPCS ITSM Categorization"   
    fileName = os.path.join(saveDir , utils.getStampedStr("{} - Uncategorized tickets.xlsx",utils.YYmMM_DD))
    dateStr = utils.getStampedStr("{}","%b %d")
    excel.openPBR(pars.pbReportsABT, "")
    excel.runReportByName(reportName=RPT_ITMS_CAT)
    # excel.saveReportToFile(fileName)
    
    excel.gotoCell("D14")                     #selecting data in the pivot table  
    excel.endSelect("down")                   #selecting data in the pivot table
    excel.endSelect("right")                  #selecting data in the pivot table
    pyautogui.hotkey("shift","left")          #selecting data in the pivot table
    excel.copy()                              #selecting data in the pivot table

    excel.open(template,4)
    excel.gotoCell("B2")    

    excel.pasteValues()
    excel.AutofitColumns()
    pyautogui.press("up")
    excel.endSelect("right")
    excel.activateAutoFilter()
    excel.gotoCell("A1")
    excel.setColumnSize("F",95)
    utils.moveFilesToArchive(fileName)
    excel.saveAs(fileName)
    excel.close()

    excel.runReportByName(reportName="ITSM - Uncategorized Tickets Summary - Latam")
    excel.findText("Grand Total",0)
    pyautogui.press("right")
    pyautogui.press("right")
    excel.copy()

    ticketCount = pyperclip.paste().replace('\n', '').replace('\r', '')
    excel.copyReportToClipBoard()
    subject = ("Weekly Ticket Categorization - There are {} tickets pending categorization as of " + dateStr).format(ticketCount)
    body = utils.getEmailTemplate("Categorized Tickets.txt")
    sendVerseEmail(subject, pars.distroBPCS+pars.distroADAM, body.format(ticketCount), "", 20,True)

pyautogui.FAILSAFE = True 


# updateSLATrackerFile()
# pasteSLATrackerScreenShotToTheEmail()