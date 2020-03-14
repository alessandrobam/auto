import pyautogui
import excel, utils, time, datetime, os, ntpath, sys, pars

dataFile  = r'"C:\Users\ALESSANDROAlves\Box\Ticket Data J&J Latam\Automation\Weekly Productive Decks\Productivity View.xlsx"'

class pyTicketStatusDeckUpate:
# 2019m10_12 - This python script will update the data sheet and send it over by email 
# to Jnj leadership
   dataSheet =r'C:\Users\ALESSANDROAlves\Box\Ticket Data J&J Latam\Status Report\Masters'
   pbrjnj = r'"C:\Users\ALESSANDROAlves\Box\Ticket Data J&J Latam\Automation\PB Reports JnJ.xlsm'
   
   def openDataSheet(self, file):
       
       try:
          utils.bringWindowToFront(ntpath.basename(file))
          print("it worked")
          time.sleep(1)
       except:
          excel.open( '"' +os.path.join(self.dataSheet, file) + '"')
        
        

   def ts(self):
        ts1 = time.time()
        return datetime.datetime.fromtimestamp(ts1).strftime('%Y-%m-%d %H:%M:%S')

   def openPBR(self):
    #    excel.open(self.pbrjnj, 10)
       utils.bringWindowToFront("PB Reports JnJ.xlsm")

   def bringDataSheet(self):
    #    excel.open(self.dataSheet, 10)
        
        utils.bringWindowToFront("DataForSlides.xlsx")

   def bringPBR(self):
    #    excel.open(self.pbrjnj, 10)
       utils.bringWindowToFront("PB Reports JnJ.xlsm")

   def cleanDataFile(self, file):
       self.openDataSheet(file)
       excel.selectRange("A56:M1000")
       pyautogui.press("delete")

       excel.selectRange("B29:P34")
       pyautogui.press("delete")

       excel.selectRange("U55:AF73")
       pyautogui.press("delete")



    



   def filteReport(self, cell, area):
       self.openPBR() 
       excel.gotoCell(cell)
       
       pyautogui.hotkey("alt","down")
       time.sleep(0.5)
       pyautogui.typewrite(area)
       time.sleep(0.5)
       pyautogui.press("tab")
       pyautogui.press("tab")
       pyautogui.press("tab")
       pyautogui.press("right")
       pyautogui.press("enter")


      #  pyautogui.hotkey("alt","down")
      #  time.sleep(0.5)
      #  pyautogui.press("tab")
      #  pyautogui.press("tab")
      #  pyautogui.press("right")
      #  pyautogui.typewrite(area)
      #  pyautogui.press("enter")



   def updateDataSlide(self,area, file):
       self.cleanDataFile(file)
       
      # Copying Weekly Throughput
       self.openPBR() 
       excel.runReportByName ("Weekly Throughput")
       if area!="ALL":
          self.filteReport("E10", area)

       excel.copyTable("D15",lr_offsetY=-1)

       self.openDataSheet(file)
       excel.gotoCell("A56")
       excel.pasteValues()

     
      # Copying Aging 
       self.openPBR() 
       if area=="ALL":
          excel.runReportByName ("Aging by Area")
       else:
          excel.runReportByName ("Aging by Module")
          self.filteReport("E10", area)
       
      # Copying Aging 

       excel.copyTable("D13", lr_offsetY=-1)
       
       self.openDataSheet(file)
       
       excel.gotoCell("B29")
       excel.pasteValues()
       excel.end("right",0,0)
       excel.deleteContent()
              
       self.openPBR() 
       
       if area!="ALL":
            excel.runReportByName ("Monthly Throughput")
            self.filteReport("E10", area)
            excel.copyTable("D13", lr_offsetY=-1)
            self.openDataSheet(file)
            excel.gotoCell("U54")
            excel.pasteValues()

       self.openPBR()  
       excel.runReportByName ("Monthly MTTR (INC, SR)")
       
       if area!="ALL":
            self.filteReport("E10", area)
            

       excel.copyTable("D13", lr_offsetY=-1)
       
       excel.copy() 
       self.openDataSheet(file)
       
       excel.gotoCell("Z54")
       excel.pasteValues()


def pasteValuesToSheet(sourceFile, targetfile, cell):
    excel.openDataSheet(targetfile)
    excel.gotoCell(cell)
    excel.pasteValues()
    excel.openPBR(sourceFile)
    pyautogui.press("esc")

def postMonthlyReport(reportName, targetCell):
    excel.openPBR(pars.pbReportsJNJ)
    excel.runReportByName(reportName)
    excel.copyTable("D13")
    pasteValuesToSheet(pars.pbReportsJNJ, dataFile, targetCell)

def updateMonthlySummary():
    excel.openDataSheet(dataFile)
    excel.nextTab()

    excel.deleteContent("A4:O26")
    excel.deleteContent("A29:O53")
    excel.deleteContent("A56:O180")

    postMonthlyReport("MTD Inflow All","A4")
    postMonthlyReport("MTD Closed All","D4")
    postMonthlyReport("MTD Backlog All","G4")
    postMonthlyReport("MTD Inflow Productive", "A29")
    postMonthlyReport("MTD Closed Productive", "D29")
    postMonthlyReport("MTD Backlog Productive", "G29") 
    postMonthlyReport("MTD SLA Predictor" , "A56")
    postMonthlyReport("MTD Breached Tickets" , "G56")



def updateDeckData():
   update = pyTicketStatusDeckUpate()
   update.updateDataSlide("ALL","DataForStatusReport - ALL.xlsx")
   update.updateDataSlide("BW","DataForStatusReport - BW.xlsx")
   update.updateDataSlide("Finance","DataForStatusReport - FINANCE.xlsx")
   update.updateDataSlide("Inbound","DataForStatusReport - INBOUND.xlsx")
   update.updateDataSlide("Outbound - BTB Latam","DataForStatusReport - OUTBOUND - BTB.xlsx")
   update.updateDataSlide("Outbound - Consumer Latam","DataForStatusReport - OUTBOUND - Consumer.xlsx")
   update.updateDataSlide("Security","DataForStatusReport - SECURITY.xlsx")


updateDeckData()
