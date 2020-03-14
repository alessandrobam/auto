import pyautogui
import time
import subprocess
import os
import pars
import sys
import datetime
import string
import excel
# Reading an excel file using Python 
import xlrd 
from utils import *
import pyperclip
from unittest import result





# pbReports = "\"C:\\Users\\ALESSANDROAlves\\Box\\Plan & Build\\Delivery\\PB Reports\\PB Reports 2019.xlsm\""

    
def goLastColWithData(Row, offsetX=0, offsetY=0):
    gotoCell("XFD"+ Row)
    end("left", offsetX, offsetY)
    # return getAddress()

def goLastRowWithData(Col,offsetX=0, offsetY=0):
    gotoCell(Col + "50000")
    end("up", offsetX, offsetY)
    # return getAddress()


def open(fileName, screenSignal=""):
    excel = subprocess.Popen(pars.excelApp + " " + fileName)
    if screenSignal!="":
       utils.waitUntil(utils.img(screenSignal),attempts=100)
    else:
        time.sleep(5)
    return excel 


def cellValue(cell=""):
    copy()
    if (cell==""): #get current cells
       cellValue = pyperclip.paste().replace('\n', '').replace('\r', '')
    else:
        gotoCell(cell)
        cellValue = pyperclip.paste().replace('\n', '').replace('\r', '')
    pyautogui.press("esc")
    return cellValue
    
    
def openDataSheet(file):
    openPBR(file,"")
    #    print("opening " + file)
    #    start_time = time.time()
    #    try:
    #        utils.bring(ntpath.basename(file))
    #        time.sleep(0.5)
    #    except:
    #         excel.open(file)
    #    elapsed_time = time.time() - start_time
    #    print("Open complete " + str(elapsed_time))


def nextTab():
    pyautogui.hotkey("ctrl","pagedown")

def previousTab():
    pyautogui.hotkey("ctrl","pageup")


def openPBR(pbreport, screenSignal=""):
    
    print("opening " + pbreport)
    start_time = time.time()
    # 
    file_name = utils.extractFileName(pbreport)[:-1]
    try:
        print("abreee:" + file_name)
        utils.bring(file_name)
        # time.sleep(0.5)
    except:
    #    utils.msgbox("","Please open " + pbreport + " first", 0)
    #    sys.exit('Forced error')
        excel = subprocess.Popen(pars.excelApp + " " + pbreport)

        # utils.waitUntil(utils.img("excel\PBR_ModelInitiatization.PNG"), attempts=100)
        # pyautogui.press("enter")
        # utils.waitUntil(utils.img("excel\current_cell_is_b6.PNG"),attempts=100)
        time.sleep(5)
    elapsed_time = time.time() - start_time
    print("Open complete " + str(elapsed_time))

    
# def open(pbreport, screenSignal=""):
    
#     excel = subprocess.Popen(pars.excelApp + " " + pbreport)
#     utils.waitUntil(utils.img(screenSignal),attempts=100)
    
      

      


    #   if screenSignal!="":
        #   time.sleep(10)
    #   else:
        # utils.waitUntil(screenSignal)
    #   utils.msgbox("Info","File is Loaded",0)
    #   pyautogui.press('esc')
    #   return excel
      
    


def runReport(reportLine, wait):
    gotoCell ("b" + str(reportLine))
    time.sleep(wait)


def findText(text, wait=2):
    pyautogui.hotkey("ctrl","f")
    pyautogui.typewrite(text)
    # gotoCell ("b" + str(reportName))
    pyautogui.press("enter")
    
    # time.sleep(wait)
    # pyautogui.press("esc")
    
    

def runReportByName(reportName, saveto="", timeStampFileFormat="today" ):
    gotoCell ("B11")
    findText(reportName)
    utils.waitUntil (r"excel\current_cell_is_b6.PNG")
    pyautogui.press("esc")
    time.sleep(1)
    fileName = ""
    if saveto != "":
        fileName = excel.saveReportToFile(saveto, excel.AutofitColumns, timeStampFileFormat)
    return fileName
    
    
        



    
def activateAutoFilter():
    pyautogui.hotkey("ctrl","shift","l")
    pyautogui.press("home")

def copyReportToClipBoard():
    gotoCell ("d14")
    pyautogui.hotkey("ctrl","a")
    pyautogui.hotkey("ctrl","c")
    time.sleep(1)

def paste():
    pyautogui.hotkey("ctrl","v")
    time.sleep(1)

def pasteValues():
    pyautogui.hotkey("alt","e")
    pyautogui.press("s")
    pyautogui.press("t")
    pyautogui.press("v")
    pyautogui.press("enter")
    time.sleep(1)

    
def close():
    pyautogui.hotkey("ctrl","w") #closing file
    time.sleep(1)
    
def copy():
    pyautogui.hotkey("ctrl","c")

def deleteContent(range=""):
    if range!="":
       excel.selectRange(range)
    pyautogui.press("delete")

    pyautogui.press("del")

def saveAs(fileName):
    if os.path.isfile(fileName ): 
        os.remove(fileName) #deleting existing file
    pyautogui.press("f12")
    pyautogui.typewrite(fileName)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)

def moveSelection(sentido, times=1):
    for x in range(times):
        pyautogui.hotkey("shift",sentido)


def enterText(pText):
    pyautogui.typewrite(pText)
    pyautogui.press("enter")



def saveReportToFile(fileName, formatting, timeStampFileFormat="", csv=False):
    def saveNew(fileName, csv):
        pyautogui.hotkey("f12")
        time.sleep(1)
        if csv:
            pyautogui.click(536, 428, interval=1) # Selecting CSV Formart 1 of 2
            pyautogui.click(508, 670, interval=1) # Selecting CSV Formart 2 of 2
            pyautogui.click(182, 404, interval=1)  #clicking on the file name box again
        
        path = utils.extractPath(fileName)
        fileName = utils.extractFileName(fileName)
        
        fileStamp = utils.getTimeStampStr(utils.YYmMM_DD, timeStampFileFormat )
      
        if fileStamp != "":
           fileName =  fileStamp + " - " + fileName 
           
        fileName = os.path.join(path, fileName )
        utils.deleteIfExists(fileName)
        pyautogui.typewrite ( fileName )
        pyautogui.press("enter")
        time.sleep(1)
        return fileName

    copyReportToClipBoard()
    
    if os.path.isfile(fileName ): 
        os.remove(fileName) #deleting existing file
    
    newExcelFile()
    paste()
    
    # pyautogui.click (153,208,clicks=2)
    AutofitColumns()
    formatting()
    gotoCell ("B11")
    fileName = saveNew(fileName, csv)
    close()
    return fileName
    
    
    # runReport(reportLinePaste)
    # copyReportToClipBoard()
    return  (fileName)
    
def gotoCell(cell):
    pyautogui.hotkey("ctrl","g")
    pyautogui.typewrite(cell)
    


    pyautogui.press("enter")
    # time.sleep(1)

def newExcelFile():
    pyautogui.hotkey("alt","f","n","l")
    time.sleep(1)
    # pyautogui.click(33,45, interval=1)
    # pyautogui.click(30,130, interval=1)
    # pyautogui.click(276,338, interval=1)
        
def test(reportLine, fileName):
    openPBR()
    gotoCell ("b" + str(reportLine))
    gotoCell ("d14")
    time.sleep(1)
    pyautogui.hotkey("ctrl","a")
    time.sleep(1)
    pyautogui.hotkey("ctrl","c")
    time.sleep(1)

def getlimitY(cell, offset=0):
    gotoCell(cell)
    
    end("right")
    right(offset)
    return getCurrentCol()

def getlimitX(cell, offset=0):
    gotoCell(cell)
    end("down")
    down(offset)
    return getCurrentRow()

def colLetter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def colNameToNum(name):
    pow = 1
    colNum = 0
    for letter in name[::-1]:
            colNum += (int(letter, 36) -9) * pow
            pow *= 26
    return colNum


def nextColLetter(cell):
    return colLetter(colNameToNum(cell)+1)



def enterFormula(formula):
    pyautogui.typewrite(formula)    
    pyautogui.press("enter")
    pyautogui.press("up")
    # time.sleep(1)

def getCurrentCol():
    a = result = getCurrentCell()
    result = ''.join([i for i in a if not i.isdigit()])
    return result

def getAddress():
    result = getCurrentCell()
    return result

def selectRange(range):
    pyautogui.hotkey("ctrl","g")
    pyautogui.typewrite(range)
    pyautogui.hotkey("enter")


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def getCurrentRow():
    a = getCurrentCell()
    result = ''.join([i for i in a if i.isdigit()])
    return result


def up(times=1):
    pyautogui.press("up", presses=times)

def down(times=1):
    pyautogui.press("down", presses=times)

def right(times=1):
    pyautogui.press("right", presses=times)

def left(times=1):
    pyautogui.press("left", presses=times)


def endSelect(type):
    if type=="down":
       pyautogui.hotkey("ctrl","shift","down")
    elif type=="right":
       pyautogui.hotkey("ctrl","shift","right")
    elif type=="up":
       pyautogui.hotkey("ctrl","shift","up")

 

def end2(cell, type, offsetX=0, offsetY=0):
    gotoCell(cell)
    end(type, offsetX, offsetY)
    return getAddress()


def filterPivot(cell, area):
       excel.gotoCell(cell)
       pyautogui.hotkey("alt","down")
       time.sleep(0.5)
       if area == "All":
            pyautogui.typewrite("*") 
            time.sleep(0.5)
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyautogui.press("enter")
       else:
         pyautogui.typewrite(area)    
         time.sleep(0.5)
         pyautogui.press("tab")
         pyautogui.press("tab")
         pyautogui.press("tab")
         pyautogui.press("right")
         pyautogui.press("enter")

def copyTable(upperLeft, ul_offsetX=0, ul_offsetY=0, lr_offsetX=0, lr_offsetY=0):
    excel.goLastRowWithData(upperLeft[0:1])
    excel.goLastColWithData(excel.getCurrentRow(),offsetX=lr_offsetX, offsetY=lr_offsetY)
    lr_cell = excel.getCurrentCell()
    excel.selectRange(excel.applyOffsets(upperLeft,ul_offsetX,ul_offsetY) + ":" + lr_cell)
    excel.copy()


def applyOffsets(cell, offsetX=0, offsetY=0):
    if offsetY==0 and offsetX==0:
        return cell
    else:
        gotoCell(cell)
        end("", offsetX, offsetY)
        return getCurrentCell()

def end(type, offsetX=0, offsetY=0):
    if type=="down":
       pyautogui.hotkey("ctrl","down")
    elif type=="right":
       pyautogui.hotkey("ctrl","right")
    elif type=="left":
       pyautogui.hotkey("ctrl","left")
    elif type=="up":
       pyautogui.hotkey("ctrl","up")
    elif type=="":
        pass #does nothing, only applies offset

    if offsetX>0:
        right(offsetX)
    else:
        left(offsetX * -1)

    if offsetY>0:
        down(offsetY)
    else:
        up(offsetY * -1)
    




def AutofitColumns():
    pyautogui.hotkey("alt","h","o","i")
    time.sleep(2)
    
    

def setColumnSize(col,size):
    print("Set column executed" )
    gotoCell(col + "1")
    pyautogui.click(91,50, interval=1)   #Home Ribbon
    pyautogui.click(1500, 131, interval=1) #Format
    pyautogui.click(1538, 235, interval=1) #Col size
    pyautogui.typewrite(str(size))
    pyautogui.press("enter")

def addCol():
    pyautogui.hotkey("alt","i")
    pyautogui.press("c")
    
def getCurrentCell():
    fileName = utils.addPath(r"excel\namebox.png", "imgPattern")
    x, y = pyautogui.locateCenterOnScreen(fileName, confidence=0.8)
    pyautogui.click(x-100,y)
    copy()
    result  = pyperclip.paste()
    pyautogui.press("enter")
    return result
        



def deleteCol():
    pyautogui.hotkey("ctrl","-")
    pyautogui.press("c")
    pyautogui.press("enter")


def PBReportsRefresh():
    utils.automate("PBReports Refresh")

def getXLDRValueAsText(book, row, col):

    type =  book.sheet_by_index(0).cell(row,col).ctype
    value =  book.sheet_by_index(0).cell(row,col).value
    
    if type == xlrd.XL_CELL_DATE:
        year, month, day, hour, min, sec = xlrd.xldate_as_tuple(value, book.datemode)
        # print (year)
        if year == 0:
            year=1
            month=1
            day = 1
        return datetime.datetime(year, month, day, hour, min)
    else:
        return value

def open_workbook(fileName):
    return xlrd.open_workbook(fileName)
    
        # return "{0}/{1}/{2}".format(month, day, year)


