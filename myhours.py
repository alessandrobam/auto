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



class pyMyHours (pyWebLocation):
        url = "https://w3-01.ibm.com/services/tools/marketplace/findaPract.wss"
        
        def __init__(self):
            pass

        def start(self):
            subprocess.Popen(pars.firefoxApp  + " " + self.url)
        
        def clickClearAllFiltes(self):
            utils.clickImg("myhours\clearall.PNG", x_offset=0,y_offset=0)
            
        
        def ClickOnSearchBox(self):
            utils.clickImg("myhours\lupa.PNG", x_offset=-100,y_offset=0)
        
        def ClickSearch(self):
            utils.clickImg("myhours\lupa.PNG", x_offset=0,y_offset=0)

        def ClickResource(self):
            utils.clickImg(r"myhours\resource_arrow.PNG", x_offset=0,y_offset=0)




def findProfessional():
    serial = pyautogui.prompt(text='Search Who?', title='Serial' , default='')    
    myhours = pyMyHours()
    myhours.start()
    utils.waitUntil("myhours\lupa.PNG")
    
    
    
    
    
    time.sleep(3)
    pyautogui.press("pagedown")
    time.sleep(1)

    # pyautogui.scroll(-100)


    
    pyautogui.click(utils.waitUntil("myhours\clearall.PNG"))
    time.sleep(3)

    

    # print(utils.waitUntil("myhours\lupa.PNG"))

    # utils.clickImg("myhours\lupa.PNG", x_offset=-100,y_offset=0)


    myhours.ClickOnSearchBox()

    pyautogui.typewrite(serial)
    myhours.ClickSearch()
    
    time.sleep(3)
    pyautogui.press("pagedown")
    time.sleep(2)
    
    myhours.ClickResource()





    # myhours.start(wait=2)

    # try:
    #     myhours.clickClearAllFiltes(4)
    # except:
    #     pass
    
    # myhours.ClickOnSearchBox()
    # pyautogui.typewrite(serial)
    # print ("------------>" + serial)
    # myhours.ClickSearch(wait=4)
    # pyautogui.scroll(-4000)
    # time.sleep(2)
    
    






            
                
                

        
                
