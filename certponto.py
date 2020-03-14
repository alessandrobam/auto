import subprocess
import time
import pyautogui
import utils
import os
import excel, verse
import csv
import pars

url = "https://secure.certponto.com.br/"
firefox = "C:\\Program Files\\Mozilla Firefox\\firefox.exe "
# pbReportsCPG = "\"C:\\AlessandroBAM\\2017m01 - Abbott DPE-PgM-PM\\IBM Brazil\\RH\\People Management Misc\\PB Reports CPG\\PB Reports CPG.xlsm\""
csvfile = "C:\\AlessandroBAM\\2017m01 - Abbott DPE-PgM-PM\\IBM Brazil\\RH\\Aprovacao Mensal de Horas Extras e Standby\\2019m03 - Analise de Overtime acima de 2 horas\\input.csv"

def start():
    subprocess.Popen(firefox  + url)
    time.sleep(3) 
    
def login():
    pyautogui.click(1263, 369)
    pyautogui.typewrite("abarbosa@br.ibm.com")
    pyautogui.press("tab")
    pyautogui.typewrite("27809819895")
    pyautogui.press("enter")
    time.sleep(5) 

def getMeThere():
    start()    
    login()

def approvePendingRequests():
    start()    
    login()
    pyautogui.click(115, 185, interval=7) # 1 - click on "Tratamento de Ponto" - Wait 2
    pyautogui.click(1184, 454, interval=1) # 2 - click on Aprov Pendentes - Wait 1
    pyautogui.click(1191, 507, interval=1) # 3 - click on Sim - Wait 2
        
    pyautogui.click(utils.waitUntil (r"Certponto\Gerar_button.PNG"))
    pyautogui.click(utils.waitUntil (r"Certponto\Checkbox_gray_back.PNG"))
    pyautogui.click(utils.waitUntil (r"Certponto\Acoes_button.PNG"))
    pyautogui.click(utils.waitUntil (r"Certponto\Aprovar_Linhas_Selecionadas.PNG"))
    pyautogui.click(utils.waitUntil (r"Certponto\sim_button.PNG"))

    


    

    # pyautogui.click(320, 701 , interval=4) # 4 - click on Gerar - 2
    # pyautogui.click(306, 851 , interval=1) # 5 - click on check box to select everything - 1
    # pyautogui.click(1837,808 , interval=2) # 6 - click on Acoes - 
    
    # pyautogui.click(1734, 847, interval=2) # 7 - Aprovar selecionados - 
    # pyautogui.click(442, 830, interval=2) # 8 - Confirmar- 

def geraEspelhoDePontoThisMonth():
    geraEspelhoDePonto(False)

def geraEspelhoDePontoPrevMonth():
    geraEspelhoDePonto(True)

def geraEspelhoDePonto(prev):
    start()    
    login()
    saveDir = "C:\\AlessandroBAM\\2017m01 - Abbott DPE-PgM-PM\\IBM Brazil\\RH\\Aprovacao Mensal de Horas Extras e Standby\\2019m01 - Melhoria do processo de validacao de Certponto vs ILC\\Certponto"

    if prev:
        fileName = saveDir + "\\" + utils.prevMonth().strftime("%Ym%m") + " - Espelho Certponto.csv"
    else:
        fileName = saveDir + "\\" + utils.today().strftime("%Ym%m") + " - Espelho Certponto.csv"

    pyautogui.click(125,192, interval=7) #click em tratamento de ponto
    pyautogui.click(112,349,  interval=3) #click espelho de ponto
    pyautogui.click(561,489,  interval=1) #click on periodo simplificado
    
    if prev:
        pyautogui.click(303,632,  interval=1) #click on mes passado
    else:
        pyautogui.click(297,607,  interval=1) #click on mes atual
    
    
    pyautogui.click(297,699,  interval=2) #click on gerar
    pyautogui.click(1852,770,  interval=1) #click on Acoes
    pyautogui.click(1773,882,  interval=1) #click Exportar CSV reduzido-
    pyautogui.click(545,219,  interval=10) #click SIM
    pyautogui.hotkey("alt","s")
    pyautogui.press("enter",interval=6)
    if os.path.isfile(fileName ): 
        os.remove(fileName) #deleting existing file
    pyautogui.typewrite(fileName)
    time.sleep(2)
    pyautogui.press("tab",interval=1)
    pyautogui.press("enter",interval=1)
    excel.openPBR(pars.pbReportsCPG,"")

def runPBReportRefresh():
    excel.openPBR(pars.pbReportsCPG,"")
    # excel.PBReportsRefresh()

def sendEmailsToWhoExceeded2HoursLimite():
    verse.start(10)
    with open(csvfile,'r') as csv_file:
        csv_reader  = csv.reader(csv_file)
        for line in csv_reader:
            print (line[1])
            subject = "Limite de 2 horas Extras Excedido em " + line[2] +   ". Qual o Motivo?"
            verse.newEmail(subject ,[line[1]],5)
            verse.newBodyLine("Oi " + line[0].split()[0] + ",")
            verse.newBodyLine("")
            verse.newBodyLine("De acordo com o Certponto, voce trabalhou " + line[8] + " horas na " + line[3] + ", dia "+ line[2] + ". Qual o motivo?")
            verse.newBodyLine("")
            verse.newBodyLine("No Aguardo")


# sendEmailsToWhoExceeded2HoursLimite()