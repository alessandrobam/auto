import excel, pars, utils, verse, pyautogui




def refreshPBReports(file):
    excel.openPBR(file)
    excel.PBReportsRefresh()



def sendPBReportsByEmail (pbrfile, reportName, emailSubject, distro, body, saveExcelReportTo, timeStampFileFormat="today", IncludeReportInBodyReport=False):
    # timeStampFileFormat - today, last_friday, none
    
    excel.openPBR(pbrfile)
    fileName  = excel.runReportByName(reportName, saveExcelReportTo, timeStampFileFormat )
    # utils.alert(fileName)
    if IncludeReportInBodyReport:
        excel.copyReportToClipBoard()

    emailSubject = emailSubject + " as of " + utils.getTimeStampStr("%b %d",timeStampFileFormat)        
    verse.sendVerseEmail(emailSubject, distro,body,fileName, True)



# sendPBReportsByEmail(pars.pbReportsABT,"CGP - Labor vs Non Labor","Test Report", ["abarbosa@br.ibm.com"],"This is a automatic email", r"C:\Temp\report01.xlsx","today", True)

# sendPBReportsByEmail(pbrfile = pars.pbReportsABT,
#                      reportName = "CP - Hours, Cost, Rev, Profit",
#                      emailSubject = "CP - Hours, Cost, Rev, Profit  - WE", 
#                      distro = ["abarbosa@br.ibm.com"],
#                      body = utils.getEmailTemplate("IPPF vs Cost Email.txt"),
#                      saveExcelReportTo = r"C:\AlessandroBAM\2017m01 - Abbott DPE-PgM-PM\Governance\Plan & Build\AUTO - Weekly CP - Hours, Cost, Rev, Profit report\CP - Hours, Cost, Rev, Profit.xlsx",
#                      timeStampFileFormat = "last_friday",
#                      IncludeReportInBodyReport = True)





# excel.AutofitColumns()