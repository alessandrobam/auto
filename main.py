# Import the necessary packages
from consolemenu import *
from consolemenu.items import *
import etime, excel, ilc, ippf, weeklyreports,pyautogui, certponto, ippf, connections, verse, checkpoint, sys, outlook, notes, utils, myhours

options = { "Etime - Pending Requests Report": etime.getETimeRequests,
            "Certponto - Worked Hours Hours Past Month": certponto.geraEspelhoDePontoPrevMonth,
            "Certponto - Worked Hours Hours Current Month": certponto.geraEspelhoDePontoThisMonth,
            "Certponto - Approve Requests": certponto.approvePendingRequests,
            "Certponto - Get me There": certponto.getMeThere,
            "Certponto - Email to who exceeded 2 hours limite": certponto.sendEmailsToWhoExceeded2HoursLimite,
            "PBReports - Cost Non Billable Weekly Report": weeklyreports.sendCNBReport,
            "PBReports - IPPF vs Actuals Weekly Report": weeklyreports.sendIPPFActuals,
            "PBReports - CP - Hours, Cost, Rev, Profit Weekly Report": weeklyreports.sendCPHoursCostRevProfit,
            "PBReports - 2019 DMC Approved PRs": weeklyreports.send2019DMCApprovedPRs,
            "PBReports - Waiting Claim Weekly Report": weeklyreports.sendClaimWatingReport,
            "PBReports - Latin America PRs Weekly Report": weeklyreports.sendPRReports,
            "PBReports - Roster Status": weeklyreports.sendMissingRoster,
            "PBReports - BPCS ITSM Categorization": weeklyreports.sendBPCSCategorizationReport,
            "PBReports - Wrong Contract": weeklyreports.sendWrongContractFlag,
            "PBReports - Update Roster Rates": weeklyreports.updateRosterWithCGPRates,
            "PBReports - JnJ Breached Tickets": weeklyreports.sendJNJBreachedTickets,
            "PBReports - Local ABT Actuals": weeklyreports.sendLocalActuals,
            "Notes - Get me There": notes.getMeThere,
            "Checkpoint - Get me There": checkpoint.openAll,
            "Checkpoint - Go to Employee": checkpoint.openByName,
            "KTLO - Save Weekly Ticket Report": connections.downloadFilesKTLO,
            "Open Url": utils.openURL,
            "Find Profile": myhours.findProfessional,
            # "Verse - Create Mass Invites": verse.createCalendarInvite,
            # "Verse - Send Direct Mail": verse.sendDirectMail,
            "Outlook - Get me There": outlook.getMeThere,
            "IPPF - Labor Details Report": ippf.getActualsHours,
            "Save PS Files": verse.savePSFile,
            "Verse - Save Email As PDF": verse.saveCurrentEmailAsPDF, #verse.saveCurrentEmailAsPDF,
            "Verse - Save Attachments": verse.saveAttachments, #verse.saveCurrentEmailAsPDF,
            "Avoid Missing Claim": ilc.avoidMissingClaim,
             0: "Testing"}


if len(sys.argv)>1:
    pyautogui.FAILSAFE = True
    if (len(sys.argv)==2):
        # print (sys.argv[1])
        options[sys.argv[1]]()
    if (len(sys.argv)==3):
        options[sys.argv[1]](sys.argv[2])

# utils.msgbox("Info","Automation is complete",0)
