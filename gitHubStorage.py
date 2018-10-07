import win32com.client
import re

class color:
   PURPLE = '\033[95m'
   CYAN = '\033[96m'
   DARKCYAN = '\033[36m'
   BLUE = '\033[94m'
   GREEN = '\033[92m'
   YELLOW = '\033[93m'
   RED = '\033[91m'
   BOLD = '\033[1m'
   UNDERLINE = '\033[4m'
   END = '\033[0m'

virusCleaned = "Cleaned"
spywareCleaned = "Successful"
# callbackCleaned = ""
spyware = "Spyware/Grayware"
callback = "C&C callback"
virus = "Virus/Malware"

inbox = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
sendMail = win32com.client.Dispatch("Outlook.Application")

toDo = inbox.GetDefaultFolder(6).Folders["Virus"]
Done = inbox.GetDefaultFolder(6).Folders["Virus Done"]


def getMail():
    message = toDo.Items.GetLast()
    checkClean = message.Body
    checkSubject = message.Subject
    if checkSubject == virus:
        if virusCleaned in checkClean:
            moveMail(message)
            return "Mail Clean No further action Required"
        else:
            body = grabContentsVirus(checkClean)
    if checkSubject == callback:
        # if callbackCleaned in checkClean:
        #     moveMail(message)
        #     return "Mail Clean No further action Required"
        # else:
        body = grabContentsCallback(checkClean)
    if checkSubject == spyware:
        if spywareCleaned in checkClean:
            moveMail(message)
            return "Mail Clean No further action Required"
        else:
            body = grabContentsSpyware(checkClean)
    # The below can be disabled for testing
    attachment = saveMail(message)
    sent = sendingMail(body, attachment, checkSubject)
    if sent == True:
        moveMail(message)
        print("Email Sent")
        return
    print("Email Failed to Send")


def grabContentsVirus(email):
    splited = re.split(': |\n|\r', email)
    print(splited)
    computer = splited[8]
    time = splited[17]
    threat = splited[5]
    file = splited[14]
    user = splited[23]
    # v = 0
    # for i in splited:
    #     print(i, v)
    #     v +=1
    #
    mustSendBody = writingMailVirus(computer, time, threat, file, user)
    return mustSendBody

def grabContentsCallback(email):
    splited = re.split(': |\n|\r', email)
    print(splited)
    computer = splited[5]
    ipAddress = splited[8]
    domain = splited[11]
    time = splited[14]
    callbackAddress = splited[17]
    riskLevel = splited[20]
    listSource = splited[23]
    action = splited[26]
    # v = 0
    # for i in splited:
    #     print(i, v)
    #     v +=1
    mustSendBody = writingMailCallback(computer, ipAddress, domain, time, callbackAddress, riskLevel, listSource, action)
    return mustSendBody

def grabContentsSpyware(email):
    splited = re.split(': |\n|\r', email)
    print(splited)
    computer = splited[5]
    domain = splited[8]
    time = splited[11]
    spywareThreat = splited[16]
    user = splited[23]
    # v = 0
    # for i in splited:
    #     print(i, v)
    #     v +=1
    mustSendBody = writingMailSpyware(computer, domain, spywareThreat, time, user)
    return mustSendBody

def writingMailSpyware(computer, domain, spywareThreat, time, user):
    email = ("***Infection Details***\n\t"
      "Computer Name: " + computer + "\n\t"
      "Domain: " + domain + "\\\n\t"
      "Spyware/Grayware and Result:  " + spywareThreat + "\n\t"
      "Date/Time: " + time + "\n\t"
      "User: " + user + "\n\n\t")
    return email


def writingMailCallback(computer, ipAddress, domain, time, callbackAddress, riskLevel, listSource, action):
    email = ("***Infection Details***\n\t"
      "Compromised Host: " + computer + "\n\t"
      "IP Address: " + ipAddress + "\n\t"
      "Domain: " + domain + "\n\t"
      "Date/Time: " + time + "\n\t"
      "Callback address: " + callbackAddress + "\n\t"
      "C&C risk level: " + riskLevel + "\n\t"
      "C&C list source: " + listSource + "\n\t"
      "Action: " + action + "\n\n\t")
    return email

def writingMailVirus(computer, time, threat, file, user):
    email =("***Infection Details***\n\t"
          "Computer Name: " + computer + "\n\t"
          "Threat Name: " + threat + "\n\t"
          "Threat File Name/Location 1: " + file + "\n\t"
          "Username: " + user + "\n\t"
          "Time Threat was Detected on Computer: " + time + "\n\n\t")
    return email

def sendingMail(body, attachment, checkSubject):  # This sends the email
    mail = sendMail.CreateItem(0x0)
    mail.To = "<MYEMAIL>"
    mail.Subject = "New Case - Desktop Support Team - " + checkSubject
    mail.Body = body
    mail.Attachments.Add(attachment)
    mail.Send()
    return True

def saveMail(message):
    subject = message.Subject
    subject = subject.replace("/", "_")
    subject = u"<PATHWAY>" + subject + ".msg"
    message.SaveAs(subject)
    print("Message Saved")
    return subject



def moveMail(message):
    message.UnRead = False
    print("moving mail to done folder")
    message.Move(Done)
    return print("MOVED")


getMail()