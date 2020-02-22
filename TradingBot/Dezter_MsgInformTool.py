# -*- coding: utf-8 -*-
"""
Created on Sun Sep 22 18:31:55 2019
Update on Sun Nov 24 2019 --> Fix: msg's words over Line's maxium will not allow to send msg
Update on Sun Dec 01 2019 --> Fix: Line Notify has words limit, Add msg split solution to fit their rule

@author: Dezter
Setup Message Inform Tool:
    Lib need to install:
        email, smtplib, skpy, requests, datetime
        line notify token="*********"
"Print" *3, "Error" *3
"""
#%% Lib Import %%#
import datetime
# Line Notify Use
import requests
# Email Use
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
# Skype Use
#from skpy import Skype

#%% Message Inform tool %%#
def InformTool(Command, OrderBack, Error):
    CommandStrlist = []; CommandStrtext = ""; OrderBackStrlist = []; OrderBackStrtext = ""; MsgText=""
    if Error == "No Error":
        # Command MSG Text
        CommandStrlist = [', '.join(map(str, Command[i])) for i in range(0,2)]
        for i in range(len(Command[2])):
            CommandStrlist.append(', '.join(map(str, Command[2][i]))) 
        CommandStrtext = '\n'.join(map(str, CommandStrlist))
        # Order Feedback MSG Text
        if OrderBack[0] != "Order: None":
            for i in range(len(OrderBack)):
                OrderBackStrlist.append(', '.join(map(str, OrderBack[i]))) 
                OrderBackStrtext = '\n'.join(map(str, OrderBackStrlist))
        elif OrderBack[0] == "Order: None":
            OrderBackStrtext = '\n'.join(map(str, OrderBack))
        MsgText = CommandStrtext + '\n'+'\n' + OrderBackStrtext
    elif OrderBack != [] :
        # Command MSG Text
        CommandStrlist = [', '.join(map(str, Command[i])) for i in range(0,2)]
        for i in range(len(Command[2])):
            CommandStrlist.append(', '.join(map(str, Command[2][i]))) 
        CommandStrtext = '\n'.join(map(str, CommandStrlist))
        # Order Feedback MSG Text
        if OrderBack[0] != "Order: None":
            for i in range(len(OrderBack)):
                OrderBackStrlist.append(', '.join(map(str, OrderBack[i]))) 
                OrderBackStrtext = '\n'.join(map(str, OrderBackStrlist))
        elif OrderBack[0] == "Order: None":
            OrderBackStrtext = '\n'.join(map(str, OrderBack))
        MsgText = CommandStrtext + '\n'+'\n' + OrderBackStrtext + '\n'+'\n' + Error
    elif Command != [] :
        # Command MSG Text
        CommandStrlist = [', '.join(map(str, Command[i])) for i in range(0,2)]
        for i in range(len(Command[2])):
            CommandStrlist.append(', '.join(map(str, Command[2][i]))) 
        CommandStrtext = '\n'.join(map(str, CommandStrlist))
        MsgText = CommandStrtext + '\n'+'\n' + 'System test, Inform Only' + '\n'+'\n' + Error # --> Test Use
        #MsgText = CommandStrtext + '\n'+'\n' + 'Check Future Order Fun' + '\n'+'\n' + Error # --> Formal Use
    else:
        MsgText = Error
    #%% Sending Msg pass through LineNotify by request %%#
    def lineNotify(token, msg):
        headers = {
        "Authorization": "Bearer " + token, 
        "Content-Type" : "application/x-www-form-urlencoded"
        }
        payload = {'message': "\n"+ msg}
        r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
        return r.status_code
    
    #Split Msg to fit Line Notify's maxium words
    TextLen = []; TextLenSort = []; MsgSplit = []
    TextLen = [Len for Len,Tab in enumerate(MsgText) if Tab=='\n']
    TextLenSort = sorted(TextLen + [0,len(MsgText)])
    Split = 40
    if (len(TextLenSort)-1)%Split == 0:
        for j in range(0,len(TextLenSort),Split):
            MsgSplit.append(TextLenSort[j])
    else:
        for j in range(0,len(TextLenSort),Split):
            MsgSplit.append(TextLenSort[j])
        MsgSplit = MsgSplit + [TextLenSort[-1]]
    token = '*********'
    try:
        for k in range(1,len(MsgSplit)):
            if k == 1:
                Massage = MsgText[0:MsgSplit[k]]
                lineNotify(token, Massage)
                #print('Send Line Notify Msg')
            else:
                Massage = MsgText[MsgSplit[k-1]+1:MsgSplit[k]]
                lineNotify(token, Massage)
                #print('Send Line Notify Msg')
    except Exception as e:
        print("Error: unable to send Line Notify Msg " + str(e))
    #%% Sending Msg pass through Gmail %%#
    Sender = '*********@mail.com'
    Passwd = '*********'
    Receivers = '*********@mail.com'
    
    msg = MIMEMultipart()
    msg['Subject'] = "Dezter's Trading Bot"
    msg['From'] = Sender
    msg['To'] = Receivers
    msg.preamble = 'Multipart massage.\n'
    msg.attach(MIMEText(MsgText))
    try:
        smtp = smtplib.SMTP("smtp.gmail.com:587")
        smtp.ehlo()
        smtp.starttls()
        smtp.login(Sender, Passwd)
        smtp.sendmail(Sender, Receivers, msg.as_string())
        #print('Send mails to',msg['To'])
    except Exception as e:
        print("Error: unable to send email " + str(e))
    #%% Sending Msg pass through Skype %%#
"""
    SkyID = "*********"
    SkyPW = "*********"
    SkyReceiver = "live:*********"
    try:
        SKUser = Skype(SkyID, SkyPW)
        Chat = SKUser.contacts[SkyReceiver].chat
        Chat.sendMsg(MsgText)
        #print('Send Skype Msg to',SkyReceiver)
    except Exception as e:
        print("Error: unable to send Skype Msg " + str(e))
"""