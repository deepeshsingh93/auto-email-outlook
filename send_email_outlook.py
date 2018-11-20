# -*- coding: utf-8 -*-
"""
Created on Wed Sep 21 15:36:00 2016

@author: Deepesh.Singh
"""

import win32com.client as win32
import psutil
import os
import subprocess

# Drafting and sending email notification to senders. You can add other senders' email in the list
def send_notification():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'abc@xzs.com; bhm@ert.com', 
    mail.Subject = 'Sent through Python'
    mail.body = 'This email alert is auto generated. Please do not respond.'
    mail.send
    
# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
    
def open_outlook():
    try:
        subprocess.call(['C:\Program Files\Microsoft Office\Office15\Outlook.exe'])
        os.system("C:\Program Files\Microsoft Office\Office15\Outlook.exe");
    except:
        print("Outlook didn't open successfully")

# Checking if outlook is already opened. If not, open Outlook.exe and send email
for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0

if (flag == 1):
    send_notification()
else:
    open_outlook()
    send_notification()

    
    
import win32ui
def process_running():
    if win32ui.FindWindow ("Microsoft Outlook","Microsoft Outlook"):
        print "Already running"
        return True


process_running()

import psutil

pythons_psutil = []
for p in psutil.process_iter():
    try:
        if p.name() == 'Outlook.exe':
            pythons_psutil.append(p)
    except psutil.Error:
        pass
    
print(psutil.pids())

lst = psutil.pids()

p = psutil.Process(692)  # The pid of desired process
print(p.name()) # If the name is "python.exe" is called by python
print(p.cmdline()) # Is the command line this process has been called with


for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        print p.name
    else:
        print "no match found"      
 
