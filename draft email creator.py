# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 13:05:44 2023

@author: nebedarkar
"""

import pandas as pd
import numpy as np
import os
import win32com.client as win32
# construct Outlook application instance
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct the email item object

student_list = pd.read_excel(r'C:\Users\nebedarkar\Downloads\AJ Reminder Email List.xlsx')



def Emailer(text, subject, recipient,cc):
    import win32com.client as win32   
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.CC = cc  
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Save()

print(student_list.columns)

mail_content_start = ""
mail_content_end = ""
with open(r'C:\Users\nebedarkar\Downloads\email.txt', "r", encoding="utf-8") as f:
    mail_content_start = f.read()
    
with open(r'C:\Users\nebedarkar\Downloads\emailend.txt', "r", encoding="utf-8") as f:
    mail_content_end = f.read()

for ind in student_list.index:
     msg = "Howdy {}!".format(student_list['First Name'][ind])
     final_msg = mail_content_start + msg + mail_content_end
     print(student_list['First Name'][ind])
     print(student_list['Email'][ind])
     Emailer(final_msg, "Reminder – This is the subject", 
             student_list["Email"][ind],
             student_list["CC"][ind])
    
print("Done")
