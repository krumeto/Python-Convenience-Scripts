# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 12:56:56 2019

@author: arnaudok
"""

def Emailer(text, subject, recipient, attachment = None):
    import win32com.client as win32   

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    if attachment != None:
        mail.Attachments.Add(attachment)
    mail.send