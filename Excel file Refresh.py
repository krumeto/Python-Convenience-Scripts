#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pathlib import Path
import win32com.client
import time
import datetime


# In[4]:

# Save links in the format:"C:\\Users\\username\\OneDrive\\FileName.xlsx" or
# For Sharepoint "https://company.sharepoint.com/sites/Example/PowerBi%20Data/Team1.xlsx" - remember to delete the ?web at the end of the link

Team1 = "C:\\Users\\arnaudok\\OneDrive - company Inc\\Python Scripts\\REPORTS\\data\\Sharepoints data\\Team1.xlsx"
Team2 = "C:\\Users\\arnaudok\\OneDrive - company Inc\\Python Scripts\\REPORTS\\data\\Sharepoints data\\Team2.xlsx"
pTeam2 = "C:\\Users\\arnaudok\\OneDrive - company Inc\\Python Scripts\\REPORTS\\data\\Sharepoints data\\pTeam2.xlsx"
Team3 = "C:\\Users\\arnaudok\\OneDrive - company Inc\\Python Scripts\\REPORTS\\data\\Team3 data\\Team3 Proactiveness.xlsx"
Team1_online = "https://company.sharepoint.com/sites/Division/Reports/PowerBi%20Data/Team1.xlsx"
Team2_online = "https://company.sharepoint.com/sites/Division/Reports/PowerBi%20Data/Team2.xlsx"
pTeam2_online = "https://company.sharepoint.com/sites/Division/Reports/PowerBi%20Data/pTeam2.xlsx"
Team3_online = "https://company.sharepoint.com/sites/Division/Reports/PowerBi%20Data/Team3 Proactiveness.xlsx"
links = [Team1, Team2, pTeam2,Team3,Team1_online,Team2_online,pTeam2_online,Team3_online]


# In[5]:


def update_excel(path):
    # Start an instance of Excel
    xlapp = win32com.client.DispatchEx("Excel.Application")
    
    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(path)
    # Optional, e.g. if you want to debug
    # xlapp.Visible = True
    time.sleep(5)
    # Refresh all data connections.
    wb.RefreshAll()
    #xlapp.CalculateUntilAsyncQueriesDone()
    time.sleep(5)
    wb.Save()
    time.sleep(10)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Updated on {now}: ", path)
    # Quit
    xlapp.Quit()


# In[6]:


for link in links:
    update_excel(link)

