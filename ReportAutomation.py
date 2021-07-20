import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.ttk import Combobox
from tkinter import filedialog
from TkinterDnD2 import *

from email.message import EmailMessage
import csv
import xlrd
import excel as excelTab

import calender as calender
import email, smtplib


CLICK_VIEWS = []
PAGES = []
DATES = []
UNIQUEDATES = []
IMPORTNAMES = []
IMPORTVIEWS = []
IMPORTHOVERS = []
IMPORTCLICKS = []
ADPROGRAMS = ['A More Just NYC Kivvit', 'AARP', 'Adiply', 'AFL-CIO', 'Aid in Dying'
                  ,'AMERICAN CHEMISTRY COUNCIL','AMERICAN INVESTMENT COUNCIL', 'American Legal Finance Association', 'American Progressive Plastic Bag Alliance'
                  ,'API', 'ASPCA', 'Astorino', 'Avangrid - Arch Street Communications', 'Back to Bowling'
                  ,'BASK', 'BE FAIR TO DIRECT CARE', 'BERLINROSEN ANTI-FRAUD LAWS', 'Bet on NY'
                  ,'BLUE COLLAR COALITION', 'BP AMERICAS', 'Building & Construction Trades of Greater New York'
                  ,'Bull Moose Club', 'Business Council of Westchester', 'Butler Associates', 'Catskills Renewable Connector'
                  ,'Cats Round Table', 'CENTRO CROMINAL JUSTICE', 'CENTRO Taxpayers for Affordable New York'
                  ,'Charter Spectrum', 'child victims act GREENBERG', 'CITIZENS FOR PROGRESS', 'Claudia Tenney for Congress'
                  ,'Clean Fuels NY Kivvit', 'Clean Path NY', 'Clyde Group NY for Pest Policy', 'Coalition for the Homeless', 'Coalition to Help Families (JACK BONNER)'
                  ,'Common Cause NY', 'Community Pharmacy Association of NYS', 'COMPASSION & CHOICES'
                  ,'Congressional Candidate', 'Cruelty Free International', 'CUNY', 'CUOMO FOR GOVERNOR'
                  ,'CWA - BERLINROSEN', 'Dev Site Test By Saad', "DON'T BLOCK NY BUILDING", 'Education Equity Campaign'
                  ,'Elise Stefanik', 'EMPIRE CITY CASINO', 'Empire Report', 'Erica Video', 'Farm Bureau'
                  ,'Friends of the BQX', 'Frontier', 'FWD.us', 'GNYHA', 'Google Adwords', 'GSG Congestion Pricing'
                  ,'GSG Criminal Justice Reform', 'GSG NYABA', 'HANYS', 'Healthcare Education Project', 'House', 'HTC ADAMS IE', 'IPPNY'
                  ,'JUUL Labs', 'KWATRA', 'Linnea Empire Test', 'Long Island Association', 'Manhattan Chamber of Commerce'
                  ,'MARATHON', 'MASK UP AMERICA', 'Metropolitan Public Strategies', 'MOLINARO', 'MPAA'
                  ,'NEW YORK YANKEES', 'New Yorkers for Clean Water and Jobs', 'New Yorkers for Responsible Gaming', 'New Yorkers United for Justice', 'NY Auto Brokers Association'
                  ,'NY GAMING ASSOCIATION', 'NY HEALTH ACT', 'NY League of Conservation Voters', 'NY State Industries for the Disabled'
                  ,'NY STATE WEAR A MASK CAMPAIGN', 'NYC CHARTER SCHOOLS', 'NYS Health Foundation GSG', 'NYSANA Nurse Anesthetists'
                  ,'NYSCOP L POLITI', 'NYSCOPBA', 'NYSPSP', 'NYSUT', 'NYTHA', 'Ostroff Associates', 'PARTNERSHIP FOR NYC'
                  ,'PARTNERSHIP FOR SAFE MEDICINE', 'Patrick B. Jenkins & Associates', 'PEF', 'PHRMA', 'Project Guardianship'
                  ,'psacentral.org', 'QUEENS Chamber of Commerce', 'REALITIES OF SINGLE PAYER', 'REBNY', 'Rebuild NY Now'
                  ,'Rechler Kivvit', 'Reclaim NY', 'Retail Council', 'Rochester AAU', 'SANDS Kivvit', "Saratoga Harness Horseperson's Assocation"
                  ,'Saratoga Mentoring', 'SEIU', 'Shenker Russo Clark', 'SIEMENS', "Sizmek's services ads", 'SKD FLEXIBLE WORK'
                  ,'SKD- RESORTS WORLD CASINO', 'SMART APPROACHES TO MARIJUANA', 'Strong Leadership NYC - Eric Adams', 'SUNY Empire State College'
                  ,'The Airbnb Tax', 'The Brooklyn Hospital Center', 'TRANSPORT WORKERS UNION', 'Trucking Association of New York'
                  ,'TRUTH ABOUT ORSTED', 'United Way', 'United Way Greater Capital Region', 'VALCOUR WIND ENERGY', 'VINCENZO GARDINO', 'WAMC', 'WESTERN OTB BATAVIA DOWNS']
   

# creating and organizing tkinter window and frames within UI
root = TkinterDnD.Tk()
root.title("Report Automation")

canvas = tk.Canvas(root, height=700, width=700, bg="#4a98f0")
canvas.pack()

frame = tk.Frame(root, bg="lightgrey")
frame.place(relwidth=0.9, relheight=0.4, relx=0.05, rely=0.05)

tabControl = ttk.Notebook(frame)
tab2 = ttk.Frame(tabControl)

tabControl.add(tab2, text ='Excel Report')
tabControl.grid(padx=50)
tabControl.pack(expand = 1, fill ="both")

emailFrame = tk.Frame(root, bg="white")
emailFrame.place(relwidth=0.9, relheight=0.45, relx=0.05, rely=0.5)
emailLabel=Label(emailFrame, text="Drafted Email", bg="white")
emailLabel.place(x=10, y=5)

email = Text(emailFrame, bg="lightgrey")
email.pack(padx=40, pady=40)

data=("Current Story", "Past Story", "Ad Report")
ads=ADPROGRAMS
 
all = '/' 

##Excel Tab
AdvertiserLabel=Label(tab2, text="Program")
AdvertiserLabel.place(x=20, y=105)
nameInput = Combobox(tab2, values=ads, width=15)
nameInput.place(x=85, y=105)

# Drag and Drop Event Handling
def drop(event):
        entry_sv.set(event.data)
dir = tk.Label(tab2, text="Drop CSV file from Broadstreet Ads and select program name and info.\nFill out any remaining stats, save file, and press construct email.")
dir.pack(pady=10)
entry_sv = StringVar()
entry_sv.set('Drop Here...')
entry = Entry(tab2, textvar=entry_sv, width=80)
entry.pack(fill=X, padx=40, pady=0)
entry.drop_target_register(DND_FILES)
entry.dnd_bind('<<Drop>>', drop)


# Method to create excel file with given data, returns data needed to find file and recognize stats
addEmail = False
addLink = False
addTweets = False
title = ""
videoAds = False
totalsImpressions = 0
totalsHovers = 0
totalsClicks = 0
fileName = ""
uniqueEmail = False
folder_path = ""
def excelReport():
  global addEmail,addLink,addTweets,uniqueEmail,title,videoAds,totalsImpressions,totalsHovers,totalsClicks,fileName,folder_path,storyTitle,DATES,UNIQUEDATES
  IMPORTNAMES.clear()
  IMPORTVIEWS.clear()
  IMPORTHOVERS.clear()
  IMPORTCLICKS.clear()
  importData()
  addEmail = emailVar.get()
  addLink = linkVar.get()
  addTweets = tweetVar.get()
  email.delete(1.0, END)
  title = nameInput.get()
  videoAds, fileName = excelTab.createReport(title, IMPORTNAMES, IMPORTVIEWS, IMPORTHOVERS, IMPORTCLICKS, addEmail, addLink, addTweets, videoAds, folder_path, DATES, UNIQUEDATES, storyTitle)
  DATES = []
  UNIQUEDATES = []

# Method to construct the email being sent to client, retrieves all 
# data needed and sorts through what info is needed to add to the email draft
def constructEmail(addEmail, addLink, addTweets, title, videoAds, totalAdImp, totalAdHovers, totalAdClicks, totalEmailImp, totalEmailClicks, totalLinkImp, totalLinkClicks, totalTweetImp, totalTweetClicks, grandTotalImp, grandTotalHovers, grandTotalClicks, totalUniqueImp, totalUniqueClicks):
  email.delete(1.0, END)
  global uniqueEmail
  
  email.insert(END, "Recipient,\n\n")
  email.insert(END, "I hope that you are well!\n")
  email.insert(END, "I wanted to give you an update on the most recent campaign stats for "+ title + ":\n\n")
  # Checks for video ad vs banner ad  
  if(videoAds == True):
      email.insert(END, "Thus-far the video ads have generated " + totalAdImp + " impressions, " + totalAdHovers + " hovers, and " + totalAdClicks + " link clicks.\n")
      videoAds = False
  else:
      email.insert(END, "Thus-far the banner ads have generated " + totalAdImp + " impressions, " + totalAdHovers + " hovers, and " + totalAdClicks + " link clicks.\n")
      videoAds = False
  if((addEmail == True) and (addLink == False) and (addTweets == False)):
      # If there is a unique email, drafts different grammar and stats opposed to no unique emails
      if(uniqueEmail == True):
        email.insert(END, "The sponsored message in the daily email newsletter and unique email have generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
        email.insert(END, "The sponsored email blast alone has generated " + totalUniqueImp + " impressions and " + totalUniqueClicks + " link clicks.\n")
      else:
        email.insert(END, "The sponsored message in the daily email newsletter has generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == False) and (addLink == True) and (addTweets == False)):
      email.insert(END, "The sponsored story on Empire Report has generated " + totalLinkImp + " impressions and " + totalLinkClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == False) and (addLink == False) and (addTweets == True)):
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + totalTweetImp + " impressions and " + totalTweetClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == True) and (addLink == True) and (addTweets == False)):
      if(uniqueEmail == True):
        email.insert(END, "The sponsored message in the daily email newsletter and unique email have generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
        email.insert(END, "The sponsored email blast alone has generated " + totalUniqueImp + " impressions and " + totalUniqueClicks + " link clicks.\n")
      else:
        email.insert(END, "The sponsored message in the daily email newsletter has generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
      email.insert(END, "The sponsored story on Empire Report has generated " + totalLinkImp + " impressions and " + totalLinkClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == False) and (addLink == True) and (addTweets == True)):
      email.insert(END, "The sponsored story on Empire Report has generated " + totalLinkImp + " impressions and " + totalLinkClicks + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + totalTweetImp + " impressions and " + totalTweetClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == True) and (addLink == False) and (addTweets == True)):
      if(uniqueEmail == True):
        email.insert(END, "The sponsored message in the daily email newsletter and unique email have generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
        email.insert(END, "The sponsored email blast alone has generated " + totalUniqueImp + " impressions and " + totalUniqueClicks + " link clicks.\n")
      else:
        email.insert(END, "The sponsored message in the daily email newsletter has generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + totalTweetImp + " impressions and " + totalTweetClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")
  elif((addEmail == True) and (addLink == True) and (addTweets == True)):
      if(uniqueEmail == True):
        email.insert(END, "The sponsored message in the daily email newsletter and unique email have generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
        email.insert(END, "The sponsored email blast alone has generated " + totalUniqueImp + " impressions and " + totalUniqueClicks + " link clicks.\n")
      else:
        email.insert(END, "The sponsored message in the daily email newsletter has generated " + totalEmailImp + " impressions and " + totalEmailClicks + " link clicks.\n")
      email.insert(END, "The sponsored story on Empire Report has generated " + totalLinkImp + " impressions and " + totalLinkClicks + " link clicks.\n")
      email.insert(END, "The sponsored tweets on Empire Report's page have generated " + totalTweetImp + " impressions and " + totalTweetClicks + " link clicks.\n")
      email.insert(END, "TOTAL: " + grandTotalImp + " impressions and " + grandTotalClicks + " link clicks\n")

  email.insert(END, "Full data report is attached.\n\n")
  email.insert(END, "Thank you for working with me on this project!!\n\n")
  email.insert(END, "Best Regards,\n")
  email.insert(END, "JP Miller\n")
  email.insert(END, "Empire Report\n")
  email.insert(END, "917-565-3378")

# Initialization
totalAdImp,totalAdHovers,totalAdClicks,totalEmailImp,totalEmailClicks,totalLinkImp,totalLinkClicks,totalTweetImp,totalTweetClicks,grandTotalImp,grandTotalHovers,grandTotalClicks,totalUniqueImp,totalUniqueClicks = 0,0,0,0,0,0,0,0,0,0,0,0,0,0

# Method to retrieve new total data from excel sheet after input to assign to variables to use for email draft
def updateEmail():
  global folder_path,fileName,totalAdImp,totalAdHovers,totalAdClicks,totalEmailImp,totalEmailClicks,totalLinkImp,totalLinkClicks,totalTweetImp,totalTweetClicks,grandTotalImp,grandTotalHovers,grandTotalClicks,totalUniqueImp,totalUniqueClicks,addLink,addEmail,addTweets, uniqueEmail

  print(folder_path + "/" + fileName)
  wholePath = folder_path + "/" + fileName
  wb = xlrd.open_workbook(wholePath)
  sheet = wb.sheet_by_index(0)

  sheet.cell_value(0, 0)

  for i in range(sheet.nrows):
    # Searches for constant strings within sheet row to find correct values
    if("Advertisement" in sheet.cell_value(i, 0)):
      if(addLink == False and addEmail == False and addTweets == False):
        while("TOTAL:" not in sheet.cell_value(i, 0)):
          i = i+1
        totalAdImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        totalAdHovers = '{:,.0f}'.format(float(sheet.cell_value(i, 2)))
        totalAdClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
      else:
        while("SUBTOTAL:" not in sheet.cell_value(i, 0)):
          i = i+1
        totalAdImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        totalAdHovers = '{:,.0f}'.format(float(sheet.cell_value(i, 2)))
        totalAdClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
    if("Email" in sheet.cell_value(i, 0)):
        while("SUBTOTAL:" not in sheet.cell_value(i, 0)):
          if("Unique" in sheet.cell_value(i, 0)):
            uniqueEmail = True
            totalUniqueImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
            totalUniqueClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
          i = i+1
        totalEmailImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        totalEmailClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
    if("Link" in sheet.cell_value(i, 0)):
        i = i+1
        totalLinkImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        totalLinkClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
    if("Tweets" in sheet.cell_value(i, 0)):
        while("SUBTOTAL:" not in sheet.cell_value(i, 0)):
          i = i+1
        totalTweetImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        totalTweetClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
    if("GRAND" in sheet.cell_value(i, 0)):
        grandTotalImp = '{:,.0f}'.format(float(sheet.cell_value(i, 1)))
        grandTotalHovers = '{:,.0f}'.format(float(sheet.cell_value(i, 2)))
        grandTotalClicks = '{:,.0f}'.format(float(sheet.cell_value(i, 3)))
  constructEmail(addEmail, addLink, addTweets, title, videoAds, totalAdImp, totalAdHovers, totalAdClicks, totalEmailImp, totalEmailClicks, totalLinkImp, totalLinkClicks, totalTweetImp, totalTweetClicks, grandTotalImp, grandTotalHovers, grandTotalClicks, totalUniqueImp, totalUniqueClicks)

# Reads in data numbers from advertisement csv from broadstreet
def importData():
  with open(entry_sv.get(), 'r') as file:
    reader = csv.reader(file)
    m = 0
    for row in reader:
      if m == 0:
        m = m+1
        continue
      if row[1] != str(0):
          IMPORTNAMES.append(row[0])
          IMPORTVIEWS.append(row[1])
          IMPORTHOVERS.append(row[2])
          IMPORTCLICKS.append(row[3])


to_email = tk.StringVar()
# Sends email to recipient
def sendEmail():
  global folder_path,fileName,email
  sender_email = "ERautotesting@gmail.com"  # Enter your address
  receiving_email = to_email.get()
  password = "EmpireReport"
  subject = "Report Testing"
  body = email.get("1.0", END)

  msg = EmailMessage()
  msg['Subject'] = subject
  msg['From'] = sender_email
  msg['To'] = receiving_email
  msg.set_content(body)

  with open(folder_path + "/" + fileName, 'rb') as f:
      file_data = f.read()
  msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=fileName)

  with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
      smtp.login(sender_email, password)
      smtp.send_message(msg)

emailVar = BooleanVar()
storyTitle = StringVar()
# Add custom title for sponsored story
def addTitle():
  global storyTitle
  storyTitle.set("Enter Title..")
  Entry(tab2, textvariable=storyTitle, width=15).place(x=340, y=125)

# Opens file explorer in C drive to choose save location
def save():
  global folder_path
  folder_path = filedialog.askdirectory(initialdir='C:\\')
  dirLabel.insert(END, folder_path)


calenderPhoto = PhotoImage(file = r'images/icons8-calendar-24.png')
dateBtn = tk.Button(tab2, text="date", image=calenderPhoto, width=20, height=20, borderwidth=0, command=lambda:  calender.main(UNIQUEDATES, DATES))
dateBtn.place(x=320, y=105)

emailCheck = Checkbutton(tab2, text="Email", variable=emailVar)
emailCheck.place(x=255, y=105)
linkVar = BooleanVar()
linkCheck = Checkbutton(tab2, text="Sponsored Link", variable=linkVar, command= lambda: addTitle())
linkCheck.place(x=350, y=105)
tweetVar = BooleanVar()
tweetCheck = Checkbutton(tab2, text="Tweets", variable=tweetVar)
tweetCheck.place(x=480, y=105)

saveLabel = tk.Label(tab2, text='Choose save location prior to reporting:')
saveLabel.place(x=25, y=160)

browseBtn = tk.Button(tab2, text='Browse', command=save)
browseBtn.place(x=25, y=190)

dirLabel = tk.Text(tab2, height=1, width=25)
dirLabel.place(x=80, y=190)

excelBtn = tk.Button(tab2, text='Report', command=excelReport)
excelBtn.place(x=350, y=175)

constructBtn = tk.Button(tab2, text='Construct Email', command= lambda: updateEmail())
constructBtn.place(x=425, y=175)

toLabel = tk.Label(emailFrame, text="To Email:", bg="white")
toLabel.place(x=210, y=280)

emailEntry = tk.Entry(emailFrame, textvariable=to_email, width=18)
emailEntry.place(x=275, y=280)


sendEmailBtn = tk.Button(emailFrame, text='Send', command= lambda: sendEmail())
sendEmailBtn.place(x=465, y=280)


root.mainloop()