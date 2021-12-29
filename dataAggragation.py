import gspread
import statistics
import math
import pandas as pd
import numpy

#Spreadsheet Set Up
credentials = {
  "type": "service_account",
  "project_id": "vacation-pricing",
  "client_email": "vacation-pricing@vacation-pricing.iam.gserviceaccount.com",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/vacation-pricing%40vacation-pricing.iam.gserviceaccount.com"}

worksheets = ["LAX-SFO","MIA-JFK","ORD-JFK","ORD-LAX","SFO-JFK"]
gc = gspread.service_account_from_dict(credentials)
sh = gc.open("Data")

#First Empty Row/Column
ws = sh.worksheet("LAX-SFO")
emptyRows = list(filter(None, ws.col_values(1)))
emptyRow = int(len(emptyRows)+1)
emptyCols = list(filter(None, ws.row_values(2)))
emptyCol = int(len(emptyCols)+1)

#Set up Pandas Dataframe
df = pd.DataFrame()

#Read Data into DataFrame
for i in range(0,5):
  ws = sh.worksheet(worksheets[i])
  if i==0:
    values = ws.get_all_values()[1:]
  else:
    values = ws.get_all_values()[1:]
  df = df.append(values)


#Swtich into Data Aggregation Spreadsheet
ws = sh.worksheet("Data")

#Find Averages per Days in Advance
def daysinAdvanceAvgs():
  avgs = []
  daysInAdvance = df[4].unique()
  for days in daysInAdvance:
    tempData = df[df[4] == days]
    #avgsTable = tempData.groupby([0, 1])[5].mean()
    avgsTable = tempData.groupby([0, 1]).mean()
    avgs.append(avgsTable[5].tolist())

  updateDaysinAdvance(avgs)


#Write Averages per Days in Advance to Spreadsheet
def updateDaysinAdvance(avgs):
  for i in range(0, len(avgs)):
    for j in range(0,len(avgs[i])):
      ws.update_cell(i+2, j+2, f'{avgs[i][j]}')

  colourDaysinAdvance(avgs)


#Colour Code the Averages per Days in Adavance
def colourDaysinAdvance(avgs):
  for i in range(0,len(avgs)):
    temp = []
    for j in range(0,len(avgs[0])):
      temp.append(avgs[i][j])
    temp.sort()

    for k in range(0,len(temp)//2):
      currents = ws.findall(str(avgs[k]))
      for cell in currents:
        if cell.col == i+2:
          colValue = chr(ord('@') + (cell.col))
          ws.format(f'{colValue}{cell.row}', {"backgroundColor": {"green": f'0.{(k * 2) + 3}'}})

    for l in range(1, math.ceil(len(temp)/2)):
      currents = ws.findall(str(avgs[-l]))
      for cell in currents:
        if cell.col == i:
          colValue = chr(ord('@') + (cell.col))
          ws.format(f'{colValue}{cell.row}', {"backgroundColor": {"red": f'0.{(l * 2) + 2}'}})

  ws.format("C2:B2", {
    "backgroundColor": {
      "red": 3.0,
      "green": 0.0,
      "blue": 0.0
    }})




daysinAdvanceAvgs()

print("done")
