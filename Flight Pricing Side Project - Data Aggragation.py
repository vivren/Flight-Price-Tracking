#To Do
#add summary of data (overall order of days before/weekday in terms of price)
#track a flight on a set day every day (flight on sept 1, 2021) see rpice change every day
#add hotel pricing
import gspread
import statistics
import math

#Spreadsheet Set Up
credentials = {
  "type": "service_account",
  "project_id": "vacation-pricing",
  "client_email": "vacation-pricing@vacation-pricing.iam.gserviceaccount.com",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/vacation-pricing%40vacation-pricing.iam.gserviceaccount.com"}

worksheets = ["ORD-JFK","LAX-SFO","SFO-JFK","ORD-LAX","MIA-JFK"]
gc = gspread.service_account_from_dict(credentials)
sh = gc.open("Data")


#First Empty Row
worksheet = sh.worksheet('ORD-JFK')
emptys = list(filter(None, worksheet.col_values(1)))
empty = int(len(emptys)+1)


#Data for Avg Price/Num of Days Before and Avg Price/Weekday
daysBefore = []
weekday = [[],[],[],[],[]]
for i in range(0,5):	
  ws = sh.worksheet(worksheets[i])
  daysBefore.append(ws.batch_get([f'C4:C{empty}', f'D4:D{empty}',f'E4:E{empty}',f'F4:F{empty}',f'G4:G{empty}',f'H4:H{empty}',f'I4:I{empty}']))
  for j in range(4,empty):
    weekday[i].append(ws.get(f'B{j}:G{j}')[0])


#Switch to Data Aggregation Worksheet
ws = sh.worksheet("Data")


#Days Before 
def newDaysBeforeData():
  daysBeforeFormatted = []
  for flights in daysBefore:
    for numOfDays in flights:
      flatList = []
      for priceList in numOfDays:
        for prices in priceList:
          flatList.append(prices)
      daysBeforeFormatted.append(flatList)
  
  avgPriceDaysBefore(daysBeforeFormatted)


def avgPriceDaysBefore(daysBeforeFormatted):
  avgs = []
  for numOfDays in daysBeforeFormatted:
    prices = []   
    for price in numOfDays:
      prices.append(float(price))
    avgs.append(statistics.mean(prices))

  updateAvgPriceDaysBefore(avgs)


def updateAvgPriceDaysBefore(avgs):
  for i in range(0,5):
    for j in range(0,7):
      ws.update_cell(j+2,i+2,f'{avgs[i*7+j]:.2f}')

  lowestHighestDaysBefore()


def lowestHighestDaysBefore():
  for i in range(2,7):
    avgs = ws.col_values(i)
    avgs.sort()
    avgs.pop(-1)
    avgs = list(map(float, avgs))
    avgs.sort()

    for j in range(0,3):


      currents = ws.findall(str(avgs[j]))

      for cell in currents:
        if cell.col == i:
          colValue = chr(ord('@')+(cell.col))
          ws.format(f'{colValue}{cell.row}',{"backgroundColor":{"green":f'0.{(j*2)+3}'}})
    for k in range(1,4):

      currents = ws.findall(str(avgs[-k]))

      for cell in currents:
        if cell.col == i:
          colValue = chr(ord('@')+(cell.col))
          ws.format(f'{colValue}{cell.row}',{"backgroundColor":{"red":f'0.{(k*2)+2}'}})

  newWeekdayData()


#Price Per Weekday

def newWeekdayData():
  for flights in weekday:
    for entry in flights:
      del entry[1:3]
  
  weekdayDataFormatted()


def weekdayDataFormatted():
  currentWeekday = ""
  weekdayFormatted = [[],[],[],[],[]]
  for h in range(0,5): #for flights in weekday:
    existingWeekdays = []
    for i in range(0,len(weekday[h])): #for entry in flights
      if weekday[h][i][0] not in existingWeekdays:
        weekdayIndexes = []
        currentWeekday = weekday[h][i][0]
        existingWeekdays.append(weekday[h][i][0])
        for j in range(0,len(weekday[h])):
          if weekday[h][j][0] == currentWeekday:
            weekdayIndexes.append(j)

        for k in range(1,4):
          prices = []
          for index in weekdayIndexes:
            prices.append(float(weekday[h][index][k]))
          weekday[h][weekdayIndexes[0]][k] = statistics.mean(prices)
        weekdayFormatted[h].append(weekday[h][weekdayIndexes[0]])

  avgPricePerWeekday(weekdayFormatted)


def updateavgPricePerWeekday(weekdayFormatted):
  weekdaysValues = ws.get('H3:H9')
  weekdays = [val for sublist in weekdaysValues for val in sublist]
  for i in range(0,5): #for flights in weekdayFormatted
    for entries in weekdayFormatted[i]:
      newEntry = []
      newEntry.append(entries[1:4])
      weekdayIndex = weekdays.index(entries[0])
      colValue1 = chr(ord('@')+(9+(i*3)))
      colValue2 = chr(ord('@')+(9+((i+1)*3))-1)
      ws.update(f'{colValue1}{weekdayIndex+3}:{colValue2}{weekdayIndex+3}',newEntry)

  lowestHighestWeekday()


def lowestHighestWeekday():
  for i in range(9,24):
    avgs = ws.col_values(i)
    avgs.sort()
    avgs.pop(-1)
    avgs = [a for a in avgs if a != ""]
    avgs = list(map(float, avgs))
    avgs = list(filter(lambda b: b%30!=0,avgs))
    avgs.sort()
  
    for j in range(0,math.floor(len(avgs)/2)):
      currents = ws.findall(str(avgs[j]))
      for cell in currents:
        if cell.col == i:
          colValue = chr(ord('@')+(cell.col))
          ws.format(f'{colValue}{cell.row}',{"backgroundColor":{"green":f'0.{(j*2)+3}'}})
    for k in range(1,(math.floor(len(avgs)/2))+1):
      currents = ws.findall(str(avgs[-k]))
      for cell in currents:
        if cell.col == i:
          colValue = chr(ord('@')+(cell.col))
          ws.format(f'{colValue}{cell.row}',{"backgroundColor":{"red":f'0.{(k*2)+2}'}})


newDaysBeforeData()
