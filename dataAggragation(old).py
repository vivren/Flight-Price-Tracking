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
  "private_key_id": "d76a7c653096db915c115387b835a027d9d5f938",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDUUN51iy0A/dC3\n89RRutZasdWrwinjn+iEYfjNFd6+RstlkZ9fVsTW113jhbuORv/Yu4Hsbv63eCvR\n+K+/Y6Kh8bzp/SJTmBz90nCgRY/28fEMRqqDwwDUCgOaIpTnJSe4a5X586//5ZB+\nFxuLQvHi3U2hvhbE7Tz9prOBWHkBSYmW4fpR2t8UUdv3EZc+CCfrylf+Jhahr2CP\nMOAxz+mfG4ZzYB/1dUEwy9ubztXNGfyCJ4z3+ZsSmE8ftJ2+dcAmmYz3V4LUgaj3\npH2ARCElI1be+j2LmLFDmpMUQYo/tGWi3R1j7q8/z9dh5GjLh9DzoVFPEje2yLZM\nn77BOJiLAgMBAAECggEAS8yhV1q8UvXK3dm9y/4FyDRkQeFcfikZZKXroyBzOQBi\nXl0qhmmedctG+FNG3ilH7zMmU+hepNbQ4jJEYHJsXh/FX85hQnH0q8PFNNsQ5kuD\nUOwrtUXZ1lnK5m0BNzKjJyq1zRpsk9H8D6AlU8cvP1zd+eg5Dc5a733j6UJSVRwq\n2QmX64kcrnZCRBRlhm9y96NoAS3WK1AnBdx+QbH4PsCYT5TjOPxAUrjVll0iHsj+\n65UBfQEbdbYBmguv7FyeOjviC0+UmErAJn4HjxdCENAqNFFRrZvWT1EUzaXE3Zzk\nB9OBMU9ulMZJ38K1JOevK2KGLPEifEkO4Fhmq3VbiQKBgQDy35nIGoPy+BUQVNSI\nLfLT214i4Zuo8WaVOC7thk8rYektyR7kAfgcqHoYEi9uoeVlsTC0jUaj3WubuCSD\ne9C26yewagg4zMg1GwEWiM8VA5me13GFIwxp8TA4okh3qnv9oDPn4BF9vM4udoRa\n7j06Lb8g1sAOkx823MamDv1wmQKBgQDfynlI8gU0XFJR7fTOelnys6FssUujRa3q\ne/efnLaUpKNJopwxEkRhka2QUaqO0Mu90kArAeQU68DMonv8vXSN05Jrf5wb/vIY\ntl511EFlpvtDLnwnfbrwV5Hr3TKA59zpyenP3dmHfI1x3LDn/sYT2+2guoBwGalG\n9tg/k+r0wwKBgGZscbf3pqmygX+Ppsl/RyadHofrSO/oTfCj7vZwa7bJ5cJPTN2v\nDulXC5ZXJPWtBMbsALRD2ASG/jh/YbapYo2hge1d6fW6NrxelQjhzSL5j7Fy/ga9\njyCmfEWF+rWsifmdtAYXcojqNLFXfd2zFtMQM8gviHMdDE7gsE2biLoRAoGAbwGD\n/8btBff1bWZXVEQkcdrga4XWOvrJNdKgFbAqcLjpthDO4Rhqiusz6K5Zp9Wx+kpc\nQfkCJc57KAZA8jEXq3IS4ve9e7WOaOutF01d0wps//oC46PeInGNlC14a7CXR/A7\n5jvpNud1UdFifvFFV3xz9pIMO46/BNBUm9THavkCgYEA1311grq52YEEhvHIq1A3\n3oq666hUKzpAN0/RG1JWjdo9Jrdn+9uFV9i7kIeeAm7B977AgfrM6kbatOMGrEd0\nC84xGBxpDBmYq+hNQUcVdEzHbWTwtFhredwrPeYnd3eCBbYL9y6i337L+QmQKySQ\nbVrOPXqqpmGWzeB5u9vRJr4=\n-----END PRIVATE KEY-----\n",
  "client_email": "vacation-pricing@vacation-pricing.iam.gserviceaccount.com",
  "client_id": "103815442999153432258",
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
