import gspread
import statistics
import math
import pandas as pd
import numpy

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
