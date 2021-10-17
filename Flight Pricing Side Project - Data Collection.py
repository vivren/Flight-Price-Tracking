from datetime import datetime, timedelta
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
import gspread

#Chrome Web Scrapping Set Up
path = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(path)
actionchains = ActionChains(driver)


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
ws = sh.get_worksheet(1)

#Current Date Set Up
currentDate = datetime.today()
formattedDate = currentDate.strftime("%m/%d/%Y")
dayOfWeek = datetime.today().strftime('%A')


#Locations Set Up
departureLocation = ws.acell('A2').value
arrivalLocation = ws.acell('B2').value


def getDepartureDate(daysToDeparture):
	departureDate = currentDate+timedelta(daysToDeparture)
	departureDate = departureDate.strftime('%Y%m%d')
	openLink(departureLocation,arrivalLocation,departureDate)


def openLink(departureLocation,arrivalLocation,departureDate):
	driver.get(f"https://secure.flightcentre.ca/search/{departureLocation}/{arrivalLocation}/{departureDate}/1/0/0/ECONOMY")
	time.sleep(15)
	getData()


def getData():
	#oneWay = driver.find_element_by_class_name("jss515")
	#oneWay.click()

	prices = driver.find_elements_by_xpath("//*[contains(@class,'test-priceWholeValue')]")
	price = prices[0].text
	cents = driver.find_elements_by_xpath("//*[contains(@class,'test-priceCentsValue')]")
	cent = cents[0].text
	totalPrice = float(f'{price}.{cent}')
	writeToSpreadSheet(totalPrice)
	print(totalPrice)


def writeToSpreadSheet(totalPrice):
	ws.update_cell(empty,1,formattedDate)
	ws.update_cell(empty,2,dayOfWeek)
	ws.update_cell(empty,i,totalPrice)


#To Run the Program
for i in range(0,5):	
	ws = sh.worksheet(worksheets[i])
	emptys = list(filter(None, ws.col_values(1)))
	empty = str(len(emptys)+1)
	for i in range(3,10):
		daysToDeparture = int(ws.cell(2,i).value)
		getDepartureDate(daysToDeparture)


print("done")

