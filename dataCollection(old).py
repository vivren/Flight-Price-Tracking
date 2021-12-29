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
  "client_email": "vacation-pricing@vacation-pricing.iam.gserviceaccount.com",
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

