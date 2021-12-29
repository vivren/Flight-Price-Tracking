from datetime import datetime, timedelta
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import gspread

#Chrome Web Scrapping Set Up
path = "C:\Program Files (x86)\chromedriver.exe"
#driver = webdriver.Chrome(path)
driver = webdriver.Chrome(ChromeDriverManager().install())
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


worksheets = ["LAX-SFO","MIA-JFK","ORD-JFK","ORD-LAX","SFO-JFK"]
gc = gspread.service_account_from_dict(credentials)
sh = gc.open("Data")
ws = sh.get_worksheet(1)

#Current Date Set Up
currentDate = datetime.today()
formattedDate = currentDate.strftime("%m/%d/%Y")
dayOfWeek = datetime.today().strftime('%A')


def getDepartureDate(daysToDeparture):
	departureDate = currentDate+timedelta(daysToDeparture)
	departureDate = departureDate.strftime('%Y%m%d')
	openLink(departureLocation,arrivalLocation,departureDate)


def openLink(departureLocation,arrivalLocation,departureDate):
	driver.get(f"https://secure.flightcentre.ca/search/{departureLocation}/{arrivalLocation}/{departureDate}/1/0/0/ECONOMY")
	time.sleep(13)
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


def writeToSpreadSheet(totalPrice):
	ws.update_cell(empty,1,departureLocation)
	ws.update_cell(empty,2,arrivalLocation)
	ws.update_cell(empty,3,formattedDate)
	ws.update_cell(empty,4,dayOfWeek)
	ws.update_cell(empty,5,daysToDeparture[j])
	ws.update_cell(empty,6,totalPrice)

#To Run the Program
daysToDeparture = [2,7,30,60,91,182,330]

for i in range(0,5):	
	ws = sh.worksheet(worksheets[i])
	departureLocation = ws.acell('A2').value
	arrivalLocation = ws.acell('B2').value
	emptys = list(filter(None, ws.col_values(1)))
	empty = len(emptys)+1
	for j in range(0,len(daysToDeparture)):
		getDepartureDate(daysToDeparture[j])
		empty+=1


print("done")

