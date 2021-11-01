from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
from email import encoders

# opt = Options()
# opt.add_argument("--headless")
# driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options = opt)
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
driver.get("https://www.amazon.in/")
driver.implicitly_wait(20)

driver.find_element(By.XPATH,"//input[contains(@id,'search') and @type='text']").send_keys("Samsung phones")
driver.find_element(By.XPATH,"//input[@value='Go']").click()
driver.find_element(By.XPATH,"//*[text()='Brand']/following::*[@class='a-icon a-icon-checkbox'][1]").click()
phoneNames = driver.find_elements(By.XPATH,"//span[contains(@class,'a-color-base a-text-normal')]")
prices = driver.find_elements(By.XPATH,"//*[@class='a-section']//child::*[contains(@class,'price-whole')]")

phoneList = []
priceList = []
for phone in phoneNames:
	# print(phone.text)
	phoneList.append(phone.text)

print("*"*50)

for price in prices:
	# print(price.text)
	priceList.append(price.text)

#Zipping lists to one list to use it for further operations	
finalList = zip(phoneList, priceList)	
# for data in list(finalList):
# 		print(data)
		
#Storing scraped data to excel file
wb = Workbook()
wb['Sheet'].title = 'Samsung Phones Data'
sh1 = wb.active
sh1.append(['Phone Name','Prices'])
for x in list(finalList):
	sh1.append(x)
wb.save("FinalRecords.xlsx")	

driver.quit()

msg = EmailMessage()
msg['Subject'] = 'Web Scraping Demo'
msg['From'] = 'Web Scraping Team'
msg['To'] = 'akshaybamne1@gmail.com'

#EmailTemplate.txt its an email template file
with open('EmailTemplate.txt') as myfile:
	data = myfile.read()
	msg.set_content(data)
	# msg.set_content('Test Email from scraped test')

# FinalRecords.xlsx is an scraped data file
with open("FinalRecords.xlsx","rb") as f:
# with open("TestExcel.xlsx","rb") as f:
	file_data = f.read()
	# print("File data in binary", file_data)
	file_name = f.name
	print("File name is", file_name)
	# msg.add_attachment(file_data, maintype = "application", subtype = "xlsx", filename = file_data)
	# encoders.encode_base64(file_data)
	msg.add_attachment(file_data,maintype="application", subtype="xlsx", filename=file_name)
	
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
	server.login("akshaybamne5@gmail.com","****Gmail Password****")
	server.send_message(msg)
	# server.quit()

print("Email sent !!!") 
