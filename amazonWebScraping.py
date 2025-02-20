from selenium import webdriver # type: ignore
from selenium.webdriver.common.by import By # type: ignore
from openpyxl import Workbook # type: ignore
from openpyxl import load_workbook # type: ignore
import smtplib
from email.message import EmailMessage
from email import encoders
from selenium.webdriver.chrome.options import Options
import os

# Create an instance of Options
options = Options()

# Add your desired options
options.add_argument("--headless")  # Run in headless mode
options.add_argument("--disable-gpu")  # Disable GPU usage
options.add_argument("--no-sandbox")  # Bypass OS security model
options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems

# Pass the options to the WebDriver
driver = webdriver.Chrome(options=options)

driver.maximize_window()
driver.get("https://www.amazon.in/")
driver.implicitly_wait(20)

driver.find_element(By.XPATH,"//input[contains(@id,'search') and @type='text']").send_keys("Samsung phones")
driver.find_element(By.XPATH,"//input[@value='Go']").click()
driver.find_element(By.XPATH,"//*[text()='Brands']/following::*[@class='a-icon a-icon-checkbox'][1]").click()
phoneNames = driver.find_elements(By.XPATH,"//a[contains(@class,'a-link-normal s-line-clamp-2 s-link-style a-text-normal')]/child::h2")
prices = driver.find_elements(By.XPATH,"//*[@class='a-section']//child::*[contains(@class,'price-whole')]")

phoneList = []
priceList = []
for phone in phoneNames:
	# print(phone.text)
	phoneList.append(phone.get_attribute("aria-label"))

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

# Get the current working directory
current_path = os.getcwd()
print("Current Working Directory:", current_path)
try:
    wb.save(current_path + "\FinalRecords.xlsx")
    print("Workbook saved successfully!")
except PermissionError:
    # Handle the error by using a different filename
    new_file_path = current_path+"\FinalRecords_backup.xlsx"
    try:
        wb.save(new_file_path)
        print(f"Original file could not be saved. Saved as {new_file_path} instead.")
    except PermissionError:
        print("PermissionError: Unable to save the file. Please check file permissions and ensure it is not in use by another application.")	

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
