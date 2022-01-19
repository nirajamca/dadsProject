from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
import time
import openpyxl
from selenium.webdriver.common.keys import Keys

# Ask user for month and year
inputYear = input("Year: ")
inputMonth = input("Month: ")

# Set driver to use Chrome
chromedriver_autoinstaller.install()
myDriver = webdriver.Chrome()
myDriver.maximize_window()

# Launch URL to open the desired website
myDriver.get("https://dwcdataportal.fldfs.com/ProofOfCoverage.aspx")

# First drop down menu, select "Expiration Date"
expirationDate = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddlSearchOptions"]/option[4]')
expirationDate.click()

# Select year user typed
match inputYear:
    case "2023":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[1]')
    case "2022":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[2]')
    case "2021":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[3]')
    case "2020":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[4]')
    case "2019":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[5]')
    case "2018":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[6]')
    case "2017":
        selectYear = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddYear"]/option[7]')

# Click respective year
selectYear.click()

# Select month user typed
match inputMonth:
    case "JAN":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[1]')
    case "FEB":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[2]')
    case "MAR":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[3]')
    case "APR":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[4]')
    case "MAY":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[5]')
    case "JUN":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[6]')
    case "JUL":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[7]')
    case "AUG":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[8]')
    case "SEP":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[9]')
    case "OCT":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[10]')
    case "NOV":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[11]')
    case "DEC":
        selectMonth = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddMonth"]/option[12]')

# Click respective year
selectMonth.click()

# Select FLORIDA WC JOINT ASSOC
selectInsurer = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ddInsurer"]/option[393]')
selectInsurer.click()

# Click search button
clickSearch = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnSearch2"]')
clickSearch.click()

# Refresh page to grab the of the "Export to Excel" button
myDriver.refresh()

# Click on "Export to Excel" button
clickExportToExcel = myDriver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_Button2"]')
clickExportToExcel.click()

# Set timer for window to 15 seconds
time.sleep(15)

# Close window when done
myDriver.quit()

# Workbook object
wb_obj = openpyxl.load_workbook("/Users/richardthomas/Downloads/ProofOfCoverageReport.xlsx")

# Sheet object created
sheet_obj = wb_obj.active

# Test loop to see if I can add elements to the workbook
for i in range(22, 29):
    sheet_obj.cell(row=1, column=i).value = "Agent"
