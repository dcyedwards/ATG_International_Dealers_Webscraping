#! python 3
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Script: ADA_Dealers
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the eighteenth day of the sixth month of the two-thousandth and seveenteeth year of our Lord (or just 18/06/2017)
# --------------------------------------------------------------------------------------------------------------------------------------------------
# The third script meant to scrape data from dealer/auctioneer websites for an International Dealers Database that ATG is creating.
# -------------------------------------------------------------------------------------------------------------------------------------------------

# Necessary modules (But not all used in this instance)
# ------------------------------------------------
import bs4
import requests
from bs4 import BeautifulSoup
import win32com.client as win32
import glob, os, win32com.client, pythoncom, datetime, time, getpass
import comtypes, comtypes.client
from datetime import timedelta
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.common.by import By
import selenium.webdriver.support.ui as ui
from selenium.webdriver.support.wait import WebDriverWait

# Creating time variables for naming my Excel workbook...
t = datetime.datetime.today()
today = t.strftime('%Y-%m-%d')

# Specifying the Google Chrome driver path..
# chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver"
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

# Specifying the default location for saving the workbook..
# cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                      # Initializing the workbook path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application') # Initializing the Excel workbook class in memory. 
xl.DisplayAlerts = False                                # Putting annoying alerts off.
wb = xl.Workbooks.Add()                                 # Creating an instance of the workbook
wb.SaveAs(wkbk +" ADA_Dealers - "+today+".xlsx")        # Saving it to disk
xl.Visible = True                                       # Making it visible to the user

# Adding a worksheet and naming columns...
sht = wb.Worksheets('Sheet1')
sht.Cells(1,1).Value = 'Dealer First Name'
sht.Cells(1,2).Value = 'Last name'
sht.Cells(1,3).Value = 'Compnay'
sht.Cells(1,4).Value = 'Website'
sht.Cells(1,5).Value = 'Email address'
sht.Cells(1,6).Value = 'Mailing address'
sht.Cells(1,7).Value = 'Phone number'
sht.Cells(1,8).Value = 'Location'
sht.Cells(1,9).Value = 'Specialism'

lastRow = sht.UsedRange.Rows.Count  # ...to always return the last used row
global nrow                         # ...global variable to iterate through rows.
nrow = lastRow + 1                  # ...incrementing nrow by 1 for appending new data when we get there.

global counter # variable to count results from webpage.
counter = 0    # Beginning with the first result.

br = webdriver.Chrome(chromeDriver) # Initializing an instance of Chrome
url = 'http://adadealers.com/antiques/index.php?page=search&s_res=AND&cid=20&s_by=f_sorting&a_d=asc' # Our URL of interest

br.get(url)                                                                # Browsing to the URL
dealers = br.find_elements_by_xpath(".//div[@class='col-sm-4 col-xs-6']")  # Accessing 'div' tag which contains links to all dealers.
dealer_range = range(0,len(dealers))                                       # Creating a range from the number of dealer results.
dealer_list = list(dealer_range)                                           # Converting that range to a list to iterate through each dealer.

# Function to grab dealers
def GetDealer(dealer):
    global counter            # Counter variable to give me dealer count
    global nrow               # Variable to increase row count and append data to next row
    dealers[dealer].click()   # Clicking on the relevant dealer link
    table_detail = br.find_element_by_xpath(".//table[@class='table table-striped tcart']")  # Finding the main dealer table in the HTML
    details = table_detail.find_elements_by_xpath(".//tbody/tr")                             # Dealer links in the table body
    company = br.find_element_by_xpath(".//div[@class='col-md-5 col-sm-5']/h2")              # Getting the company name
    print(company.text)
    sht.Range('C'+str(nrow)).Value = str(company.text)                                       # I meant, getting the company name as text
    sht.Range('H'+str(nrow)).Value = 'USA'                                                   # Sticking that in Excel
    try:                     # Very important to trap errors and work with lists.            # Doing the same for other data of interest
        contact_name = details[0].text
        print(contact_name)
        sht.Range('A'+str(nrow)).Value = str(contact_name)                                   # The dealer's name
    except NoSuchElementException:
        print('N/A')
        sht.Range('A'+str(nrow)).Value = 'N/A'
        pass
    try:
        address = details[1].text                                                            # The dealer's address
        print(address)
        sht.Range('F'+str(nrow)).Value = str(address)
    except NoSuchElementException:
        print('N/A')
        sht.Range('F'+str(nrow)).Value = 'N/A'
        pass
    try:
        phone = details[2].text                                                              # The dealer's phone number
        print(phone)
        sht.Range('G'+str(nrow)).Value = str(phone)
    except NoSuchElementException:
        print('N/A')
        sht.Range('G'+str(nrow)).Value = 'N/A'
        pass
    try:
        email = details[3].text                                                              # The dealer's email address
        print(email)
        sht.Range('E'+str(nrow)).Value = str(email)
    except NoSuchElementException:
        print('N/A')
        sht.Range('E'+str(nrow)).Value = 'N/A'
        pass
    try:
        website = details[4].text                                                           # The dealer's website
        print(website)
        sht.Range('D'+str(nrow)).Value = str(website)
    except NoSuchElementException:
        print('N/A')
        sht.Range('D'+str(nrow)).Value = 'N/A'
    try:
        expertise = br.find_element_by_xpath(".//div[@class='tab-pane active']")            # The dealer's expertise
        specialty = expertise.text.split('Contact:',1)
        print(specialty)
        sht.Range('I'+str(nrow)).Value = str(specialty)
    except NoSuchElementException:
        print('N/A')
        sht.Range('I'+str(nrow)).Value = 'N/A'
    br.execute_script("window.history.go(-1)")                                               # Returning to the previous page and continuing the loop.
    print(counter)
    print('\nNext...')
    nrow+=1
    counter+=1
    

start_time = time.monotonic()

for dealer in dealer_list:
    dealers = br.find_elements_by_xpath(".//div[@class='col-sm-4 col-xs-6']")
    GetDealer(dealer)                                                                       # Calling our GetDealer function


# Tidying and wrapping up
# -----------------------
br.close()
wb.Save()
xl.DisplayAlerts = True
wb.Close(True)
end_time = time.monotonic()
print('Done')
print(timedelta(seconds=end_time - start_time))

