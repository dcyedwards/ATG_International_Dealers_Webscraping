#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------
# Script: winterantiqueshow_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the twentieth day of the sixth month of the two-thousandth and seventeenth year of our Lord (or just 20/06/2017)
# ------------------------------------------------------------------------------------------------------------------------------------
# The sixth of eight scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves comany names, dealer names, websites, email addresses and address details of 68 dealers in about 4 minutes.
# ------------------------------------------------------------------------------------------------------------------------------------

# Modules and stuff
# -----------------
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

t = datetime.datetime.today()     # Our timer variable
today = t.strftime('%Y-%m-%d')    # The way we want to see it

# Our file path to our Google Chrome driver
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver"
#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

# Specifying the location to our soon-to-be-saved Excelw workbook
cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
#cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                               # Initializing the file path in memory
xl = win32.gencache.EnsureDispatch('Excel.Application')          # Initializing the Excel workbook class in memory
xl.DisplayAlerts = False                                         # Putting alerts off
wb = xl.Workbooks.Add()                                          # Adding a workbook to our instance of Excel
wb.SaveAs(wkbk +" Winterantiquesshow_Dealers - "+today+".xlsx")  # Naming and saving our workbook
xl.Visible = True                                                # Making it visible so that we see what's going on

# Adding worksheets and naming columns
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

lastRow = sht.UsedRange.Rows.Count                             # Our Excel last row variable so that we know the last used row
global nrow                                                    # My variable to indicate the next row in Excel...
nrow = lastRow + 1                                             # ...which is always the next row after the last used range of rows

url = 'http://winterantiquesshow.com/exhibitor/'
br = webdriver.Chrome(chromeDriver)
br.get(url)

archive = br.find_element_by_xpath(".//div[@id='content']")
dealers = archive.find_elements_by_tag_name('a')
Dealer_range = range(23,len(dealers))
Dealer_List = list(Dealer_range)

global counter
counter = 0

def GetDealer(dealer):
    global counter
    global nrow
    dealers[dealer].click()
    try:
        company = br.find_element_by_xpath(".//h1[@class='exhibitor-name text-uppercase']")
        print(company.text)
        sht.Range('C'+str(nrow)).Value = str(company.text)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('C'+str(nrow)).Value = 'N/A'
        pass
    try:
        address = br.find_element_by_xpath(".//div[@class='exhibitor-address']")
        print(address.text)
        sht.Range('F'+str(nrow)).Value = str(address.text)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('F'+str(nrow)).Value = 'N/A'
        pass
    try:
        contact = br.find_element_by_xpath(".//div[@class='exhibitor-contact']")
        contact_names = contact.text.split()
        first_name = contact_names[0]
        last_name = contact_names[1]
        print(first_name)
        print(last_name)
        sht.Range('A'+str(nrow)).Value = str(first_name)
        sht.Range('B'+str(nrow)).Value = str(last_name)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('A'+str(nrow)).Value = 'N/A'
        sht.Range('B'+str(nrow)).Value = 'N/A'
        pass
    try:
        website = br.find_element_by_xpath(".//div[@class='exhibitor-website']")
        print(website.text)
        sht.Range('D'+str(nrow)).Value = str(website.text)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('D'+str(nrow)).Value = 'N/A'
        pass
    info = br.find_element_by_xpath(".//div[@class='exhibitor-text']")
    try:
        specialty = info.find_elements_by_tag_name('p')
        expertise = specialty[-1].text
        print(expertise)
        sht.Range('I'+str(nrow)).Value = str(expertise)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('I'+str(nrow)).Value = 'N/A'
        pass
    br.execute_script("window.history.go(-1)")
    print(counter)
    print('\nNext...')
    nrow+=1
    counter+=1

start_time = time.monotonic()

for dealer in Dealer_List[::3]:
    archive = br.find_element_by_xpath(".//div[@id='content']")
    dealers = archive.find_elements_by_tag_name('a')
    GetDealer(dealer)

br.close()
wb.Save()
xl.DisplayAlerts = True
wb.Close(True)
end_time = time.monotonic()
print('Done')
print(timedelta(seconds=end_time - start_time))

