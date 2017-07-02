#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------------
# Script: CINOA_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the twelfth day of the sixth month of the two-thousandth and seveenteeth year of our Lord (or just 12/06/2017)
# -------------------------------------------------------------------------------------------------------------------------------------------
# The first amongst a series of scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves comany names, addresses, websites, phone numbers, email addresses and specializations of over 3,122 dealers in about 100 minutes.
# -------------------------------------------------------------------------------------------------------------------------------------------

# Necessary modules (But not all used in this instance)
# -----------------------------------------------------
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

t = datetime.datetime.today()              # Creating my timer variable
today = t.strftime('%Y-%m-%d')             # Formatting the time output

#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"


br = webdriver.Chrome(chromeDriver)           # Starting an instance of Google Chrome
url = 'https://www.cinoa.org/cinoa/dealers?'  # Our CINOA dealers URL

#cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                       # Initializing the workbook path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application')  # Initializing the Excel workbook class in memory.
xl.DisplayAlerts = False                                 # Putting annoying alerts off.
wb = xl.Workbooks.Add()                                  # Creating an instance of the workbook
wb.SaveAs(wkbk +" CINOA_Dealers - "+today+".xlsx")       # Saving it to disk
xl.Visible = True                                        # Making it visible to the user

# Adding worksheets and naming columns
sht = wb.Worksheets('Sheet1')
sht.Cells(1,1).Value = 'Dealer Name'
sht.Cells(1,2).Value = 'Company'
sht.Cells(1,3).Value = 'Website'
sht.Cells(1,4).Value = 'Mailing Address'
sht.Cells(1,5).Value = 'Phone no.'
sht.Cells(1,6).Value = 'Specialization'

lastRow = sht.UsedRange.Rows.Count  # My last row Excel variable
global nrow                         # My 'next row' Excel variable
nrow = lastRow + 1                  # Incrementing my nrow variable by 1 for data entry onto the next row


br.get(url)                         # Navigating to our URL

dealers = br.find_elements_by_link_text('VIEW DEALER')    # The list of all dealers in the HTML of the current page.
dealer_list = range(0, len(dealers))                      # Converting that list into a Python range object

pg = 1      # A page counter for iterating through pages on the website's HTML


# Function to iterate through pages on the website HTML
# -----------------------------------------------------
def Pager(pg):  # Using the page number as the function's argument
    next_page = br.find_element_by_xpath(".//a[@href='/cinoa/dealers?page=%s']" %pg).click() # ...and clicking on the next page
    print('Moving to next page...')

# Function to scrape dealer information per page
# ----------------------------------------------
def GetDealer(dealer): # Passing the dealer number from our dealer_list range object as the function's arugment
    try:               # Error handling for when things go wrong (which they will)...also, lots of error handling
        global nrow
        br.implicitly_wait(8)               # Selenium is so fast at browsing that I have to slow it down a bit for the elements to have a chance of generating in the HTML
        link = dealers[dealer].click()      # Anyway, click on the relevant dealer's link...
        print("|------------------------|")
        try:
            company_name = br.find_element_by_xpath(".//h1[@class='text-center']")   # ...get their company name,
            sht.Range('B'+str(nrow)).Value = str(company_name.text)                  # their text attribute
            print(company_name.text)
        except NoSuchElementException:
            print("N/A")
            company_name = "N/A"
            sht.Range('B'+str(nrow)).Value = company_name                            # Paste the company_name into Excel
            pass
        try:
            address = br.find_element_by_xpath(".//div[@class='dealer-address']")    # Get the dealer's address
            sht.Range('D'+str(nrow)).Value = str(address.text)                       # Stick the address into Excel
            print(address.text)
        except NoSuchElementException:
            print("N/A")
            address = "N/A"
            sht.Range('D'+str(nrow)).Value = address                                 
            pass
        try:      
            website = br.find_element_by_xpath(".//div[@class='web']")               # We want their website too
            sht.Range('C'+str(nrow)).Value = str(website.text)                       # Stick it into Excel
            print(website.text)
        except NoSuchElementException:
            print("N/A")
            website = "N/A"
            sht.Range('C'+str(nrow)).Value = website
            pass
        try:        
            phone = br.find_element_by_xpath(".//div[@class='phone']")               # Can't foreget their phone number now, can we?
            sht.Range('E'+str(nrow)).Value = str(phone.text)                         # Let's put that too in Excel
            print(phone.text)
        except NoSuchElementException:
            print("N/A")
            phone = "N/A"
            sht.Range('E'+str(nrow)).Value = phone
            pass
        try:
            br.implicitly_wait(2)                                                    # Pausing for just a bit to allow the webpage to catch up with the driver
            contact = br.find_element_by_xpath(".//div[@class='contact-name text-uppercase']")  # Alright, now we want their contact name
            sht.Range('A'+str(nrow)).Value = str(contact.text)                                  # And we want it in Excel
            print(contact.text)
        except NoSuchElementException:
            print("N/A")
            contact = "N/A"
            sht.Range('A'+str(nrow)).Value = contact
            pass
        try:
            specialization = br.find_element_by_xpath(".//div[@class='members-specialties text-center']") # It's also helpful to know what they specialize in
            sht.Range('F'+str(nrow)).Value = str(specialization.text)                                     # And of course we want to store this in Excel no matter how much they do
            print(specialization.text)
        except NoSuchElementException:
            print("N/A")
            specialization = "N/A"
            sht.Range('F'+str(nrow)).Value = specialization
            pass
        nrow +=1    # Increasing our Excel row counter by 1 for the next dealer's data
    except IndexError:                                       # And when we get to the end of all pages in the list, we get an error...
        print('Done: End of list')                           # So we know we're done
        end_time = time.monotonic()                          # Stopping my timer and wrapping things up below.
        print(timedelta(seconds=end_time - start_time))
        wb.Save()
        xl.DisplayAlerts = True
        wb.Close(True)
        raise SystemExit

start_time = time.monotonic()             # Starting my time. Don't get confused. This line is here because the lines below actually execute our GetDealer function

while pg < 106:                           # Do this for all 105 pages because there aren't any more.
    try:
        for dealer in dealer_list:        # Sourcing our dealer numbers from our dealer_list range object.
            dealers = br.find_elements_by_link_text('VIEW DEALER')    # Refreshing the HTML elements dealers list to prevent the 'Stale Element' error in the code 
            GetDealer(dealer)                                         # The acutal line that executes the GetDealer function
            br.execute_script("window.history.go(-1)")                # Once dealer details have been read from page, return to the menu
            if dealer == len(dealer_list)-1:                          # Some logic to mark the page breaker
                pg+=1                                                 # Increasing the page marker pg by 1
                Pager(pg)                                             # Calling our Pager function and passing the pg argument to it.
    except IndexError:                                                # When there's an index error, you'll know you've run out of pages to turn...
        print('Done: End of list')                                    # ...in which case we'll be done with getting all the info we need.
        end_time = time.monotonic()
        print(timedelta(seconds=end_time - start_time))
        wb.Save()
        xl.DisplayAlerts = True
        wb.Close(True)
        raise SystemExit

