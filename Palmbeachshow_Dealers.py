#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------
# Script: Palmbeachshow_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the twentieth day of the sixth month of the two-thousandth and seventeenth year of our Lord (or just 20/06/2017)
# ------------------------------------------------------------------------------------------------------------------------------------
# The seventh of eight scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves comany names, dealer names, websites, email addresses and address details of 161 dealers in about 15 minutes.
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

t = datetime.datetime.today()    # Our timer 
today = t.strftime('%Y-%m-%d')   # And how we want to see our time

chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

# Specifying our Excel file path (the actual file will be created soon)
cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
#cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                             # Creating the file path in memory
xl = win32.gencache.EnsureDispatch('Excel.Application')        # Initializing an instance of Excel
xl.DisplayAlerts = False                                       # We don't like display alerts (at least not now anyway)
wb = xl.Workbooks.Add()                                        # Adding a workbook to our Excel session
wb.SaveAs(wkbk +" Palmbeachshow.com_Dealers - "+today+".xlsx") # Saving it to disk
xl.Visible = True                                              # Making it visible to the user - you

sht = wb.Worksheets('Sheet1')                                  # Adding sheets and naming columns and stuff 
sht.Cells(1,1).Value = 'Dealer First Name'
sht.Cells(1,2).Value = 'Last name'
sht.Cells(1,3).Value = 'Compnay'
sht.Cells(1,4).Value = 'Website'
sht.Cells(1,5).Value = 'Email address'
sht.Cells(1,6).Value = 'Mailing address'
sht.Cells(1,7).Value = 'Phone number'
sht.Cells(1,8).Value = 'Location'
sht.Cells(1,9).Value = 'Specialism'

lastRow = sht.UsedRange.Rows.Count                             # Our variable to tell us the last used row of data in Excel 
global nrow                                                    # Another variable to serve as a row counter
nrow = lastRow + 1                                             # Naturally, this variable must be global to work within the -  
                                                               # - scope of the function which follows soon.
url = 'https://www.palmbeachshow.com/exhibitors-list/'         # Our URL of interest
br = webdriver.Chrome(chromeDriver)                            # Starting up Chrome   
br.get(url)                                                    # Navigating to our website

dealer_list = br.find_element_by_xpath(".//div[@class='ex-listn ']") # Getting the list of the parent element that contains the elements we want
dealer_list_EX = dealer_list.find_elements_by_xpath(".//ul/li/a")    # Getting all 'a' tags that are descendants of the dealer_list 'div' tag above
len(dealer_list_EX)                                                  # Just checking to make sure we have a sizeable number otherwise we'll be scraping nothing  
all_dealers = range(0,len(dealer_list_EX))                           # Creating a range object from which we'll pass values to our GetDealer function below

global counter  # A coutner variable to count our dealers as we scrape them
counter = 0     # Setting it to zero (which is actually the first in programming)

# Our function which does all the dirty work
def GetDealer(dealer):
    global counter  # Remember our counter variable?
    global nrow     # And our Excel variable to increment rows?
    br.get(url)     # In this case, rather than going to the previous page in the browser, I simply 'requery' the URL because it's easier and I'm lazy.
    dealer_list = br.find_element_by_xpath(".//div[@class='ex-listn ']") # And of course that means finding the parent element again to avoid the Stale Element error
    dealer_list_EX = dealer_list.find_elements_by_xpath(".//ul/li/a")    # Same as above
    company = dealer_list_EX[dealer].find_element_by_tag_name('div')     # Now we get the dealer's company name...
    print(company.text)
    sht.Range('C'+str(nrow)).Value = str(company.text)                   # ...and write it to disk...
    dealer_list_EX[dealer].click()                                       # ...before clicking on the dealer's link
    br.implicitly_wait(10)                                               # Let's wait 10 seconds (which is long, I know) to make sure he webpage is fully loaded
    Frame = br.find_element_by_xpath("//iframe[contains(@class, 'boxer-iframe')]") # Now there's an iframe here so beware. First let's identify it...
    br.switch_to_frame(Frame)                                                      # ...and then switch to it
    aa = br.find_element_by_xpath("//div[@class='intcontainer-exhibitor']")        # Now we can get the container element we want...
    divs = aa.find_elements_by_tag_name('div')                                     # ...and all it's 'div' children
    try:
        dealer_name = divs[0].text                                      # The first item in that list is the dealer's actual name
        dealer_full_name = dealer_name.split()                          # However we want their first and last names separately so le'ts split the string...
        dealer_first_name = dealer_full_name[1]                         # ...and take the second element which is the dealer's first name...
        dealer_last_name = dealer_full_name[-1]                         # ...and the last element which is their last name
        sht.Range('A'+str(nrow)).Value = str(dealer_first_name)         # And we just write to Excel
        sht.Range('B'+str(nrow)).Value = str(dealer_last_name)
        print(dealer_name)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('A'+str(nrow)).Value = 'N/A'
        sht.Range('B'+str(nrow)).Value = 'N/A'
        pass
    try:
        country = divs[2].text                                         # The third (yes third, two is three when reading Python lists) element tells us their country of domicile or whatever
        sht.Range('H'+str(nrow)).Value = str(country)                  # We write that to disk
        print(country)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('H'+str(nrow)).Value = 'N/A'
        pass
    try:
        phone = divs[3].text                                          # Let's get their phone number as well
        sht.Range('G'+str(nrow)).Value = str(phone)                   # Write that to disk
        print(phone)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('G'+str(nrow)).Value = 'N/A'
        pass
    try:
        email = divs[4].text                                         # And if they have an email address then we want it too
        email_strip = email.split()                                  # However it's got some other unwanted text...  
        EmaiL = email_strip[-1]                                      # ...which we strip out
        sht.Range('E'+str(nrow)).Value = str(EmaiL)                  # Finally we write that to Excel too
        print(email)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('E'+str(nrow)).Value = 'N/A'
        pass
    try:
        website = divs[5].text                                      # We want their website too which is the fourth div tag
        sht.Range('D'+str(nrow)).Value = str(website)               # We write that to Excel too
        print(website)
    except (NoSuchElementException, IndexError):
        print('N/A')
        sht.Range('D'+str(nrow)).Value = 'N/A'
        pass
    sht.Range('I'+str(nrow)).Value = 'Fine art, antiques and jewellery' # Now these guys don't individually say what they specialize in, however the website they're on says they do this so we're going with that.
    print(counter)
    nrow+=1
    counter +=1
    print('\nNext...')
    print('\n')

start_time = time.monotonic()	# Kicking off our timer

for dealer in all_dealers:      # Passing values from our all_dealers range object...                             
    GetDealer(dealer)           # ... to our function
print('\nDone')                 # Let the user (us) know when it's done scraping

# Closing stuff and exiting gracefully like a swan
br.close()
wb.Save()
xl.DisplayAlerts = True
wb.Close(True)
end_time = time.monotonic()    # Stopping the timer
print(timedelta(seconds=end_time - start_time)) # Let's know how long it took
