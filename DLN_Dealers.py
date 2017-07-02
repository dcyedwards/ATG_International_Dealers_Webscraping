#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------
# Script: DLN_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the sixteenth day of the sixth month of the two-thousandth and seveenteeth year of our Lord (or just 16/06/2017)
# ------------------------------------------------------------------------------------------------------------------------------------
# The fourth amongst a series of scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves comany names, websites and specializations of 76 dealers in about a minute.
# ------------------------------------------------------------------------------------------------------------------------------------

# Some modules I need and some of which I don't but hey
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

t = datetime.datetime.today()    # As usual, creating my timer variable
today = t.strftime('%Y-%m-%d')   # Formatting the time output

#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"


br = webdriver.Chrome(chromeDriver)                                 # And starting an instance of Google Chrome
url = 'https://designlifenetwork.com/list-of-top-antique-dealers/'  # Our dealers URL on Designlifenetwork.com

cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
#cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                             # Initializing the workbook path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application')        # Initializing the Excel workbook class in memory.
xl.DisplayAlerts = False                                       # Putting annoying alerts off.
wb = xl.Workbooks.Add()                                        # Creating an instance of the workbook
wb.SaveAs(wkbk +" DLN_Dealers - "+today+".xlsx")               # Saving it to disk
xl.Visible = True                                              # Making it visible to the user

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

lastRow = sht.UsedRange.Rows.Count  # My last row Excel variable
global nrow                         # My 'next row' Excel variable
nrow = lastRow + 1                  # Incrementing my nrow variable by 1 for data entry onto the next


br.get(url)                         # Navigating to our URL
br.implicitly_wait(10)              # Slowing things down a little bit because the webdriver is too fast for the HTML to load

try:
    close_pop_up = br.find_element_by_xpath(".//a[@id='fancybox_ns-close']").click()   # Closing some pop-up window which appears when you browse to the site
except NoSuchElementException:
    pass

dealer_table = br.find_element_by_xpath(".//table[@class='easy-table easy-table-default ']") # Finding the HTML container which links to all dealers
tbody = dealer_table.find_elements_by_xpath(".//tbody/tr")                                   # Accessing all the dealer links in the table body of the container above
dealer_list = range(0,len(tbody))                                                            # Creating our dealer_list range object for the purposes of iteration.

# Our GetDealer function gets ready:
# ---------------------------------
def GetDealer(dealer):
    global nrow                                          # Not forgetting our Excel 'next row' variable for moving to the next row and data entry
    row = tbody[dealer].find_elements_by_tag_name('td')  # Clicking on the first row in the HTML which corresponds to the first dealer
    company = row[0].text                                # Extracting the company name from the link's text
    sht.Range('C'+str(nrow)).Value = str(company)        # Saving that in Excel
    print(company)
    try:
        website = row[0].find_element_by_css_selector('a').get_attribute('href') # Getting the dealer website from their href link attribute
        sht.Range('D'+str(nrow)).Value = str(website)                            # Sticking that in Excel
        print(website)
    except NoSuchElementException:                      # Error handling for when things go wrong
        sht.Range('D'+str(nrow)).Value = "N/A"
        print("N/A")
        pass
    specialty = row[1].text                             # Getting their specialty...
    sht.Range('I'+str(nrow)).Value = str(specialty)
    print(specialty)
    country = row[3].text                               # ...and their country
    sht.Range('H'+str(nrow)).Value = str(country)
    print(country)
    print('\nNext')
    nrow+=1

start_time = time.monotonic()                           # Kicking off the timer

# Actually calling the function
# -----------------------------
for dealer in dealer_list:                              # Remember our dealer_list range object thing?
    GetDealer(dealer)

# Wrapping up and closing
# -----------------------
br.close()
wb.Save()
xl.DisplayAlerts = True
wb.Close(True)
end_time = time.monotonic()
print('Done')
print(timedelta(seconds=end_time - start_time))
