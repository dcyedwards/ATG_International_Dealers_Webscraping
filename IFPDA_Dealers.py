#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------
# Script: IFPDA_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the seventeenth day of the sixth month of the two-thousandth and seventeenth year of our Lord (or just 17/06/2017)
# ------------------------------------------------------------------------------------------------------------------------------------
# The fifth of eight scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves comany names, websites and specializations of 160 dealers in about 15 minutes.
# ------------------------------------------------------------------------------------------------------------------------------------

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

t = datetime.datetime.today()   # As usual, creating my timer variable
today = t.strftime('%Y-%m-%d')  # Formatting the time output

chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

br = webdriver.Chrome(chromeDriver)    # And starting an instance of Google Chrome
url = 'http://www.ifpda.org/dealers'   # Our URL of interest

cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
#cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                        # Initializing the workbook path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application')   # Initializing the Excel workbook class in memory.
xl.DisplayAlerts = False                                  # Putting annoying alerts off.
wb = xl.Workbooks.Add()                                   # Creating an instance of the workbook
wb.SaveAs(wkbk +" IFPDA_Dealers - "+today+".xlsx")        # Saving to disk
xl.Visible = True                                         # Making it visible to the user

# Adding worksheets and naming columns
sht = wb.Worksheets('Sheet1')
sht.Cells(1,1).Value = 'Dealer Name'
sht.Cells(1,2).Value = 'Company'
sht.Cells(1,3).Value = 'Website'
sht.Cells(1,4).Value = 'Email'
sht.Cells(1,5).Value = 'Mailing Address'
sht.Cells(1,6).Value = 'Phone no.'
sht.Cells(1,7).Value = 'Specialization'

lastRow = sht.UsedRange.Rows.Count   # My last row Excel variable
global nrow                          # My 'next row' Excel variable
nrow = lastRow + 1                   # Incrementing my nrow variable by 1 for data entry onto the next

br.get(url)                          # Navigating to our URL
dealer_results = br.find_element_by_xpath(".//div[@class='dealer-total-result no-padding col-xs-12']") # Accessing the parent element of the dealer links below
dealers = dealer_results.find_elements_by_xpath(".//h3/a")                                             # Acessing the list of all 'a' tags which are descendants of the parent element above

dealer_range = range(0,len(dealers)) # Creating a range object from the list of dealers above

# Our GetDealer function which will do most of the heavy lifting
# --------------------------------------------------------------
def GetDealer(dealer):
    global nrow                                              # Prepping our nrow variable
    company = dealers[dealer].get_attribute('title')         # Getting the dealer's company name from the title attribute of their link
    dealers[dealer].click()
    sht.Range('A'+str(nrow)).Value = 'N/A'
    sht.Range('B'+str(nrow)).Value = str(company)            # Writing to Excel
    print(company)
    website = br.find_element_by_xpath(".//div[@class='galleries-link no-padding col-sm-6 col-xs-6']")   # We need the dealer's website too
    sht.Range('C'+str(nrow)).Value = str(website.text)                                                 
    print(website.text)
    print('')
    toggle = br.find_element_by_xpath(".//img[@class='pull-left toggle-location']").click()              # Have to click on this element to get to the address details
    address_details = br.find_elements_by_xpath("//div[contains(@class, 'galleries-location-details')]") # Right, so now we've got to the address details but...
    address_detail_List = [str(i.text) for i in address_details]                                         # ...we have to put all the address elements into a list...
    address_detail_extract = "\n".join(address_detail_List)                                              # ...because there could be several so it's easier to concatenate them as a single text with spaces.
    print('')
    sht.Range('E'+str(nrow)).Value = str(address_detail_extract)                                         # And now we write the address details to Excel
    print(address_detail_extract)
    expertise = br.find_element_by_xpath(".//div[@class='dealer-specialties']")                          # Likewise, let's get the specialties container in the HTML
    expertise_list = expertise.find_elements_by_xpath("//div[contains(@class, 'specialties-top-details col-lg-2 col-md-2 col-sm-2 col-xs-6')]")  # And like above, we could have a list of specialties...
    expertise_LIST = [str(i.text) for i in expertise_list]                                               # Putting the text attributes of these elements into a list.                                        # So we put them into a list for concatenation...
    expertise_extract = "\n".join(expertise_LIST)                                                                                                # ...which works quite well, you see?
    sht.Range('G'+str(nrow)).Value = str(expertise_extract)                                                                                      # And finally we just write them to Excel. Easy, isn't it?
    print('')
    print(expertise_extract)
    del address_detail_List[:]                              # But hang on? What's going on, you say? Well I'll be re-using these lists over and over again and I delete their contents just to prevent any possible issues...
    del expertise_LIST[:]                                   # Of course, it's entirely possible that after each iteration, the contents of each list are cleared from memory but this is just to be very explicit because reasons.
    nrow+=1                                                 # And finally we can't forget to increment our Excel row counter, can we?
    print('next')

start_time = time.monotonic()

# Calling our function
for dealer in dealer_range:                                                                                 # Using our dealer_range object as 'value' container
    dealer_results = br.find_element_by_xpath(".//div[@class='dealer-total-result no-padding col-xs-12']")  # Getting the main container for the list of dealers in a loop to avoid the 'State Element' Selenium error
    dealers = dealer_results.find_elements_by_xpath(".//h3/a")                                              # Getting 'href' links for each dealer which are children of h3.
    GetDealer(dealer)                                                                                       # The actual function
    br.execute_script("window.history.go(-1)")                                                              # Once we've got details for a dealer, we go back a page and do the same process for the next dealer. Simple.
    print('ok')

# Wrapping up
br.close()
end_time = time.monotonic()
print(timedelta(seconds=end_time - start_time))

wb.Save()                    # Save stuff
xl.DisplayAlerts = True      # Enable all annoying alerts
wb.Close(True)               # Close the workbook
print("\nDone")              # Inform user
print("\nDavid...because you are a lazy fool.") # Insult myself
