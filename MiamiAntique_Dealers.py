#! python 3
# ------------------------------------------------------------------------------------------------------------------------------------
# Script: MiamiAntique_Dealers
# ------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# ------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the twenty-second day of the sixth month of the two-thousandth and seventeenth year of our Lord (or just 18/06/2017)
# ------------------------------------------------------------------------------------------------------------------------------------
# The last of eight scripts meant to scrape data from dealer/auctioneer websites for ATG's International Dealers Database.
# Retrieves dealer company names and specializations/expertise of 382 dealers in about one 80 minutes.
# This is the slowest running script because it has to wait for the javascript on the page to execute before it runs. As such lots of
# takes place. This script is also the only script to be partially automatic in that the user is required to scroll down the page when
# the script reaches the end of it. Alright, let's begin:
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


t = datetime.datetime.today()    # As usual, creating my timer variable
today = t.strftime('%Y-%m-%d')   # Formatting the time output

chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//' # Specifying location for saving Excel file
#cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                                         # Initializing the file path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application')                    # Initializing the Excel workbook class in memory.
xl.DisplayAlerts = False                                                   # Putting annoying alerts off.
wb = xl.Workbooks.Add()                                                    # Creating an instance of the workbook
wb.SaveAs(wkbk +" Original_Miami_Antique_Dealers - "+today+".xlsx")        # Saving to disk
xl.Visible = True                                                          # Making it visible to the user

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

lastRow = sht.UsedRange.Rows.Count                                       # My last row Excel variable
global nrow                                                              # My 'next row' Excel variable
nrow = lastRow + 1                                                       # Incrementing my nrow variable by 1 for data entry onto the next
global counter                                                           # So this is new, it's a variable to count the number of dealers I get. A bit superflous yes but hey
counter = 0                                                              # Setting the counter to 0


url = 'http://antique.a2zinc.net/OMAS2018/Public/eventmap.aspx'          # Our URL beause we can't do anything without it.
br = webdriver.Chrome(chromeDriver)                                      # Launching Google Chrome
br.get(url)                                                              # Navigating to our URL
time.sleep(20)                                                           # Pausing the programme because the website takes a while to load all elements and Selenium's too fast

Map = br.find_element_by_xpath(".//i[@class='fa fa-list fa-lg']").click()# Right, so here we click on the exhibitor seating map
time.sleep(30)                                                           # We pause a bit for the map to load.
list_body = br.find_element_by_xpath(".//div[@class='listTableBody']")   # We get the table body for all dealer links on the map
dealers = list_body.find_elements_by_xpath(".//a[@class='exhibitorName']") # Within that table body, we get all the dealer links

dealer_range = range(322,len(dealers))                                   # Now this used to start from 0 but I changed it to 322 because the internet connectivity broke at 322 and I couldn't be bothered starting all over again.
dealer_list = list(dealer_range)                                         # The power of lists

# Our function
def GetDealer(dealer):
    global nrow                             # Our global nrow variable
    global counter                          # And our counter      
    company = dealers[dealer].text          # Getting the dealer name from the link text
    print(counter)
    print(company)
    sht.Range('C'+str(nrow)).Value = str(company) # Writing to disk
    time.sleep(5)                                 # Pausing to allow the website to catch up
    dealers[dealer].click()                       # Now we click on the link
    time.sleep(5)                                 # And wait 5 seconds for the website to catch up
    Frame = br.find_element_by_xpath("//iframe[contains(@style, 'overflow: scroll;')]")  # The data's within an iframe which we must find...
    br.switch_to_frame(Frame)                                                            # ...and then switch to...
    try:
        time.sleep(2)                                                                    #...and then wait two seconds for the page to load
        specialty = br.find_element_by_xpath(".//div[@class='panel-body small']")        # We want the dealer specialty or specialties...
        specialties = specialty.find_elements_by_xpath(".//div[@class='ProductCategoryContainer']/li")  # ...but they can be quite a few...
        specialty_list = [str(i.text) for i in specialties]                                             # ...whose text attributes we put in a list
    except (NoSuchElementException, Exception):
        specialty_list = []
        pass
    try:
        Li = br.find_elements_by_xpath(".//div[@class='col-sm-8']/ul/li")               # Plus, they could have other specialties unbeknownst to us (in the third 'li' tag)...
        other_specialty = [str(i.text) for i in Li[3:]]                                 # Which we also put in a list named accordingly
    except (NoSuchElementException, Exception):
        other_specialty = []
        pass                                                                            # Some logic below to determine how to format the specialty(-ies?)
    if len(specialty_list) == 0 and len(other_specialty) > 0:                           # Ok, so if one of these lists has data...
        expertise_extract = "\n".join(other_specialty)                                  # Concatenate the specialties with a new row as delimiter
        print(expertise_extract)
        sht.Range('I'+str(nrow)).Value = str(expertise_extract)                         # Write to disk
    elif len(specialty_list) > 0 and len(other_specialty) == 0:                         # Same as above: if one list contains data - in this case, the specialty_list then...
        expertise_extract = "\n".join(specialty_list)                                   # ...concatenate its values with new lines and...
        print(expertise_extract)
        sht.Range('I'+str(nrow)).Value = str(expertise_extract)                         # ...write to disk
    elif len(specialty_list) > 0 and len(other_specialty) > 0:                          # If both lists contain data...
        expertise_list = specialty_list + other_specialty                               # Put them together,
        expertise_extract = "\n".join(expertise_list)                                   # concatenate with new lines...
        print(expertise_extract)                                
        sht.Range('I'+str(nrow)).Value = str(expertise_extract)                         # ...and write to disk (and by disk I mean Excel just to be clear)
    elif len(specialty_list) == 0 and len(other_specialty) == 0:                        # If no lists contain data then...
        print('N/A')                                                                    # ...inform the user
    br.switch_to_default_content()                                                      # When we're done with that, switch back from the iframe to the default window
    time.sleep(2)                                                                       
    close = WebDriverWait(br,15).until(
        EC.presence_of_element_located((By.XPATH,".//button[@class='close']"))).click()# Wait a few seconds or till when Selenium detects the close button and close
    nrow+=1         # Let's increment the Excel row counter...
    counter+=1      # ...and our dealer counter too 
        
                    
            
start_time = time.monotonic() # Kicking off the timer

for dealer in dealer_list:    # Using the items in dealer list to iterate through...
    GetDealer(dealer)         # ...our function
    print('\n')

# Wrapping up
print('Done')
br.close()
wb.Save()
xl.DisplayAlerts = True
wb.Close(True)
end_time = time.monotonic()
print('Done')                                  # Let the user know when we're done
print(timedelta(seconds=end_time - start_time))# Tell the user the time it took 
