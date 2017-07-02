#! python 3
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Script: AADLA_Dealers
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Author: David Edwards
# --------------------------------------------------------------------------------------------------------------------------------------------------
# Date: Written on the sixteenth day of the sixth month of the two-thousandth and seveenteeth year of our Lord (or just 16/06/2017)
# --------------------------------------------------------------------------------------------------------------------------------------------------
# The second in a couple of scripts meant to scrape data from dealer/auctioneer websites for an International Dealers Database that ATG is creating.
# Retrieves the geographic locations, phone numbers, email addresses and specializations of 85 dealers in about a minute.
# -------------------------------------------------------------------------------------------------------------------------------------------------

# Necessary modules (...well, some of them anyway)
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

t = datetime.datetime.today()
today = t.strftime('%Y-%m-%d')

#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver" # Specifying the location of our Chrome driver
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

br = webdriver.Chrome(chromeDriver)
url = 'http://aadla.com/alphabetical-list/'

#cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'

wkbk = os.path.expanduser(cFolder)                                # Creating and naming my Excel file and stuff
xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
wb.SaveAs(wkbk +" AADLA_Dealers - "+today+".xlsx")                # Using today's date 
xl.Visible = True                                                 # I like to see what's going on

sht = wb.Worksheets('Sheet1')
sht.Cells(1,1).Value = 'Dealer Name'
sht.Cells(1,2).Value = 'Company'
sht.Cells(1,3).Value = 'Email'
sht.Cells(1,4).Value = 'Mailing Address'
sht.Cells(1,5).Value = 'Phone no.'
sht.Cells(1,6).Value = 'Specialization'

lastRow = sht.UsedRange.Rows.Count                               # Last row of our Excel file
global nrow                                                      # A row counter which I use to increment the rows to be populated. Must be a global variable to be used in functions.
nrow = lastRow + 1                                               

Iterator = list(range(0,85))                                     # A simple list to iterate through each link 

br.get(url)                                                      # Going to the AADLA website

# Where things start to get real...
# ---------------------------------------------------------------------------------------------------------------------------
dealer_widget = br.find_element_by_xpath(".//div[@class='textwidget']") 
dealer_list = dealer_widget.find_elements_by_css_selector("li")  # A list of all dealers on the site

# My GetDealer function to run through each dealer link
# -----------------------------------------------------
def GetDealer(dealer):
    global nrow
    dealer_list[dealer].find_element_by_css_selector('a').click()                     # Right! So remember our dealer and iterator lists? Well they really play a role here.
    company_name = br.find_element_by_xpath(".//div[@class='inside']")                # Dealer's company name
    sht.Range('B'+str(nrow)).Value = str(company_name.text)
    print(company_name.text)
    location = br.find_element_by_xpath(".//span[@style='color: #6a6a6a;']")          # Dealer's location (country)
    sht.Range('D'+str(nrow)).Value = str(location.text)
    print(location.text)
    contact = br.find_elements_by_xpath(".//span[@style='color: #6a6a6a;']")
    phone = contact[1].text                                                           # Dealer's phone number (for harassing and stalking purposes)
    sht.Range('E'+str(nrow)).Value = str(contact[1].text)
    print(phone)
    email = contact[2].text                                                           # Dealer's email address
    sht.Range('C'+str(nrow)).Value = str(contact[2].text)
    print(email)
    exp = br.find_elements_by_xpath(".//div/ul/li")                                   # Dealer's specialization located under up to 6 'ul' tags
    expL = [str(i.text) for i in exp]                                                 # Storing all the specializations in a list in string (text) format.
    expertise = '-'.join(expL)                                                        # A concatenation of all the strings delimited by a '-'
    sht.Range('F'+str(nrow)).Value = str(expertise)                                   
    print(expertise)
    del expL[:]                                                                       # Explicitly emptying the specialization list for the next dealer (just in case the list's contents are stored in memory.
    nrow+=1                                                                           # Incrementing the row counter by 1 and we're done here.
    print("|--------------------------------|")

start_time = time.monotonic()

for dealer in Iterator:                                                               # Beginning a loop in which the function gets called.
    dealer_widget = br.find_element_by_xpath(".//div[@class='textwidget']")           
    dealer_list = dealer_widget.find_elements_by_css_selector("li")                   # Getting and refreshing the list of links to dealers to avoid the Stale Element Error.
    GetDealer(dealer)                                                                 # Calling the function and passing the 'dealer' argument from the 'Iterator' list.
    br.execute_script("window.history.go(-1)")                                        # Returning to the previous page after getting the dealer's details and continuing the loop.
    print("Next")

# Wrapping up
# -----------
br.close()                                                                            # Closing the browser 
end_time = time.monotonic()                                                           # Stopping the timer
print(timedelta(seconds=end_time - start_time))                                       # Printing out how long it all took
wb.Save()                                                                             # Saving the workbook
xl.DisplayAlerts = True                                                               
wb.Close(True)
print("\nDone")
print("\nO David...you lazy fool.")                                                
    
