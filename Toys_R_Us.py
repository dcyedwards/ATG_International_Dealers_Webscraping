#! python 3                                                                                                                                                                             |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Script: Toys_R_Us.py                                                                                                                                                                  |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Author: David Edwards                                                                                                                                                                 |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Date: Written on the thirty-first day of the seventh month of the two-thousandth and seveenteeth year of our Lord and Saviour Jesus Christ (or just 31/07/2017)                       |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Purpose: To retrieve brand names and item types from the Toys R Us website for the ATG Categorization exercise.                                                                       |
#          This script retrieves brand names and item types from two different sections of the website:                                                                                 |
#          1. All toys                                                                                                                                                                  |
#          2. Bikes                                                                                                                                                                     |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Note:                                                                                                                                                                                 |
#       This script exists mainly because no accessible API for Toys R US exists which would huge downloads of data very easy. For this reason,                                         |
#       to modify this script, one must be familiar with some basic HTML (tags, href links, CSS Selectors, parent elements and their descendants) as                                    |
#       well as some basic Python programming (variable assignments, lists, functions, passing variables between functions and loops).                                                  |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
# Execution Plan:                                                                                                                                                                       |
#                                                                                                                                                                                       |
# 1. The Toys R Us website has a 3-level hierarchy of menus from which to filter down to the data required.                                                                             |
#    It's basic structure is represented somewhat below:                                                                                                                                |
#                                                                                                                                                                                       |
#    -----------------------------------------------                                                                                                                                    |
#    || Master Category --> Sub-Category --> Data ||                                                                                                                                    |
#    -----------------------------------------------                                                                                                                                    |
#                                                                                                                                                                                       |
# 2. I choose to write three functions (inspired by the structure of the website and my belief in the God of the Bible) to interact with the site.                                      |
#    I like functions because they keep things tidy, logical and allow for changes to be made easily at different points in the programming execution.                                  |
#    These functions are:                                                                                                                                                               |
#                                                                                                                                                                                       |
#    i.  GetCategory()    -   The master function responsible for looping through all categories of interest and passing arguments to the other functions.                              |
#                            (This function reminds me of God the Father who out of love for mankind, gives the command to the Son and Holy Spirit to save mankind)                     |
#                                                                                                                                                                                       |
#    ii. GetSubCategory() -  The function which receives arguments from GetCategory() but also passes its own arguments on its own authority (like Jesus Christ) to                     |
#                            the function GetBrands().                                                                                                                                  |
#                                                                                                                                                                                       |
#   iii. GetBrands()     -   The function responsible for actually getting the brand names at the desired webpage level. This function reads the elements at                            |
#                            the parent, child and descendant levels and may access their full or relative XPATHs where necessary (Works kind of like the Holy Spirt).                  |
# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|



# Necessary Modules (Some of which aren't used but are there in case the developer may want them later)
# ----------------------------------------------------------------------------------------------------
import bs4                       
from bs4 import BeautifulSoup
import win32com.client as win32
import glob, os, win32com.client, pythoncom, datetime, time, getpass
import comtypes, comtypes.client
import string
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

#chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\Downloads\\chromedriver_win32\\chromedriver"
chromeDriver = "C:\\Users\\"+getpass.getuser()+"\\chromedriver_win32\\chromedriver"

#cFolder = 'https://atgmedia-my.sharepoint.com/personal//'+getpass.getuser()+'_auctiontechnologygroup_com//Documents//Laptop//Desktop//'
cFolder = 'C:\\Users\\'+getpass.getuser()+'\\Desktop\\Desktop'


wkbk = os.path.expanduser(cFolder)                                 # 1. Initializing the workbook path in memory.
xl = win32.gencache.EnsureDispatch('Excel.Application')            # 2. Initializing the Excel workbook class in memory.
xl.DisplayAlerts = False                                           # 3. Putting annoying alerts off.
wb = xl.Workbooks.Add()                                            # 4. Creating an instance of the workbook
#wb.SaveAs(wkbk +" Toys_R_Us (All Items) - "+today+".xlsx")        # 5. Saving to disk
wb.SaveAs(wkbk +" Toys_R_Us (Bikes) - "+today+".xlsx")
xl.Visible = True                                                  # 6. Making it visible to the user

# URLs (Two URLs here - one for All items and the other for bikes)
# ----------------------------------------------------------------
url = 'http://www.toysrus.co.uk/toys/browse/toys/_/N-102329'              # 7. All toys
url2 = 'http://www.toysrus.co.uk/toys/browse/bikes-ride-ons/_/N-102228'   # 8. Bikes

# Google Chrome options
# ---------------------
options = webdriver.ChromeOptions()                         # 9. Accessing Chrome's options
options.add_argument("--start-maximized")                   # 10. Start Chrome maximized
br = webdriver.Chrome(chromeDriver, chrome_options=options) # 11. Yes, start it maximized because some elements do not appear unless the browser is maximized.

# Calling the webdriver and passing the URL to it.
# ------------------------------------------------
br.get(url)                   # 12. Comment out the URL you're not interested in at the moment. You'll have to run the script independently for each URL.
#br.get(url2)

Letters = string.ascii_letters # 13. Now this variable contains a string of upper and lowercase characters which GetBrands() will need to select the list of data we want.
                               # 14. So please keep this in mind. 

# Declaring some global variables to pass between our functions.
# -------------------------------------------------------------
global nrow   # 15. Our Excel row counter to specify which row data should be written to from our GetBrands() function.
global cNum   # 16. Our Category Number (cNum) variable to tell us which category we're currently scraping.
cNum = 0      # 17. And we'll start with the first category (0 means 1st in Python)
global sht    # 18. Our Excel worksheet variable which is crucial to actually getting the data written to disk.

# Okay let's write our functions:
# ------------------------------

# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# 1.

def GetCategory(): # 19. The master function which loops through each category and primarily governs the program flow.
    global nrow    # 20. Our Excel row counter variable
    global cNum    # 21. Our category counter
    print(cNum)
    br.execute_script("window.scrollTo(479,848)")                            # 22. Scrolling down the page just a little bit for the webdriver to 'see' some links.
    categories = br.find_elements_by_xpath(".//a[@class='f12 color-2']")     # 23. Accessing the list of categories on the main menu page. 
    while cNum < len(categories):                                            # 24. Now let's loop through each category until we get to the end...
        sht = wb.Worksheets.Add()                                            # 25. First of all, we'll create create a new Excel sheet.
        categories = br.find_elements_by_xpath(".//a[@class='f12 color-2']") # 26. Next we'll access the categories list again to preven the Stale Element Exception from occurring
        ct = categories[cNum].text                                           # 27. Let's store the category name in a variable which we'll need later and ultimately in our Excel book.
        shtname = ct.split(' ',1)                                            # 28. Let's split the category name by spaces and just used the first (because Excel doesn't like long sheet names you see).
        sht.Name = shtname[0]                                                # 29. So, using the first part of the split name as the name of the Excel sheet suffices.
        print('--------------------------------')                   
        print('Category: '+ categories[cNum].text)                           # 30. Let's tel the user what category they're currently scraping.
        categories[cNum].click()                                             # 31. Okay let's finally select the category...
        GetSubCategory(cNum, sht, ct)                                        # 32. ...and hand over execution of the program to GetSubCategory
        cNum +=1                                                             # 33. When that's done, let's increment the category number by 1 and select that next.


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# 2.

def GetSubCategory(cNum, sht, ct):                                                   # 34. This category accepts the category number, sheet name and category name from above. It also reads the HTML and determines...
    global nrow                                                                      # 35. ...if there's a sub-category or not.
    print('Category number :'+str(cNum))                                             # 36. Let's tell the user what category number we're on
    br.execute_script("window.scrollTo(479,848)")                                    # 37. Scrolling down the page just a little bit for the webdriver to 'see' stuff. Ridiculous I know but hey.
    sub_cats = br.find_elements_by_xpath(".//a[@class='f12 color-2']")               # 38. Let's access the list(s) of sub-categories which is the exact same command for getting the list of categories.
    if len(sub_cats) == 0:                                                           # 39. Now for some logic: if there are no elements in the list, then there's no sub-category - in which case...
        sc = 'N/A'                                                                   # 40. sc (sub-category) will be null or 'N/A' and...
        GetBrands(sht, ct, sc)                                                       # 41. ...we simply call our GetBrands() function.
    else:                                                                            # 42. Otherwise...
        try:                                                                         # 43. Let's loop through each sub-category with some error handling.
            for sub in range(0,len(sub_cats)):                                       # 44. Let's start our loop.
                sub_cats = br.find_elements_by_xpath(".//a[@class='f12 color-2']")   # 45. And yes, we'll always have to source the list of sub-cateogry elements on each iteration of our loop (Stale Element Exception you see)
                print('--------------------------------')
                print('Sub Category: '+ sub_cats[sub].text)                          # 46. Tell us what sub-category we're on
                print('--------------------------------')
                sc = sub_cats[sub].text                                              # 47. Store the sub-category name - we'll need it later in GetBrands()
                sub_cats[sub].click()                                                # 48. Let's click into that sub-category...
                GetBrands(sht, ct, sc)                                               # 49. ...and call our GetBrands() function passing the sheet, category and sub-category names to it.
            br.execute_script("window.history.go(-1)")                               # 50. Once GetBrands() finishes, it returns execution to this function which goes back a page and moves to the next available sub-category.
        except IndexError:                                                           # 51. Should there be no more sub-categories, execution is given back to GetCategory() which then moves on to the next Category.
            br.execute_script("window.history.go(-1)")                               # 52. In case of an error which is most likely to be an Index one...go back a page and...
            for sub in range(0,len(sub_cats)):
                sub_cats = br.find_elements_by_xpath(".//a[@class='f12 color-2']")
                print('--------------------------------')
                print('Sub Category: '+ sub_cats[sub].text)
                print('--------------------------------')
                sc = sub_cats[sub].text
                sub_cats[sub].click()            
                GetBrands(sht, ct, sc)                                              # 53. ...call our GetBrands() function again.
            br.execute_script("window.history.go(-1)")

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# 3.

def GetBrands(sht, ct, sc):                                                             # 54.GetBrands() which gets the brands from the site and writes them to Excel
    global nrow                                                                         # 55. Remember our nrow Excel row number counter variable?
    sht.Activate()                                                                      # 56. Let's give focus to the current sheet we want to write data to...
    sht.Cells(1,1).Value = 'Category: '+ str(ct)                                        # 57. Let's write the category we're looking at to the sheet...
    sht.Cells(1,2).Value = 'Sub Category'                                               # 58. ...and the Sub-category header field...
    sht.Cells(1,3).Value = 'Brand'                                                      # 59. ...and the brand header field.
    lastRow = sht.UsedRange.Rows.Count                                                  # 60. Now we need some marker to let our programme know the last row of used data and for that we use a lastRow variable...
    br.execute_script("window.scrollTo(479,848)")                                       # 61. ...which is exactly the same way you'd do it in VBA...for example: lastRow = ActiveSheet.UsedRange.Rows.Count
    view_more = br.find_elements_by_xpath(".//li[@class='filter-more']/a")              # 62. Now, the HTML requires some filters to be expanded but leaves no unique identifier for each filter because there can be up to five and...
    vrange = range(0,len(view_more))                                                    # 63. ...we are only interested in the Brands filter. So let's just get the complete list of all filters.
    for v in vrange:                                                                    # 64. Let's begin a loop through them to simply open them all.
        vw = view_more[v]
        loc = vw.location["y"]-150                                                      # 65. But because some filters may not be visible to Selenium, let's get their respective locations...
        time.sleep(4)                                                                   # 66. ...while we wait a few seconds for the webpage to catch up with our instructions...
        br.execute_script("window.scrollTo(0,%d);" %loc)                                # 67. ...before scrolling to each filter element one element at a time...
        vw.click()                                                                      # 68. ...and then opening them up
        time.sleep(4)                                                                   # 69. I pause the programme here again for 4 seconds to allow the Javascript on the page to finish executing after the above command.
    filter_section = br.find_elements_by_xpath(".//li[@class='filter-type closed']")    # 70. Right, now that all filters are open...let's get the lists of data on the webpage
    filter_range = range(0,len(filter_section))
    brand_list = []                                                                     # 71. Naturally, I like to have lists ready to hold just the element attributes I want from webpages so I present to you brand_list.
    for f in filter_range:                                                              # 72. Let's begin our loop through the data lists
        try:           
            lastRow = sht.UsedRange.Rows.Count                                          # 73. Our last row in Excel variable
            nrow = lastRow + 1                                                          # 74. And the row we'll write data to which is our lastRow plus one row
            if filter_section[f].find_elements_by_xpath(".//ul/li")[1].text[1] in Letters: # 75. Now since the lists have no unique identifier, we need to idenfity based on the actual data contained in them. We compare this with Letters...
                brands = filter_section[f].find_elements_by_xpath(".//ul/li")           # 76. ...where we identify our list as the one whose second character in its second row contains a letter. You'll need to look at the HTML for more detail.
                print('')
                for brand in brands:                                                    # 77. Now that we've successfully got our list, all we want are the brand names and not the quantities
                    brand_list.append(brand.text)                                       # 78. So let's put brand names and quantities in our brand_list because the names and quantities are concatenated together.
                brand_splits = [x.split(' (') for x in brand_list]                      # 78. But no worries, we'll create a new list and split them out into brand_splits (Lazy programming? Yeah maybe but hey)
                sht.Activate()                                                          # 79. Let's give focus but to the Excel sheet we're interested in
                for x in range(0,len(brand_splits)):                                    # 80. And from our list of brand names only (brand_splits), let's...
                    sht.Range('B'+str(nrow)).Value = str(sc)                            # 81. ...write the first sub-category name to the first available row in column B of our Excel sheet...
                    sht.Range('C'+str(nrow)).Value = str(brand_splits[x][0])            # 82. ...and take the first item (the stripped out name) of each item (which is a mini list) in our brand_splits list. We stick that into column C.      
                    print(brand_splits[x][0])                                           # 83. Let's show the user what that item is.
                    nrow +=1                                                            # 84. Let's increment the nrow counter by 1 for data entry into the next row.
                break                                                                   # 85. This break clause stops the loop in 71. from continuing when the condition in 75. is met.
        except IndexError:                                                              # 86. Should an index error occur for whatever reason...
            continue                                                                    # 87. The programme ignores it and continues with the loop.
    br.execute_script("window.history.go(-1)")                                          # 88a. When all data from brand_splits have been written to Excel, GetBrands() moves back to the sub-category (or category) page...
                                                                                        # 88b. ...and hands execution back to either GetSubCategory or GetCategory.
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
start_time = time.monotonic()

# Kicking everything off
# ----------------------

GetCategory()

# Wrapping up and closing
# -----------------------

br.close()                                             # 89. End our Google Chrome session
wb.Save()                                              # 90. Save our Workbook (It's recommended to use autosave as well in Excel's options)
xl.DisplayAlerts = True                                # 91. Display alerts and stuff.
wb.Close(True)                                         # 92. Close the workbook but save again (.Close(True)) before you do so.
end_time = time.monotonic()                            # 93. Stop our timer
print('\nDone')                                        # 94. Let the user know it's all done.
print(timedelta(seconds=end_time - start_time))        # 95. Show the user the total time taken to run everything.
