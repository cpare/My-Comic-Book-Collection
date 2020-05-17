from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import  By
#import xlwings as xw
from pprint import pprint
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys
import xlrd
import numpy

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

driver = webdriver.Chrome()
driver.maximize_window()
#workbook = xw.Book("Comics.xlsx")
#sheet = workbook.sheets['My Collection']

ExcelWorkbookName = 'Comics.xlsx'

sheet = pd.read_excel(open(ExcelWorkbookName, 'rb'),
              sheet_name='My Collection')

# Create an empty Dataframe
dfResults = pd.DataFrame(columns = ['title','issue','grade','cgc','publisher',
                                    'volume','published','keyIssue','price_paid',
                                    'cover_price','price','comic_age','notes',
                                    'characters_info','story','url_link'])


driver.get("https://comicspriceguide.com/login")


# Login Elements
input_login_username = driver.find_element_by_xpath('//input[@id="user_username"]')
input_login_password = driver.find_element_by_xpath('//input[@id="user_password"]')
button_login_submit = driver.find_element_by_id("btnLogin")

# Logging in code
input_login_username.send_keys("botbotbot")
input_login_password.send_keys("iforgotit!")
driver.execute_script("arguments[0].click();",button_login_submit)

# Waiting time for page to load.
time.sleep(10)

# If the sheet doesnt exist, then create it. If it exists, then don't.
#date = date.today().strftime("%Y-%m-%d")
#sheets_list = [sh.name for sh in workbook.sheets]
#if date not in sheets_list :
 #   xw.sheets.add(date)

for comic_num in sheet.iterrows():
    try:
        # Navigating to search page
        if(driver.current_url != "https://comicspriceguide.com/Search"):
            driver.get("https://comicspriceguide.com/Search")

        # Wait for the search page to load.
        driver.implicitly_wait(30)

        # The Search elements.
        input_search_title = driver.find_element_by_id("search")
        input_search_issue = driver.find_element_by_id("issueNu")
        button_search_submit = driver.find_element_by_id("btnSearch")

        # Get the Comic
        comic = comic_num[1].to_list()
                
        # Appending the title to the end of list. Example, "The Amazing Spider-Man #101"
        comic.append(comic[1].upper() + " #" + str(comic[3]))

        # This is to convert "101.0" to just "101"
        if isinstance(comic[3], float):
            comic[3] = int(comic[3])

        # Input the search parameters.
        input_search_title.send_keys(str(comic[1]))
        time.sleep(1)

        # If the issue of the comic has any alphabet in it, then remove it and then search.
        if not str(comic[3]).isdigit(): 
            input_search_issue.send_keys(str(comic[3][:-1]))
        else: 
            input_search_issue.send_keys(str(comic[3]))
        time.sleep(1)
        driver.execute_script("arguments[0].click();",button_search_submit)

        # Wait for 20 seconds for results to show up
        time.sleep(12)

        # Source Code of the search result page.
        source_code = driver.page_source

        # Instantiate BS4 using the source code.
        soup = bs4.BeautifulSoup(source_code,'html.parser')

        # Initial similarity. This similarity is between the given title and the hyperlink comic.
        similarity = 0

        # Link of the comic.
        comic_link = ''

        # Check for all the hyperlinks on the results page which lead to the comic
        for link in soup.find_all('a', attrs={'class':'grid_issue'}):

            # Replace the superscript "#" in the comic name
            a = str(link.text).replace("<sup>#</sup>","#")

            # Check for similarity between the hyperlink comic and my comic title. If more,
            percentage = similar(a,comic[-1])
            if percentage > similarity:
                similarity = similar(a,comic[-1])
                final_link = 'https://comicspriceguide.com' + str(link["href"])
                comic_link = final_link

        # Goto the comic page if result is found.
        if comic_link != '':
            driver.get(comic_link)

        # Wait 5 seconds for page to load and get its source code
        time.sleep(5)
        source_code = driver.page_source

        # New BS4 Instance with the comic's page's source code
        soup = bs4.BeautifulSoup(source_code,'html.parser')

        # Finding out all details
        publisher = soup.find('a',attrs={'id':'hypPub'}).text
        title = comic[1]
        volume = soup.find('span',attrs={'id':'lblVolume'}).text
        issue = comic[3]
        grade = comic[4]
        cgc = "No" if comic[5] == None else comic[5]
        notes = soup.find('span',attrs = {'id':'spQComment'}).text
        price_paid = "$" + (str(comic[7]) if str(comic[7]) != None else "0")
        keyIssue = "Yes" if "Key Issue" in soup.text else "No"
        basic_info = []
        for s in soup.find_all('div',attrs={"class":"m-0 f-12"}):
            basic_info.append(s.parent.find('span',attrs={"class":"f-11"}).text.replace("   ", " "))

        published = basic_info[0] if basic_info[0] != " ADD" else "Unknown"
        comic_age = basic_info[1] if basic_info[1] != " ADD" else "Unknown"
        cover_price = basic_info[2] if basic_info[2] != " ADD" else "Unknown"   

        priceTable = soup.find(name='table',attrs={"id":"pricetable"})

        for td in priceTable.find_all('td'):
            if(cgc == "Yes"):
                id = 'lblGraded' + str(grade).replace('.','')
            else:
                id = 'lblValue' + str(grade).replace('.','')

            if(td.find('span',attrs={'id':id}) != None):
                price =  td.find('span',attrs={'id':id}).text

        characters_info = soup.find('div',attrs={'id':'dvCharacterList'}).text if soup.find('div',attrs={'id':'dvCharacterList'}) != None else "No Info Found"
        story = soup.find('div',attrs={'id':'dvStoryList'}).text.replace("Stories may contain spoilers","")
        url_link = driver.current_url

        # Uncomment below to debug
        #print(publisher,title,volume,issue,grade,cgc,notes,price_paid,published,comic_age,cover_price,price,characters_info,story)

        # Data to be put into excel file
        dfResults = dfResults.append({'title' : title,
                                      'issue' : issue,
                                      'grade':grade,
                                      'cgc':cgc,
                                      'publisher':publisher,
                                      'volume':volume,
                                      'published':published,
                                      'keyIssue':keyIssue,
                                     'price_paid':price_paid,
                                     'cover_price':cover_price,
                                     'price':price,
                                     'comic_age':comic_age,
                                     'notes':notes,
                                     'characters_info':characters_info,
                                     'story':story,
                                     'url_link':url_link}, ignore_index=True)

        
    except Exception as e:
        print("Oops there was an error." + str(e))
        dfResults = dfResults.append({'title' : comic_num[1][1],
                                      'issue' : comic_num[1][3],
                                      'grade': comic_num[1][4],
                                      'cgc': comic_num[1][5]}, ignore_index=True)
        driver.get("https://comicspriceguide.com/Search")
        continue

writer = ExcelWriter(ExcelWorkbookName)    
dfResults.to_excel(writer,date.today().strftime("%Y-%m-%d"))
print("Work is complete.")
