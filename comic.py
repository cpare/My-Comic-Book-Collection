from __future__ import print_function
from selenium import webdriver
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys
import random
import gspread
import locale

locale.setlocale( locale.LC_ALL, '' )

# =============================================================================
#   Variables
# =============================================================================
rundate = date.today().strftime("%Y-%m-%d")
htmlBody = ''
# The error codes
NO_SEARCH_RESULTS_FOUND = 1

User_Name = input('ComicsPriceGuide.com Username:  ')
User_Pass = input('ComicsPriceGuide.com Password:  ')
Google_Workbook = input('Google Workbook Name:    ')
Google_Sheet = input('Google Worksheet Name:    ')


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

driver = webdriver.Chrome()

#driver.maximize_window()


def LoginComicsPriceGuide(User_Name, User_Pass):
    driver.get("https://comicspriceguide.com/login")
    # Login Page Elements
    input_login_username = driver.find_element_by_xpath('//input[@id="user_username"]')
    input_login_password = driver.find_element_by_xpath('//input[@id="user_password"]')
    button_login_submit = driver.find_element_by_id("btnLogin")
    # Fill out login page
    input_login_username.send_keys(User_Name)
    input_login_password.send_keys(User_Pass)
    driver.execute_script("arguments[0].click();",button_login_submit)
    # Waiting between 5 and 20 seconds to look like a user
    time.sleep(random.uniform(5, 20))
    return()
    

def SearchComic(Title, Issue):
    print(fullName + " - Searching...")
    # Navigating to search page
    if(driver.current_url != "https://comicspriceguide.com/Search"):
        driver.get("https://comicspriceguide.com/Search")
    # Wait for the search page to load.
    driver.implicitly_wait(15)
    # Search Page HTML elements
    input_search_title = driver.find_element_by_id("search")
    input_search_issue = driver.find_element_by_id("issueNu")
    button_search_submit = driver.find_element_by_id("btnSearch")
    # Fill out the search fields
    input_search_title.send_keys(Title)
    input_search_issue.send_keys(Issue)
    #sleep to prevent overloading the site...
    time.sleep(random.uniform(2, 15))
    driver.execute_script("arguments[0].click();",button_search_submit)
    # Wait for results to show up
    time.sleep(random.uniform(5, 30))
    # Capture resulting page source
    source_code = driver.page_source
    # Instantiate BS4 using the source code.
    soup = bs4.BeautifulSoup(source_code,'html.parser')
    # Initial similarity. This similarity is between the given title and the hyperlink comic.
    similarity = 0
    # Link of the comic.
    comic_link = ''
    percentage = 0
    #Determine the best match for the comic that was just serached for
    for candidate in soup.find_all('a', attrs={'class':'grid_issue'}):
        # Replace the superscript "#" in the comic name
        a = str(candidate.text).replace("<sup>#</sup>","#").upper()
       # Check for similarity between the hyperlink comic and my comic title. If more,
        percentage = similar(a,fullName)
        if percentage > similarity:
            similarity = similar(a,fullName)
            final_link = 'https://comicspriceguide.com' + str(candidate["href"])
            comic_link = final_link
    if percentage > 0 :
        print("     Found a match, confidence: " + str(int(percentage*100)) + "% - " + comic_link) 
    else:
        percentage = None
        print(str(thisComic['Title']) + " #" + str(thisComic['Issue']) + " - " + str(thisComic['Book Link']))
        comic_link = thisComic['Book Link']
    
    array = [comic_link, percentage]
    return(array)


def generate_HTMLPage(sortedsheet):
    
    for index, thisComic in sortedsheet.iterrows():
        title = str(thisComic['Title']).strip().upper()
        notes = str(thisComic['Notes']).strip()
        issue = int(str(thisComic['Issue']).strip())
        value = int(str(thisComic['Value']).strip())
        image = str(thisComic['Cover Image']).strip().upper()
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        Key = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()
        if cgc.upper() =='NO':
            cgcdiv = ''
        else: 
            cgcdiv = "<div class='cgc'>CGC</div>"
        if key.upper() =='NO':
            keydiv = ''
        else: 
            keydiv = "<div class='key'>KEY</div>"
        htmlBody = htmlBody + "<div class='hvrbox'><img src='"  +  str(image) + "' alt='Cover' class='hvrbox-layer_bottom'><div class='hvrbox-layer_top'><div class='hvrbox-text'>" 
        htmlBody = htmlBody + "<a href='" + str(url) + "'>" + str(title) + " #" + str(issue) + str(variant) +"<br><br>Grade: " + str(grade) + "<br><br>Value: " + locale.currency(value, grouping=True) + "<br><br>" + str(notes) + "</a></div>" + str(cgcdiv) + str(keydiv) +"</div></div>"

    with open("comics.html",'w') as f:
        f.write("""<style type'"text/css">
        body {background-color: 282828;}
        a {color: whitesmoke;text-decoration: none;}
        .cgc {background-color: rgb(148, 7, 35);z-index: inherit 5;font-family: Arial, Helvetica, sans-serif;position: absolute;bottom: 0;}
        .hvrbox,
        .hvrbox * {box-sizing: border-box; padding: 5px;}
        .hvrbox {position: relative;display: inline-block;overflow: hidden;width: 250px;height: 400px;}
        .hvrbox img {width: 250px;height: 400px;}
        .hvrbox .hvrbox-layer_bottom {display: block;}
        .hvrbox .hvrbox-layer_top {
        	opacity: 0;
        	position: absolute;
        	top: 0;
        	left: 0;
        	right: 0;
        	bottom: 0;
        	width: 250px;
        	height: 400px;
        	background: rgba(0, 0, 0, 0.6);
        	color: #fff;
        	padding: 15px;
        	-moz-transition: all 0.4s ease-in-out 0s;
        	-webkit-transition: all 0.4s ease-in-out 0s;
        	-ms-transition: all 0.4s ease-in-out 0s;
        	transition: all 0.4s ease-in-out 0s;
        }
        .hvrbox:hover .hvrbox-layer_top,
        .hvrbox.active .hvrbox-layer_top {opacity: 1;}
        .hvrbox .hvrbox-text {
        	font-family: Arial, Helvetica, sans-serif;
            text-align: center;
        	font-size: 18px;
        	display: inline-block;
        	position: absolute;
        	top: 50%;
        	left: 50%;
        	-moz-transform: translate(-50%, -50%);
        	-webkit-transform: translate(-50%, -50%);
        	-ms-transform: translate(-50%, -50%);
        	transform: translate(-50%, -50%);
        }
        .hvrbox .hvrbox-text_mobile {
        	font-size: 15px;
        	border-top: 1px solid rgb(179, 179, 179); /* for old browsers */
        	border-top: 1px solid rgba(179, 179, 179, 0.7);
        	margin-top: 5px;
        	padding-top: 2px;
        	display: none;
        }
        .hvrbox.active .hvrbox-text_mobile {display: block;}
        </style>
        """)
        f.write(htmlBody)
    

def ReadGoogleSheet(Google_Workbook, Google_Sheet):
    # =============================================================================
    #   Read Google sheet into pandas Dataframe - Requires Service Account in Google API
    #   file stored in ~/.config/gspread/service_account.json 
    # =============================================================================
    gc = gspread.service_account()
    sh = gc.open(Google_Workbook)
    worksheet = sh.worksheet(Google_Sheet)
    Starting_DF = pd.DataFrame(worksheet.get_all_records())
    sortedsheet = Starting_DF.sort_values(by=['Title','Volume','Issue'])
    return Starting_DF, sortedsheet, worksheet
    
def BackupGoogleSheet(Sheetname):
    # =============================================================================
    #  Make a backup of the current sheet in the event it all goes to shit
    # =============================================================================
    starting_rows = Starting_DF.shape[0] 
    starting_cols = Starting_DF.shape[1] 
    backup = sh.add_worksheet(title="Backup " + rundate, rows=starting_rows, cols=starting_cols)
    backup.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())





SheetData = ReadGoogleSheet(Google_Workbook, Google_Sheet)
StartingDF = SheetData[0]
sortedsheet = SheetData[1]
worksheet = SheetData[2]
generate_HTMLPage(sortedsheet)

BackupGoogleSheet(StartingDF)       #  Backup current sheet


LoginComicsPriceGuide(User_Name, User_Pass)

for index, thisComic in sortedsheet.iterrows():
    try:
        # =============================================================================
        #  Fetch required data fields
        # =============================================================================
        title = str(thisComic['Title']).strip().upper()
        issue = int(str(thisComic['Issue']).strip())
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()
        
        price_paid = float(0.00)
        #  Prepare the Price Paid field as an float so we can do math stuff
        print(type(thisComic['Price Paid']))
        if type(thisComic['Price Paid']) == str:
            price_paid = float(str(thisComic['Price Paid']).strip().replace('$','') if float(str(thisComic['Price Paid']).strip().replace('$','')) != None else "0")
        elif type(thisComic['Price Paid']) ==  float:
            price_paid = thisComic['Price Paid']
        else:
            print('WARNING: Price Paid: ' + type(price_paid))
            
        if thisComic['Price Paid'] == 0:
            price_paid = float(0.01)
            
            
        sortedsheet.at[index,'Price Paid'] = price_paid
        fullName = title + " #" + str(issue) + variant
        confidence = ''
        print('Gathering : ' + fullName)
        
        if url == '' :
            print('No direct URL - Calling search')
            search_results_Array = SearchComic(title, issue)
            url = search_results_Array[0]
            confidence = search_results_Array[1]
            

# =============================================================================
#  A match has been determined - get the details
# =============================================================================
        if url != '':
            driver.get(url)
        else:
            raise ValueError(NO_SEARCH_RESULTS_FOUND,"Looks like the search gave no result. Try searching the title and issue manually to confirm the issue.",thisComic['Title'],thisComic['Issue'])
        
        # Wait 5 seconds for page to load and get its source code
        time.sleep(random.uniform(60, 240))
        source_code = driver.page_source
        
        # New BS4 Instance with the comic's page's source code
        soup = bs4.BeautifulSoup(source_code,'html.parser')
        
        # Finding out all details
        publisher = soup.find('a',attrs={'id':'hypPub'}).text
        volume = soup.find('span',attrs={'id':'lblVolume'}).text
        notes = soup.find('span',attrs = {'id':'spQComment'}).text
        keyIssue = "Yes" if "Key Issue" in soup.text else "No"
        image = soup.find('img',attrs={'id':'imgCoverMn'})['src']
        if image[0:4] != 'http':
            #  In some cases the URL is relative
            image = 'https://comicspriceguide.com/' + image
        basic_info = []
        for s in soup.find_all('div',attrs={"class":"m-0 f-12"}):
            basic_info.append(s.parent.find('span',attrs={"class":"f-11"}).text.replace("   ", " "))
        published = basic_info[0] if basic_info[0] != " ADD" else "Unknown"
        comic_age = basic_info[1] if basic_info[1] != " ADD" else "Unknown"
        cover_price = basic_info[2] if basic_info[2] != " ADD" else "Unknown"   

# =============================================================================
#  Get prices into a dataframe to find our grade
#  Known Defect: Comics graded as 10 fail due to the DF having '10.'not '10.0'
# =============================================================================
        if len(grade) < 3:
            grade = grade + ".0"       
        priceTable = soup.find(name='table',attrs={"id":"pricetable"})
        # Load the priceTable into a dataframe
        pricesdf = pd.read_html(priceTable.prettify())[0]
        # Truncate Condition column values to allow matching
        pricesdf['Condition'] = pricesdf['Condition'].str[:3]
        pricesdf = pricesdf.rename(columns={'Graded Value  *': 'Graded Value'})
        thisbooksgrade = pricesdf.loc[pricesdf['Condition'] == grade]
        RawValue = float(thisbooksgrade['Raw Value'].iloc[0].replace('$',''))
        GradedValue = float(thisbooksgrade['Graded Value'].iloc[0].replace('$',''))
        value = RawValue if cgc.upper() == 'NO' else GradedValue   
        characters_info = soup.find('div',attrs={'id':'dvCharacterList'}).text if soup.find('div',attrs={'id':'dvCharacterList'}) != None else "No Info Found"
        story = soup.find('div',attrs={'id':'dvStoryList'}).text.replace("Stories may contain spoilers","")
        url_link = driver.current_url
        
# =============================================================================
#  Determine Book Price change from last scan using the "Value" field
# =============================================================================
# =============================================================================
#         if 'Value' in thisComic:
#             print('     Value Field Exists')
#             LastScanValue = float(LastScanValue.replace('$',''))
#             CurrentScanValue = float(value.replace('$',''))
#             priceshift = round((CurrentScanValue - LastScanValue),2)
#             print('     PriceShift: ' + str(priceshift))
#             sortedsheet.at[index,'PriceShift'] = priceshift
#         else:
#             print('     Value Field does not exist, populating with current value')
#             sortedsheet.at[index,'Value'] = RawValue
# =============================================================================
            
# =============================================================================
#  update the DF
# =============================================================================
        sortedsheet.at[index,'Publisher'] = publisher
        sortedsheet.at[index,'Volume'] = volume
        sortedsheet.at[index,'Published'] = published
        sortedsheet.at[index,'KeyIssue'] = keyIssue
        sortedsheet.at[index,'Cover Price'] = cover_price
        sortedsheet.at[index,'Comic Age'] = comic_age
        sortedsheet.at[index,'Notes'] = notes
        if confidence !='':
            sortedsheet.at[index,'Confidence'] = confidence
        else:
            sortedsheet.at[index,'Confidence'] = None
        sortedsheet.at[index,'Book Link'] = url_link
        sortedsheet.at[index,'Graded'] = GradedValue
        sortedsheet.at[index,'Ungraded'] = RawValue
        sortedsheet.at[index,'Cover Image'] = image
        sortedsheet.at[index, rundate] = value


    except ValueError as ve:
        if(ve.args[0] == NO_SEARCH_RESULTS_FOUND):
            print("     Unable to find Match for " + str(ve.args[2]) + " #" + str(ve.args[3]))
        #driver.get("https://comicspriceguide.com/Search")
        
    except Exception as e:
        print("Error while working on " + title + ' ' + str(e))
        #driver.get("https://comicspriceguide.com/Search")
        continue

# =============================================================================
#  Commit results back to Google Sheet
# =============================================================================
# with pd.ExcelWriter(Excel_Workbook_Name, mode='w') as writer:  
#     sortedsheet.to_excel(writer, sheet_name=Excel_Sheet_Name)
    sortedsheet.fillna('', inplace=True)
    worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())
# =============================================================================
    
generate_HTMLPage(sortedsheet)

print("Work is complete.")
