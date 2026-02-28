from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By   # Fix: deprecated find_element_by_* removed in Selenium 4
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys
import random
import gspread
import locale
from urllib.parse import quote_plus

locale.setlocale(locale.LC_ALL, '')

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


# All valid CGC grades in ascending order
GRADE_SCALE = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5,
               5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
               9.2, 9.4, 9.6, 9.8, 10.0]


def _format_grade(g):
    """Format a float grade as eBay searchers would type it (e.g. 9.8, 1.0)."""
    return f"{g:.1f}"


def _ebay_sold_prices(query):
    """
    Search eBay sold/completed listings for *query* and return a list of
    sale prices (floats) sorted newest-first.
    """
    search_url = (
        "https://www.ebay.com/sch/i.html?"
        f"_nkw={quote_plus(query)}"   # Fix: use proper URL encoding for special chars
        "&LH_Sold=1&LH_Complete=1&_sop=13"
    )
    driver.get(search_url)
    time.sleep(random.uniform(3, 7))
    soup = bs4.BeautifulSoup(driver.page_source, 'html.parser')

    prices = []
    for span in soup.find_all('span', attrs={'class': 's-item__price'}):
        raw = span.text.strip().replace('$', '').replace(',', '')
        try:
            if ' to ' in raw:
                parts = raw.split(' to ')
                prices.append((float(parts[0]) + float(parts[1])) / 2)
            else:
                prices.append(float(raw))
        except ValueError:
            pass
    return prices


def _ebay_price_for_grade(title, issue, grade_float, cgc, variant=''):
    """Return the most-recent eBay sold price for a specific grade, or None."""
    cgc_str = "CGC" if cgc.upper() != 'NO' else ""
    grade_str = _format_grade(grade_float)
    query = f"{title} #{issue} {cgc_str} {grade_str} {variant}".strip()
    prices = _ebay_sold_prices(query)
    if prices:
        print(f"     eBay [{grade_str}]: ${prices[0]:.2f}  ({len(prices)} sales found)")
        return prices[0]
    return None


def GetEbayPrice(title, issue, grade, cgc, variant=''):
    """
    Look up the most recent eBay sold price for the given comic + grade.

    Strategy:
      1. Try exact grade match.
      2. If no sales, walk outward on GRADE_SCALE to find the nearest lower
         and higher grades that have sales, then linearly interpolate.
      3. Return None if no eBay data at all.
    """
    try:
        target = float(grade)
    except ValueError:
        print(f"     eBay: Cannot parse grade '{grade}'")
        return None

    # 1. Exact match
    price = _ebay_price_for_grade(title, issue, target, cgc, variant)
    if price is not None:
        return price

    # 2. Interpolate
    print(f"     eBay: No sales for grade {grade} – searching nearby grades…")

    if target not in GRADE_SCALE:
        print(f"     eBay: Grade {target} not in standard scale – cannot interpolate.")
        return None

    idx = GRADE_SCALE.index(target)

    lower_grade, lower_price = None, None
    for g in reversed(GRADE_SCALE[:idx]):
        p = _ebay_price_for_grade(title, issue, g, cgc, variant)
        if p is not None:
            lower_grade, lower_price = g, p
            break

    upper_grade, upper_price = None, None
    for g in GRADE_SCALE[idx + 1:]:
        p = _ebay_price_for_grade(title, issue, g, cgc, variant)
        if p is not None:
            upper_grade, upper_price = g, p
            break

    if lower_price is not None and upper_price is not None:
        ratio = (target - lower_grade) / (upper_grade - lower_grade)
        interpolated = round(lower_price + ratio * (upper_price - lower_price), 2)
        print(
            f"     eBay: Interpolated {grade} from "
            f"{lower_grade}=${lower_price:.2f} ↔ {upper_grade}=${upper_price:.2f} "
            f"→ ${interpolated:.2f}"
        )
        return interpolated
    elif lower_price is not None:
        print(f"     eBay: Using nearest lower grade {lower_grade} → ${lower_price:.2f}")
        return lower_price
    elif upper_price is not None:
        print(f"     eBay: Using nearest upper grade {upper_grade} → ${upper_price:.2f}")
        return upper_price

    print(f"     eBay: No usable sales data found for {title} #{issue}.")
    return None


driver = webdriver.Chrome()


def LoginComicsPriceGuide(User_Name, User_Pass):
    driver.get("https://comicspriceguide.com/login")
    # Fix: find_element_by_xpath / find_element_by_id removed in Selenium 4
    input_login_username = driver.find_element(By.ID, "user_username")
    input_login_password = driver.find_element(By.ID, "user_password")
    button_login_submit = driver.find_element(By.ID, "btnLogin")
    input_login_username.send_keys(User_Name)
    input_login_password.send_keys(User_Pass)
    driver.execute_script("arguments[0].click();", button_login_submit)
    time.sleep(random.uniform(5, 20))


def SearchComic(Title, Issue, fullName, thisComic):
    # Fix: fullName and thisComic are now explicit parameters instead of globals
    print(fullName + " - Searching...")
    if driver.current_url != "https://comicspriceguide.com/Search":
        driver.get("https://comicspriceguide.com/Search")
    driver.implicitly_wait(15)
    # Fix: use By.ID instead of deprecated find_element_by_id
    input_search_title = driver.find_element(By.ID, "search")
    input_search_issue = driver.find_element(By.ID, "issueNu")
    button_search_submit = driver.find_element(By.ID, "btnSearch")
    input_search_title.send_keys(Title)
    input_search_issue.send_keys(Issue)
    time.sleep(random.uniform(2, 15))
    driver.execute_script("arguments[0].click();", button_search_submit)
    time.sleep(random.uniform(5, 30))
    source_code = driver.page_source
    soup = bs4.BeautifulSoup(source_code, 'html.parser')
    similarity = 0
    comic_link = ''
    percentage = 0
    for candidate in soup.find_all('a', attrs={'class': 'grid_issue'}):
        a = str(candidate.text).replace("<sup>#</sup>", "#").upper()
        percentage = similar(a, fullName)
        if percentage > similarity:
            similarity = similar(a, fullName)
            final_link = 'https://comicspriceguide.com' + str(candidate["href"])
            comic_link = final_link
    if percentage > 0:
        print("     Found a match, confidence: " + str(int(percentage * 100)) + "% - " + comic_link)
    else:
        percentage = None
        print(str(thisComic['Title']) + " #" + str(thisComic['Issue']) + " - " + str(thisComic['Book Link']))
        comic_link = thisComic['Book Link']

    return [comic_link, percentage]


def generate_HTMLPage(sortedsheet):
    global htmlBody   # Fix: UnboundLocalError — must declare global before assignment
    htmlBody = ''     # Reset so re-runs don't append to stale content

    for index, thisComic in sortedsheet.iterrows():
        title = str(thisComic['Title']).strip().upper()
        notes = str(thisComic['Notes']).strip()
        issue = int(str(thisComic['Issue']).strip())
        # Fix: value is float, not int — int() truncates cents (e.g. $19.99 → $19)
        try:
            value = float(str(thisComic['Value']).strip().replace('$', '').replace(',', ''))
        except (ValueError, KeyError):
            value = 0.0
        image = str(thisComic['Cover Image']).strip()  # Fix: removed .upper() — uppercasing breaks URLs
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] is None else thisComic['CGC Graded']
        # Fix: Key was incorrectly reading CGC Graded field instead of KeyIssue
        key = "No" if thisComic['KeyIssue'] is None else thisComic['KeyIssue']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()

        cgcdiv = '' if cgc.upper() == 'NO' else "<div class='cgc'>CGC</div>"
        # Fix: was checking 'key' (lowercase) but variable was 'Key' (uppercase) — NameError
        keydiv = '' if key.upper() == 'NO' else "<div class='key'>KEY</div>"

        htmlBody += (
            f"<div class='hvrbox'>"
            f"<img src='{image}' alt='Cover' class='hvrbox-layer_bottom'>"
            f"<div class='hvrbox-layer_top'><div class='hvrbox-text'>"
            f"<a href='{url}'>{title} #{issue}{variant}<br><br>"
            f"Grade: {grade}<br><br>"
            f"Value: {locale.currency(value, grouping=True)}<br><br>"
            f"{notes}</a></div>{cgcdiv}{keydiv}</div></div>"
        )

    with open("comics.html", 'w') as f:
        f.write("""<style type="text/css">
        body {background-color: #282828;}
        a {color: whitesmoke; text-decoration: none;}
        .cgc {background-color: rgb(148, 7, 35); z-index: 5; font-family: Arial, Helvetica, sans-serif; position: absolute; bottom: 0;}
        .key {background-color: rgb(7, 100, 35); z-index: 5; font-family: Arial, Helvetica, sans-serif; position: absolute; bottom: 0; right: 0;}
        .hvrbox, .hvrbox * {box-sizing: border-box; padding: 5px;}
        .hvrbox {position: relative; display: inline-block; overflow: hidden; width: 250px; height: 400px;}
        .hvrbox img {width: 250px; height: 400px;}
        .hvrbox .hvrbox-layer_bottom {display: block;}
        .hvrbox .hvrbox-layer_top {
            opacity: 0; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            width: 250px; height: 400px; background: rgba(0,0,0,0.6); color: #fff; padding: 15px;
            transition: all 0.4s ease-in-out 0s;
        }
        .hvrbox:hover .hvrbox-layer_top, .hvrbox.active .hvrbox-layer_top {opacity: 1;}
        .hvrbox .hvrbox-text {
            font-family: Arial, Helvetica, sans-serif; text-align: center; font-size: 18px;
            display: inline-block; position: absolute; top: 50%; left: 50%;
            transform: translate(-50%, -50%);
        }
        .hvrbox .hvrbox-text_mobile {
            font-size: 15px; border-top: 1px solid rgba(179,179,179,0.7);
            margin-top: 5px; padding-top: 2px; display: none;
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
    sortedsheet = Starting_DF.sort_values(by=['Title', 'Volume', 'Issue'])
    return Starting_DF, sortedsheet, worksheet, sh


def BackupGoogleSheet(sh, Starting_DF, sortedsheet):
    # Fix: was using globals sh/Starting_DF/sortedsheet; now passed as explicit parameters
    # Fix: Sheetname parameter was unused; removed it
    starting_rows = Starting_DF.shape[0]
    starting_cols = Starting_DF.shape[1]
    backup = sh.add_worksheet(title="Backup " + rundate, rows=starting_rows, cols=starting_cols)
    backup.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())


# =============================================================================
#  Main
# =============================================================================
SheetData = ReadGoogleSheet(Google_Workbook, Google_Sheet)
StartingDF = SheetData[0]
sortedsheet = SheetData[1]
worksheet = SheetData[2]
sh = SheetData[3]

generate_HTMLPage(sortedsheet)
BackupGoogleSheet(sh, StartingDF, sortedsheet)

LoginComicsPriceGuide(User_Name, User_Pass)

for index, thisComic in sortedsheet.iterrows():
    try:
        # =============================================================================
        #  Fetch required data fields
        # =============================================================================
        title = str(thisComic['Title']).strip().upper()
        issue = int(str(thisComic['Issue']).strip())
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] is None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()

        # Fix: simplified and corrected price_paid conversion logic
        # Previous code had a dead-code else branch (float() is never None) and
        # could double-parse the string causing an error on malformed input.
        price_paid = 0.0
        raw_paid = thisComic['Price Paid']
        try:
            if isinstance(raw_paid, str):
                price_paid = float(raw_paid.strip().replace('$', '').replace(',', '')) or 0.0
            elif isinstance(raw_paid, (int, float)):
                price_paid = float(raw_paid)
        except (ValueError, TypeError):
            print(f'WARNING: Could not parse Price Paid value: {raw_paid!r}')

        if price_paid == 0:
            price_paid = 0.01   # Avoid divide-by-zero in any ROI calculations

        sortedsheet.at[index, 'Price Paid'] = price_paid
        fullName = title + " #" + str(issue) + variant
        confidence = ''
        print('Gathering : ' + fullName)

        if url == '':
            print('No direct URL - Calling search')
            # Fix: pass fullName and thisComic explicitly (no longer implicit globals)
            search_results_Array = SearchComic(title, issue, fullName, thisComic)
            url = search_results_Array[0]
            confidence = search_results_Array[1]

        # =============================================================================
        #  A match has been determined - get the details
        # =============================================================================
        if url != '':
            driver.get(url)
        else:
            raise ValueError(NO_SEARCH_RESULTS_FOUND,
                             "Looks like the search gave no result.",
                             thisComic['Title'], thisComic['Issue'])

        time.sleep(random.uniform(60, 240))
        source_code = driver.page_source
        soup = bs4.BeautifulSoup(source_code, 'html.parser')

        publisher = soup.find('a', attrs={'id': 'hypPub'}).text
        volume = soup.find('span', attrs={'id': 'lblVolume'}).text
        notes = soup.find('span', attrs={'id': 'spQComment'}).text
        keyIssue = "Yes" if "Key Issue" in soup.text else "No"
        image = soup.find('img', attrs={'id': 'imgCoverMn'})['src']
        if image[0:4] != 'http':
            image = 'https://comicspriceguide.com/' + image

        basic_info = []
        for s in soup.find_all('div', attrs={"class": "m-0 f-12"}):
            basic_info.append(s.parent.find('span', attrs={"class": "f-11"}).text.replace("   ", " "))
        published = basic_info[0] if basic_info[0] != " ADD" else "Unknown"
        comic_age = basic_info[1] if basic_info[1] != " ADD" else "Unknown"
        cover_price = basic_info[2] if basic_info[2] != " ADD" else "Unknown"

        # =============================================================================
        #  Get prices from CPG price table
        #  Fix: pad grade to 3 chars BEFORE eBay lookup so both sources use same format
        #  Known Defect: grade 10.0 stored as '10.' on CPG — truncation to [:3] gives '10.'
        # =============================================================================
        if len(grade) < 3:
            grade = grade + ".0"

        priceTable = soup.find(name='table', attrs={"id": "pricetable"})
        pricesdf = pd.read_html(priceTable.prettify())[0]
        pricesdf['Condition'] = pricesdf['Condition'].str[:3]
        pricesdf = pricesdf.rename(columns={'Graded Value  *': 'Graded Value'})
        thisbooksgrade = pricesdf.loc[pricesdf['Condition'] == grade]
        RawValue = float(thisbooksgrade['Raw Value'].iloc[0].replace('$', ''))
        GradedValue = float(thisbooksgrade['Graded Value'].iloc[0].replace('$', ''))
        cpg_value = RawValue if cgc.upper() == 'NO' else GradedValue

        # --- eBay pricing (issue #5) ---
        # Prefer real market data (last eBay sale) over guide prices.
        ebay_value = GetEbayPrice(title, issue, grade, cgc, variant)
        if ebay_value is not None:
            value = ebay_value
            print(f"     Using eBay price: ${value:.2f} (CPG guide: ${cpg_value:.2f})")
        else:
            value = cpg_value
            print(f"     eBay price unavailable – falling back to CPG guide: ${value:.2f}")

        characters_info = (
            soup.find('div', attrs={'id': 'dvCharacterList'}).text
            if soup.find('div', attrs={'id': 'dvCharacterList'}) is not None
            else "No Info Found"
        )
        story = soup.find('div', attrs={'id': 'dvStoryList'}).text.replace("Stories may contain spoilers", "")
        url_link = driver.current_url

        # =============================================================================
        #  update the DataFrame
        # =============================================================================
        sortedsheet.at[index, 'Publisher'] = publisher
        sortedsheet.at[index, 'Volume'] = volume
        sortedsheet.at[index, 'Published'] = published
        sortedsheet.at[index, 'KeyIssue'] = keyIssue
        sortedsheet.at[index, 'Cover Price'] = cover_price
        sortedsheet.at[index, 'Comic Age'] = comic_age
        sortedsheet.at[index, 'Notes'] = notes
        sortedsheet.at[index, 'Confidence'] = confidence if confidence != '' else None
        sortedsheet.at[index, 'Book Link'] = url_link
        sortedsheet.at[index, 'Graded'] = GradedValue
        sortedsheet.at[index, 'Ungraded'] = RawValue
        sortedsheet.at[index, 'Cover Image'] = image
        sortedsheet.at[index, rundate] = value

    except ValueError as ve:
        if ve.args[0] == NO_SEARCH_RESULTS_FOUND:
            print("     Unable to find Match for " + str(ve.args[2]) + " #" + str(ve.args[3]))

    except Exception as e:
        print("Error while working on " + title + ' ' + str(e))
        continue

# =============================================================================
#  Commit results back to Google Sheet
#  Fix: worksheet.update() was inside the for loop — moved outside so we make
#       one API call at the end instead of one per comic (major performance fix)
# =============================================================================
sortedsheet.fillna('', inplace=True)
worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())

generate_HTMLPage(sortedsheet)

print("Work is complete.")
