from __future__ import print_function
import requests
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
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
NO_SEARCH_RESULTS_FOUND = 1

User_Name = input('ComicsPriceGuide.com Username:  ')
User_Pass = input('ComicsPriceGuide.com Password:  ')
Google_Workbook = input('Google Workbook Name:    ')
Google_Sheet = input('Google Worksheet Name:    ')

# =============================================================================
#   Requests session — replaces Selenium/Chrome entirely
# =============================================================================
session = requests.Session()
session.headers.update({
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
})


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


# All valid CGC grades in ascending order
GRADE_SCALE = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5,
               5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
               9.2, 9.4, 9.6, 9.8, 10.0]


def _format_grade(g):
    """Format a float grade as eBay searchers would type it (e.g. 9.8, 1.0)."""
    return f"{g:.1f}"


def _get_page(url, retries=3, delay_range=(2, 6)):
    """GET a URL with retry logic. Returns BeautifulSoup and final URL."""
    for attempt in range(retries):
        try:
            time.sleep(random.uniform(*delay_range))
            resp = session.get(url, timeout=30, allow_redirects=True)
            resp.raise_for_status()
            soup = bs4.BeautifulSoup(resp.text, 'html.parser')
            return soup, resp.url
        except requests.RequestException as e:
            print(f"     Request error (attempt {attempt + 1}/{retries}): {e}")
            if attempt == retries - 1:
                raise
    return None, url


def _ebay_sold_prices(query):
    """
    Search eBay sold/completed listings for *query* and return a list of
    sale prices (floats) sorted newest-first.
    """
    search_url = (
        "https://www.ebay.com/sch/i.html?"
        f"_nkw={quote_plus(query)}"
        "&LH_Sold=1&LH_Complete=1&_sop=13"
    )
    try:
        soup, _ = _get_page(search_url, delay_range=(1, 3))
    except requests.RequestException:
        return []

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


def LoginComicsPriceGuide(User_Name, User_Pass):
    """
    Log into ComicsPriceGuide.com using a requests session.
    Discovers the login form dynamically to handle CSRF tokens or hidden fields.
    """
    print("Logging into ComicsPriceGuide.com...")
    resp = session.get("https://comicspriceguide.com/login", timeout=30)
    soup = bs4.BeautifulSoup(resp.text, 'html.parser')

    # Build payload from all hidden form fields (captures CSRF tokens etc.)
    form = soup.find('form')
    payload = {}
    action = 'https://comicspriceguide.com/login'
    if form:
        for hidden in form.find_all('input', {'type': 'hidden'}):
            if hidden.get('name'):
                payload[hidden['name']] = hidden.get('value', '')
        raw_action = form.get('action', '/login')
        action = raw_action if raw_action.startswith('http') else f"https://comicspriceguide.com{raw_action}"

    payload['user_username'] = User_Name
    payload['user_password'] = User_Pass

    time.sleep(random.uniform(1, 3))
    resp = session.post(action, data=payload, timeout=30, allow_redirects=True)

    # Verify login succeeded by checking for logout link or user-specific content
    if 'logout' in resp.text.lower() or 'sign out' in resp.text.lower():
        print("Login successful.")
    else:
        print(f"WARNING: Login may have failed (HTTP {resp.status_code}). Continuing anyway.")

    time.sleep(random.uniform(1, 3))


def SearchComic(Title, Issue, fullName, thisComic):
    """
    Search ComicsPriceGuide for a comic by title and issue.
    Returns [best_match_url, confidence_percentage].
    """
    print(fullName + " - Searching...")

    # GET the search page to discover the form structure
    resp = session.get("https://comicspriceguide.com/Search", timeout=30)
    soup = bs4.BeautifulSoup(resp.text, 'html.parser')

    form = soup.find('form')
    payload = {}
    action = 'https://comicspriceguide.com/Search'
    if form:
        for hidden in form.find_all('input', {'type': 'hidden'}):
            if hidden.get('name'):
                payload[hidden['name']] = hidden.get('value', '')
        raw_action = form.get('action', '/Search')
        action = raw_action if raw_action.startswith('http') else f"https://comicspriceguide.com{raw_action}"

    payload['search'] = Title
    payload['issueNu'] = str(Issue)

    time.sleep(random.uniform(2, 8))
    resp = session.post(action, data=payload, timeout=30, allow_redirects=True)
    soup = bs4.BeautifulSoup(resp.text, 'html.parser')

    similarity = 0
    comic_link = ''
    percentage = 0

    for candidate in soup.find_all('a', attrs={'class': 'grid_issue'}):
        a = str(candidate.text).replace("<sup>#</sup>", "#").upper()
        pct = similar(a, fullName)
        if pct > similarity:
            similarity = pct
            percentage = pct
            comic_link = 'https://comicspriceguide.com' + str(candidate["href"])

    if percentage > 0:
        print("     Found a match, confidence: " + str(int(percentage * 100)) + "% - " + comic_link)
    else:
        percentage = None
        print(str(thisComic['Title']) + " #" + str(thisComic['Issue']) + " - " + str(thisComic['Book Link']))
        comic_link = thisComic['Book Link']

    return [comic_link, percentage]


def generate_HTMLPage(sortedsheet):
    global htmlBody
    htmlBody = ''  # Reset so re-runs don't append to stale content

    for index, thisComic in sortedsheet.iterrows():
        title = str(thisComic['Title']).strip().upper()
        notes = str(thisComic['Notes']).strip()
        issue = int(str(thisComic['Issue']).strip())
        try:
            value = float(str(thisComic['Value']).strip().replace('$', '').replace(',', ''))
        except (ValueError, KeyError):
            value = 0.0
        image = str(thisComic['Cover Image']).strip()
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] is None else thisComic['CGC Graded']
        key = "No" if thisComic['KeyIssue'] is None else thisComic['KeyIssue']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()

        cgcdiv = '' if cgc.upper() == 'NO' else "<div class='cgc'>CGC</div>"
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
    """
    Read Google Sheet into a pandas DataFrame.
    Requires service account at ~/.config/gspread/service_account.json
    """
    gc = gspread.service_account()
    sh = gc.open(Google_Workbook)
    worksheet = sh.worksheet(Google_Sheet)
    Starting_DF = pd.DataFrame(worksheet.get_all_records())
    sortedsheet = Starting_DF.sort_values(by=['Title', 'Volume', 'Issue'])
    return Starting_DF, sortedsheet, worksheet, sh


def BackupGoogleSheet(sh, Starting_DF, sortedsheet):
    """Create a dated backup worksheet in the same Google Workbook."""
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
            price_paid = 0.01  # Avoid divide-by-zero in ROI calculations

        sortedsheet.at[index, 'Price Paid'] = price_paid
        fullName = title + " #" + str(issue) + variant
        confidence = ''
        print('Gathering : ' + fullName)

        if url == '':
            print('No direct URL - Calling search')
            search_results_Array = SearchComic(title, issue, fullName, thisComic)
            url = search_results_Array[0]
            confidence = search_results_Array[1]

        # =============================================================================
        #  A match has been determined - get the details
        # =============================================================================
        if url == '':
            raise ValueError(NO_SEARCH_RESULTS_FOUND,
                             "Looks like the search gave no result.",
                             thisComic['Title'], thisComic['Issue'])

        time.sleep(random.uniform(5, 15))
        soup, final_url = _get_page(url)

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

        # --- eBay pricing ---
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

        # =============================================================================
        #  Update the DataFrame
        # =============================================================================
        sortedsheet.at[index, 'Publisher'] = publisher
        sortedsheet.at[index, 'Volume'] = volume
        sortedsheet.at[index, 'Published'] = published
        sortedsheet.at[index, 'KeyIssue'] = keyIssue
        sortedsheet.at[index, 'Cover Price'] = cover_price
        sortedsheet.at[index, 'Comic Age'] = comic_age
        sortedsheet.at[index, 'Notes'] = notes
        sortedsheet.at[index, 'Confidence'] = confidence if confidence != '' else None
        sortedsheet.at[index, 'Book Link'] = final_url
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
#  Commit results back to Google Sheet (one API call at the end)
# =============================================================================
sortedsheet.fillna('', inplace=True)
worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())

generate_HTMLPage(sortedsheet)

print("Work is complete.")
