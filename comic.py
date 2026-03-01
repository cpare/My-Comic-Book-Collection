from __future__ import print_function
from curl_cffi import requests
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import random
import gspread
import locale
import os
from urllib.parse import quote_plus
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '.env'))

locale.setlocale(locale.LC_ALL, '')

# =============================================================================
#   Variables
# =============================================================================
rundate = date.today().strftime("%Y-%m-%d")
htmlBody = ''
NO_SEARCH_RESULTS_FOUND = 1

User_Name = os.getenv('CPG_USERNAME') or input('ComicsPriceGuide.com Username:  ')
User_Pass = os.getenv('CPG_PASSWORD') or input('ComicsPriceGuide.com Password:  ')
Google_Workbook = os.getenv('GOOGLE_WORKBOOK') or input('Google Workbook Name:    ')
Google_Sheet = os.getenv('GOOGLE_SHEET') or input('Google Worksheet Name:    ')
CV_API_KEY = os.getenv('CV_API_KEY') or input('Comic Vine API Key:    ')

# =============================================================================
#   HTTP Sessions
#   curl_cffi impersonates Chrome TLS fingerprint — bypasses most bot detection
# =============================================================================
# Session for Comic Vine + eBay
session = requests.Session(impersonate="chrome120")

CV_BASE = "https://comicvine.gamespot.com/api"
CV_HEADERS = {"User-Agent": "MyComicCollection/1.0"}


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


# All valid CGC grades in ascending order
GRADE_SCALE = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5,
               5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
               9.2, 9.4, 9.6, 9.8, 10.0]


def _format_grade(g):
    """Format a float grade as eBay searchers would type it (e.g. 9.8, 1.0)."""
    return f"{g:.1f}"


# =============================================================================
#   Comic Vine API
# =============================================================================

def _cv_get(endpoint, params):
    """Make a Comic Vine API call. Returns the parsed JSON results or None."""
    params.update({
        'api_key': CV_API_KEY,
        'format': 'json',
    })
    url = f"{CV_BASE}/{endpoint}/"
    try:
        time.sleep(random.uniform(1, 2))  # CV rate limit: 200 req/hour
        resp = session.get(url, params=params, headers=CV_HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data.get('status_code') == 1:
            return data.get('results')
        else:
            print(f"     Comic Vine API error: {data.get('error')} (code {data.get('status_code')})")
            return None
    except Exception as e:
        print(f"     Comic Vine request failed: {e}")
        return None


def SearchComicVine(title, issue, variant=''):
    """
    Search Comic Vine for a specific issue.
    Returns a dict with metadata or None if not found.

    Fields returned:
      publisher, volume, published, cover_price, comic_age,
      notes, key_issue, cover_image, book_link, characters
    """
    # Search issues by volume name + issue number
    results = _cv_get('issues', {
        'filter': f'volume.name:{title},issue_number:{issue}',
        'field_list': (
            'id,name,issue_number,volume,description,deck,'
            'image,cover_date,store_date,cover_price,'
            'site_detail_url,character_credits,story_arc_credits,'
            'person_credits'
        ),
        'limit': 10,
    })

    if not results:
        # Fallback: broader name search
        print(f"     CV: No exact match — trying name search for '{title}'")
        results = _cv_get('issues', {
            'filter': f'name:{title},issue_number:{issue}',
            'field_list': (
                'id,name,issue_number,volume,description,deck,'
                'image,cover_date,store_date,cover_price,'
                'site_detail_url,character_credits,story_arc_credits'
            ),
            'limit': 10,
        })

    if not results:
        print(f"     CV: No results found for {title} #{issue}")
        return None

    # Pick best match by similarity on volume name
    full_name = f"{title} #{issue}{variant}".upper()
    best = None
    best_score = 0

    for r in results:
        vol_name = r.get('volume', {}).get('name', '') if r.get('volume') else ''
        candidate = f"{vol_name} #{r.get('issue_number', '')}".upper()
        score = similar(candidate, full_name)
        if score > best_score:
            best_score = score
            best = r

    if best is None or best_score < 0.3:
        print(f"     CV: Best match confidence too low ({int(best_score*100)}%) — skipping")
        return None

    print(f"     CV: Matched '{best.get('volume', {}).get('name', '')} #{best.get('issue_number')}' "
          f"(confidence {int(best_score*100)}%)")

    # Extract cover date / comic age
    cover_date = best.get('cover_date') or best.get('store_date') or 'Unknown'
    comic_age = _classify_age(cover_date)

    # Key issue: flag if description/deck mentions it, or has notable story arcs
    description = best.get('description') or ''
    deck = best.get('deck') or ''
    key_issue = 'Yes' if any(kw in (description + deck).lower() for kw in [
        'key issue', 'first appearance', '1st appearance', 'origin', 'death of',
        'first app', '1st app'
    ]) else 'No'

    # Cover image
    image = ''
    if best.get('image'):
        image = best['image'].get('original_url') or best['image'].get('medium_url') or ''

    # Characters
    chars = best.get('character_credits') or []
    characters = ', '.join([c.get('name', '') for c in chars[:10]])

    # Publisher
    publisher = ''
    if best.get('volume') and best['volume'].get('api_detail_url'):
        # Publisher is on the volume — fetch it if not already in result
        vol = best.get('volume', {})
        publisher = vol.get('publisher', {}).get('name', '') if vol.get('publisher') else ''
        if not publisher:
            # Volume detail has publisher — make a quick extra call
            vol_results = _cv_get(f"volume/{vol.get('id', '')}", {
                'field_list': 'publisher'
            }) if vol.get('id') else None
            if vol_results and isinstance(vol_results, dict):
                publisher = vol_results.get('publisher', {}).get('name', '') or ''

    return {
        'publisher': publisher,
        'volume': best.get('volume', {}).get('name', '') if best.get('volume') else '',
        'published': cover_date,
        'cover_price': best.get('cover_price') or 'Unknown',
        'comic_age': comic_age,
        'notes': deck or _strip_html(description[:500]) if description else '',
        'key_issue': key_issue,
        'cover_image': image,
        'book_link': best.get('site_detail_url') or '',
        'characters': characters,
        'confidence': best_score,
    }


def _strip_html(text):
    """Strip HTML tags from a string."""
    return bs4.BeautifulSoup(text, 'html.parser').get_text(separator=' ').strip()


def _classify_age(cover_date):
    """Classify a comic into Golden/Silver/Bronze/Copper/Modern Age by cover date."""
    if not cover_date or cover_date == 'Unknown':
        return 'Unknown'
    try:
        year = int(str(cover_date)[:4])
        if year < 1956:
            return 'Golden Age'
        elif year < 1970:
            return 'Silver Age'
        elif year < 1985:
            return 'Bronze Age'
        elif year < 1992:
            return 'Copper Age'
        else:
            return 'Modern Age'
    except (ValueError, TypeError):
        return 'Unknown'


# =============================================================================
#   eBay Pricing
# =============================================================================

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
        time.sleep(random.uniform(1, 3))
        resp = session.get(search_url, timeout=30)
        soup = bs4.BeautifulSoup(resp.text, 'html.parser')
    except Exception as e:
        print(f"     eBay request failed: {e}")
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


# =============================================================================
#   HTML Report
# =============================================================================

def generate_HTMLPage(sortedsheet):
    global htmlBody
    htmlBody = ''

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


# =============================================================================
#   Google Sheets
# =============================================================================

def ReadGoogleSheet(Google_Workbook, Google_Sheet):
    gc = gspread.service_account()
    sh = gc.open(Google_Workbook)
    worksheet = sh.worksheet(Google_Sheet)
    Starting_DF = pd.DataFrame(worksheet.get_all_records())
    sortedsheet = Starting_DF.sort_values(by=['Title', 'Volume', 'Issue'])
    return Starting_DF, sortedsheet, worksheet, sh


def BackupGoogleSheet(sh, Starting_DF, sortedsheet):
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

for index, thisComic in sortedsheet.iterrows():
    try:
        # =====================================================================
        #  Fetch required data fields
        # =====================================================================
        title = str(thisComic['Title']).strip().upper()
        issue = int(str(thisComic['Issue']).strip())
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] is None else thisComic['CGC Graded']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        fullName = title + " #" + str(issue) + variant

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
            price_paid = 0.01

        sortedsheet.at[index, 'Price Paid'] = price_paid
        print('Gathering : ' + fullName)

        # =====================================================================
        #  Comic Vine — metadata lookup
        # =====================================================================
        cv_data = SearchComicVine(title, issue, variant)

        if cv_data:
            publisher    = cv_data['publisher']
            volume       = cv_data['volume']
            published    = cv_data['published']
            cover_price  = cv_data['cover_price']
            comic_age    = cv_data['comic_age']
            notes        = cv_data['notes']
            keyIssue     = cv_data['key_issue']
            image        = cv_data['cover_image']
            url_link     = cv_data['book_link']
            confidence   = cv_data['confidence']
        else:
            # Preserve existing sheet values if CV lookup fails
            publisher   = str(thisComic.get('Publisher', ''))
            volume      = str(thisComic.get('Volume', ''))
            published   = str(thisComic.get('Published', 'Unknown'))
            cover_price = str(thisComic.get('Cover Price', 'Unknown'))
            comic_age   = str(thisComic.get('Comic Age', 'Unknown'))
            notes       = str(thisComic.get('Notes', ''))
            keyIssue    = str(thisComic.get('KeyIssue', 'No'))
            image       = str(thisComic.get('Cover Image', ''))
            url_link    = str(thisComic.get('Book Link', ''))
            confidence  = None
            print(f"     CV lookup failed — preserving existing sheet values")

        # =====================================================================
        #  eBay — market pricing
        # =====================================================================
        if len(grade) < 3:
            grade = grade + ".0"

        value = GetEbayPrice(title, issue, grade, cgc, variant)
        if value is None:
            # Fall back to existing sheet value if eBay has nothing
            try:
                value = float(str(thisComic.get('Value', '0')).replace('$', '').replace(',', '')) or 0.0
            except (ValueError, TypeError):
                value = 0.0
            print(f"     eBay unavailable — keeping existing value: ${value:.2f}")

        # =====================================================================
        #  Update the DataFrame
        # =====================================================================
        sortedsheet.at[index, 'Publisher']   = publisher
        sortedsheet.at[index, 'Volume']      = volume
        sortedsheet.at[index, 'Published']   = published
        sortedsheet.at[index, 'KeyIssue']    = keyIssue
        sortedsheet.at[index, 'Cover Price'] = cover_price
        sortedsheet.at[index, 'Comic Age']   = comic_age
        sortedsheet.at[index, 'Notes']       = notes
        sortedsheet.at[index, 'Confidence']  = confidence
        sortedsheet.at[index, 'Book Link']   = url_link
        sortedsheet.at[index, 'Cover Image'] = image
        sortedsheet.at[index, rundate]       = value

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
