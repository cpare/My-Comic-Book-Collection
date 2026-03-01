"""
test_run.py — Runs comic.py logic against the first 10 records only.
Reads credentials from .env, no interactive prompts needed.
"""
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

rundate = date.today().strftime("%Y-%m-%d")
NO_SEARCH_RESULTS_FOUND = 1
TEST_LIMIT = 10

Google_Workbook = os.getenv('GOOGLE_WORKBOOK') or input('Google Workbook Name: ')
Google_Sheet    = os.getenv('GOOGLE_SHEET')    or input('Google Sheet Name: ')
CV_API_KEY      = os.getenv('CV_API_KEY')      or input('Comic Vine API Key: ')

session = requests.Session(impersonate="chrome120")
CV_BASE = "https://comicvine.gamespot.com/api"
CV_HEADERS = {"User-Agent": "MyComicCollection/1.0"}

GRADE_SCALE = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5,
               5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0,
               9.2, 9.4, 9.6, 9.8, 10.0]


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


def _format_grade(g):
    return f"{g:.1f}"


def _strip_html(text):
    return bs4.BeautifulSoup(text, 'html.parser').get_text(separator=' ').strip()


def _classify_age(cover_date):
    if not cover_date or cover_date == 'Unknown':
        return 'Unknown'
    try:
        year = int(str(cover_date)[:4])
        if year < 1956:   return 'Golden Age'
        elif year < 1970: return 'Silver Age'
        elif year < 1985: return 'Bronze Age'
        elif year < 1992: return 'Copper Age'
        else:             return 'Modern Age'
    except (ValueError, TypeError):
        return 'Unknown'


# =============================================================================
#   Comic Vine
# =============================================================================

def _cv_get(endpoint, params):
    params.update({'api_key': CV_API_KEY, 'format': 'json'})
    url = f"{CV_BASE}/{endpoint}/"
    try:
        time.sleep(random.uniform(1, 2))
        resp = session.get(url, params=params, headers=CV_HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data.get('status_code') == 1:
            return data.get('results')
        else:
            print(f"     CV API error: {data.get('error')} (code {data.get('status_code')})")
            return None
    except Exception as e:
        print(f"     CV request failed: {e}")
        return None


def SearchComicVine(title, issue, variant=''):
    results = _cv_get('issues', {
        'filter': f'volume.name:{title},issue_number:{issue}',
        'field_list': (
            'id,name,issue_number,volume,description,deck,'
            'image,cover_date,store_date,cover_price,'
            'site_detail_url,character_credits'
        ),
        'limit': 10,
    })

    if not results:
        print(f"     CV: No exact match — trying name search")
        results = _cv_get('issues', {
            'filter': f'name:{title},issue_number:{issue}',
            'field_list': (
                'id,name,issue_number,volume,description,deck,'
                'image,cover_date,store_date,cover_price,'
                'site_detail_url,character_credits'
            ),
            'limit': 10,
        })

    if not results:
        print(f"     CV: No results for {title} #{issue}")
        return None

    full_name = f"{title} #{issue}{variant}".upper()
    best, best_score = None, 0

    for r in results:
        vol_name = r.get('volume', {}).get('name', '') if r.get('volume') else ''
        candidate = f"{vol_name} #{r.get('issue_number', '')}".upper()
        score = similar(candidate, full_name)
        if score > best_score:
            best_score = score
            best = r

    if best is None or best_score < 0.3:
        print(f"     CV: Best match too low ({int(best_score*100)}%) — skipping")
        return None

    print(f"     CV: '{best.get('volume', {}).get('name', '')} #{best.get('issue_number')}' "
          f"({int(best_score*100)}% confidence)")

    cover_date = best.get('cover_date') or best.get('store_date') or 'Unknown'
    description = best.get('description') or ''
    deck = best.get('deck') or ''
    key_issue = 'Yes' if any(kw in (description + deck).lower() for kw in [
        'key issue', 'first appearance', '1st appearance', 'origin',
        'death of', 'first app', '1st app'
    ]) else 'No'

    image = ''
    if best.get('image'):
        image = best['image'].get('original_url') or best['image'].get('medium_url') or ''

    chars = best.get('character_credits') or []
    characters = ', '.join([c.get('name', '') for c in chars[:10]])

    publisher = ''
    vol = best.get('volume') or {}
    if vol.get('id'):
        vol_data = _cv_get(f"volume/4050-{vol['id']}", {'field_list': 'publisher,name,start_year'})
        if vol_data and isinstance(vol_data, dict):
            publisher = (vol_data.get('publisher') or {}).get('name', '')

    return {
        'publisher':   publisher,
        'volume':      vol.get('name', ''),
        'published':   cover_date,
        'cover_price': best.get('cover_price') or 'Unknown',
        'comic_age':   _classify_age(cover_date),
        'notes':       deck or _strip_html(description[:500]),
        'key_issue':   key_issue,
        'cover_image': image,
        'book_link':   best.get('site_detail_url') or '',
        'characters':  characters,
        'confidence':  best_score,
    }


# =============================================================================
#   eBay Pricing
# =============================================================================

def _ebay_sold_prices(query):
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
    cgc_str = "CGC" if cgc.upper() != 'NO' else ""
    grade_str = _format_grade(grade_float)
    query = f"{title} #{issue} {cgc_str} {grade_str} {variant}".strip()
    prices = _ebay_sold_prices(query)
    if prices:
        print(f"     eBay [{grade_str}]: ${prices[0]:.2f}  ({len(prices)} sales found)")
        return prices[0]
    return None


def GetEbayPrice(title, issue, grade, cgc, variant=''):
    try:
        target = float(grade)
    except ValueError:
        print(f"     eBay: Cannot parse grade '{grade}'")
        return None

    price = _ebay_price_for_grade(title, issue, target, cgc, variant)
    if price is not None:
        return price

    print(f"     eBay: No sales for grade {grade} – searching nearby grades…")
    if target not in GRADE_SCALE:
        print(f"     eBay: Grade {target} not in standard scale.")
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
        print(f"     eBay: Interpolated → ${interpolated:.2f}")
        return interpolated
    elif lower_price is not None:
        return lower_price
    elif upper_price is not None:
        return upper_price

    print(f"     eBay: No usable data for {title} #{issue}.")
    return None


# =============================================================================
#  Main — TEST MODE
# =============================================================================
print(f"\n{'='*60}")
print(f"TEST RUN — first {TEST_LIMIT} records only")
print(f"{'='*60}\n")

gc = gspread.service_account()
sh = gc.open(Google_Workbook)
worksheet = sh.worksheet(Google_Sheet)
StartingDF = pd.DataFrame(worksheet.get_all_records())
sortedsheet = StartingDF.sort_values(by=['Title', 'Volume', 'Issue']).copy()

print(f"Loaded {len(sortedsheet)} total records. Processing first {TEST_LIMIT}...\n")
test_sheet = sortedsheet.head(TEST_LIMIT).copy()

for index, thisComic in test_sheet.iterrows():
    try:
        title   = str(thisComic['Title']).strip().upper()
        issue   = int(str(thisComic['Issue']).strip())
        grade   = str(thisComic['Grade']).strip()
        cgc     = "No" if thisComic['CGC Graded'] is None else thisComic['CGC Graded']
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
            print(f'WARNING: Could not parse Price Paid: {raw_paid!r}')
        if price_paid == 0:
            price_paid = 0.01

        test_sheet.at[index, 'Price Paid'] = price_paid
        print(f"\n[{list(test_sheet.index).index(index)+1}/{TEST_LIMIT}] {fullName}")

        # --- Comic Vine ---
        cv_data = SearchComicVine(title, issue, variant)
        if cv_data:
            publisher   = cv_data['publisher']
            volume      = cv_data['volume']
            published   = cv_data['published']
            cover_price = cv_data['cover_price']
            comic_age   = cv_data['comic_age']
            notes       = cv_data['notes']
            keyIssue    = cv_data['key_issue']
            image       = cv_data['cover_image']
            url_link    = cv_data['book_link']
            confidence  = cv_data['confidence']
        else:
            publisher = volume = published = cover_price = comic_age = ''
            notes = keyIssue = image = url_link = ''
            confidence = None
            print("     CV lookup failed — values will be blank")

        # --- eBay ---
        if len(grade) < 3:
            grade = grade + ".0"

        value = GetEbayPrice(title, issue, grade, cgc, variant)
        if value is None:
            try:
                value = float(str(thisComic.get('Value', '0')).replace('$', '').replace(',', '')) or 0.0
            except (ValueError, TypeError):
                value = 0.0
            print(f"     eBay unavailable — keeping existing: ${value:.2f}")

        test_sheet.at[index, 'Publisher']   = publisher
        test_sheet.at[index, 'Volume']      = volume
        test_sheet.at[index, 'Published']   = published
        test_sheet.at[index, 'KeyIssue']    = keyIssue
        test_sheet.at[index, 'Cover Price'] = cover_price
        test_sheet.at[index, 'Comic Age']   = comic_age
        test_sheet.at[index, 'Notes']       = notes
        test_sheet.at[index, 'Confidence']  = confidence
        test_sheet.at[index, 'Book Link']   = url_link
        test_sheet.at[index, 'Cover Image'] = image
        test_sheet.at[index, rundate]       = value

        print(f"     ✓ Publisher: {publisher or 'N/A'} | Age: {comic_age} | Value: ${value:.2f}")

    except Exception as e:
        print(f"     ERROR: {e}")
        continue

# Write the 10 updated rows back
print(f"\n{'='*60}")
print("Writing results back to Google Sheet...")

for index, row in test_sheet.iterrows():
    for col in test_sheet.columns:
        sortedsheet.at[index, col] = row[col]

sortedsheet.fillna('', inplace=True)
worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())

print("✓ Google Sheet updated.")
print(f"{'='*60}")
print("TEST RUN COMPLETE")
