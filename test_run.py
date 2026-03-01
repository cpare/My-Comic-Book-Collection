"""
test_run.py — Integration test against live Google Sheet + APIs.
Processes first N records, writes results back, reports success rate.

Input fields (user-owned, never overwritten):
  Title, Issue, Grade, CGC Graded, Variant, Price Paid, Book Link (if pre-populated)

Output/enrichment fields (written by this script):
  Publisher, Volume, Published, KeyIssue, Cover Price, Comic Age, Notes,
  Confidence, Book Link (only if blank), Cover Image,
  Graded, Ungraded, Graded Gain, Value, <rundate>
"""
from __future__ import print_function
from curl_cffi import requests
import re
from datetime import date
import pandas as pd
import gspread
import locale
import os
from dotenv import load_dotenv

from core import (
    SearchComicVine,
    GetEbayPrice,
    GetEbayPriceGraded,
    GetEbayPriceUngraded,
    safe_fillna,
    normalise_grade,
    rundate,
)

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '.env'))
locale.setlocale(locale.LC_ALL, '')

TEST_LIMIT      = 20
Google_Workbook = os.getenv('GOOGLE_WORKBOOK') or input('Google Workbook Name: ')
Google_Sheet    = os.getenv('GOOGLE_SHEET')    or input('Google Sheet Name: ')
CV_API_KEY      = os.getenv('CV_API_KEY')      or input('Comic Vine API Key: ')

session = requests.Session(impersonate="chrome120")

# =============================================================================
#  Load sheet
# =============================================================================
print(f"\n{'='*60}")
print(f"TEST RUN — first {TEST_LIMIT} records")
print(f"{'='*60}\n")

gc        = gspread.service_account()
sh        = gc.open(Google_Workbook)
worksheet = sh.worksheet(Google_Sheet)
StartingDF   = pd.DataFrame(worksheet.get_all_records())
sortedsheet  = StartingDF.sort_values(by=['Title', 'Volume', 'Issue']).copy()

print(f"Loaded {len(sortedsheet)} total records. Processing first {TEST_LIMIT}...\n")
test_sheet = sortedsheet.head(TEST_LIMIT).copy()

# =============================================================================
#  Metrics
# =============================================================================
stats = {'cv_hit': 0, 'cv_miss': 0, 'ebay_hit': 0, 'ebay_miss': 0, 'errors': 0}

# =============================================================================
#  Main loop
# =============================================================================
for index, thisComic in test_sheet.iterrows():
    n = list(test_sheet.index).index(index) + 1
    try:
        # =====================================================================
        #  READ-ONLY inputs — user's source of truth, never overwritten
        # =====================================================================
        title   = str(thisComic['Title']).strip().upper()
        issue   = int(str(thisComic['Issue']).strip())
        grade   = str(thisComic['Grade']).strip()
        cgc     = str(thisComic.get('CGC Graded', 'No')).strip()
        variant = str(thisComic.get('Variant', '')).strip()
        variant = '' if variant in ('nan', 'None') else variant

        # Book Link: if user pre-populated, use it and don't overwrite
        existing_book_link = str(thisComic.get('Book Link', '')).strip()
        existing_book_link = '' if existing_book_link in ('nan', 'None') else existing_book_link

        # Price Paid
        price_paid = 0.0
        try:
            raw_paid = thisComic.get('Price Paid', 0)
            if isinstance(raw_paid, str):
                price_paid = float(raw_paid.strip().replace('$', '').replace(',', '')) or 0.0
            elif isinstance(raw_paid, (int, float)):
                price_paid = float(raw_paid)
        except (ValueError, TypeError):
            print(f'     WARNING: Could not parse Price Paid: {thisComic.get("Price Paid")!r}')
        if price_paid == 0:
            price_paid = 0.01

        # =====================================================================
        #  Enrichment inputs — read for matching context, will be written back
        # =====================================================================
        existing_publisher = str(thisComic.get('Publisher', '')).strip()

        # Volume column format: "Volume 1", "Vol. 2", or just "1"
        raw_volume = str(thisComic.get('Volume', '')).strip()
        vol_match = re.search(r'\d+', raw_volume)
        volume_number = int(vol_match.group()) if vol_match else None

        fullName = f"{title} #{issue}{variant}"
        print(f"\n[{n}/{TEST_LIMIT}] {fullName}")

        # =====================================================================
        #  Comic Vine metadata lookup
        # =====================================================================
        cv_data = SearchComicVine(
            session, CV_API_KEY, title, issue, variant,
            volume_number=volume_number,
            publisher=existing_publisher,
        )

        if cv_data:
            stats['cv_hit'] += 1
            cv_publisher = cv_data['publisher']
            cv_volume    = cv_data['volume']
            published    = cv_data['published']
            cover_price  = cv_data['cover_price']
            comic_age    = cv_data['comic_age']
            notes        = cv_data['notes']
            key_issue    = cv_data['key_issue']
            cover_image  = cv_data['cover_image']
            cv_book_link = cv_data['book_link']
            confidence   = cv_data['confidence']
        else:
            stats['cv_miss'] += 1
            cv_publisher = cv_volume = published = cover_price = comic_age = ''
            notes = key_issue = cover_image = cv_book_link = ''
            confidence = None
            print("     CV: no confident match — enrichment fields unchanged")

        # =====================================================================
        #  eBay pricing: graded, ungraded, and current-state value
        # =====================================================================
        grade_norm = normalise_grade(grade)

        graded_price   = GetEbayPriceGraded(session, title, issue, grade_norm or grade, variant)
        ungraded_price = GetEbayPriceUngraded(session, title, issue, grade_norm or grade, variant)

        is_cgc = cgc.upper() not in ('NO', 'N', 'FALSE', '', 'NAN', 'NONE')
        if is_cgc and graded_price is not None:
            current_value = graded_price
        elif ungraded_price is not None:
            current_value = ungraded_price
        else:
            current_value = None

        if current_value is not None:
            stats['ebay_hit'] += 1
        else:
            stats['ebay_miss'] += 1
            try:
                current_value = float(
                    str(thisComic.get('Value', '0')).replace('$', '').replace(',', '')
                ) or 0.0
            except (ValueError, TypeError):
                current_value = 0.0
            print(f"     eBay: no sales — keeping existing value: ${current_value:.2f}")

        graded_gain = round((graded_price or 0.0) - price_paid, 2)

        # =====================================================================
        #  Write enrichment fields back
        #  Rules:
        #    - Title, Issue, Grade, CGC Graded, Variant, Price Paid: NEVER touched
        #    - Publisher, Volume: written only if CV found them
        #    - Book Link: written only if user left it blank
        #    - All other enrichment fields: written if CV matched
        # =====================================================================
        if cv_data:
            test_sheet.at[index, 'Publisher']   = cv_publisher
            test_sheet.at[index, 'Volume']      = cv_volume
            test_sheet.at[index, 'Published']   = published
            test_sheet.at[index, 'KeyIssue']    = key_issue
            test_sheet.at[index, 'Cover Price'] = cover_price
            test_sheet.at[index, 'Comic Age']   = comic_age
            test_sheet.at[index, 'Notes']       = notes
            test_sheet.at[index, 'Confidence']  = round(confidence, 4)
            test_sheet.at[index, 'Cover Image'] = cover_image
            if not existing_book_link:
                test_sheet.at[index, 'Book Link'] = cv_book_link

        test_sheet.at[index, 'Price Paid']   = price_paid
        test_sheet.at[index, 'Graded']       = graded_price   if graded_price   is not None else ''
        test_sheet.at[index, 'Ungraded']     = ungraded_price if ungraded_price is not None else ''
        test_sheet.at[index, 'Graded Gain']  = graded_gain
        test_sheet.at[index, 'Value']        = current_value
        test_sheet.at[index, rundate]        = current_value

        cv_tag   = f"CV ✓ ({int(confidence*100)}%)" if cv_data else "CV ✗"
        ebay_tag = f"eBay ✓ ${current_value:.2f}" if stats['ebay_hit'] > stats['ebay_miss'] else f"eBay ✗ (kept ${current_value:.2f})"
        print(f"     {cv_tag} | {ebay_tag} | {comic_age or 'Age unknown'}")

    except Exception as e:
        stats['errors'] += 1
        print(f"     ERROR: {e}")
        import traceback; traceback.print_exc()
        continue

# =============================================================================
#  Write results back to Google Sheet
# =============================================================================
print(f"\n{'='*60}")
print("Writing results back to Google Sheet...")

for index, row in test_sheet.iterrows():
    for col in test_sheet.columns:
        sortedsheet.at[index, col] = row[col]

safe_fillna(sortedsheet)
worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())
print("✓ Google Sheet updated.")

# =============================================================================
#  Results summary
# =============================================================================
total     = TEST_LIMIT
cv_rate   = stats['cv_hit']   / total * 100
ebay_rate = stats['ebay_hit'] / total * 100

print(f"\n{'='*60}")
print(f"RESULTS ({total} records)")
print(f"  Comic Vine hits : {stats['cv_hit']:>3}/{total}  ({cv_rate:.0f}%)")
print(f"  eBay hits       : {stats['ebay_hit']:>3}/{total}  ({ebay_rate:.0f}%)")
print(f"  Errors          : {stats['errors']:>3}")
print(f"{'='*60}")
print("TEST RUN COMPLETE")
