"""
test_run.py — Integration test. Processes first N records.

Two independent tasks:
  1. Book Identification (Comic Vine) — runs only if 'Identification Date' is blank.
     On success, writes enrichment fields + sets Identification Date = today.
  2. Book Valuation (eBay) — always runs regardless of identification status.

Input fields (user-owned, never overwritten):
  Title, Issue, Grade, CGC Graded, Variant, Price Paid, Book Link (if pre-populated)

Output/enrichment fields:
  Publisher, Volume, Published, KeyIssue, Cover Price, Comic Age, Notes,
  Confidence, Identification Date, Book Link (only if blank), Cover Image,
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
    ID_DATE_COL,
)

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '.env'))
locale.setlocale(locale.LC_ALL, '')

TEST_LIMIT      = 100
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

# Add Identification Date column if it doesn't exist
if ID_DATE_COL not in sortedsheet.columns:
    print(f"Adding '{ID_DATE_COL}' column to sheet...")
    sortedsheet[ID_DATE_COL] = ''

print(f"Loaded {len(sortedsheet)} total records. Processing first {TEST_LIMIT}...\n")
test_sheet = sortedsheet.head(TEST_LIMIT).copy()

# =============================================================================
#  Metrics
# =============================================================================
stats = {
    'cv_hit':      0,   # CV returned a confident match
    'cv_miss':     0,   # CV returned nothing / below threshold
    'cv_skipped':  0,   # Already identified — skipped
    'ebay_hit':    0,
    'ebay_miss':   0,
    'errors':      0,
}

# =============================================================================
#  Main loop
# =============================================================================
for index, thisComic in test_sheet.iterrows():
    n = list(test_sheet.index).index(index) + 1
    try:
        # =====================================================================
        #  READ-ONLY inputs
        # =====================================================================
        title   = str(thisComic['Title']).strip().upper()
        issue   = int(str(thisComic['Issue']).strip())
        grade   = str(thisComic['Grade']).strip()
        cgc     = str(thisComic.get('CGC Graded', 'No')).strip()
        variant = str(thisComic.get('Variant', '')).strip()
        variant = '' if variant in ('nan', 'None') else variant

        existing_book_link = str(thisComic.get('Book Link', '')).strip()
        existing_book_link = '' if existing_book_link in ('nan', 'None') else existing_book_link

        price_paid = 0.0
        try:
            raw_paid = thisComic.get('Price Paid', 0)
            if isinstance(raw_paid, str):
                price_paid = float(raw_paid.strip().replace('$', '').replace(',', '')) or 0.0
            elif isinstance(raw_paid, (int, float)):
                price_paid = float(raw_paid)
        except (ValueError, TypeError):
            pass
        if price_paid == 0:
            price_paid = 0.01

        # Identification Date — controls whether CV runs
        id_date = str(thisComic.get(ID_DATE_COL, '')).strip()
        id_date = '' if id_date in ('nan', 'None') else id_date

        # Enrichment context for matching
        existing_publisher = str(thisComic.get('Publisher', '')).strip()
        raw_volume = str(thisComic.get('Volume', '')).strip()
        vol_match = re.search(r'\d+', raw_volume)
        volume_number = int(vol_match.group()) if vol_match else None

        fullName = f"{title} #{issue}{variant}"
        print(f"\n[{n}/{TEST_LIMIT}] {fullName}")

        # =====================================================================
        #  Task 1: Book Identification (Comic Vine)
        #  Only runs if Identification Date is blank
        # =====================================================================
        if id_date:
            stats['cv_skipped'] += 1
            print(f"     CV: Already identified on {id_date} — skipping")
            cv_data = None
        else:
            cv_data = SearchComicVine(
                session, CV_API_KEY, title, issue, variant,
                volume_number=volume_number,
                publisher=existing_publisher,
            )
            if cv_data:
                stats['cv_hit'] += 1
            else:
                stats['cv_miss'] += 1

        # =====================================================================
        #  Task 2: Book Valuation (eBay) — always runs
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
            print(f"     eBay: no sales — keeping existing: ${current_value:.2f}")

        graded_gain = round((graded_price or 0.0) - price_paid, 2)

        # =====================================================================
        #  Write back
        # =====================================================================
        test_sheet.at[index, 'Price Paid'] = price_paid

        if cv_data:
            test_sheet.at[index, 'Publisher']    = cv_data['publisher']
            # Volume is user-owned (volume number e.g. "Volume 1") — never overwrite
            test_sheet.at[index, 'Published']    = cv_data['published']
            test_sheet.at[index, 'KeyIssue']     = cv_data['key_issue']
            test_sheet.at[index, 'Cover Price']  = cv_data['cover_price']
            test_sheet.at[index, 'Comic Age']    = cv_data['comic_age']
            test_sheet.at[index, 'Notes']        = cv_data['notes']
            test_sheet.at[index, 'Confidence']   = round(cv_data['confidence'], 4)
            test_sheet.at[index, 'Cover Image']  = cv_data['cover_image']
            test_sheet.at[index, ID_DATE_COL]    = rundate   # Mark as identified
            if not existing_book_link:
                test_sheet.at[index, 'Book Link'] = cv_data['book_link']

        test_sheet.at[index, 'Graded']      = graded_price   if graded_price   is not None else ''
        test_sheet.at[index, 'Ungraded']    = ungraded_price if ungraded_price is not None else ''
        test_sheet.at[index, 'Graded Gain'] = graded_gain
        test_sheet.at[index, 'Value']       = current_value
        test_sheet.at[index, rundate]       = current_value

        # Status line
        if cv_data:
            cv_tag = f"CV ✓ ({int(cv_data['confidence']*100)}%)"
        elif id_date:
            cv_tag = f"CV — (done {id_date})"
        else:
            cv_tag = "CV ✗"
        ebay_tag = f"eBay ✓ ${current_value:.2f}" if current_value else "eBay ✗"
        print(f"     {cv_tag} | {ebay_tag}")

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
cv_ran  = stats['cv_hit'] + stats['cv_miss']
print(f"\n{'='*60}")
print(f"RESULTS ({TEST_LIMIT} records)")
print(f"  CV identified   : {stats['cv_hit']:>3}/{cv_ran} ran  ({int(stats['cv_hit']/cv_ran*100) if cv_ran else 0}% hit rate)")
print(f"  CV skipped      : {stats['cv_skipped']:>3}  (already identified)")
print(f"  CV missed       : {stats['cv_miss']:>3}  (below confidence threshold)")
print(f"  eBay hits       : {stats['ebay_hit']:>3}/{TEST_LIMIT}  ({int(stats['ebay_hit']/TEST_LIMIT*100)}%)")
print(f"  Errors          : {stats['errors']:>3}")
print(f"{'='*60}")
print("TEST RUN COMPLETE")
