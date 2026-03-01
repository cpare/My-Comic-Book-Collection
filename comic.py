"""
comic.py — Main script. Reads Google Sheet, enriches with Comic Vine + eBay,
writes results back and generates an HTML collection page.

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
import locale
import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from datetime import date

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

Google_Workbook = os.getenv('GOOGLE_WORKBOOK') or input('Google Workbook Name:    ')
Google_Sheet    = os.getenv('GOOGLE_SHEET')    or input('Google Worksheet Name:    ')
CV_API_KEY      = os.getenv('CV_API_KEY')      or input('Comic Vine API Key:    ')

session = requests.Session(impersonate="chrome120")


# =============================================================================
#   HTML Report
# =============================================================================

def generate_HTMLPage(sortedsheet):
    htmlBody = ''

    for index, thisComic in sortedsheet.iterrows():
        title   = str(thisComic['Title']).strip().upper()
        notes   = str(thisComic['Notes']).strip()
        issue   = int(str(thisComic['Issue']).strip())
        try:
            value = float(str(thisComic['Value']).strip().replace('$', '').replace(',', ''))
        except (ValueError, KeyError):
            value = 0.0
        image   = str(thisComic['Cover Image']).strip()
        grade   = str(thisComic['Grade']).strip()
        cgc     = str(thisComic.get('CGC Graded', 'No'))
        key     = str(thisComic.get('KeyIssue', 'No'))
        variant = '' if str(thisComic.get('Variant', '')).strip() in ('nan', '') else str(thisComic['Variant']).strip()
        url     = '' if str(thisComic.get('Book Link', '')).strip() in ('nan', '') else str(thisComic['Book Link']).strip()

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

def ReadGoogleSheet(workbook, sheet):
    gc = gspread.service_account()
    sh = gc.open(workbook)
    worksheet = sh.worksheet(sheet)
    df = pd.DataFrame(worksheet.get_all_records())
    if ID_DATE_COL not in df.columns:
        df[ID_DATE_COL] = ''
    sorted_df = df.sort_values(by=['Title', 'Volume', 'Issue'])
    return df, sorted_df, worksheet, sh


def BackupGoogleSheet(sh, starting_df, sorted_df):
    rows = starting_df.shape[0]
    cols = starting_df.shape[1]
    backup = sh.add_worksheet(title="Backup " + rundate, rows=rows, cols=cols)
    backup.update([sorted_df.columns.values.tolist()] + sorted_df.values.tolist())


# =============================================================================
#  Main
# =============================================================================
SheetData   = ReadGoogleSheet(Google_Workbook, Google_Sheet)
StartingDF  = SheetData[0]
sortedsheet = SheetData[1]
worksheet   = SheetData[2]
sh          = SheetData[3]

generate_HTMLPage(sortedsheet)
BackupGoogleSheet(sh, StartingDF, sortedsheet)

for index, thisComic in sortedsheet.iterrows():
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
            print(f'WARNING: Could not parse Price Paid: {thisComic.get("Price Paid")!r}')
        if price_paid == 0:
            price_paid = 0.01

        # =====================================================================
        #  Enrichment context (used for matching, will be written back)
        # =====================================================================
        existing_publisher = str(thisComic.get('Publisher', '')).strip()
        raw_volume = str(thisComic.get('Volume', '')).strip()
        vol_match = re.search(r'\d+', raw_volume)
        volume_number = int(vol_match.group()) if vol_match else None

        fullName = f"{title} #{issue}{variant}"
        print(f'Gathering : {fullName}')

        # =====================================================================
        #  Task 1: Book Identification (Comic Vine)
        #  Only runs if Identification Date is blank
        # =====================================================================
        id_date = str(thisComic.get(ID_DATE_COL, '')).strip()
        id_date = '' if id_date in ('nan', 'None') else id_date

        if id_date:
            print(f"     CV: Already identified on {id_date} — skipping")
            cv_data = None
        else:
            cv_data = SearchComicVine(
            session, CV_API_KEY, title, issue, variant,
            volume_number=volume_number,
                publisher=existing_publisher,
            )

        if cv_data:
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
            cv_publisher = cv_volume = published = cover_price = comic_age = ''
            notes = key_issue = cover_image = cv_book_link = ''
            confidence = None
            print(f"     CV: no confident match — enrichment fields unchanged")

        # =====================================================================
        #  eBay pricing
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
            try:
                current_value = float(
                    str(thisComic.get('Value', '0')).replace('$', '').replace(',', '')
                ) or 0.0
            except (ValueError, TypeError):
                current_value = 0.0
            print(f"     eBay unavailable — keeping existing value: ${current_value:.2f}")

        graded_gain = round((graded_price or 0.0) - price_paid, 2)

        # =====================================================================
        #  Write enrichment fields back
        # =====================================================================
        sortedsheet.at[index, 'Price Paid'] = price_paid

        if cv_data:
            sortedsheet.at[index, 'Publisher']   = cv_publisher
            # Volume is user-owned (volume number e.g. "Volume 1") — never overwrite
            sortedsheet.at[index, 'Published']   = published
            sortedsheet.at[index, 'KeyIssue']    = key_issue
            sortedsheet.at[index, 'Cover Price'] = cover_price
            sortedsheet.at[index, 'Comic Age']   = comic_age
            sortedsheet.at[index, 'Notes']       = notes
            sortedsheet.at[index, 'Confidence']  = round(confidence, 4)
            sortedsheet.at[index, 'Cover Image'] = cover_image
            sortedsheet.at[index, ID_DATE_COL]   = rundate
            if not existing_book_link:
                sortedsheet.at[index, 'Book Link'] = cv_book_link

        sortedsheet.at[index, 'Graded']      = graded_price   if graded_price   is not None else ''
        sortedsheet.at[index, 'Ungraded']    = ungraded_price if ungraded_price is not None else ''
        sortedsheet.at[index, 'Graded Gain'] = graded_gain
        sortedsheet.at[index, 'Value']       = current_value
        sortedsheet.at[index, rundate]       = current_value

    except Exception as e:
        print(f"Error while working on {title}: {e}")
        continue

# =============================================================================
#  Commit results back to Google Sheet (one API call at the end)
# =============================================================================
safe_fillna(sortedsheet)
worksheet.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())

generate_HTMLPage(sortedsheet)

print("Work is complete.")
