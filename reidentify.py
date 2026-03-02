"""
reidentify.py — Force re-identification of a subset of comics via Comic Vine.

Clears the Identification Date for the target rows so SearchComicVine() runs,
then writes corrected metadata back to the sheet. eBay pricing is NOT re-run
(to keep runtime short and avoid unnecessary API calls).

Usage:
    python3 reidentify.py              # re-identify first 100 comics (sorted order)
    python3 reidentify.py --limit 50   # first 50
    python3 reidentify.py --all        # entire sheet
"""
from __future__ import print_function
import argparse
import os
import re
import sys

from curl_cffi import requests
from dotenv import load_dotenv
import gspread
import pandas as pd

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '.env'))
sys.path.insert(0, os.path.dirname(__file__))

import time
from core import SearchComicVine, safe_fillna, rundate, ID_DATE_COL

# ---------------------------------------------------------------------------
#  Args
# ---------------------------------------------------------------------------
parser = argparse.ArgumentParser()
group = parser.add_mutually_exclusive_group()
group.add_argument('--limit', type=int, default=100, help='Number of comics to re-identify (default: 100)')
group.add_argument('--all',   action='store_true',   help='Re-identify all comics in the sheet')
args = parser.parse_args()

# ---------------------------------------------------------------------------
#  Connect (with retry on transient Google API errors)
# ---------------------------------------------------------------------------
Google_Workbook = os.getenv('GOOGLE_WORKBOOK')
Google_Sheet    = os.getenv('GOOGLE_SHEET')
CV_API_KEY      = os.getenv('CV_API_KEY')

def _gsheet_connect(workbook, sheet, retries=3):
    for attempt in range(retries):
        try:
            gc        = gspread.service_account()
            sh        = gc.open(workbook)
            worksheet = sh.worksheet(sheet)
            return gc, sh, worksheet
        except Exception as e:
            if attempt < retries - 1:
                wait = 10 * (attempt + 1)
                print(f"Google Sheets error (attempt {attempt+1}): {e} — retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise

gc, sh, worksheet = _gsheet_connect(Google_Workbook, Google_Sheet)

df = pd.DataFrame(worksheet.get_all_records())
if ID_DATE_COL not in df.columns:
    df[ID_DATE_COL] = ''

# Build column-name → sheet column index map (1-based)
headers = worksheet.row_values(1)
col_index = {name: i+1 for i, name in enumerate(headers)}

sorted_df = df.sort_values(by=['Title', 'Volume', 'Issue'])

limit = len(sorted_df) if args.all else args.limit
target = sorted_df.head(limit)

print(f"Re-identifying {len(target)} comics (of {len(sorted_df)} total)...")
print(f"Run date: {rundate}\n")

session = requests.Session(impersonate="chrome120")

fixed   = 0
skipped = 0
failed  = 0

COOLDOWN_EVERY  = 100  # pause every N comics to let CV rate limit recover
COOLDOWN_SECS   = 60   # seconds to pause
CHECKPOINT_EVERY = 250  # flush changes to Google Sheets every N fixed rows

# Columns we update on each identified row
UPDATE_COLS = [
    'Publisher', 'Published', 'KeyIssue', 'Cover Price',
    'Comic Age', 'Notes', 'Confidence', 'Cover Image',
    'Book Link', ID_DATE_COL,
]

# Track pending row updates: sheet_row (1-based) → list of (col_index, value)
pending = {}  # sheet_row -> {col: value}

def _sheet_row(df_index):
    """Convert DataFrame index to 1-based Google Sheet row (header is row 1)."""
    # sorted_df positions map back to original df positions; sheet row = original df index + 2
    return df_index + 2

def send_whatsapp(msg):
    """Fire a WhatsApp message to Chris via openclaw CLI."""
    import subprocess
    try:
        subprocess.run(
            ['openclaw', 'message', 'send',
             '--channel', 'whatsapp',
             '--to', '+14079214706',
             '--message', msg],
            timeout=15, check=False
        )
    except Exception as e:
        print(f"  [WA notify failed: {e}]")

def flush_pending(pending, worksheet, label='', fixed_total=0, total=0):
    """Write all pending updates to the sheet in one batch call per row."""
    if not pending:
        return
    # Build a list of Cell objects for batch update
    cells = []
    for sheet_row, col_vals in pending.items():
        for col_idx, value in col_vals.items():
            cells.append(gspread.Cell(sheet_row, col_idx, value))
    for attempt in range(3):
        try:
            worksheet.update_cells(cells, value_input_option='USER_ENTERED')
            msg = f"✅ Checkpoint: {len(pending)} rows written to sheet ({fixed_total}/{total} comics identified so far)"
            print(f"  ✓ {msg}")
            send_whatsapp(msg)
            pending.clear()
            return
        except Exception as e:
            wait = 10 * (attempt + 1)
            print(f"  Sheet write error (attempt {attempt+1}): {e} — retrying in {wait}s...")
            time.sleep(wait)
    err_msg = f"⚠️ Checkpoint write FAILED after 3 attempts — {len(pending)} rows not saved yet"
    print(f"  ✗ {err_msg}")
    send_whatsapp(err_msg)

for index, thisComic in target.iterrows():
    processed = fixed + skipped + failed
    if processed > 0 and processed % COOLDOWN_EVERY == 0:
        print(f"\n--- Cooldown pause {COOLDOWN_SECS}s after {processed} comics (CV rate limit) ---\n")
        time.sleep(COOLDOWN_SECS)
    title   = str(thisComic['Title']).strip().upper()
    issue   = str(thisComic['Issue']).strip()
    variant = str(thisComic.get('Variant', '')).strip()
    variant = '' if variant in ('nan', 'None') else variant

    existing_publisher = str(thisComic.get('Publisher', '')).strip()
    raw_volume = str(thisComic.get('Volume', '')).strip()
    vol_match  = re.search(r'\d+', raw_volume)
    volume_number = int(vol_match.group()) if vol_match else None

    existing_book_link = str(thisComic.get('Book Link', '')).strip()
    existing_book_link = '' if existing_book_link in ('nan', 'None') else existing_book_link

    fullName = f"{title} #{issue}{variant}"
    print(f"[{fixed+skipped+failed+1}/{len(target)}] {fullName}")

    try:
        cv_data = SearchComicVine(
            session, CV_API_KEY,
            title=title,
            issue=issue,
            variant=variant,
            volume_number=volume_number,
            publisher=existing_publisher,
        )
    except Exception as e:
        print(f"     ERROR: {e}")
        failed += 1
        continue

    if not cv_data:
        print(f"     No confident match — leaving row unchanged")
        skipped += 1
        continue

    # Show what's changing
    old_link = existing_book_link or sorted_df.at[index, 'Book Link']
    new_link = cv_data['book_link']
    if old_link != new_link:
        print(f"     Book Link: {old_link}")
        print(f"           → : {new_link}")

    old_pub = sorted_df.at[index, 'Publisher']
    new_pub = cv_data['publisher']
    if old_pub != new_pub:
        print(f"     Publisher: '{old_pub}' → '{new_pub}'")

    old_date = sorted_df.at[index, 'Published']
    new_date = cv_data['published']
    if str(old_date) != str(new_date):
        print(f"     Published: '{old_date}' → '{new_date}'")

    # Stage corrections in DataFrame
    new_vals = {
        'Publisher':   cv_data['publisher'],
        'Published':   cv_data['published'],
        'KeyIssue':    cv_data['key_issue'],
        'Cover Price': cv_data['cover_price'],
        'Comic Age':   cv_data['comic_age'],
        'Notes':       cv_data['notes'],
        'Confidence':  round(cv_data['confidence'], 4),
        'Cover Image': cv_data['cover_image'],
        'Book Link':   cv_data['book_link'],
        ID_DATE_COL:   rundate,
    }
    for col, val in new_vals.items():
        sorted_df.at[index, col] = val

    # Queue for sheet write
    sheet_row = _sheet_row(index)
    pending[sheet_row] = {
        col_index[col]: val
        for col, val in new_vals.items()
        if col in col_index
    }

    fixed += 1

    # Checkpoint every CHECKPOINT_EVERY fixed rows
    if fixed % CHECKPOINT_EVERY == 0:
        flush_pending(pending, worksheet, label=f"after {fixed} fixed", fixed_total=fixed, total=len(target))

# ---------------------------------------------------------------------------
#  Final flush
# ---------------------------------------------------------------------------
print(f"\nResults: {fixed} fixed, {skipped} skipped (no match), {failed} errors")
if pending:
    print("Writing final checkpoint to Google Sheet...")
    flush_pending(pending, worksheet, label="final", fixed_total=fixed, total=len(target))

print("Done.")
