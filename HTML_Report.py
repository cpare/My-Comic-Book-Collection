#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML Report Generator for Comic Book Collection
Reads from Google Sheets and generates comics.html
Non-interactive: workbook/sheet names read from .env
"""

import os
import locale
import gspread
import pandas as pd
from datetime import date
from dotenv import load_dotenv

load_dotenv(dotenv_path='.env')

locale.setlocale(locale.LC_ALL, '')

rundate = date.today().strftime("%Y-%m-%d")

GOOGLE_WORKBOOK = os.getenv('GOOGLE_WORKBOOK', 'Comics')
GOOGLE_SHEET    = os.getenv('GOOGLE_SHEET', 'Real')

HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Comic Book Collection</title>
<style>
body {{
    color: whitesmoke;
    background-color: #282828;
    font-family: Arial, Helvetica, sans-serif;
    margin: 0;
    padding: 10px;
}}
a {{ color: whitesmoke; text-decoration: none; }}
.cgc {{
    background-color: rgb(148, 7, 35);
    z-index: 5;
    position: absolute;
    bottom: 0;
    left: 0;
    padding: 2px 6px;
    font-size: 12px;
    font-weight: bold;
}}
.key {{
    background-color: rgb(212, 175, 55);
    z-index: 5;
    position: absolute;
    bottom: 0;
    right: 0;
    padding: 2px 6px;
    font-size: 12px;
    font-weight: bold;
    color: #222;
}}
.title {{
    position: relative;
    font-size: large;
    font-weight: bold;
    width: 210px;
    margin: 0 auto;
    margin-top: 8px;
}}
.published {{ color: lightgray; font-size: 13px; margin-top: 4px; }}
.notes {{
    position: absolute;
    top: 0;
    font-size: 14px;
    text-align: center;
    display: flex;
    justify-content: center;
    align-items: center;
    height: 400px;
    width: 210px;
    padding: 10px;
}}
.grade {{
    position: absolute;
    bottom: 50px;
    left: 0;
    font-size: 18px;
    width: 250px;
    text-align: center;
}}
.value {{
    position: absolute;
    bottom: 10px;
    left: 0;
    font-size: x-large;
    width: 250px;
    text-align: center;
}}
.hvrbox, .hvrbox * {{ box-sizing: border-box; padding: 0; }}
.hvrbox {{
    position: relative;
    display: inline-block;
    overflow: hidden;
    width: 250px;
    height: 400px;
    margin: 4px;
    vertical-align: top;
}}
.hvrbox img {{ width: 250px; height: 400px; object-fit: cover; }}
.hvrbox .hvrbox-layer_bottom {{ display: block; }}
.hvrbox .hvrbox-layer_top {{
    opacity: 0;
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0, 0, 0, 0.82);
    padding: 0;
    transition: opacity 0.3s ease-in-out;
}}
.hvrbox:hover .hvrbox-layer_top,
.hvrbox.active .hvrbox-layer_top {{ opacity: 1; }}
.hvrbox .hvrbox-text {{
    text-align: center;
    font-size: 15px;
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    padding: 10px 15px;
}}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def safe_str(val, default=''):
    """Return string value, substituting default for nan/None/empty."""
    s = str(val).strip()
    return default if s.lower() in ('nan', 'none', '') else s


def safe_float(val):
    """Parse a currency/numeric string to float, returning 0.0 on failure."""
    try:
        v = str(val).replace('$', '').replace(',', '').strip()
        return float(v) if v else 0.0
    except (ValueError, TypeError):
        return 0.0


def _flag(val):
    """Return True if value represents a truthy yes/cgc/key flag."""
    return str(val).strip().upper() not in ('NO', 'N', 'FALSE', '', 'NAN', 'NONE', '0')


def _issue_label(issue_raw, variant):
    """
    Build the display issue string, preserving letter suffixes (e.g. '500B', '275B').
    Falls back to raw string if it can't be parsed as int.
    """
    s = safe_str(issue_raw, '0')
    # Try pure int first
    try:
        return str(int(float(s))) + safe_str(variant)
    except (ValueError, TypeError):
        pass
    # Has a letter suffix like '500B' — keep as-is
    return s + safe_str(variant)


def generate_html_page(sortedsheet, outfile='comics.html'):
    """Render sorted DataFrame to a hover-card HTML gallery."""
    cards = []
    for _, comic in sortedsheet.iterrows():
        title     = safe_str(comic.get('Title', ''), 'Unknown').upper()
        notes     = safe_str(comic.get('Notes', ''))
        published = safe_str(comic.get('Published', ''))
        issue_lbl = _issue_label(comic.get('Issue', '0'), comic.get('Variant', ''))
        value     = safe_float(comic.get('Value', 0))
        image     = safe_str(comic.get('Cover Image', ''))   # NOT uppercased — it's a URL
        grade     = safe_str(comic.get('Grade', ''))
        url       = safe_str(comic.get('Book Link', ''))
        cgc       = _flag(comic.get('CGC Graded', 'No'))
        key       = _flag(comic.get('KeyIssue', 'No'))

        try:
            value_str = locale.currency(value, grouping=True)
        except Exception:
            value_str = f'${value:.2f}'

        cgc_badge = "<div class='cgc'>CGC</div>" if cgc else ''
        key_badge = "<div class='key'>KEY</div>" if key else ''

        card = (
            f"<div class='hvrbox'>"
            f"<img src='{image}' alt='{title} #{issue_lbl}' class='hvrbox-layer_bottom'>\n"
            f"\t<div class='hvrbox-layer_top'>\n"
            f"\t\t<div class='hvrbox-text'>\n"
            f"\t\t\t<div class='title'><a href='{url}'>{title} #{issue_lbl}</a></div>\n"
            f"\t\t\t<div class='published'>{published}</div>\n"
            f"\t\t\t<div class='notes'>{notes}</div>\n"
            f"\t\t\t<div class='grade'>Grade: {grade}</div>\n"
            f"\t\t\t<div class='value'>{value_str}</div>\n"
            f"\t\t\t{cgc_badge}\n"
            f"\t\t\t{key_badge}\n"
            f"\t\t</div>\n"
            f"\t</div>\n"
            f"</div>\n"
        )
        cards.append(card)

    with open(outfile, 'w', encoding='utf-8') as f:
        f.write(HTML_TEMPLATE.format(body='\n'.join(cards)))

    print(f"Generated {outfile} ({len(sortedsheet)} comics)")


def read_google_sheet(workbook, sheet):
    gc = gspread.service_account()
    sh = gc.open(workbook)
    worksheet = sh.worksheet(sheet)
    df = pd.DataFrame(worksheet.get_all_records())
    sorted_df = df.sort_values(by=['Title', 'Issue'])
    return df, sorted_df, worksheet, sh


def backup_sheet(sh, df, label=None):
    """Create a backup worksheet. Skips if one with this label already exists."""
    label = label or rundate
    title = f"Backup {label}"
    existing = [w.title for w in sh.worksheets()]
    if title in existing:
        print(f"Backup '{title}' already exists — skipping")
        return None
    rows = df.shape[0] + 2
    cols = df.shape[1]
    backup_ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
    backup_ws.update([df.columns.values.tolist()] + df.values.tolist())
    print(f"Backup created: '{title}'")
    return backup_ws


if __name__ == '__main__':
    print(f"Reading sheet '{GOOGLE_SHEET}' from workbook '{GOOGLE_WORKBOOK}'...")
    df, sorted_df, worksheet, sh = read_google_sheet(GOOGLE_WORKBOOK, GOOGLE_SHEET)
    print(f"  {len(df)} comics loaded")

    print("Checking backup...")
    backup_sheet(sh, df)

    print("Generating HTML report...")
    generate_html_page(sorted_df)
    print("Done.")
