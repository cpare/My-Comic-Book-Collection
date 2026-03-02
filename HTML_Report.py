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


def safe_str(val, default=''):
    s = str(val).strip()
    return default if s.lower() == 'nan' or s.lower() == 'none' else s


def safe_float(val):
    try:
        v = str(val).replace('$', '').replace(',', '').strip()
        return float(v) if v else 0.0
    except (ValueError, TypeError):
        return 0.0


def generate_html_page(sortedsheet, outfile='comics.html'):
    html_body = ''
    for _, comic in sortedsheet.iterrows():
        title     = safe_str(comic.get('Title', ''), '').upper()
        notes     = safe_str(comic.get('Notes', ''))
        published = safe_str(comic.get('Published', ''))
        issue_raw = safe_str(comic.get('Issue', '0'), '0')
        try:
            issue = int(float(issue_raw))
        except (ValueError, TypeError):
            issue = 0
        value     = safe_float(comic.get('Value', 0))
        image     = safe_str(comic.get('Cover Image', ''), '').upper()
        grade     = safe_str(comic.get('Grade', ''))
        cgc       = safe_str(comic.get('CGC Graded', 'No'), 'No')
        key       = safe_str(comic.get('KeyIssue', 'No'), 'No')
        variant   = safe_str(comic.get('Variant', ''))
        url       = safe_str(comic.get('Book Link', ''))

        cgc_div = '' if cgc.upper() in ('NO', 'N', 'FALSE', '', 'NAN', 'NONE') else "<div class='cgc'>CGC</div>"
        key_div = '' if key.upper()  in ('NO', 'N', 'FALSE', '', 'NAN', 'NONE') else "<div class='key'>KEY</div>"

        published_div = f"<div class='published'>{published}</div>"
        title_div     = f"<div class='title'><a href='{url}'>{title} #{issue}{variant}</a></div>"
        notes_div     = f"<div class='notes'>{notes}</div>"
        grade_div     = f"<div class='grade'>Grade: {grade}</div>"
        try:
            value_str = locale.currency(value, grouping=True)
        except Exception:
            value_str = f'${value:.2f}'
        value_div = f"<div class='value'>{value_str}</div>"

        html_body += f"<div class='hvrbox'><img src='{image}' alt='Cover' class='hvrbox-layer_bottom'>\n"
        html_body += "\t<div class='hvrbox-layer_top'>\n"
        html_body += "\t\t<div class='hvrbox-text'>\n"
        html_body += f"\t\t\t{title_div}\n"
        html_body += f"\t\t\t{published_div}\n"
        html_body += f"\t\t\t{notes_div}\n"
        html_body += f"\t\t\t{grade_div}\n"
        html_body += f"\t\t\t{value_div}\n"
        html_body += f"\t\t\t{cgc_div}\n"
        html_body += f"\t\t\t{key_div}\n"
        html_body += "\t\t</div>\n"
        html_body += "\t</div>\n"
        html_body += "</div>\n"

    css = """
<html>
<head>
<meta charset="utf-8">
<title>Comic Book Collection</title>
<style type="text/css">
body {
    color: whitesmoke;
    background-color: #282828;
    font-family: Arial, Helvetica, sans-serif;
    margin: 0;
    padding: 10px;
}
a { color: whitesmoke; text-decoration: none; }
.cgc {
    background-color: rgb(148, 7, 35);
    z-index: 5;
    position: absolute;
    bottom: 0;
    left: 0;
    padding: 2px 6px;
    font-size: 12px;
    font-weight: bold;
}
.key {
    background-color: rgb(212, 175, 55);
    z-index: 5;
    position: absolute;
    bottom: 0;
    right: 0;
    padding: 2px 6px;
    font-size: 12px;
    font-weight: bold;
    color: #222;
}
.title {
    position: relative;
    font-size: large;
    font-weight: bold;
    width: 210px;
    margin: 0 auto;
    margin-top: 8px;
}
.published { color: lightgray; font-size: 13px; margin-top: 4px; }
.notes {
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
}
.grade {
    position: absolute;
    bottom: 50px;
    left: 0;
    font-size: 18px;
    width: 250px;
    text-align: center;
}
.value {
    position: absolute;
    bottom: 10px;
    left: 0;
    font-size: x-large;
    width: 250px;
    text-align: center;
}
.hvrbox, .hvrbox * { box-sizing: border-box; padding: 0; }
.hvrbox {
    position: relative;
    display: inline-block;
    overflow: hidden;
    width: 250px;
    height: 400px;
    margin: 4px;
    vertical-align: top;
}
.hvrbox img { width: 250px; height: 400px; object-fit: cover; }
.hvrbox .hvrbox-layer_bottom { display: block; }
.hvrbox .hvrbox-layer_top {
    opacity: 0;
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0, 0, 0, 0.82);
    padding: 0;
    transition: all 0.3s ease-in-out 0s;
}
.hvrbox:hover .hvrbox-layer_top,
.hvrbox.active .hvrbox-layer_top { opacity: 1; }
.hvrbox .hvrbox-text {
    text-align: center;
    font-size: 15px;
    display: inline-block;
    padding: 10px 15px;
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
}
</style>
</head>
<body>
"""

    with open(outfile, 'w', encoding='utf-8') as f:
        f.write(css)
        f.write(html_body)
        f.write("\n</body></html>\n")

    print(f"Generated {outfile} ({len(sortedsheet)} comics)")


def read_google_sheet(workbook, sheet):
    gc = gspread.service_account()
    sh = gc.open(workbook)
    worksheet = sh.worksheet(sheet)
    df = pd.DataFrame(worksheet.get_all_records())
    sorted_df = df.sort_values(by=['Title', 'Issue'])
    return df, sorted_df, worksheet, sh


def backup_sheet(sh, df, label=None):
    label = label or rundate
    title = f"Backup {label}"
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

    print("Creating backup...")
    backup_sheet(sh, df)

    print("Generating HTML report...")
    generate_html_page(sorted_df)
    print("Done.")
