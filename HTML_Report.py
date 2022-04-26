#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Apr 16 16:27:20 2022

@author: cpare
"""


from __future__ import print_function
from selenium import webdriver
import bs4
import time
from difflib import SequenceMatcher
from datetime import date
import pandas as pd
import sys
import random
import gspread
import locale

locale.setlocale( locale.LC_ALL, '' )

# =============================================================================
#   Variables
# =============================================================================
rundate = date.today().strftime("%Y-%m-%d")
# The error codes
NO_SEARCH_RESULTS_FOUND = 1

Google_Workbook = input('Google Workbook Name:    ')
Google_Sheet = input('Google Worksheet Name:    ')


def generate_HTMLPage(sortedsheet):
    htmlBody = ''
    for index, thisComic in sortedsheet.iterrows():
        title = str(thisComic['Title']).strip().upper()
        notes = str(thisComic['Notes']).strip()
        published = str(thisComic['Published']).strip()
        issue = int(str(thisComic['Issue']).strip())
        value = thisComic['Value']
        if (value == ''):
            value = 0
        if (type(value) != float and type(value) != int):
            value = value.replace("$", "")
            value = float(value)
        image = str(thisComic['Cover Image']).strip().upper()
        grade = str(thisComic['Grade']).strip()
        cgc = "No" if thisComic['CGC Graded'] == None else thisComic['CGC Graded']
        key = "No" if thisComic['KeyIssue'] == None else thisComic['KeyIssue']
        variant = '' if str(thisComic['Variant']).strip() == 'nan' else str(thisComic['Variant']).strip()
        url = '' if str(thisComic['Book Link']).strip() == 'nan' else str(thisComic['Book Link']).strip()
        if cgc.upper() =='NO':
            cgc_div = ''
        else: 
            cgc_div = "<div class='cgc'>CGC</div>"
        if key.upper() =='NO':
            key_div = ''
        else: 
            key_div = "<div class='key'>KEY</div>"
        published_div = "<div class='published'>" + str(published) + "</div>"
        title_div = "<div class='title'><a href='" + str(url) + "'>" + str(title) + " #" + str(issue) + str(variant) +"</a></div>"
        notes_div = "<div class='notes'>" + str(notes) + "</div>"
        grade_div = "<div class='grade'>Grade:  " + str(grade) + "</div>"
        value_div = "<div class='value'>" + locale.currency(value, grouping=True) + "</div>"
        
        htmlBody += "<div class='hvrbox'><img src='"  +  str(image) + "' alt='Cover' class='hvrbox-layer_bottom'>\n"
        htmlBody += "\t<div class='hvrbox-layer_top'>\n"
        htmlBody += "\t\t<div class='hvrbox-text'>\n" 
        htmlBody += "\t\t\t" + str(title_div) + "\n"
        htmlBody += "\t\t\t" + str(published_div) + "\n"        
        htmlBody += "\t\t\t" + str(notes_div) + "\n"
        htmlBody += "\t\t\t" + str(grade_div) + "\n"
        htmlBody += "\t\t\t" + str(value_div) + "\n"
        htmlBody += "\t\t\t" + str(cgc_div) + "\n"
        htmlBody += "\t\t\t" + str(key_div) + "\n"

        #htmlBody +=  str(title_div) + str(notes_div) + str(grade_div) + str(value_div) + str(cgc_div) + str(key_div)
        htmlBody += "\t\t</div>\n"
        htmlBody += "\t</div>\n"
        htmlBody += "</div>\n"
        
    with open("comics.html",'w') as f:
        f.write(""" 
    <html>
    <head>
    <style type'"text/css">
    body {color: whitesmoke; background-color: 282828;font-family: Arial, Helvetica, sans-serif;}
    a {color: whitesmoke; text-decoration: none;}
    .cgc {background-color: rgb(148, 7, 35);z-index: inherit 5;position: absolute; bottom:0; left:0;}
    .key {background-color: rgb(212, 175, 55);z-index: inherit 5;position: absolute;bottom:0; right: 0;}
    .title {position: relative; left: -20; font-size:x-large; align-content: center; width: 250px;}
    .published {color:lightgray; position: relative;}
    .notes {top:0; position: absolute; font-size:16px; text-align: center;display: flex;
    justify-content: center;
    align-items: center;
    height: 400px;
    width: 210px;
    }
    .grade {position: absolute;bottom: 50; left: 0; font-size:20px;align-content: center; width: 250px;}
    .value {position: absolute;bottom: 0; left: 0; font-size:x-large; align-content: center; width: 250px;}
    .hvrbox,
    .hvrbox * {box-sizing: border-box; padding: 0px;}
    .hvrbox {position: relative;display: inline-block;overflow: hidden;width: 250px;height: 400px;}
    .hvrbox img {width: 250px;height: 400px;}
    .hvrbox .hvrbox-layer_bottom {display: block;}
    .hvrbox .hvrbox-layer_top {
        opacity: 0;
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0, 0, 0, 0.75);
        padding: 0px;
        -moz-transition: all 0.4s ease-in-out 0s;
        -webkit-transition: all 0.4s ease-in-out 0s;
        -ms-transition: all 0.4s ease-in-out 0s;
        transition: all 0.4s ease-in-out 0s;
    }
    .hvrbox:hover .hvrbox-layer_top,
    .hvrbox.active .hvrbox-layer_top {opacity: 1;}
    .hvrbox .hvrbox-text {
        text-align: center;
        font-size: 17px;
        display: inline-block;
        padding-left: 20px;
        padding-right: 20px;
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
    }
    .hvrbox .hvrbox-text_mobile {
    font-size: 15px;
    border-top: 1px solid rgb(179, 179, 179); /* for old browsers */
    border-top: 1px solid rgba(179, 179, 179, 0.7);
    margin-top: 5px;
    padding-top: 2px;
    display: none;
    }
    .hvrbox.active .hvrbox-text_mobile {display: block;}
    </style>
    </head>
    <body>
        """)
        f.write(htmlBody)
        f.write("</body></html>")
    

def ReadGoogleSheet(Google_Workbook, Google_Sheet):
    # =============================================================================
    #   Read Google sheet into pandas Dataframe - Requires Service Account in Google API
    #   file stored in ~/.config/gspread/service_account.json 
    # =============================================================================
    gc = gspread.service_account()
    sh = gc.open(Google_Workbook)
    worksheet = sh.worksheet(Google_Sheet)
    Starting_DF = pd.DataFrame(worksheet.get_all_records())
    sortedsheet = Starting_DF.sort_values(by=['Title','Volume','Issue'])
    return Starting_DF, sortedsheet, worksheet
    
def BackupGoogleSheet(Sheetname):
    # =============================================================================
    #  Make a backup of the current sheet in the event it all goes to shit
    # =============================================================================
    starting_rows = Starting_DF.shape[0] 
    starting_cols = Starting_DF.shape[1] 
    backup = sh.add_worksheet(title="Backup " + rundate, rows=starting_rows, cols=starting_cols)
    backup.update([sortedsheet.columns.values.tolist()] + sortedsheet.values.tolist())


SheetData = ReadGoogleSheet(Google_Workbook, Google_Sheet)
StartingDF = SheetData[0]
sortedsheet = SheetData[1]
worksheet = SheetData[2]
generate_HTMLPage(sortedsheet)