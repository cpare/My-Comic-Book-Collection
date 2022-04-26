#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 15:20:40 2020

@author: cpare
"""


import gspread
import pandas as pd
    

gc = gspread.oauth()

sh = gc.open("Comics")

#worksheet = sh.sheet1
worksheet = sh.worksheet("My Collection")

dataframe = pd.DataFrame(worksheet.get_all_records())

worksheet.update([dataframe.columns.values.tolist()] + dataframe.values.tolist())