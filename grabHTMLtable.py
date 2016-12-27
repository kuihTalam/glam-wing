# -*- coding: utf-8 -*-
"""
Created on Mon Dec 19 21:48:11 2016
 for row in message2.find_all('tr')[1:]:
        col = row.find_all('td')[1:]
@author: songheng
"""

#! python3
# lucky.py - Opens several Google search results.

import pandas as pd
import html5lib
import time
import sys
from sqlalchemy import create_engine

baconFile = open('C:\\Users\\songheng\\Desktop\\Zach\\automate\\FFL.txt', 'r')
baconContent = baconFile.read()
baconList = baconContent.splitlines()
length = float(len(baconList))
r=0

# Loop to grab html table data by using pandas
for m in baconList:
    for i, df1 in enumerate(pd.read_html(r'C:\Users\songheng\Desktop\KLSE\%s.html' % m, attrs = {'id': 'financials_table_pl'})):
        spam1 = df1

    for j, df2 in enumerate(pd.read_html(r'C:\Users\songheng\Desktop\KLSE\%s.html' % m, attrs = {'id': 'financials_table_bs'})):
        spam2 = df2
    
    for k, df3 in enumerate(pd.read_html(r'C:\Users\songheng\Desktop\KLSE\%s.html' % m, attrs = {'id': 'financials_table_cf'})):
        spam3 = df3

    for l, df4 in enumerate(pd.read_html(r'C:\Users\songheng\Desktop\KLSE\%s.html' % m, attrs = {'id': 'financials_table_ratio'})):
        spam4 = df4
    
    # Create xlsx file based on stockCode
    writer = pd.ExcelWriter(r'C:/Users/songheng/Desktop/KLSE/Excel/%s.xlsx' % m) 
    spam1.to_excel(writer, 'PL')   # Write to PL worksheet
    spam2.to_excel(writer, 'BS')   # Write to BS worksheet
    spam3.to_excel(writer, 'CF')   # Write to CF worksheet
    spam4.to_excel(writer, 'Ratios') # Write to Ratios worksheet
    
    # Code for progress bar
    progress = (int(r)+1)/length
    prog_bar = '#' * int(progress * 50)
    prog_pct = int(progress * 100)
    sys.stdout.write("\rGenerating : [%s] %d%%" %(prog_bar, prog_pct))
    sys.stdout.flush()
    r+=1
    print ("Generate %s.xlsx Done. No. %s !!!" % (m, r) + "\n")
        







