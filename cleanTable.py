# -*- coding: utf-8 -*-
"""
Created on Mon Dec 26 22:19:03 2016

@author: songheng
"""

#! python3
# updateProduce.py - Corrects costs in produce sales spreadsheet.

import openpyxl
import sys

baconFile = open('C:\\Users\\songheng\\Desktop\\Zach\\automate\\FFL.txt', 'r')
baconContent = baconFile.read()
baconList = baconContent.splitlines()
length = float(len(baconList))
r=0

# Loop over the xlsx files and clean up
for m in baconList:
    wb = openpyxl.load_workbook(r'C:\Users\songheng\Desktop\KLSE\Excel\%s.xlsx' % m)
    sheet = wb.get_sheet_by_name('PL')

    sheet['B1'] = ''
    sheet['C1'] = ''
    sheet['D1'] = 'Trailing 2016'
    sheet['E1'] = 'FY2016'
    sheet['F1'] = 'FY2015'
    sheet['G1'] = 'FY2014'
    sheet['H1'] = 'FY2013'
    sheet['I1'] = 'FY2012'
    sheet['J1'] = 'FY2011'
    sheet['K1'] = 'FY2010'
    sheet['L1'] = 'FY2009'
    sheet['M1'] = 'FY2008'
    for cellObj in sheet.columns[2]:
        sheet['%s' % cellObj.coordinate] = ''
        
    
    sheet = wb.get_sheet_by_name('BS')
    
    sheet['B1'] = ''
    sheet['B2'] = ''
    sheet['C1'] = ''
    sheet['D1'] = 'Trailing 2016'
    sheet['E1'] = 'FY2016'
    sheet['F1'] = 'FY2015'
    sheet['G1'] = 'FY2014'
    sheet['H1'] = 'FY2013'
    sheet['I1'] = 'FY2012'
    sheet['J1'] = 'FY2011'
    sheet['K1'] = 'FY2010'
    sheet['L1'] = 'FY2009'
    sheet['M1'] = 'FY2008'
    for cellObj in sheet.columns[2]:
        sheet['%s' % cellObj.coordinate] = ''
    
    sheet = wb.get_sheet_by_name('CF')
    
    sheet['B1'] = ''
    sheet['C1'] = ''
    sheet['D1'] = 'Trailing 2016'
    sheet['E1'] = 'FY2016'
    sheet['F1'] = 'FY2015'
    sheet['G1'] = 'FY2014'
    sheet['H1'] = 'FY2013'
    sheet['I1'] = 'FY2012'
    sheet['J1'] = 'FY2011'
    sheet['K1'] = 'FY2010'
    sheet['L1'] = 'FY2009'
    sheet['M1'] = 'FY2008'
    for cellObj in sheet.columns[2]:
        sheet['%s' % cellObj.coordinate] = ''
    
    sheet = wb.get_sheet_by_name('Ratios')
    
    sheet['B1'] = ''
    sheet['C1'] = ''
    sheet['D1'] = 'Trailing 2016'
    sheet['E1'] = 'FY2016'
    sheet['F1'] = 'FY2015'
    sheet['G1'] = 'FY2014'
    sheet['H1'] = 'FY2013'
    sheet['I1'] = 'FY2012'
    sheet['J1'] = 'FY2011'
    sheet['K1'] = 'FY2010'
    sheet['L1'] = 'FY2009'
    sheet['M1'] = 'FY2008'
    for cellObj in sheet.columns[2]:
        sheet['%s' % cellObj.coordinate] = ''
    
    wb.save(r'C:\Users\songheng\Desktop\KLSE\Excel\%s.xlsx' % m)
    progress = (int(r)+1)/length
    prog_bar = '#' * int(progress * 50)
    prog_pct = int(progress * 100)
    sys.stdout.write("\rGenerating : [%s] %d%%" %(prog_bar, prog_pct))
    sys.stdout.flush()
    r+=1