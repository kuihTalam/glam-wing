# -*- coding: utf-8 -*-
"""
Created on Fri Dec 30 09:06:07 2016

@author: TSO1PG1
"""

from openpyxl import load_workbook
import sqlite3
conn = sqlite3.connect('qtrqtr.db')
c = conn.cursor()
 
wb = load_workbook(r'D:\Python\stockCodeTable.xlsx', read_only=True)
ws = wb['Sheet1']
data = []

for row in ws.rows:
    for cell in row:
        data.append(cell.value)
        
SC = []
SN = []
for num in range(len(data)):
    if num%2 == 0:
        SC.append(data[num])
    
    else:
        SN.append(data[num])
        
for i, j in zip(SC, SN):
    if i == '5235SS':
        pass
    else:
        istLine = "INSERT OR IGNORE INTO code VALUES(" + str(i) +", " + str(j) + ")"
        c.execute(istLine)

conn.commit()
conn.close()
