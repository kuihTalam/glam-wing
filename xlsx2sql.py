# -*- coding: utf-8 -*-
"""
Created on Fri Dec 30 09:06:07 2016

@author: TSO1PG1
"""

from openpyxl import load_workbook
import sqlite3
conn = sqlite3.connect('qtrqtr.db')
c = conn.cursor()
 
wb = load_workbook(r'D:\Python\foo.xlsx', read_only=True)
ws = wb['Sheet2']
data = []

for row in ws.rows:
    for cell in row:
        data.append(cell.value)

tupper = []
for i in range (12):
    tupper.append(data[i]) 

istLine = "INSERT OR IGNORE INTO code VALUES(" + str(tupper[0].date()) + ", " \
            + str(tupper[1].date()) + ", " + str(tupper[2]) + ", " \
            + str(tupper[3].date()) + ", " + str(tupper[4]) + ", " + str(tupper[5]) \
            + ", " + str(tupper[6]) + ", " + str(tupper[7]) + "," \
            + str(tupper[8]) + ", " + str(tupper[9]) + ", " + str(tupper[10]) \
            + ", " + str(tupper[11]) +")"

#SC = []
#SN = []
#for num in range(len(data)):
#    if num%2 == 0:
#        SC.append(data[num])
    
#    else:
#        SN.append(data[num])
        
#for i, j in zip(SC, SN):
#    if i == '5235SS':
#        pass
#    else:
#        istLine = "INSERT OR IGNORE INTO code VALUES(" + str(i) +", " + str(j) + ")"
#        c.execute(istLine)

conn.commit()
conn.close()

