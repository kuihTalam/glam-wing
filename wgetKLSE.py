# -*- coding: utf-8 -*-
"""
Created on Wed Dec 21 21:27:21 2016

@author: songheng

"""
import time
import sys

baconFile = open('C:\\Users\\songheng\\Desktop\\Zach\\automate\\FFL.txt', 'r')
stockCode = open('C:\\Users\\songheng\\Desktop\\Zach\\automate\\ProgressCode.txt', 'w')
baconContent = baconFile.read()
baconList = baconContent.splitlines()
length = float(len(baconList))
r=0

# Loop over list in text file and print out each 'wget' command line 
# to ProgressCode.txt
for i in baconList:
    stockCode.write('wget -x --load-cookies cookies.txt --output-document=%s.html --wait=10 "http://www.shareinvestor.com/fundamental/financials.html?counter=%s.MY&period=fy&cols=10"\n' % (i, i))
    
    # Progress bar 
    time.sleep(0.1)
    progress = (int(r)+1)/length
    prog_bar = '#' * int(progress * 50)
    prog_pct = int(progress * 100)
    sys.stdout.write("\rInstalling : [%s] %d%%" %(prog_bar, prog_pct))
    sys.stdout.flush()
    r+=1
stockCode.close()
