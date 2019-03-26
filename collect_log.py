#!/usr/bin/python
import os
from os import listdir
import sys
import json
import re
import csv
import time
import datetime

# argv[1]: file
if len(sys.argv) == 2:
    path = sys.argv[1]
else:
    print (sys.argv[0] + " file_list")

if sys.argv[1] == "--":
    files = [f.strip() for f in sys.stdin]
else:
    file_list = open (sys.argv[1], "r") 
    files = [f.strip() for f in file_list]
    file_list.close()

with open('output.csv', 'w') as outputfile:
    with open(files[0]+'/measurements.csv', 'r') as measurefile:
        measurecsv = csv.reader(measurefile)
        writer = csv.writer(outputfile)

        row_data = []
        row_title = []
        row_lsl = ['LSL','','','','']
        row_usl = ['USL','','','','']
        all = []
        row_title.append('Filename')
        row_title.append('DUT_ID')
        row_title.append('START_DATE_TIME')
        row_title.append('STATUS')
        row_title.append('FAILURE_CODE')

        next(measurecsv) # bypass title
        for row_i in measurecsv:
            row_title.append(row_i[0])  # TEST ITEM NAME
            row_usl.append(row_i[6]) # NUMBER_MAX
            row_lsl.append(row_i[7]) # NUMBER_MIN
            #if row_i[5] in (None, "", ' '): # Test value
            #    row_data.append(row_i[9])  
            #else:
            #    row_data.append(row_i[5])

        all.append(row_title)
        all.append(row_usl)
        all.append(row_lsl)
        #all.append(row_data)
        writer.writerows(all)

    for f in files:
        with open(f+'/metadata.csv','r') as metafile:
            metacsv = csv.reader(metafile)
            writer = csv.writer(outputfile)
            row = []
            all = []

            row.append(f)
            next(metacsv) # bypass title
            row_i = next(metacsv) 
            row.append(row_i[0]) # DUT_ID
            row.append(row_i[16]) # START_DATE_TIME
            row.append(row_i[13]) # STATUS
            row.append(row_i[14]) # FAILURE_CODE

            with open(f+'/measurements.csv', 'r') as measurefile:
                measurecsv = csv.reader(measurefile)
                next(measurecsv) # bypass title
                for row_i in measurecsv:
                    if row_i[5] in (None, "", ' '):
                        row.append(row_i[9])  
                    else:
                        row.append(row_i[5])

            all.append(row)
            writer.writerows(all)
