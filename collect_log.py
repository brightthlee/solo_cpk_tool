#!/usr/bin/python
import os
from os import listdir
import sys
import json
import re
import csv
from time import asctime, gmtime
import time
import datetime
import xlsxwriter

# argv[1]: file
if len(sys.argv) == 2:
    path = sys.argv[1]
else:
    print ("Usage: " + sys.argv[0] + " file_list")
    exit()

if sys.argv[1] == "--":
    files = [f.strip() for f in sys.stdin]
else:
    file_list = open (sys.argv[1], "r") 
    files = [f.strip() for f in file_list]
    file_list.close()

workbook = xlsxwriter.Workbook('test.xlsx', {'strings_to_numbers': True})

format_green = workbook.add_format({'bg_color': '#C6EFCE'})
format_yellow = workbook.add_format({'bg_color': '#FFFF00'})
format_pink = workbook.add_format({'bg_color': '#FF8AD8'})
format_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'})

worksheet_all = workbook.add_worksheet('all_data')
worksheet_cpk = workbook.add_worksheet('cpk')

worksheet_all.write_string(0,0,'Filename')
worksheet_all.write_string(0,1,'DUT_ID')
worksheet_all.write_string(0,2,'START_DATE_TIME')
worksheet_all.write_string(0,3,'STATUS')
worksheet_all.write_string(0,4,'FAILURE_CODE')

worksheet_cpk.write_string(0,0,'Filename')
worksheet_cpk.write_string(0,1,'DUT_ID')
worksheet_cpk.write_string(0,2,'START_DATE_TIME')
worksheet_cpk.write_string(0,3,'STATUS')
worksheet_cpk.write_string(0,4,'FAILURE_CODE')
worksheet_cpk.write_string(1,0,'LSL')
worksheet_cpk.write_string(2,0,'USL')

worksheet_cpk.write_string(3,0,'MIN')
worksheet_cpk.write_string(4,0,'MAX')
worksheet_cpk.write_string(5,0,'CPK')
worksheet_cpk.write_string(6,0,'CPK-')
worksheet_cpk.write_string(7,0,'CPK+')
worksheet_cpk.write_string(8,0,'STDEV')
worksheet_cpk.write_string(9,0,'AVERAGE')
worksheet_cpk.write_string(10,0,'Useful COUNT')

worksheet_cpk.write_formula(3,5,'=MIN(F12:F2000)')
worksheet_cpk.write_formula(4,5,'=MAX(F12:F2000)')
worksheet_cpk.write_formula(5,5,'=MIN(F7:F8)')
worksheet_cpk.write_formula(6,5,'=ABS(F10-F2)/(3*F9)')
worksheet_cpk.write_formula(7,5,'=ABS(F3-F10)/(3*F9)')
worksheet_cpk.write_formula(8,5,'=STDEV(F12:F2000)')
worksheet_cpk.write_formula(9,5,'=AVERAGE(F12:F2000)')
worksheet_cpk.write_formula(10,5,'=COUNT(F12:F2000)')

worksheet_cpk.conditional_format('F6:F8',{'type': 'cell',
                                     'criteria': '>=',
                                     'value': 2,
                                     'format': format_green})
worksheet_cpk.conditional_format('F6:F8',{'type': 'cell',
                                     'criteria': 'between',
                                     'minimum': 1,
                                     'maximum': 1.33,
                                     'format': format_yellow})
worksheet_cpk.conditional_format('F6:F8',{'type': 'cell',
                                     'criteria': 'between',
                                     'minimum': 0.67,
                                     'maximum': 1,
                                     'format': format_pink})
worksheet_cpk.conditional_format('F6:F8',{'type': 'cell',
                                     'criteria': '<',
                                     'value': 0.67,
                                     'format': format_red})

if os.path.isdir(files[0]):
    for d in files:
        with open(d+'/metadata.csv','r') as metafile:
            metacsv = csv.reader(metafile)

            next(metacsv) # bypass title
            row = next(metacsv) 
            if row[13] != 'PASS': # if not pass, measurements will be lost some items
                continue
        with open(d+'/measurements.csv', 'r') as measurefile:
            print "Using " + d + "/measurements.csv to list items"
            measurecsv = csv.reader(measurefile)

            next(measurecsv) # bypass title
            column_num = 5
            for row_i in measurecsv:
                worksheet_all.write_string(0,column_num,row_i[0]) # Test item name
                worksheet_cpk.write_string(0,column_num,row_i[0]) # Test item name
                cpk_enable = 0
                if row_i[7] not in (None,"", ' ', '--'): 
                    worksheet_cpk.write_number(1,column_num,float(row_i[7])) # Number Min
                if row_i[7] not in (None,"", ' ', '--'): 
                    worksheet_cpk.write_number(2,column_num,float(row_i[6])) # Number Max
                column_num += 1
            break

    all_row_num = 1
    cpk_row_num = 11
    for f in files:
        if not os.path.isfile(f+'/metadata.csv'):
            print ("Missing " + f + '/metadata.csv')
            continue
        if not os.path.isfile(f+'/measurements.csv'):
            print ("Missing " + f + '/measurements.csv')
            continue
        with open(f+'/metadata.csv','r') as metafile:
                metacsv = csv.reader(metafile)
                row = []

                next(metacsv) # bypass title
                row = next(metacsv) 
                worksheet_all.write_string(all_row_num, 0, f) # filename / folder
                worksheet_all.write_string(all_row_num, 1, row[0]) # DUT_ID
                worksheet_all.write_string(all_row_num, 2, row[16]) # START_DATE_TIME
                worksheet_all.write_string(all_row_num, 3, row[13]) # STATUS
                worksheet_all.write_string(all_row_num, 4, row[14]) # FAILURE_CODE

                worksheet_cpk.write_string(cpk_row_num, 0, f) # filename / folder
                worksheet_cpk.write_string(cpk_row_num, 1, row[0]) # DUT_ID
                worksheet_cpk.write_string(cpk_row_num, 2, row[16]) # START_DATE_TIME
                worksheet_cpk.write_string(cpk_row_num, 3, row[13]) # STATUS
                worksheet_cpk.write_string(cpk_row_num, 4, row[14]) # FAILURE_CODE

                with open(f+'/measurements.csv', 'r') as measurefile:
                    measurecsv = csv.reader(measurefile)
                    next(measurecsv) # bypass title
                    column_num = 5
                    for row_i in measurecsv:
                        if row_i[5] in (None, "", ' '):
                            worksheet_all.write_string(all_row_num, column_num, row_i[9])
                        else:
                            try:
                                worksheet_all.write_number(all_row_num, column_num, float(row_i[5]))
                            except:
                                worksheet_all.write_string(all_row_num, column_num, row_i[5])
                        if row[13] == 'PASS':
                            if row_i[5] in (None, "", ' '):
                                worksheet_cpk.write_string(cpk_row_num, column_num, row_i[9])
                            else:
                                try:
                                    worksheet_cpk.write_number(cpk_row_num, column_num, float(row_i[5]))
                                except:
                                    worksheet_cpk.write_string(cpk_row_num, column_num, row_i[5])
                        column_num += 1
                    if row[13] == 'PASS':
                        cpk_row_num += 1
                all_row_num += 1

elif '.scj' in files[0]:
    for f in files:
        with open(f,'r') as json_file:
            data = json.load(json_file)
            if data['status'] != 'PASS':
                continue
            print "Using " + files[0] + " to list items"
            column_num = 5
            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    worksheet_all.write_string(0,column_num, mea['name'])
                    worksheet_cpk.write_string(0,column_num, mea['name'])
                    if 'numeric_min' in mea:
                        worksheet_cpk.write_number(1,column_num,mea['numeric_min']) # Number Min
                    if 'numeric_max' in mea:
                        worksheet_cpk.write_number(2,column_num,mea['numeric_max']) # Number Max
                    column_num += 1
            break

    all_row_num = 1
    cpk_row_num = 11
    for f in files:
        #try:
            with open(f,'r') as json_file:
                data = json.load(json_file)

                worksheet_all.write_string(all_row_num, 0, f) # filename
                worksheet_all.write_string(all_row_num, 1, data['dut_id']) # DUT_ID
                worksheet_all.write_string(all_row_num, 2, asctime(gmtime(data['start_time_ms']/1000))) # START_DATE_TIME
                worksheet_all.write_string(all_row_num, 3, data['status']) # STATUS
                # worksheet_all.write_string(all_row_num, 4, ) # FAILURE_CODE

                worksheet_cpk.write_string(cpk_row_num, 0, f) # filename
                worksheet_cpk.write_string(cpk_row_num, 1, data['dut_id']) # DUT_ID
                worksheet_cpk.write_string(cpk_row_num, 2, asctime(gmtime(data['start_time_ms']/1000))) # START_DATE_TIME
                worksheet_cpk.write_string(cpk_row_num, 3, data['status']) # STATUS
                # worksheet_cpk.write_string(all_row_num, 4, ) # FAILURE_CODE

                column_num = 5
                for i in range (0, len(data['phases'])):
                    for mea in data['phases'][i]['measurements']:
                        if 'numeric_value' in mea and mea['text_value'] != "inf":
                            worksheet_all.write_number(all_row_num, column_num, mea['numeric_value'])
                        else:
                            worksheet_all.write_string(all_row_num, column_num, mea['text_value'])
                        if data['status'] == 'PASS':
                            if 'numeric_value' in mea and mea['text_value'] != "inf":
                                worksheet_cpk.write_number(cpk_row_num, column_num, mea['numeric_value'])
                            else:
                                worksheet_cpk.write_string(cpk_row_num, column_num, mea['text_value'])
                        column_num += 1
                if data['status'] == 'PASS':
                    cpk_row_num += 1
                all_row_num += 1
        #except:
        #    print ("Bypass file: %s" % f)

workbook.close()
