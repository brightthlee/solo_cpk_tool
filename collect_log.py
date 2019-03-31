#!/usr/bin/python
import os
import sys
import json
import re
import csv
from time import asctime, gmtime
import time
import datetime
import xlsxwriter
import zipfile
from StringIO import StringIO
from xlsxwriter.utility import xl_range,xl_col_to_name

# argv[1]: file
if len(sys.argv) == 3:
    path = sys.argv[1]
    output = sys.argv[2]
    print ('Writing {}.xlsx'.format(output))
else:
    print ("Usage: " + sys.argv[0] + " file_list" + " output_file_name")
    exit()

if sys.argv[1] == "--":
    files = [f.strip() for f in sys.stdin]
else:
    file_list = open (sys.argv[1], "r") 
    files = [f.strip() for f in file_list]
    file_list.close()

workbook = xlsxwriter.Workbook('{}.xlsx'.format(output), {'strings_to_numbers': True})

format_green = workbook.add_format({'bg_color': '#C6EFCE'})
format_yellow = workbook.add_format({'bg_color': '#FFFF00'})
format_pink = workbook.add_format({'bg_color': '#FF8AD8'})
format_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'})

worksheet_all = workbook.add_worksheet('all_data')

worksheet_all.write_string(0,0,'Filename')
worksheet_all.write_string(0,1,'DUT_ID')
worksheet_all.write_string(0,2,'START_DATE_TIME')
worksheet_all.write_string(0,3,'STATUS')
worksheet_all.write_string(0,4,'FAILURE_CODE')

worksheet_all.write_string(1,0,'LSL')
worksheet_all.write_string(2,0,'USL')
worksheet_all.write_string(3,0,'MIN')
worksheet_all.write_string(4,0,'MAX')
worksheet_all.write_string(5,0,'CPK')
worksheet_all.write_string(6,0,'CPK-')
worksheet_all.write_string(7,0,'CPK+')
worksheet_all.write_string(8,0,'STDEV')
worksheet_all.write_string(9,0,'AVERAGE')
worksheet_all.write_string(10,0,'Useful COUNT')

all_row_num = 11
column_num = 0
cpk_columns = []

if '.zip' in files[0]:
    for z in files:
        archive = zipfile.ZipFile(z,'r')
        metafile = StringIO(archive.read('metadata.csv'))
        metacsv = csv.reader(metafile,delimiter=',')

        row = metacsv.next() # bypass title
        row = metacsv.next()
        if row[13] != 'PASS': # if not pass, measurements will be lost some items
            archive.close()
            continue
        measurefile = StringIO(archive.read('measurements.csv'))
        print "Using " + z + "/measurements.csv to list items"
        measurecsv = csv.reader(measurefile,delimiter=',')

        next(measurecsv) # bypass title
        column_num = 5
        for row_i in measurecsv:
            worksheet_all.write_string(0,column_num,row_i[0]) # Test item name
            if row_i[7] not in (None,"", ' ', '--'): 
                worksheet_all.write_number(1,column_num,float(row_i[7])) # Number Min
            if row_i[6] not in (None,"", ' ', '--'): 
                worksheet_all.write_number(2,column_num,float(row_i[6])) # Number Max
            column_num += 1
        archive.close()
        break

    for z in files:
        archive = zipfile.ZipFile(z,'r')
        try:
            metafile = StringIO(archive.read('metadata.csv'))
        except:
            print ("Missing " + f + '/metadata.csv')
            archive.close()
            continue
        try:
            measurefile = StringIO(archive.read('measurements.csv'))
        except:
            print ("Missing " + f + '/metadata.csv')
            archive.close()
            continue

        metacsv = csv.reader(metafile,delimiter=',')
        row = []

        next(metacsv) # bypass title
        row = next(metacsv) 
        worksheet_all.write_string(all_row_num, 0, f) # filename / folder
        worksheet_all.write_string(all_row_num, 1, row[0]) # DUT_ID
        worksheet_all.write_string(all_row_num, 2, row[16]) # START_DATE_TIME
        worksheet_all.write_string(all_row_num, 3, row[13]) # STATUS
        worksheet_all.write_string(all_row_num, 4, row[14]) # FAILURE_CODE

        measurecsv = csv.reader(measurefile,delimiter=',')
        next(measurecsv) # bypass title
        column_num = 5
        for row_i in measurecsv:
            if row_i[5] in (None, "", ' '):
                worksheet_all.write_string(all_row_num, column_num, row_i[9])
            else:
                try:
                    worksheet_all.write_number(all_row_num, column_num, float(row_i[5]))
                    cpk_columns.append(column_num)
                except:
                    worksheet_all.write_string(all_row_num, column_num, row_i[5])
            column_num += 1

        all_row_num += 1
        archive.close()

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
                    if 'numeric_min' in mea:
                        worksheet_all.write_number(1,column_num,mea['numeric_min']) # Number Min
                    if 'numeric_max' in mea:
                        worksheet_all.write_number(2,column_num,mea['numeric_max']) # Number Max
                    column_num += 1
            break

    for f in files:
        with open(f,'r') as json_file:
            data = json.load(json_file)

            worksheet_all.write_string(all_row_num, 0, f) # filename
            worksheet_all.write_string(all_row_num, 1, data['dut_id']) # DUT_ID
            worksheet_all.write_string(all_row_num, 2, asctime(gmtime(data['start_time_ms']/1000))) # START_DATE_TIME
            worksheet_all.write_string(all_row_num, 3, data['status']) # STATUS
            # worksheet_all.write_string(all_row_num, 4, ) # FAILURE_CODE

            column_num = 5
            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    if 'numeric_value' in mea and mea['text_value'] != "inf":
                        worksheet_all.write_number(all_row_num, column_num, mea['numeric_value'])
                        cpk_columns.append(column_num)
                    else:
                        worksheet_all.write_string(all_row_num, column_num, mea['text_value'])
                    column_num += 1
            all_row_num += 1

for i in cpk_columns:
    column_letter = xl_col_to_name(i)
    worksheet_all.write_array_formula(3,i,3,i,'=MIN(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(4,i,4,i,'=MAX(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(5,i,5,i,'=MIN({0}7:{0}8)'.format(column_letter))
    worksheet_all.write_array_formula(6,i,6,i,'=ABS({0}10-{0}2)/(3*{0}9)'.format(column_letter))
    worksheet_all.write_array_formula(7,i,7,i,'=ABS({0}3-{0}10)/(3*{0}9)'.format(column_letter))
    worksheet_all.write_array_formula(8,i,8,i,'=STDEV(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(9,i,9,i,'=AVERAGE(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(10,i,10,i,'=COUNT(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))

    worksheet_all.conditional_format(5,i,7,i,{'type': 'cell',
                                                       'criteria': '>=',
                                                       'value': 2,
                                                       'format': format_green})
    worksheet_all.conditional_format(5,i,7,i,{'type': 'cell',
                                                       'criteria': 'between',
                                                       'minimum': 1,
                                                       'maximum': 1.33,
                                                       'format': format_yellow})
    worksheet_all.conditional_format(5,i,7,i,{'type': 'cell',
                                                       'criteria': 'between',
                                                       'minimum': 0.67,
                                                       'maximum': 1,
                                                       'format': format_pink})
    worksheet_all.conditional_format(5,i,7,i,{'type': 'cell',
                                                       'criteria': '<',
                                                       'value': 0.67,
                                                       'format': format_red})
workbook.close()
