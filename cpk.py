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
import operator
import enum
from StringIO import StringIO
from xlsxwriter.utility import xl_range,xl_col_to_name

if len(sys.argv) == 3:
    path = sys.argv[1]
    output = sys.argv[2]
    print ('Writing {0}.xlsx and {0}.csv'.format(output))
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
cpk_csv = open('{}.csv'.format(output),'wb')
csvwriter = csv.writer(cpk_csv)

format_green = workbook.add_format({'bg_color': '#C6EFCE'})
format_yellow = workbook.add_format({'bg_color': '#FFFF00'})
format_pink = workbook.add_format({'bg_color': '#FF8AD8'})
format_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'})

worksheet_all = workbook.add_worksheet('all_data')

ITEMs = ['Filename', 'DUT_ID', 'START_DATE_TIME', 'STATUS', 'FAILURE_CODE']
LSLs = ['LSL','','','','']
USLs = ['USL','','','','']
all_data_rows = []

column_num = 0
cpk_columns = []

class scz_columns(enum.IntEnum):
    TEST_ITEMS = 0
    DUT_ID = 0
    NUMERIC_VALUE = 5
    NUMERIC_MAX= 6
    NUMERIC_MIN= 7
    TEXT_VALUE = 9
    STATUS = 13
    FAILURE_CODE = 14
    START_DATE_TIME = 16

def set_conditional_format(worksheet, row, column):
    worksheet.conditional_format(row,column,row,column,{'type': 'cell',
                                             'criteria': '>=',
                                             'value': 2,
                                             'format': format_green})
    worksheet.conditional_format(row,column,row,column,{'type': 'cell',
                                             'criteria': 'between',
                                             'minimum': 1,
                                             'maximum': 1.33,
                                             'format': format_yellow})
    worksheet.conditional_format(row,column,row,column,{'type': 'cell',
                                             'criteria': 'between',
                                             'minimum': 0.67,
                                             'maximum': 1,
                                             'format': format_pink})
    worksheet.conditional_format(row,column,row,column,{'type': 'cell',
                                             'criteria': '<',
                                             'value':0.67,
                                             'format': format_red})

globals().update(scz_columns.__members__)

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
        for row_i in measurecsv:
            ITEMs.append (row_i[TEST_ITEMS])
            if row_i[NUMERIC_MIN] not in (None, "", ' ', '--'): 
                LSLs.append(row_i[NUMERIC_MIN])
            else:
                LSLs.append('')
            if row_i[NUMERIC_MAX] not in (None, "", ' ', '--'): 
                USLs.append(row_i[NUMERIC_MAX])
            else:
                USLs.append('')

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
        meta_row = []

        next(metacsv) # bypass title
        meta_row = next(metacsv) 
        mea_row = [f, meta_row[DUT_ID], meta_row[START_DATE_TIME], meta_row[STATUS], meta_row[FAILURE_CODE]]

        measurecsv = csv.reader(measurefile,delimiter=',')
        next(measurecsv) # bypass title
        for row_i in measurecsv:
            if row_i[NUMERIC_VALUE] in (None, "", ' '):
                mea_row.append(row_i[TEXT_VALUE])
            else:
                try:
                    mea_row.append(float(row_i[NUMERIC_VALUE]))
                except:
                    mea_row.append(row_i[NUMERIC_VALUE])

        all_data_rows.append(mea_row)
        archive.close()

elif '.scj' in files[0]:
    for f in files:
        with open(f,'r') as json_file:
            data = json.load(json_file)
            if data['status'] != 'PASS':
                continue
            print "Using " + files[0] + " to list items"
            mea_row = []
            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    ITEMs.append(mea['name'])
                    if 'numeric_min' in mea:
                        LSLs.append(mea['numeric_min'])
                    else:
                        LSLs.append('')
                    if 'numeric_max' in mea:
                        USLs.append(mea['numeric_max'])
                    else:
                        USLs.append('')
            break

    all_data_rows = []

    for f in files:
        with open(f,'r') as json_file:
            data = json.load(json_file)

            start_date_time = asctime(gmtime(data['start_time_ms']/1000)) # START_DATE_TIME
            mea_row = [f, data['dut_id'], start_date_time, data['status'], '']

            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    if 'numeric_value' in mea and mea['text_value'] != "inf":
                        mea_row.append(mea['numeric_value'])
                    else:
                        mea_row.append(mea['text_value'])

            all_data_rows.append(mea_row)

mea_data_sorted = sorted(all_data_rows, key = operator.itemgetter(2), reverse=False) # sort by time

worksheet_all.write_row(0,0,ITEMs)
worksheet_all.write_row(1,0,LSLs)
worksheet_all.write_row(2,0,USLs)
worksheet_all.write_column(3,0,['MIN','MAX','CPK','CPK-','CPK+','STDEV','AVERAGE','Useful Count'])

column_num = 11
for data in mea_data_sorted:
    # if data[3] == 'PASS':
        worksheet_all.write_row(column_num, 0, data)
        column_num += 1

worksheet_all.freeze_panes(11,1)

numeric_columns = []
for i in range(0, len(mea_data_sorted)):
    # find the pass one
    row = mea_data_sorted[i]
    if row[3] == 'PASS':
        for j in range(5, len(row)):
            if type(row[j]) in (int, float):
                numeric_columns.append(j)
        break

all_row_num = 11 + len(mea_data_sorted)
for i in numeric_columns:
    column_letter = xl_col_to_name(i)
    worksheet_all.write_array_formula(3,i,3,i,'=MIN(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(4,i,4,i,'=MAX(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    if LSLs[i] != '':
        worksheet_all.write_array_formula(6,i,6,i,'=ABS({0}10-{0}2)/(3*{0}9)'.format(column_letter)) # CPK-
        set_conditional_format(worksheet_all, 6, i)
    if USLs[i] != '':
        worksheet_all.write_array_formula(7,i,7,i,'=ABS({0}3-{0}10)/(3*{0}9)'.format(column_letter)) # CPK+
        set_conditional_format(worksheet_all, 7, i)
    if LSLs[i] != '' or USLs[i] != '':
        worksheet_all.write_array_formula(5,i,5,i,'=MIN({0}7:{0}8)'.format(column_letter)) # CPK
        set_conditional_format(worksheet_all, 5, i)
    worksheet_all.write_array_formula(8,i,8,i,'=STDEV(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(9,i,9,i,'=AVERAGE(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))
    worksheet_all.write_array_formula(10,i,10,i,'=COUNT(IF($D$12:$D${0}="PASS",{1},""))'.format(all_row_num,xl_range(11,i,all_row_num-1,i)))

csv_items = ['DUT_ID']
csv_lsl = ['LSL']
csv_usl = ['USL']
for i in range(1,len(ITEMs)):
    # if i in numeric_columns:
    if LSLs[i] != '' or USLs[i] != '':
        csv_items.append(ITEMs[i])
        csv_lsl.append(LSLs[i])
        csv_usl.append(USLs[i])

csvwriter.writerows([csv_items,csv_lsl,csv_usl])

for data in mea_data_sorted:
    if data[3] == 'PASS':
        row = []
        row.append(data[1])
        # for i in numeric_columns:
        for i in range(1, len(LSLs)):
            if LSLs[i] != '' or USLs[i] != '':
                row.append(data[i])
        csvwriter.writerow(row)

workbook.close()
cpk_csv.close()
