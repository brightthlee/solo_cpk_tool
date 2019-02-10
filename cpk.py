#!/usr/bin/python
from os import listdir
import sys
import json
import re
import csv
import time
import datetime

# argv[1]: folder
# argv[2]: start date/time

if len(sys.argv) >= 2:
    path = sys.argv[1]
    #list = (map(int, (sys.argv[2].split(','))))
    #start_timemilli = int(time.mktime(datetime.datetime(*list).timetuple()))*1000
    #print start_timemilli
else:
    path = '.'


files = (listdir(path))

items_list = []
for f in files:
    if '.json' in f:
        with open(f) as json_file:
            data = json.load(json_file)
            regex = '([-0-9.]*)\s?([x<>=\s]+)\s([-0-9.]*)'
            validator_regex = re.compile(regex)
            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    validator = data['phases'][i]['measurements'][mea]['validators'][0]
                    matches = validator_regex.search(validator)
                    if matches and len(matches.groups()) == 3:
                        items_list.append(mea)
        break
    if '.scj' in f:
        with open(f) as json_file:
            data = json.load(json_file)
            for i in range (0, len(data['phases'])):
                for mea in data['phases'][i]['measurements']:
                    if mea.has_key('numeric_max'):
                        items_list.append(mea['name'])
        break

with open('FATP.csv', 'w') as csvfile:
    csvwriter = csv.writer(csvfile) 
    csvwriter.writerow(['log name'] + items_list)
    for f in files:
        if '.json' in f:
            try: 
                with open(f) as json_file:
                    data = json.load(json_file)
                    values_list = []
                    for i in range (0, len(data['phases'])):
                        for mea in data['phases'][i]['measurements']:
                            if mea in items_list:
                                values_list.append(data['phases'][i]['measurements'][mea]['measured_value'])
                    csvwriter.writerow([f] + values_list)
            except:
                print ("Bypass file: %s" % f) 
        if '.scj' in f:
            try:
                with open(f) as json_file:
                    data = json.load(json_file)
                    values_list = []
                    for i in range (0, len(data['phases'])):
                        for mea in data['phases'][i]['measurements']:
                            if mea['name'] in items_list:
                                print mea
                                values_list.append(mea['numeric_value'])
                    csvwriter.writerow([f] + values_list)
            except:
                print ("Bypass file: %s" % f) 
