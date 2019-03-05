import csv
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import datetime
import json
from datetime import date
from collections import OrderedDict
import pprint
# FUNCTIONS/GLOBALS
source_types = {
    '1 - Missionary Finding': 'Source 1',
    '2 - Less-Active Member Referral': 'Source 2',
    '3 - Recent-Convert Referral': 'Source 3',
    '4 - Active Member Referral': 'Source 4',
    '5 - English Class': 'Source 5',
    '6 - Temple Tours': 'Source 6',
}


def csv_to_list(file_name, type=True):
    # Takes data from a CSV and puts it all in a list.
                # True means use dictionaries, false means use lists.
    try:
        csv_file = open('{}.csv'.format(file_name), 'r', encoding='utf-8')
    except:
        try:
            csv_file = open('{}.txt'.format(file_name), 'r', encoding='utf-8')
        except:
            print("UNABLE TO OPEN FILE")
    if type:
        read_csv = csv.DictReader(csv_file)
    else:
        read_csv = csv.reader(csv_file)
    csv_data = []
    for dataLine in read_csv:
        csv_data.append(dataLine)
    csv_file.close()
    return csv_data


#
# RETRIEVE BAPTISMAL SOURCE DATA
read_check_this_week = csv_to_list('Book1')
# RETRIEVE MISSION ORGANIZATION
try_it = open('C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\master_area_props.txt')
master_area_props = json.load(try_it)
try_it.close()
# CHECK IF THE WHOLE MISSION HAS REPORTED YET
append_and_write = True
# print(str(len(read_check_this_week)))
# print(str(len(master_area_props)))
# if len(read_check_this_week) == len(master_area_props):
# 	append_and_write = True
# else:
# 	print("OH NO!")
# IF SO, APPEND THE DATA AND WRITE IT INTO THE BAPTISMAL SOURCE ARCHIVE
pprint.pprint(master_area_props)
if append_and_write:
    formatso = []
    for report in read_check_this_week:
        # FORMAT GOOGLE SURVEY RESULTS INTO REPORT LINES
        check_zone = master_area_props[report["Area 區域"]][0]
        check_zone = check_zone.split(" ")
        check_zone = check_zone[0].upper()
        if check_zone == "CENTRAL~WEST~NORTH":
            check_zone = "NORTH"
        if check_zone == "EAST~SOUTH":
            check_zone = "SOUTH"
        #
        check_district = master_area_props[report["Area 區域"]][1]
        check_district = check_district.split(" ")
        if check_district[1] == "A" or check_district[1] == "B":
            check_district = "_".join(check_district[:2]).upper()
        else:
            check_district = check_district[0].upper()
        #
        check_area = report["Area 區域"].split(" ")
        check_area[0] = check_area[0].upper()
        check_area = "_".join(check_area)
        #
        check_ward = check_area
        if "1" in check_ward:
            check_ward = check_ward.split("_")[0]
            check_ward = check_ward + "_1"
        elif "2" in check_ward:
            check_ward = check_ward.split("_")[0]
            check_ward = check_ward + "_2"
        elif "3" in check_ward:
            check_ward = check_ward.split("_")[0]
            check_ward = check_ward + "_3"
        elif "4" in check_ward:
            check_ward = check_ward.split("_")[0]
            check_ward = check_ward + "_4"
        else:
            check_ward = check_ward.split("_")[0]
        #
        # OrderedDict([('\ufeffTimestamp', '12/9/2018 16:56'),x
        #      ('Area 區域', 'Taoyuan 1 E'),x
        #      ('Name 姓名', '何坤穎'),x
        #      ('Source 來源', '3 - Recent-Convert Referral'),              ---?
        #      ('Name 姓名(1)', ''),x
        #      ('Source 來源(1)', ''),              ---?
        #      ('Name 姓名(2)', ''),x
        #      ('Source 來源(2)', ''),              ---?
        #      ('Name 姓名(3)', ''),x
        #      ('Source 來源(3)', '')])              ---?
        # 1 - Missionary Finding
        # 2 - Less-Active Member Referral
        # 3 - Recent-Convert Referral
        # 5 - English Class
        # 4 - Active Member Referral
        # 6 - Temple Tours
        #
        input_sources = {
            'Source 1': 0,
            'Source 2': 0,
            'Source 3': 0,
            'Source 4': 0,
            'Source 5': 0,
            'Source 6': 0
        }
        for key, val in report.items():
            if 'Source' in key:
                if val in source_types.keys():
                    input_sources[source_types[val]] += 1
                    break
        formatso.append(
            OrderedDict([
                #	Source 來源(1)	Name 姓名(2)	Source 來源(2)	Name 姓名(1)	Source 來源(1)	Name 姓名(2)	Source 來源(2)	Name 姓名(3)	Source 來源(3)
                        ("Report Date", report["\ufeffTimestamp"].split(" ")[0]),
                        ("Area", check_area),
                        ("District", check_district),
                        ("Zone", check_zone),
                        ("Ward", check_ward),
                        ("Stake", check_zone),
                        ("Source 1", input_sources["Source 1"]),
                        ("Source 2", input_sources["Source 2"]),
                        ("Source 3", input_sources["Source 3"]),
                        ("Source 4", input_sources["Source 4"]),
                        ("Source 5", input_sources["Source 5"]),
                        ("Source 6", input_sources["Source 6"])
                        ])
        )
    # WRITE IN NEW INDICATORS
    kir = open("C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\baptismal_source_reports.txt", "a", newline='')
    kir_csv = csv.DictWriter(kir, fieldnames=['Report Date', 'Area', 'District', 'Zone', 'Ward',
                                              'Stake', 'Source 1', 'Source 2', 'Source 3', 'Source 4', 'Source 5', 'Source 6'])
    for week_report in formatso:
        #	Report Date,Area,District,Zone,Ward,Stake,Source 1,Source 2,Source 3,Source 4,Source 5,Source 6
        kir_csv.writerow(week_report)
    kir.close()
