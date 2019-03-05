import csv
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import datetime
import json
from datetime import date
from collections import OrderedDict
# FUNCTIONS


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


# gauth = GoogleAuth()
# drive = GoogleDrive(gauth)
# GLOBALS
# time
m_date = "{}/{}/{}".format(datetime.date.today().year,
                           datetime.date.today().month, datetime.date.today().day)
what_num_week_it_is = m_date.split("/")
for item in range(len(what_num_week_it_is)):
    what_num_week_it_is[item] = int(what_num_week_it_is[item])
what_num_week_it_is = datetime.date(
    what_num_week_it_is[0], what_num_week_it_is[1], what_num_week_it_is[2]).isocalendar()[1]
#
# RETRIEVE THIS WEEKS NEW KEY INDICATORS
# check_this_week = drive.CreateFile({'id':'1Zyep97trdvX4jWEKdFKeNL5OXv9S8YjlzWHqhQeNo64'})#ABPA google survey results
# file_name = "C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\"+str(what_num_week_it_is)+m_date[2:4]+".txt"
file_name = "feb24-2019"

# check_this_week.GetContentFile(file_name, mimetype='text/csv')
read_check_this_week = csv_to_list(file_name)
# RETRIEVE MISSION ORGANIZATION
try_it = open('C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\master_area_props.txt')
master_area_props = json.load(try_it)
try_it.close()
# CHECK IF THE WHOLE MISSION HAS REPORTED YET
append_and_write = True
# if len(read_check_this_week) == len(master_area_props):
#     append_and_write = True
# append_and_write = True
# IF SO, APPEND THE DATA AND WRITE IT INTO THE KEY INDICATOR HISTORY
print(str(master_area_props))
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
        formatso.append(
            OrderedDict([
                        ("Report Date", report["\ufeffTimestamp"].split(" ")[0]),
                        ("Area", check_area),
                        ("District", check_district),
                        ("Zone", check_zone),
                        ("Ward", check_ward),
                        ("Stake", check_zone),
                        ("BC", report["Baptized and Confirmed 已經洗禮和接受證實的朋友"]),
                        ("NW", report["Next Week Confirmations 下週接受證實的朋友"]),
                        ("A", report["A Friends A朋友"]),
                        ("B", report["B Friends B朋友"]),
                        ("C", report["C Friends C朋友"]),
                        ("D", report["D Friends D朋友"]),
                        ("CF", report["Current Number of Friends Being Taught 現有朋友"]),
                        ("BS", report["Baptismal Date Friends at Sacrament Meeting 有洗禮目標，出席聖餐聚會的朋友"]),
                        ("LA", report.get("Less Active Members at Sacrament Meeting 不活耀的成員出席聖餐聚會", 0)),
                        ("BD", report["BD"]),
                        ("SM", report["SM"]),
                        ("NF", report["NI"])
                        ])
        )
    # WRITE IN NEW INDICATORS
    kir = open("C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\key_indicator_reports.txt", "a", newline='')
    kir_csv = csv.DictWriter(kir, fieldnames=["Report Date", "Area", "District", "Zone",
                                              "Ward", "Stake", "BC", "NW", "A", "B", "C", "D", "CF", "BS", "LA", "BD", "SM", "NF"])
    print(str(formatso))
    for week_report in formatso:
        # Report Date,Area,District,Zone,Ward,Stake,BC,NW,A,B,C,D,BS,CI,BD,SM,NI
        kir_csv.writerow(week_report)
    kir.close()
