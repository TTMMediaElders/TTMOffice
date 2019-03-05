# IMPORTS
import csv
import json
from emailing.emailing import send_gmail
import datetime
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
import calendar
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import xlsxwriter
import json
import win32com.client #https://docs.microsoft.com/en-us/office/vba/api/excel.pagesetup.fittopageswide
import win32com.client.dynamic
import win32api
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth
from pathlib import Path
# FUNCTIONS
#pip install -r /path/to/requirements.txt

def last_day(d, day_name):
	days_of_week = ['sunday', 'monday', 'tuesday', 'wednesday',
					'thursday', 'friday', 'saturday']
	target_day = days_of_week.index(day_name.lower())
	delta_day = target_day - d.isoweekday()
	if delta_day >= 0:
		delta_day -= 7  # go back 7 days
	return d + timedelta(days=delta_day)


def turn_to_datetime(date):
	date = date.split("/")
	date = datetime.date(int(date[2]), int(date[0]), int(date[1]))
	return date


def get_mish_org():
	# tmo stands for the mission organization
	tmo = csv_to_list('key_indicator_reports', type=True)
	tmo = [report for report in tmo if turn_to_datetime(
		report['Report Date']) >= (report_timing - relativedelta(weeks=1))]
	mission_organization_area = {area['Area']: {'Zone': area['Zone'],
												'District': area['District'], 'Area': area['Area']} for area in tmo}
	mission_organization_zone = {area['Zone']: {'Area': area['Area'],
												'District': area['District'], 'Zone': area['Zone']} for area in tmo}
	mission_organization_district = {area['District']: {'Area': area['Area'],
														'Zone': area['Zone'], 'District': area['District']} for area in tmo}
	tmo = {**mission_organization_area, **mission_organization_zone}
	tmo = {**tmo, **mission_organization_district}
	return tmo


def merger(output_path, input_paths):
	pdf_merger = PdfFileMerger()
	for path in input_paths:
		pdf_merger.append(path)
	with open(output_path, 'wb') as fileobj:
		pdf_merger.write(fileobj)


def sort_history_list(history_data, reverse=False):
	# Sorts a list of dates chronologically
	new_list = []
	date_compare = {find_wards['Report Date'] for find_wards in history_data}
	compare_list = list(map(turn_to_datetime, date_compare))
	zipped_pairs = zip(compare_list, date_compare)
	z = [x for _, x in sorted(zipped_pairs)]
	for date in z:
		new_list.append(
			[organize_it for organize_it in history_data if organize_it['Report Date'] == date][0])
	return new_list


def sort_list(x_axis_list, reverse=False):
	# Sorts a list of dates chronologically
	compare_list = list(map(turn_to_datetime, x_axis_list))
	zipped_pairs = zip(compare_list, x_axis_list)
	if reverse:
		z = [x for _, x in sorted(zipped_pairs, reverse=reverse)]
	else:
		z = [x for _, x in sorted(zipped_pairs)]
	return z


def pretify_name(area_name):
	# Changes are name from capitalized report name to normal name.
	area_name = area_name.split("_")
	area_name[0] = area_name[0].lower().capitalize()
	area_name = " ".join(area_name)
	return area_name


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
	for dataline in read_csv:
		csv_data.append(dataline)
	csv_file.close()
	return csv_data
# CLASSES

class Abathur():
	with open('.\\master_area_props.txt') as mish_org:
		master_area_props = json.load(mish_org)
	print_color = True
	with open('settings.json') as json_file:
		settingsData = json.load(json_file)

	report_timing = turn_to_datetime(settingsData[1]['Date'])
	translate_chinese = {
		'Area': '區域',
		'Zone': '地帶',
		'District': '地區',
		'Stake': '支聯會',
		'Ward': '支會'
	}
	stake_titles = {
		'NORTH': 'North Zone 北台北地帶',
		'SOUTH': 'South Zone 南台北地帶',
		'EAST': 'East Zone 東台北地帶',
		'XINZHU': 'Xinzhu Zone 新竹地帶',
		'CENTRAL': 'Central Zone 中台北地帶',
		'TAOYUAN': 'Taoyuan Zone 桃園地帶',
		'WEST': 'West Zone 西台北地帶',
		'HUALIAN': 'Hualian Zone 花蓮地帶'
	}
	ind_titles = {
		'BC': 'Baptized and Confirmed 已接受洗禮和證實的朋友',
		'NW': 'Next Week 下週',
		'CF': 'Current Friends 現有朋友',
		'BS': 'BD Friends at Sacrament Meeting 有洗禮目標及出席聖餐聚會的朋友',
		'BD': 'Baptismal Dates 訂下洗禮日期的朋友',
		'SM': 'Total Sacrament Meeting 出席聖餐聚會的朋友',
		'NF': 'New Friends 新朋友',
		'LA': 'Less Actives Back to Church 不活耀成員出席聖餐聚會',
	}
	ind_colors = {
		'Date of Last Baptism': '#800000',
		'2019 Baptismal Count': '#800000',
		'ind_red': '#ffc7ce',
		'ind_l_gr': '#c6efce',
		'ind_d_gr': '#87dd97',
		'white': '#FFFFFF',
		'BD': '#D0312E',
		'SM': '#F59122',
		'BC': '#1AA8DE',
		'NF': '#1FB48A',
		'BS': '#B83DBA',
		'CF': '#5F9EA0',
		'LA': '#800000',
		'black': '262626'
	}


	indicator_standards = {
		'BD': settingsData[0]['BD'],
		'SM': settingsData[0]['SM'],
		'NF': settingsData[0]['NF']
	}
	source_types = {
		'Source 1': '1 - Missionary Finding',
		'Source 2': '2 - Less-Active Member Referral',
		'Source 3': '3 - Recent-Convert Referral',
		'Source 4': '4 - Active Member Referral',
		'Source 5': '5 - English Class',
		'Source 6': '6 - Temple Tours'
	}#convert source headers to source names
	headers_not_in_csv = ['Date of Last Baptism', f'{report_timing.year} Baptismal Count']
	keys = ['Report Date', 'Area', 'District', 'Zone', 'Ward', 'Stake']
	xlApp = win32com.client.Dispatch("Excel.Application")
	xlApp.Visible = False
	xlApp.DisplayAlerts = False
	def __init__(self, file_name, make_these):
		self.data = csv_to_list(file_name)
		self.baptismal_sources = csv_to_list('.\\data\\baptismal_source_reports')
		self.reports_to_make = {}
		for report_type in make_these:
			self.reports_to_make[report_type] = list(
				set(map(lambda make_list: make_list[report_type], self.return_data_range(1))))
		self.csv_headers = list(filter(lambda data_line: turn_to_datetime(
			data_line["Report Date"]) == self.report_timing, self.data))[0].keys()

	def init_workbook(self, path):
		self.excel_path = Path(path).resolve().as_posix()
		self.wrkbook = xlsxwriter.Workbook(self.excel_path,{'strings_to_numbers': True})
		self.wrksheet = self.wrkbook.add_worksheet('Table')
		self.row = 0
		self.col = 4
		self.general_format = self.wrkbook.add_format(
			{'align': 'center', 'valign': 'vcenter', 'bold': True, 'border': 1})

	def set_headers(self, report_kind, report_title):
		# set_column(first_col, last_col, width, cell_format, options)
		# for reference:
		# 0:'A',	1:'B',	2:'C',	3:'D',	4:'E',	5:'F',	6:'G',	7:'H',	8:'I',	9:'J',	10:'K',	11:'L',	12:'M',	13:'N',\
		# 	14:'O',	15:'P',	16:'Q',	17:'R',	18:'S',	19:'T',	20:'U',	21:'V',	22:'W',	23:'X',	24:'Y',	25:'Z'
		try:
			self.wrksheet.set_row(self.row, 160)
			self.wrksheet.set_column('A:A', 8)
			self.wrksheet.set_column('B:B', 19)
			self.wrksheet.set_column('C:D', 7)
			if report_kind == "Zone" or report_kind == "Stake":
				self.wrksheet.set_column('E:E', 10)
				self.wrksheet.set_column('F:G', 7)
				self.wrksheet.set_column('H:L', 3)
				self.wrksheet.set_column('M:R', 7)
				if self.print_color:
					format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
													  'font_size': 18, 'font_color': 'white', 'bg_color': self.ind_colors['black']})
				else:
					format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
													  'font_size': 18, 'font_color': 'black', 'bg_color': self.ind_colors['white']})
				format.set_text_wrap()
				self.wrksheet.merge_range(f'A{self.row+1}:D{self.row+1}', self.stake_titles.get(report_title, pretify_name(report_title)) +
										  ' Key Indicators 主要指標 {:%B %d, %Y}'.format(self.report_timing), format)
				for header in self.headers_not_in_csv:
					if self.print_color:
						format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
														  'font_color': 'white', 'bg_color': self.ind_colors.get(header, self.ind_colors['black'])})
					else:
						format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
														  'font_color': 'black', 'bg_color': self.ind_colors['white']})
					format.set_text_wrap()
					format.set_rotation(90)
					self.wrksheet.write(self.row, self.col, header, format)
					self.wrksheet.write(self.row+1, self.col, " ", self.wrkbook.add_format({'align': 'center', 'valign': 'vcenter',
																							'font_color': 'white', 'bg_color': self.ind_colors['black']}))
					self.col += 1
			else:
				self.wrksheet.set_column('E:E', 7)
				self.wrksheet.set_column('F:J', 3)
				self.wrksheet.set_column('K:P', 7)
				if self.print_color:
					format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
													  'font_size': 18, 'font_color': 'white', 'bg_color': self.ind_colors['black']})
				else:
					format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
													  'font_size': 18, 'font_color': 'black', 'bg_color': self.ind_colors['white']})
				format.set_text_wrap()
				self.wrksheet.merge_range(f'A{self.row+1}:D{self.row+1}', self.stake_titles.get(report_title, pretify_name(report_title)) +
										  ' Key Indicators 主要指標 {:%B %d, %Y}'.format(self.report_timing), format)
			for ind in self.csv_headers:
				if ind not in self.keys:
					if self.print_color:
						# self.workbook.add_format({'font_color': 'white', 'bg_color':'#D0312E','valign': 'vcenter'})
						format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
														  'font_color': 'white', 'bg_color': self.ind_colors.get(ind, self.ind_colors['black'])})
					else:
						format = self.wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
														  'font_color': 'black', 'bg_color': self.ind_colors['white']})
					format.set_text_wrap()
					format.set_rotation(90)
					self.wrksheet.write(self.row, self.col, self.ind_titles.get(ind, ind), format)
					if ind in ['BC', 'NF', 'SM', 'BD']:
						standard = self.indicator_standards.get(
							ind, 1)*int((len(self.return_data_range(1, report_kind, report_title))/2))
						if ind == 'BC':
							standard = "{}/月".format(standard)
						else:
							standard = "{}/禮拜".format(standard)
						self.wrksheet.write(self.row+1, self.col, standard, self.wrkbook.add_format({'align': 'center', 'valign': 'vcenter',
																									 'font_color': 'white', 'bg_color': self.ind_colors['black']}))
					else:
						self.wrksheet.write(self.row+1, self.col, " ", self.wrkbook.add_format({'align': 'center', 'valign': 'vcenter',
																								'font_color': 'white', 'bg_color': self.ind_colors['black']}))
					self.col += 1
			self.wrksheet.merge_range(f'A{self.row+2}:D{self.row+2}', " ", self.wrkbook.add_format({'align': 'center', 'valign': 'vcenter',
																									'font_color': 'white', 'bg_color': self.ind_colors['black']}))
			self.row += 2
			if report_kind == "Zone" or report_kind == "Stake":
				self.col = 6
			else:
				self.col = 4
		except AttributeError:
			raise Warning(
				"No Workbook initialized, use the init_workbook method to initialize a workbook.")

	def make_table_current(self, report_style, report_header):
		# self.wrksheet.write(row, col, str, format)
		# self.wrksheet.merge_range(range,str,format))
		# data_query_result = self.return_data_range(1, report_kind, report_header)
		data_query_result = [dataline for dataline in self.data if turn_to_datetime(
			dataline['Report Date']) == self.report_timing and dataline[report_style] == report_header]
		ind_headers = data_query_result[-1].keys()
		start_col = self.col
		dist_wards_dict = {district: {wards['Ward'] for wards in data_query_result if wards['District'] == district}
						   for district in {find_districts['District'] for find_districts in data_query_result}}
		baps_this_year_total = 0
		for district, wards in dist_wards_dict.items():
			baps_this_year_total += sum([int(baps['BC']) for baps in self.data if baps['District']
										 == district and baps['Report Date'].split("/")[-1] == str(self.report_timing.year)])
			if report_style == 'Zone':
				self.wrksheet.merge_range(f'A{self.row+1}:R{self.row+1}', f'{ district} DISTRICT',
										  self.wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#2F4F4F'}))
				self.row += 1
			for ward in wards:
				if report_style == 'Stake':
					self.wrksheet.write(self.row, 0, "", self.wrkbook.add_format(
						{'border': 1, 'bg_color': self.ind_colors['black']}))
					self.wrksheet.merge_range(f'B{self.row+1}:R{self.row+1}', f'{ward} WARD - 支會', self.wrkbook.add_format(
						{'border': 1, 'font_color': 'white', 'bg_color': self.ind_colors['black']}))
					self.row += 1
				missionaries = [
					find_mish for find_mish in data_query_result if find_mish['Ward'] == ward]
				area_list = {find_wards['Area'] for find_wards in missionaries}
				for area in area_list:
					missionary_data = list(
						filter(lambda find_mish: find_mish['Area'] == area, missionaries))[0]
					missionary_names = self.master_area_props.get(
						pretify_name(area), ['No Match'])[-1]
					# missionary_test = [mish_names for mish_names in mail_list if mish_names['Area'] == pretify_name(area)]
					# print(str(missionary_test))
					# missionary_test = f"{missionary_test[0]['Type'][0]}. {missionary_test[0]['Last Name']} / {missionary_test[1]['Last Name']}"
					self.wrksheet.merge_range(
						f'A{self.row+1}:B{self.row+1}', missionary_names, self.general_format)
					self.wrksheet.merge_range(
						f'C{self.row+1}:D{self.row+1}', pretify_name(area), self.general_format)
					if report_style == 'Zone' or report_style == 'Stake':
						# print(ward,str(sum([int(week['BC']) for week in self.data if week['Ward']==ward])))
						date_of_last_bap = [last_bap['Report Date']
											for last_bap in self.data if last_bap['Area'] == area and int(last_bap['BC']) >= 1]
						if date_of_last_bap:
							date_of_last_bap = date_of_last_bap[-1]
						else:
							date_of_last_bap = " "
						self.wrksheet.write(self.row, 4, date_of_last_bap, self.general_format)
						baps_this_year = sum([int(baps['BC']) for baps in self.data if baps['Area'] ==
											  area and baps['Report Date'].split("/")[-1] == str(self.report_timing.year)])
						self.wrksheet.write(self.row, 5, baps_this_year, self.general_format)
					for name in missionary_data.keys():
						if name in ind_headers and name not in self.keys:
							if name in self.indicator_standards.keys():
								self.wrksheet.write(self.row, start_col, missionary_data[name], self.wrkbook.add_format(
									{'border': 1, 'font_color': 'black', 'bg_color': self.return_ind_color(name, missionary_data[name], 1), 'align': 'center', 'valign': 'vcenter', 'bold': True}))
							else:
								self.wrksheet.write(self.row, start_col,
													missionary_data[name], self.general_format)
							start_col += 1
					self.row += 1
					start_col = self.col
				if len(area_list) >= 2 and len(wards) >= 2 and report_style == 'Stake':
					totals = {senor_key: sum([int(a_dict[senor_key]) for a_dict in missionaries])
							  for senor_key in ind_headers if senor_key not in self.keys}
					self.wrksheet.merge_range(f'A{self.row+1}:F{self.row+1}', f"{ward} WARD TOTALS - 支會總數",
											  self.wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#595959'}))
					for ein_val in totals.values():
						self.wrksheet.write(self.row, start_col, ein_val, self.wrkbook.add_format(
							{'border': 1, 'font_color': 'white', 'bg_color': '#595959', 'align': 'center', 'valign': 'vcenter'}))
						start_col += 1
					start_col = self.col
					self.row += 1
			if report_style == "Zone":
				this_week_total = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in data_query_result if a_dict2['District'] == district])
								   for senor_key2 in ind_headers if senor_key2 not in self.keys}
				self.wrksheet.merge_range(f'A{self.row+1}:F{self.row+1}', f"{district} DISTRICT TOTALS - 地區總數",
										  self.wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#008080'}))
				for ein_val2 in this_week_total.values():
					self.wrksheet.write(self.row, start_col, ein_val2, self.wrkbook.add_format(
						{'border': 1, 'font_color': 'white', 'bg_color': '#008080', 'align': 'center', 'valign': 'vcenter'}))
					start_col += 1
				self.row += 1
				start_col = self.col
		if report_style != "Area":
			this_week_total = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in data_query_result])
							   for senor_key2 in ind_headers if senor_key2 not in self.keys}
			if report_style == 'Zone' or report_style == 'Stake':
				self.wrksheet.merge_range(f'A{self.row+1}:E{self.row+1}', f"{report_style.upper()} TOTALS - {self.translate_chinese[report_style]}總數", self.wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#003366'}))
				self.wrksheet.write(self.row, start_col-1, baps_this_year_total, self.wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'valign': 'vcenter'}))
			else:
				self.wrksheet.merge_range(f'A{self.row+1}:F{self.row+1}', f"{report_style.upper()} TOTALS - {self.translate_chinese[report_style]}總數", self.wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#003366'}))
			for ein_val2 in this_week_total.values():
				self.wrksheet.write(self.row, start_col, ein_val2, self.wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'valign': 'vcenter'}))
				start_col += 1
			self.row += 1
		self.row += 1

	def make_history_table(self, r_type, r_name):
		historical_data = self.return_data_range(6, r_type, r_name)
		ind_headers = historical_data[-1].keys()
		if r_type == 'Zone' or r_type == 'Stake':
			self.wrksheet.merge_range(f'A{self.row}:R{self.row}', "HISTORICAL DATA", self.wrkbook.add_format(
				{'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': 'white', 'bg_color': self.ind_colors['black']}))
		else:
			self.wrksheet.merge_range(f'A{self.row}:P{self.row}', "HISTORICAL DATA", self.wrkbook.add_format(
				{'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': 'white', 'bg_color': self.ind_colors['black']}))
		beginning_col = self.col
		weeks = sort_list(
			(set(map(lambda find_dates: find_dates['Report Date'], historical_data))), reverse=True)[1:]
		mish_count = len(
			list(filter(lambda find_mish: find_mish['Report Date'] == weeks[0], historical_data)))
		# print(str(mish_count),str(list(set(map(lambda get_areas: get_areas['Area'],list(filter(lambda find_mish: find_mish['Report Date']==weeks[0],historical_data)))))))#Print Zones, #of Comps and Area names
		for week in weeks:
			week_data = list(
				filter(lambda find_mish: find_mish['Report Date'] == week, historical_data))
			totalize = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in week_data])
						for senor_key2 in ind_headers if senor_key2 not in self.keys}
			self.wrksheet.merge_range(f'A{self.row+1}:D{self.row+1}',
									  '{:%B %d, %Y}'.format(turn_to_datetime(week)), self.general_format)
			for ind_name, some_val in totalize.items():
				if ind_name in self.indicator_standards.keys():
					self.wrksheet.write(self.row, self.col, some_val, self.wrkbook.add_format({'border': 1, 'font_color': 'black', 'bg_color': self.return_ind_color(
						ind_name, some_val, mish_count), 'align': 'center', 'valign': 'vcenter', 'bold': True}))
				else:
					self.wrksheet.write(self.row, self.col, some_val, self.general_format)
				self.col += 1
			self.col = beginning_col
			self.row += 1
		data_history_totals = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in historical_data if turn_to_datetime(
			a_dict2['Report Date']) != self.report_timing]) for senor_key2 in ind_headers if senor_key2 not in self.keys}
		for ind_ in ind_headers:
			if ind_ not in self.keys:
				data_history_totals[ind_] = data_history_totals[ind_]/len(weeks)
		if r_type == 'Zone' or r_type == 'Stake':
			self.wrksheet.merge_range(f'A{self.row+1}:F{self.row+1}', "Historical Data Averages",
									  self.wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#551A8B'}))
		else:
			self.wrksheet.merge_range(f'A{self.row+1}:D{self.row+1}', "Historical Data Averages",
									  self.wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#551A8B'}))
		for ein_val2 in data_history_totals.values():
			self.wrksheet.write(self.row, self.col, "{:.1f}".format(ein_val2), self.wrkbook.add_format(
				{'border': 1, 'font_color': 'white', 'bg_color': '#551A8B', 'align': 'center', 'valign': 'vcenter'}))
			self.col += 1
		self.row += 1

	def make_graphs(self, rep_type, rep_name):
		# Graphs indicator Data over time
		# DATA HAS TO BE INT FOR GRAPHING
		def prettier_date_format(date):
			# change xx/xx/xxxx(12/12/2012) format to Mon XX format(Dec 12)
			date = date.split("/")
			for var in range(len(date)):
				date[var] = int(date[var])
			date[0] = datetime.date(date[2], date[0], date[1]).strftime("%b")
			for var in range(len(date)):
				date[var] = str(date[var])
			date = " ".join(date[:2])
			return(date)
		dates_to_go_back = 13
		date_list = {dates['Report Date'] for dates in self.data if turn_to_datetime(
			dates['Report Date']) >= (self.report_timing-relativedelta(weeks=dates_to_go_back))}
		if date_list:
			data_group = self.return_data_range(dates_to_go_back, rep_type, rep_name)
			multi = 1
			if rep_type != 'Area':
				multi = len(self.return_data_range(1, rep_type, rep_name))/2
				print(f"rep_type: {rep_type}, rep_name: {rep_name}, multi: {multi}")
				new_group = []
				for week in date_list:
					sum_week = [sum_it for sum_it in data_group if sum_it['Report Date'] == week]
					if sum_week:
						sum_week_headers = sum_week[0].keys()
						total = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in sum_week])
								 for senor_key2 in sum_week_headers if senor_key2 not in self.keys}
						total['Report Date'] = week
						new_group.append(total)
				data_group = sort_history_list(new_group)
			date_list = list(map(prettier_date_format, sort_list(set(date_list))))
			self.wrksheet = self.wrkbook.add_worksheet('Graph')
			# self.wrksheet.write(row, col, text, format)
			# self.wrksheet.merge_range('A1:B1',text, format)
			self.wrksheet.merge_range('A1:Q1', pretify_name(rep_name), self.wrkbook.add_format(
				{'font_size': 24, 'font_color': '#000000', 'bg_color': '#FFFFFF'}))
			self.wrksheet.merge_range('A2:Q2', 'Key Indicators 主要指標', self.wrkbook.add_format(
				{'font_size': 18, 'font_color': '#FFFFFF', 'bg_color': '#000000'}))
			# Baptismal pictures and Titles
			# BC
			self.wrksheet.insert_image('A4', Path('.\\report images\\bap_pic.jpg').resolve().as_posix())
			self.wrksheet.merge_range('C4:F7', 'Baptized and Confirmed 已經接受洗禮和證實的朋友', self.wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['BC']) for a_dict in data_group]
			self.wrksheet.insert_image('A9', self.draw_graph(
				date_list, indicator_history, 'BC', multi))
			# SM
			self.wrksheet.insert_image('A25', Path('.\\report images\\church_pic.jpg').resolve().as_posix())
			self.wrksheet.merge_range('C25:F28', 'Sacrament Meeting 出席聖餐聚會的朋友', self.wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['SM']) for a_dict in data_group]
			self.wrksheet.insert_image('A30', self.draw_graph(
				date_list, indicator_history, 'SM', multi))
			# BD
			self.wrksheet.insert_image('J4', Path('.\\report images\\date_pic.jpg').resolve().as_posix())
			self.wrksheet.merge_range('L4:O7', 'Baptismal Dates 訂下洗禮日期的朋友', self.wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['BD']) for a_dict in data_group]
			self.wrksheet.insert_image('J9', self.draw_graph(
				date_list, indicator_history, 'BD', multi))
			# NF
			self.wrksheet.insert_image('J25', Path('.\\report images\\nf_pic.jpg').resolve().as_posix())
			self.wrksheet.merge_range('L25:O28', 'New Friends 新朋友', self.wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['NF']) for a_dict in data_group]
			self.wrksheet.insert_image('J30', self.draw_graph(
				date_list, indicator_history, 'NF', multi))
			# CONVERT BAPTISM VISUALS
			self.wrksheet.merge_range('A50:Q51', 'Convert Baptisms 歸信者洗禮', self.wrkbook.add_format(
				{'font_size': 18, 'font_color': '#FFFFFF', 'bg_color': '#000000'}))
			years_to_graph = [
				turn_to_datetime(self.data[-1]['Report Date']).year,
				(turn_to_datetime(self.data[-1]['Report Date'])-relativedelta(years=1)).year,
				(turn_to_datetime(self.data[-1]['Report Date'])-relativedelta(years=2)).year
			]
			years_data = {}
			for year in years_to_graph:
				years_data[year] = []
				date_list = []
				for month in range(1, 13):
					last_sunday = max(week[-1] for week in calendar.monthcalendar(year, month))
					date_list.append(turn_to_datetime(f'{month}/{last_sunday}/{year}'))
					_, num_days = calendar.monthrange(year, month)
					month_data = [report for report in self.data if turn_to_datetime(report['Report Date']) > turn_to_datetime(
						f'{month}/{1}/{year}') and turn_to_datetime(report['Report Date']) <= turn_to_datetime(f'{month}/{num_days}/{year}') and report[rep_type] == rep_name]
					month_bap_sum = sum([int(data['BC']) for data in month_data])
					# if month_bap_sum:
					years_data[year].append(month_bap_sum)
			month_list = ['{:%b}'.format(date) for date in date_list]
			self.wrksheet.insert_image('A53', self.draw_three_year_graph(month_list, years_data))
			area_baptismal_sources = [dataline for dataline in self.baptismal_sources if turn_to_datetime(dataline['Report Date']) >= (turn_to_datetime(self.baptismal_sources[-1]['Report Date'])-relativedelta(years=1)) and\
										rep_name == dataline[rep_type]]
			self.wrksheet.insert_image('I53', self.baptismal_source_pie(area_baptismal_sources))

	def make_summary_sheet(self, rep_type):
		# self.wrksheet = self.wrkbook.add_worksheet('Totals')
		this_week = self.return_data_range(1)
		indicators = [key for key in this_week[0].keys() if key not in self.keys]
		total_of_rep_type = len({r_t_names[rep_type] for r_t_names in this_week})
		totals = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in this_week])
				  for senor_key2 in indicators if senor_key2 not in self.keys}
		averages = {senor_key2: sum([int(a_dict2[senor_key2])/total_of_rep_type for a_dict2 in this_week])
					for senor_key2 in indicators if senor_key2 not in self.keys}
		print(f'total_of_rep_type: {total_of_rep_type}')
		print(f'totals: {totals}')
		print(f'averages: {averages}')

	def export_pdf(self,rep_tipe):
		self.wrkbook.close()
		#
		wb = self.xlApp.Workbooks.Open(self.excel_path)
		count = 2
		for sheet in wb.Sheets:
			sheet.Visible = 1
			sheet.PageSetup.Zoom = False
			sheet.PageSetup.Orientation = count
			sheet.PageSetup.CenterHorizontally = True
			sheet.PageSetup.CenterVertically = True
			sheet.PageSetup.LeftMargin = 0.24
			sheet.PageSetup.RightMargin = 0.24
			sheet.PageSetup.TopMargin = 0.24
			sheet.PageSetup.BottomMargin = 0.24
			sheet.PageSetup.HeaderMargin = 0.35
			sheet.PageSetup.FooterMargin = 0.35
			print_area = f'A1:{self.col_letter(self.col)}{self.row}'
			sheet.PageSetup.PrintArea = print_area
			count -= 1
		wb.Worksheets(["Table", "Graph"]).Select()
		print(self.excel_path)
		pdf_path = self.excel_path.replace('Worksheets','PDFs')
		print(pdf_path)
		print(pdf_path);print(type(pdf_path))
		self.xlApp.ActiveSheet.ExportAsFixedFormat(0, pdf_path, 0, True, True)
		wb.Close()
		wb = None

	def return_ind_color(self, ind_name, ind_num, multiplier):
		ind_num = int(ind_num)
		if ind_num < self.indicator_standards.get(ind_name, 'NF')*multiplier-1:
			if ind_name == 'BD' and ind_num == 4 and multiplier == 1:
				return self.ind_colors['ind_l_gr']
			else:
				return self.ind_colors['ind_red']
		elif ind_num == self.indicator_standards.get(ind_name, 'NF')*multiplier-1:
			return self.ind_colors['ind_l_gr']
		elif ind_num >= self.indicator_standards.get(ind_name, 'NF')*multiplier:
			return self.ind_colors['ind_d_gr']
		else:
			return self.ind_colors['white']

	def return_data_range(self, date_back, query_key="", query_word=""):
		date_back = self.report_timing - relativedelta(weeks=date_back)
		if query_word and query_key:
			self.ranged_data = list(filter(lambda data_line: turn_to_datetime(
				data_line["Report Date"]) >= date_back and turn_to_datetime(data_line["Report Date"]) <= self.report_timing and data_line[query_key] == query_word, self.data))
		else:
			self.ranged_data = list(filter(lambda data_line: turn_to_datetime(
				data_line["Report Date"]) >= date_back and turn_to_datetime(data_line["Report Date"]) <= self.report_timing, self.data))
		return self.ranged_data

	def draw_graph(self, x_axis, indicators, ind_name, multiplyer):
		indicator_graph_colors = {
			'BC': 'blue',
			'BD': 'red',
			'SM': 'orange',
			'NF': 'green'
		}
		fig = plt.figure(figsize=(6, 3.4))
		# L #H
		ax = fig.add_subplot(111)
		a_standard = []
		for a_num in range(len(indicators)):
			indicators[a_num] = int(indicators[a_num])
			standard = self.indicator_standards.get(ind_name, 1)*multiplyer
			if ind_name == 'BC':
				standard = 1
				multiplyer = 1
			a_standard.append(standard)
		# PLOT THE DATA
		while len(x_axis) != len(indicators):
			x_axis = x_axis[1:]
		plt.plot(x_axis, a_standard, color='gray')
		plt.plot(x_axis, indicators, color=indicator_graph_colors[ind_name])
		# ANNOTATE THE DATA
		for date, intg in zip(x_axis, indicators):
			ax.annotate('%s' % intg, xy=(date, intg), xytext=(
				date, intg+0.1), textcoords='data')
		# set the range of y axis
		if sorted(indicators)[-1]+1 <= a_standard[-1]:
			plt.ylim(-0.1, a_standard[-1]+1.5)
		else:
			plt.ylim(-0.1, sorted(indicators)[-1]+1.5)
		# angle the dates on the x-axis so they fit
		plt.xticks(rotation=45)
		plt.grid(b=True, axis='y')
		file_name = Path('.\\Graphs\\{}.png'.format(ind_name)).resolve().as_posix()
		plt.savefig(file_name, bbox_inches='tight')
		plt.clf()
		plt.close(fig)
		return file_name

	def baptismal_source_pie(self,year_baptismal_sources_data):
		# BAPTISMAL SOURCES PIE CHART
		sources = [some_item for some_item in list(self.baptismal_sources[-1].keys()) if 'Source' in some_item]
		if year_baptismal_sources_data:
			bap_data = {}
			for source in sources:
				bap_data[self.source_types[source]] = 0
				for line_bap in year_baptismal_sources_data:
					bap_data[self.source_types[source]] += int(line_bap[source])
					# bap_data[self.source_types[source]] = sum([int(some_dict[source]) if some_dict[source] != '' else some_dict[source] == int('0') for some_dict in year_baptismal_sources_data])
			total = sum(bap_data.values())
			baptisms = True
			if total == 0:
				print("NO BAPTISMS IN {}!!!".format(year_baptismal_sources_data[-1]["Area"]))
				baptisms = False
				total = 1
			for par_name, param in bap_data.items():
				bap_data[par_name] = param / total
			labels = [key[4:] for key in bap_data.keys()]
			sizes = [val for val in bap_data.values() if val > 0]
			color_list = {
				'1 - Missionary Finding': '#F2A104',
				'2 - Less-Active Member Referral': '#888C46',
				'3 - Recent-Convert Referral': '#0294A5',
				'4 - Active Member Referral': '#A79C93',
				'5 - English Class': '#C1403D',
				'6 - Temple Tours': '#0438A3'
			}
			color_list = [color_list[v] for v in bap_data.keys()]
			fig, ax = plt.subplots(figsize=(6.5, 5))
										   # L    #H
			wedges, texts, autotexts = ax.pie(sizes, autopct=lambda pct: "{:.1f}%".format(pct),
											  textprops=dict(color="#d3d3d3"), radius=1,
												wedgeprops=dict(width=0.4, edgecolor='w'),counterclock=False,
												pctdistance=0.8,colors=color_list)
			ax.legend(wedges, labels,
					  title="Sources",
					  loc="center left",
					  bbox_to_anchor=(1, 0, 0.5, 1))
			plt.setp(autotexts, size=12, weight="bold")
			# ax.set_title("Baptismal Sources\n(over last 52wks)")
			file_name = Path('.\\Graphs\\pieChart.png').resolve().as_posix()
			plt.savefig(file_name,bbox_inches='tight')
			plt.clf()
			plt.close(fig)
			if baptisms:
				return file_name
			else:
				return '.\\Graphs\\no_baptisms.png'

	@staticmethod
	def col_letter(sym):
		col_letter_dict = {0: 'A',	1: 'B',	2: 'C',	3: 'D',	4: 'E',	5: 'F',	6: 'G',	7: 'H',	8: 'I',	9: 'J',	10: 'K',	11: 'L',	12: 'M',	13: 'N',
						   14: 'O',	15: 'P',	16: 'Q',	17: 'R',	18: 'S',	19: 'T',	20: 'U',	21: 'V',	22: 'W',	23: 'X',	24: 'Y',	25: 'Z'}
		try:
			return col_letter_dict[sym]
		except:
			return list(col_letter_dict.keys())[list(col_letter_dict.values()).index(sym)]

	@staticmethod
	def draw_three_year_graph(x_axis, three_yr_data):
		# Three subplots sharing both x/y axes
		fig, (ax1, ax2, ax3) = plt.subplots(3, sharex=True, sharey=True)
		axes = [ax1,ax2,ax3]
		color = ['blue', 'orange', 'red']
		i = 0
		for ano, indicators in three_yr_data.items():
			for a_num in range(len(indicators)):
				indicators[a_num] = int(indicators[a_num])
			# PLOT THE DATA
			if ano == datetime.date.today().year:
				esc_x_axis = x_axis
				cut_list = datetime.date.today().month
				esc_x_axis = esc_x_axis[:cut_list]
				indicators = indicators[:cut_list]
				esc_x_axis = esc_x_axis [:len(indicators)]
				axes[i].plot(esc_x_axis, indicators, color=color[i], label=f'{ano} ({sum(indicators)})')
			else:
				axes[i].plot(x_axis, indicators, color=color[i], label=f'{ano} ({sum(indicators)})')
			# ANNOTATE THE DATA
			for date, intg in zip(x_axis, indicators):
				axes[i].annotate('%s' % intg, xy=(date, intg), xytext=(
					date, intg+0.1), textcoords='data')
			axes[i].grid(b=True, axis='y')
			axes[i].legend()
			i+=1
		# set the range of y axis
		max_y = []
		for ind in three_yr_data.values():
			if ind:
				ind = max(ind)
			else:
				ind = 0
			max_y.append(ind)
		max_y = max(max_y)
		plt.ylim(-0.1, max_y+(max_y*0.5))
		# angle the dates on the x-axis so they fit
		plt.xticks(rotation=45)
		plt.grid(b=True, axis='y')
		plt.legend()
		file_name = Path('.\\Graphs\\3yearGraph.png').resolve().as_posix()
		plt.savefig(file_name, bbox_inches='tight')
		plt.clf()
		plt.close(fig)
		return file_name

class Zeratul():
	report_timing = Abathur.report_timing
	mail_list = csv_to_list('.\\data\\Roster-Excel')
	folder_ids = {}
	gauth = GoogleAuth()
	gauth.LoadCredentialsFile(".\\emailing\\credentials.json")
	if gauth.credentials is None:
		gauth.LocalWebserverAuth()
	elif gauth.access_token_expired:
		gauth.Refresh()
	else:
		gauth.Authorize()
	gauth.SaveCredentialsFile("credentials.json")
	drive = GoogleDrive(gauth)
	zone_name_dict = {
		"NT": "NORTH",
		"ET": "EAST",
		"WT": "WEST",
		"ST": "SOUTH",
		"CT": "CENTRAL",
		"HL": "HUALIAN",
		"XZ": "XINZHU",
		"Ta": "TAOYUAN"
	}
	def __init__(self):
		self.report_day = '{:%B %d, %Y}'.format(self.report_timing)
		file_list = self.drive.ListFile(
			{'q': "'1gDAH387mrHWNhxKXlyDmgtUKVy9OeJAA' in parents and trashed=false"}).GetList()
		file_check = [_file for _file in file_list if _file['title'] == self.report_day]
		if file_check:
			self.report_folder = file_check[0]
		else:
			self.report_folder = self.drive.CreateFile({'title': self.report_day,
														'mimeType': 'application/vnd.google-apps.folder',
														'parents': [{"kind": "drive#filelink", "id": '1gDAH387mrHWNhxKXlyDmgtUKVy9OeJAA'}]})
			self.report_folder.Upload()

	def make_report_folder(self, folder_name):
		file_list = self.drive.ListFile(
			{'q': "'{}' in parents and trashed=false".format(self.report_folder['id'])}).GetList()
		file_check = [_file for _file in file_list if _file['title'] == folder_name]
		if file_check:
			for folder in file_check:
				if folder['title'] == folder_name:
					self.folder_ids[folder_name] = folder['id']
		else:
			report_folder = self.drive.CreateFile({'title': folder_name,
												   'mimeType': 'application/vnd.google-apps.folder',
												   'parents': [{"kind": "drive#filelink", "id": self.report_folder['id']}]})
			report_folder.Upload()
			self.folder_ids[folder_name] = report_folder["id"]

	def send_report(self, file_name, file_type, file_path):
		#
		file_list = self.drive.ListFile(
			{'q': "'{0}' in parents and trashed=false".format(self.folder_ids[file_type])}).GetList()
		file_check = [_file for _file in file_list if _file['title'] == file_name]
		if not file_check:
			ki_report = self.drive.CreateFile({'title': f'{file_name}', "parents": [
				{"kind": "drive#fileLink", "id": self.folder_ids[file_type]}]})
			ki_report.SetContentFile(file_path)
			ki_report.Upload()
			permission = ki_report.InsertPermission({
				'type': 'anyone',
				'value': 'anyone',
				'role': 'reader'})
			send_from = "ttmmediaelders@gmail.com"
			password = "fiwctgjmcnatrfhb"
			link = ki_report['alternateLink'][8:]
			subject = f'{file_name} {file_type} Report/報告'
			message = f"Here is the {self.report_day} key indicator report for the {file_name} {file_type}:\n\n{link}"
			# Find the area email for current report
			email = []
			if file_type == 'Area':
				email = [find_email['Email']
						 for find_email in self.mail_list if find_email['Area'] == pretify_name(file_name)]
			elif file_type == 'District':
				dls = [cand for cand in self.mail_list if cand['Position']
					   == 'DL' or cand['Position'] == 'DT']
				email = [dl['Email'] for dl in dls if dl['District']
						 == pretify_name(file_name) == dl['District']]
			elif file_type == 'Zone':
				zls_stls = [cand for cand in self.mail_list if 'ZL' in cand['Position']
							or 'STL' in cand['Position']]
				email = [zl_stl['Email']
						 for zl_stl in zls_stls if file_name == self.zone_name_dict[zl_stl['Zone']]]
			email = set(email)
			#
			if email:
				sleep_time = 5
				for send_to in email:
					while True:
						try:
							send_gmail(send_from, send_to, password, text=message, subject=subject)
							print(
								f"Uploaded {file_name} to {file_type} folder and sent file to {send_to}!")
							break
						except:
							time.sleep(sleep_time)
							sleep_time = sleep_time * 1.5
			else:
				print(f'NOT SENT:\t{file_type}\t{file_name}\t{file_path}\n')
				with open('not_sent.txt', 'a', encoding='utf-8') as error:
					error.write(f'{file_type}\t{file_name}\t{file_path}\r\n')
		else:
			print(
				f"{file_name} {file_type} report has already been uploaded to TTM google drive and likely sent!")
