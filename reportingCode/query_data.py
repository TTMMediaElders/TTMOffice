import calendar
import csv
import datetime
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
import pprint
import xlsxwriter
import win32com.client
import win32com.client.dynamic
import win32api
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import os
#
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
	file_name = os.path.abspath('.\\Graphs\\3yearGraph.png')
	plt.savefig(file_name, bbox_inches='tight')
	plt.clf()
	plt.close(fig)
	return file_name

def baptismal_source_pie(year_baptismal_sources_data):
	# BAPTISMAL SOURCES PIE CHART
	source_types = {
		'Source 1': '1 - Missionary Finding',
		'Source 2': '2 - Less-Active Member Referral',
		'Source 3': '3 - Recent-Convert Referral',
		'Source 4': '4 - Active Member Referral',
		'Source 5': '5 - English Class',
		'Source 6': '6 - Temple Tours'
	}
	sources = [some_item for some_item in list(baptismal_sources[-1].keys()) if 'Source' in some_item]
	if year_baptismal_sources_data:
		bap_data = {}
		for source in sources:
			bap_data[source_types[source]] = 0
			for line_bap in year_baptismal_sources_data:
				bap_data[source_types[source]] += int(line_bap[source])
				# bap_data[source_types[source]] = sum([int(some_dict[source]) if some_dict[source] != '' else some_dict[source] == int('0') for some_dict in year_baptismal_sources_data])
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
		file_name = os.path.abspath('.\\Graphs\\pieChart.png')
		plt.savefig(file_name,bbox_inches='tight')
		plt.clf()
		plt.close(fig)
		if baptisms:
			return file_name
		else:
			return os.path.abspath('.\\Graphs\\no_baptisms.png')

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


def merger(output_path, input_paths):
	pdf_merger = PdfFileMerger()
	for path in input_paths:
		pdf_merger.append(path)
	with open(output_path, 'wb') as fileobj:
		pdf_merger.write(fileobj)


def col_letter(sym):
	col_letter_dict = {0: 'A',	1: 'B',	2: 'C',	3: 'D',	4: 'E',	5: 'F',	6: 'G',	7: 'H',	8: 'I',	9: 'J',	10: 'K',	11: 'L',	12: 'M',	13: 'N',
					   14: 'O',	15: 'P',	16: 'Q',	17: 'R',	18: 'S',	19: 'T',	20: 'U',	21: 'V',	22: 'W',	23: 'X',	24: 'Y',	25: 'Z'}
	try:
		return col_letter_dict[sym]
	except:
		return list(col_letter_dict.keys())[list(col_letter_dict.values()).index(sym)]


def sort_list(x_axis_list, reverse=False):
	# Sorts a list of dates chronologically
	compare_list = list(map(change_to_datetime_obj, x_axis_list))
	zipped_pairs = zip(compare_list, x_axis_list)
	if reverse:
		z = [x for _, x in sorted(zipped_pairs, reverse=reverse)]
	else:
		z = [x for _, x in sorted(zipped_pairs)]
	return z


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


def sort_history_list(history_data, reverse=False):
	# Sorts a list of dates chronologically
	new_list = []
	date_compare = {find_wards['Report Date'] for find_wards in history_data}
	compare_list = list(map(change_to_datetime_obj, date_compare))
	zipped_pairs = zip(compare_list, date_compare)
	z = [x for _, x in sorted(zipped_pairs)]
	for date in z:
		new_list.append(
			[organize_it for organize_it in history_data if organize_it['Report Date'] == date][0])
	return new_list


def change_to_datetime_obj(time_var):
	time_var = time_var.split("/")
	change = datetime.date(int(time_var[2]), int(time_var[0]), int(time_var[1]))
	return change


def return_ind_color(ind_name, ind_num, multiplier=1):
	ind_num = int(ind_num)
	if ind_num < indicator_standards.get(ind_name, 'NF')*multiplier-1:
		return ind_colors['ind_red']
	elif ind_num == indicator_standards.get(ind_name, 'NF')*multiplier-1:
		return ind_colors['ind_l_gr']
	elif ind_num >= indicator_standards.get(ind_name, 'NF')*multiplier:
		return ind_colors['ind_d_gr']
	else:
		return ind_colors['white']

def find_closest_week(some_date):
	count = 0
	while True:
		find_it = len([find_mish for find_mish in mish_data_history if find_mish['Report Date'] == change_to_datetime_obj(mish_data_history[-1]['Report Date'])-relativedelta(weeks=count) and find_mish['Zone']==zname])
		count+=1
		if find_it:
			break
	return find_it

def draw_graph(x_axis, indicators, ind_name, multiplyer):
	indicator_graph_colors = {
		'BC': 'blue',
		'BD': 'red',
		'SM': 'orange',
		'NF': 'green'
	}
	fig = plt.figure(figsize=(6, 3.4))
	# L #H
	ax = fig.add_subplot(111)
	# fig, ax = plt.subplots()
	# import matplotlib.ticker as ticker
	# Be sure to only pick integer tick locations.
	# for axis in [ax.xaxis, ax.yaxis]:
	# axis.set_major_locator(ticker.MaxNLocator(integer=True))
	a_standard = []
	for a_num in range(len(indicators)):
		indicators[a_num] = int(indicators[a_num])
		standard = indicator_standards.get(ind_name, 1)*multiplyer
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
	file_name = os.path.abspath('.\\Graphs\\{}.png').format(
		ind_name)
	plt.savefig(file_name, bbox_inches='tight')
	plt.clf()
	plt.close(fig)
	return file_name


#
all_data = csv_to_list(os.path.abspath('.\\data\\key_indicator_reports'))
baptismal_sources = csv_to_list(os.path.abspath('.\\data\\baptismal_source_reports'))
secret_keys = ['Report Date', 'Area', 'District', 'Zone', 'Ward', 'Stake']
ind_headers = all_data[-1].keys()
#
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
indicator_standards = {
	'BD': 6,
	'SM': 3,
	'NF': 4
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
	'black': '#262626'
}
year_of_this_report = 2019
date_list = []
for month in range(1, 13):
	last_sunday = max(week[-1] for week in calendar.monthcalendar(year_of_this_report, month))
	date_list.append(change_to_datetime_obj(f'{month}/{last_sunday}/{year_of_this_report}'))
# print(f'TARGET DATE LIST: {date_list}')
merge_list = []
xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.Visible = False  # Keep the excel sheet closed
xlApp.DisplayAlerts = True  # "Do you want to over write it?" Will not Pop up
for date in date_list:
	_, num_days = calendar.monthrange(date.year, date.month)
	mish_data_history = [dataline for dataline in all_data if change_to_datetime_obj(
		dataline['Report Date']) >= change_to_datetime_obj(f'{date.month}/{1}/{date.year}') and change_to_datetime_obj(dataline['Report Date']) <= change_to_datetime_obj(f'{date.month}/{num_days}/{date.year}')]
	if mish_data_history:
		report_title = date
		# path = Path(f'.\\Worksheets\\monthlyReports\\{date}.xlsx').resolve().as_posix()
		path = os.path.abspath(f'.\\Worksheets\\monthlyReports\\{date}.xlsx')
		wrkbook = xlsxwriter.Workbook(path)
		general_format = wrkbook.add_format(
			{'align': 'center', 'valign': 'vcenter', 'bold': True, 'border': 1})
		wrksheet = wrkbook.add_worksheet('REPORT')
		# SET REPORT HEADERS
		keys = ['Report Date', 'Area', 'District', 'Zone', 'Ward', 'Stake']
		row = 0
		col = 4
		wrksheet.set_row(row, 160)
		wrksheet.set_column('A:A', 8)
		wrksheet.set_column('B:B', 19)
		wrksheet.set_column('C:D', 7)
		wrksheet.set_column('E:E', 7)
		wrksheet.set_column('F:J', 5)
		wrksheet.set_column('K:P', 7)
		format = wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
									 'font_size': 18, 'font_color': 'white', 'bg_color': ind_colors['black']})
		format.set_text_wrap()
		wrksheet.merge_range(f'A{row+1}:D{row+1}',
							 '{:%B %d, %Y} Mission Report'.format(date), format)
		for ind in ind_headers:
			if ind not in keys:
				format = wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
											 'font_color': 'white', 'bg_color': ind_colors.get(ind, ind_colors['black'])})
				format.set_text_wrap()
				format.set_rotation(90)
				wrksheet.write(row, col, ind_titles.get(ind, ind), format)
				col += 1
		row += 1
		col = 4
	# WRITE TOTALS BY ZONE
		beginning_col = col
		format = wrkbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
		zone_names = {zone_name['Zone'] for zone_name in mish_data_history}
		for zname in zone_names:
			wrksheet.write(row, col-1, stake_titles.get(zname,zname), format)
			wrksheet.merge_range(f'A{row+1}:D{row+1}',stake_titles.get(zname,zname), format)
			count = -1
			while True:
				this_week = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in mish_data_history if a_dict2['Zone'] == zname and change_to_datetime_obj(a_dict2['Report Date']) == change_to_datetime_obj(mish_data_history[count]['Report Date'])])
																												for senor_key2 in ind_headers if senor_key2 not in keys}
				count-=1
				if this_week:
					break
			mish_count = len([find_mish for find_mish in mish_data_history if change_to_datetime_obj(find_mish['Report Date']) == change_to_datetime_obj(mish_data_history[-1]['Report Date']) and find_mish['Zone']==zname])
			for indekator, valyoo in this_week.items():
				if indekator in indicator_standards.keys():
					wrksheet.write(row, col, valyoo, wrkbook.add_format({'border': 1, 'font_color': 'black', 'bg_color': return_ind_color(indekator, valyoo, mish_count), 'align': 'center', 'valign': 'vcenter', 'bold': True}))
				else:
					wrksheet.write(row, col, valyoo, general_format)
				col += 1
			col = beginning_col
			row += 1
		data_history_totals = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in mish_data_history if change_to_datetime_obj(a_dict2['Report Date']) == change_to_datetime_obj(mish_data_history[-1]['Report Date'])])
											   for senor_key2 in ind_headers if senor_key2 not in keys}
		wrksheet.merge_range(f'A{row+1}:D{row+1}', "Totals",
							 wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#003366'}))
		for ein_val2 in data_history_totals.values():
			wrksheet.write(row, col, f"{ein_val2}", wrkbook.add_format(
				{'border': 1, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'valign': 'vcenter'}))
			col += 1
		col = beginning_col
		row += 1
		for ind_ in ind_headers:
			if ind_ not in keys and ind_ != 'BC':
				try:
					data_history_totals[ind_] = data_history_totals[ind_]/len(zone_names)
				except:
					pass
		wrksheet.merge_range(f'A{row+1}:D{row+1}', "Averages",
							 wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#551A8B'}))
		format = wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#551A8B', 'align': 'center', 'valign': 'vcenter'})
		for el_llave,ein_val2 in data_history_totals.items():
			if el_llave == 'BC':
				ein_val2 = " "
			else:
				ein_val2 = "{:.1f}".format(ein_val2)
			wrksheet.write(row, col, ein_val2, format)
			col += 1
		row += 1
		col = beginning_col
	###HISTORICAL DATA TABLE
		row += 1
		wrksheet.merge_range(f'A{row}:P{row}', "HISTORICAL DATA", wrkbook.add_format(
			{'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': 'white', 'bg_color': ind_colors['black']}))
		beginning_col = col
		mish_data_history = [dataline for dataline in all_data if change_to_datetime_obj(
			dataline['Report Date']) >= (date-relativedelta(weeks=5)) and change_to_datetime_obj(dataline['Report Date']) <= date]
		weeks = sort_list({find_dates['Report Date']
						  for find_dates in mish_data_history}, reverse=True)
		if weeks:
			mish_count = len([find_mish for find_mish in mish_data_history if find_mish['Report Date'] == weeks[0]])
			for week in weeks:
				week_data = [mish for mish in mish_data_history if mish['Report Date'] == week]
				totalize = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in week_data])
							for senor_key2 in ind_headers if senor_key2 not in keys}
				wrksheet.merge_range(f'A{row+1}:D{row+1}',
									 '{:%B %d, %Y}'.format(change_to_datetime_obj(week)), general_format)
				for ind_name, some_val in totalize.items():
					if ind_name in indicator_standards.keys():
						wrksheet.write(row, col, some_val, wrkbook.add_format({'border': 1, 'font_color': 'black', 'bg_color': return_ind_color(
							ind_name, some_val, mish_count), 'align': 'center', 'valign': 'vcenter', 'bold': True}))
					else:
						wrksheet.write(row, col, some_val, general_format)
					col += 1
				col = beginning_col
				row += 1
			data_history_totals = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in mish_data_history])
												   for senor_key2 in ind_headers if senor_key2 not in keys}
			wrksheet.merge_range(f'A{row+1}:D{row+1}', "Totals",
								 wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#003366'}))
			for ein_val2 in data_history_totals.values():
				wrksheet.write(row, col, f"{ein_val2}", wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#003366', 'align': 'center', 'valign': 'vcenter'}))
				col += 1
			col = beginning_col
			row += 1
			for ind_ in ind_headers:
				if ind_ not in keys:
					try:
						data_history_totals[ind_] = data_history_totals[ind_]/len(weeks)
					except:
						pass
			wrksheet.merge_range(f'A{row+1}:D{row+1}', "Historical Data Averages",
								 wrkbook.add_format({'border': 1, 'font_color': 'white', 'bg_color': '#551A8B'}))
			for ein_val2 in data_history_totals.values():
				wrksheet.write(row, col, "{:.1f}".format(ein_val2), wrkbook.add_format(
					{'border': 1, 'font_color': 'white', 'bg_color': '#551A8B', 'align': 'center', 'valign': 'vcenter'}))
				col += 1
			row += 1
			col = beginning_col
	#WRITE ZONE GOALS
			row += 2
			wrksheet.merge_range(f'F{row}:M{row}', 'MONTHLY BAPTISMAL GOALS', wrkbook.add_format({'align': 'center', 'valign': 'vcenter','font_size': 18}))
			#, 'font_color': '#FFFFFF', 'bg_color': '#000000'
			format = wrkbook.add_format({'font_size': 18, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
			zone_names_abr = {
				'NORTH':'N',
				'SOUTH':'S',
				'EAST':'E',
				'XINZHU':'XZ',
				'CENTRAL':'C',
				'TAOYUAN':'T',
				'WEST':'W',
				'HUALIAN':'H'
			}
			zone_goals = [month for month in csv_to_list(os.path.abspath('.\\data\\zone_goals')) if month['Report Month'] == str(date.month)]
			if zone_goals:
				zone_goals = zone_goals[0]
			else:
				zone_goals = {k:0 for k in csv_to_list(os.path.abspath('.\\data\\zone_goals')[0].keys())}
			col+=1
			mish_data_history = [dataline for dataline in all_data if change_to_datetime_obj(
				dataline['Report Date']) >= change_to_datetime_obj(f'{date.month}/{1}/{date.year}') and change_to_datetime_obj(dataline['Report Date']) <= change_to_datetime_obj(f'{date.month}/{num_days}/{date.year}')]
			for z_name in zone_names:
				wrksheet.write(row, col, zone_names_abr.get(z_name,z_name), format)
				zone_bc_actual = sum([int(data_piece['BC']) for data_piece in mish_data_history if data_piece['Zone']==z_name])
				wrksheet.write(row+1, col, zone_bc_actual, format)
				wrksheet.write(row+2, col, zone_goals.get(z_name,0), format)
				col+=1
			wrksheet.write(row, col, 'TOT.', format)
			zone_bc_actual_total = sum([int(data_piece['BC']) for data_piece in mish_data_history])
			wrksheet.write(row+1, col, zone_bc_actual_total, format)
			zone_goals_total = sum([int(goal) for goal in zone_goals.values()])
			wrksheet.write(row+2, col, zone_goals_total, format)
			col = beginning_col
			row+=1
			format = wrkbook.add_format({'font_size': 18, 'border': 1, 'align': 'center', 'valign': 'vcenter','bg_color':'#dddddd','font_color':'white'})
			wrksheet.write(row, col, 'MTD', format)
			row+=1
			format = wrkbook.add_format({'font_size': 18, 'border': 1, 'align': 'center', 'valign': 'vcenter','bg_color':'#aaaaaa','font_color':'white'})
			wrksheet.write(row, col, 'GOAL', format)
			# GRAPHS
			wrksheet = wrkbook.add_worksheet('GRAPHS')
			mish_data_history = [dataline for dataline in all_data if change_to_datetime_obj(
				dataline['Report Date']) >= (date-relativedelta(weeks=11)) and change_to_datetime_obj(dataline['Report Date']) <= date]
			weeks = sort_list({find_dates['Report Date']
							  for find_dates in mish_data_history}, reverse=False)
			# wrksheet.write(row, col, text, format)
			# wrksheet.merge_range('A1:B1',text, format)
	#GRAPHS
			date_list = list(map(prettier_date_format, weeks))
			multi = len({r_t_names['Area'] for r_t_names in [find_mish for find_mish in mish_data_history
				if change_to_datetime_obj(find_mish['Report Date']) == change_to_datetime_obj(weeks[-1])]})
			new_group = []
			for week in weeks:
				sum_week = [sum_it for sum_it in mish_data_history if sum_it['Report Date'] == week]
				sum_week_headers = sum_week[0].keys()
				total = {senor_key2: sum([int(a_dict2[senor_key2]) for a_dict2 in sum_week])
						 for senor_key2 in sum_week_headers if senor_key2 not in keys}
				total['Report Date'] = week
				new_group.append(total)
			data_group = sort_history_list(new_group)
			wrksheet.merge_range('A1:Q1', '{:%B %d, %Y} Mission Report'.format(report_title), wrkbook.add_format(
				{'font_size': 24, 'font_color': '#000000', 'bg_color': '#FFFFFF'}))
			wrksheet.merge_range('A2:Q2', 'Key Indicators 主要指標', wrkbook.add_format(
				{'font_size': 18, 'font_color': '#FFFFFF', 'bg_color': '#000000'}))
			# Baptismal pictures and Titles
			# BC
			wrksheet.insert_image('A4', os.path.abspath('.\\reportImages\\bap_pic.jpg'))
			wrksheet.merge_range('C4:F7', 'Baptized and Confirmed 已經接受洗禮和證實的朋友', wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['BC']) for a_dict in data_group]
			wrksheet.insert_image('A9', draw_graph(date_list, indicator_history, 'BC', multi))
			# SM
			wrksheet.insert_image('A25', os.path.abspath('.\\reportImages\\church_pic.jpg'))
			wrksheet.merge_range('C25:F28', 'Sacrament Meeting 出席聖餐聚會的朋友', wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['SM']) for a_dict in data_group]
			wrksheet.insert_image('A30', draw_graph(
				date_list, indicator_history, 'SM', multi))
			# BD
			wrksheet.insert_image('J4', os.path.abspath('.\\reportImages\\date_pic.jpg'))
			wrksheet.merge_range('L4:O7', 'Baptismal Dates 訂下洗禮日期的朋友', wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['BD']) for a_dict in data_group]
			wrksheet.insert_image('J9', draw_graph(date_list, indicator_history, 'BD', multi))
			# NF
			wrksheet.insert_image('J25', os.path.abspath('.\\reportImages\\nf_pic.jpg'))
			wrksheet.merge_range('L25:O28', 'New Friends 新朋友', wrkbook.add_format(
				{'font_size': 14, 'font_color': '#000000', 'bg_color': '#FFFFFF', 'text_wrap': True}))
			indicator_history = [int(a_dict['NF']) for a_dict in data_group]
			wrksheet.insert_image('J30', draw_graph(
				date_list, indicator_history, 'NF', multi))
			# CONVERT BAPTISM VISUALS
			wrksheet.merge_range('A50:Q51', 'Convert Baptisms 歸信者洗禮', wrkbook.add_format(
				{'font_size': 18, 'font_color': '#FFFFFF', 'bg_color': '#000000'}))
			years_to_graph = [
				change_to_datetime_obj(all_data[-1]['Report Date']).year,
				(change_to_datetime_obj(all_data[-1]['Report Date'])-relativedelta(years=1)).year,
				(change_to_datetime_obj(all_data[-1]['Report Date'])-relativedelta(years=2)).year
			]
			years_data = {}
			for year in years_to_graph:
				years_data[year] = []
				date_list = []
				for month in range(1, 13):
					last_sunday = max(week[-1] for week in calendar.monthcalendar(year, month))
					date_list.append(change_to_datetime_obj(f'{month}/{last_sunday}/{year}'))
					_, num_days = calendar.monthrange(year, month)
					month_data = [report for report in all_data if change_to_datetime_obj(report['Report Date']) > change_to_datetime_obj(
						f'{month}/{1}/{year}') and change_to_datetime_obj(report['Report Date']) <= change_to_datetime_obj(f'{month}/{num_days}/{year}')]
					month_bap_sum = sum([int(all_data['BC']) for all_data in month_data])
					# if month_bap_sum:
					years_data[year].append(month_bap_sum)
			month_list = ['{:%b}'.format(date) for date in date_list]
			wrksheet.insert_image('A54', draw_three_year_graph(month_list, years_data))
			area_baptismal_sources = [dataline for dataline in baptismal_sources if change_to_datetime_obj(dataline['Report Date']) >= (change_to_datetime_obj(baptismal_sources[-1]['Report Date'])-relativedelta(years=1))]
			wrksheet.insert_image('I54', baptismal_source_pie(area_baptismal_sources))
			wrkbook.close()
			#
			wb = xlApp.Workbooks.Open(path)
			count = 2
			for sheet in wb.Sheets:
				sheet.Visible = 1
				sheet.PageSetup.Zoom = False
				sheet.PageSetup.FitToPagesTall = 1
				sheet.PageSetup.FitToPagesWide = 1
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
				# print_area = f'A1:{col_letter(col)}{row}'
				# print(print_area)
				# sheet.PageSetup.PrintArea = print_area
				count -= 1
			wb.Worksheets(["REPORT", "GRAPHS"]).Select()
			pdf_path = path.split('.')
			pdf_path = pdf_path[0]+'.pdf'
			pdf_path = pdf_path.replace('Worksheets', 'PDFs')
			wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
			wb.Close()
			wb = None
			merge_list.append(pdf_path)
		else:
			wrkbook.close()
		merger(
			os.path.abspath(f'.\\PDFs\\masterReports\\{year_of_this_report} Master Summary {datetime.date.today().strftime("%m-%d-%Y")}.pdf'), merge_list)
xlApp.Quit
xlApp = None
