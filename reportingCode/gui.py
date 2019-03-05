import PySimpleGUI as sg
# from BlackBoxReporter import report
import report
import os
import json
# MAIN
menu_def = [
    ['Settings', ['Properties', 'Indicator Standards', 'Indicators', '---', 'Clear Reports']],
    ['Help', 'About...']
]
report_time = report.Abathur.report_timing.strftime('%m/%d/%Y')
layout = [
    [sg.Menu(menu_def, key='Menu')],
    [sg.Checkbox('Area', size=(10, 1), key='Area', default=False), sg.InputText(
        f'{report_time}', size=(10, 1), key='Date'), sg.Button('Change Date')],
    [sg.Checkbox('District', size=(10, 1), key='District', default=False), sg.Checkbox(
        'Generate Master Reports', size=(20, 1), key='Master', default=False)],
    [sg.Checkbox('Zone', size=(10, 1), key='Zone', default=False)],
    [sg.Checkbox('Stake', size=(10, 1), key='Stake', default=True)],
    [sg.Checkbox('Ward', size=(10, 1), key='Ward', default=True)],
    [sg.Button('Make Reports'), sg.Button('Exit')]
]
# ,sg.InputText(size=(20,1), key='FileName'),sg.FolderBrowse()
# size=(400, 200) \/\/
window = sg.Window('Report Maker').Layout(layout)
# sg.Print("Program Start!", location=(525, 675), size=(70, 15))
ind_win = False  # control var for Indicator Options window
prop_win = False  # control var for Properties window
del_win = False  # control var for Clear Reports window
send_reports = False  # controls whether reports are emailed out or not
# Designated report types, these cannot be changed
valid_report_types = ['Area', 'District', 'Zone', 'Stake']


while True:
    event, values = window.Read()
    # print(f"event: {event}, values: {values}")
    # MAIN WINDOW
    reports_to_make = [key for key, val in values.items() if val ==
                       True and key in valid_report_types]
    if event == 'Change Date':
        # Settings data loaded
        with open('settings.json') as json_file:
            settingsData = json.load(json_file)

        report.Abathur.report_timing = report.turn_to_datetime(values['Date'])
        report.Zeratul.report_timing = report.turn_to_datetime(values['Date'])
        settingsData[1]['Date'] = values['Date']

        #Dump to file
        with open('settings.json', 'w') as outfile:
            json.dump(settingsData, outfile)

    if event == 'Make Reports':
        reports_maker = report.Abathur('.\\data\\key_indicator_reports', reports_to_make)
        hermes = report.Zeratul()
        try:
            print(f'REPORTING DAY:\t{report.Abathur.report_timing}')
            for report_kind, reports_list in reports_maker.reports_to_make.items():
                merge_list = []
                if send_reports:
                    hermes.make_report_folder(report_kind)
                print(report_kind)
                count = 1
                check_path = os.listdir(
                    f'.\\PDFs\\{report_kind}')
                check_path = [fyle_name.split(".")[0] for fyle_name in check_path]
                for report_name in reports_list:
                    f_path = f'.\\Worksheets\\{report_kind}\\{report_name}.xlsx'
                    pdf_path = f_path.replace('Worksheets', 'PDFs').replace('xlsx', 'pdf')
                    if report_name not in check_path:
                        print(report_name+f'\t{count}/{(len(reports_list)-len(check_path))}')
                        sg.OneLineProgressMeter(f'{report_kind} Report Generation', count, (len(
                            reports_list)-len(check_path)), 'key', 'Generation Progress:', orientation='h')
                        count += 1
                        reports_maker.init_workbook(f_path)
                        reports_maker.set_headers(report_kind, report_name)
                        reports_maker.make_table_current(report_kind, report_name)
                        reports_maker.make_history_table(report_kind, report_name)
                        reports_maker.make_graphs(report_kind, report_name)
                        reports_maker.export_pdf(report_kind)
                    else:
                        print(f'{report_name} report already created.')
                    if values['Master']:
                        merge_list.append(pdf_path)
                    if send_reports:
                        hermes.send_report(report_name, report_kind, pdf_path)
                    # break
                reports_maker.make_summary_sheet(report_kind)
                if values['Master']:
                    report.merger(
                        f'.\\PDFs\\Master Reports\\{report_kind}_Master.pdf', merge_list)
        except Exception as ex:
            reports_maker.xlApp.Quit
            raise ex
    # CLEAR REPORTS WINDOW
    # if not del_win and event == 'Clear Reports':
    if event == 'Clear Reports':
        del_win = True
        layout2 = [[sg.Checkbox('Excel', size=(10, 1), key='Worksheets')],
                   [sg.Checkbox('PDFs', size=(10, 1), key='PDFs')],
                   [sg.Button('Clear Files')]]
        win2 = sg.Window('Clear Reports', location=(480, 500)).Layout(layout2)
    if del_win:
        ev2, vals2 = win2.Read()
        if ev2 is None or ev2 == 'Exit':
            del_win = False
            win2.Close()
        if ev2 == 'Clear Files':
            folders_to_wipe = [key for key, val in vals2.items() if val == True]
            for folder_type in folders_to_wipe:
                for folder in reports_to_make:
                    folder_path = f'.\\{folder_type}\\{folder}\\'
                    el_folder = os.listdir(folder_path)
                    for file in el_folder:
                        os.remove(folder_path+file)
            sg.PopupOK('Files Recycled!')
    # PROPERTIES WINDOW
    # if not prop_win and event == 'Properties':
    if event == 'Properties':
        prop_win = True
        layout3 = [
            [sg.Frame(layout=[[sg.Checkbox('Email Reports', size=(10, 1), key='Email Reports', default=True)]], title='Emailing',
                      title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='These settings will effect how reports will be emailed')],
            [sg.Frame(layout=[[sg.Checkbox('Print w/ Color', size=(10, 1), key='Color', default=True)]],
                      title='Printing', title_color='blue', relief=sg.RELIEF_SUNKEN, tooltip='Change printing settings')],
            [sg.Button('Save Changes')]
        ]
        win3 = sg.Window('Properties', location=(750, 250)).Layout(layout3)
    if event == 'Indicators':
        changeIndicatorLayout = [

        ]

        win4 = sg.Window('Change Indicators', location=(750,350)).Layout(changeIndicatorLayout)

    if prop_win:
        ev3, vals3 = win3.Read()
        if ev3 is None or ev3 == 'Exit':
            prop_win = False
            win3.Close()
        if ev3 == 'Save Changes':
            # Activate or deactivate Hermes, normally actived
            send_reports = vals3['Email Reports']
            report.Abathur.print_color = vals3['Color']
            sg.PopupOK('Changes Saved!')
    # INDICATOR CONTROL
    # if not ind_win and event == 'Indicator Standards':
    if event == 'Indicator Standards':
        ind_win = True
        layout4 = [
            [sg.Frame(layout=[
                [sg.Text('Baptismal Dates: '), sg.InputText(
                    f"{report.Abathur.indicator_standards['BD']}", size=(3, 1), key='BD')],
                [sg.Text('Sacrament Attendance: '), sg.InputText(
                    f"{report.Abathur.indicator_standards['SM']}", size=(3, 1), key='SM')],
                [sg.Text('New Friends: '), sg.InputText(
                    f"{report.Abathur.indicator_standards['NF']}", size=(3, 1), key='NF')],
            ], title='Current Standards', title_color='blue', relief=sg.RELIEF_SUNKEN, tooltip='Changes are not saved after closing.')],
            [sg.Button('Save Changes')]
        ]
        win4 = sg.Window('Indicator Standards', location=(1050, 450)).Layout(layout4)
    if ind_win:
        ev4, vals4 = win4.Read()
        if ev4 is None or ev4 == 'Exit':
            ind_win = False
            win4.Close()
        if ev4 == 'Save Changes':
            report.Abathur.indicator_standards['BD'] = vals4['BD']
            report.Abathur.indicator_standards['SM'] = vals4['SM']
            report.Abathur.indicator_standards['NF'] = vals4['NF']
            with open('settings.json', 'w') as outfile:
                json.dump({'BD': vals4['BD'], 'SM': vals4['SM'], 'NF': vals4['NF']}, outfile)
            sg.PopupOK('Changes Saved!')
    if event is None or event == 'Exit':
        break
window.Close()
