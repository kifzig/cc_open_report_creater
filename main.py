import pandas as pd
import PySimpleGUI as sg
import xlsxwriter
import os
from datetime import date
from specialty import Specialty


def is_file_available(path):
    try:
        file = os.listdir(path)[0]
    except IndexError:
        return False
    else:
        return True


def make_recruiter_folders(path, specialty):
    try:
        for folder in specialty.recruiter_folders:
            new_path = os.path.join(path, folder)
            if not os.path.isdir(new_path):
                os.mkdir(new_path)
                make_subfolders(new_path)
        print("Empty folder created in " + path)
        print("Place your downloaded files from Contact Contact in these folders.")
        return path
    except OSError as error:
        print(error)
        return path



def make_subfolders(path):
    try:
        # parent_directory = path
        # if not os.path.isdir(parent_directory):
        #     os.mkdir(path=parent_directory)
        for folder in specialty.subfolders:
            print(folder)
            print(path)
            new_path = os.path.join(path, folder)
            print(new_path)
            if not os.path.isdir(new_path):
                os.mkdir(new_path)
        return path
    except OSError as error:
        print(error)
        return path


def clear_subfolders(chosen_path):
    for folder in specialty.subfolders:
        files = os.listdir(chosen_path + f'/{folder}')
        if not files:
            continue
        else:
            file = chosen_path + f'/{folder}/{files[0]}'
            if os.path.isfile(file):
                os.remove(chosen_path + f'/{folder}/{files[0]}')


def add_contact_owner_column(num_of_rows, recruiter):
    contact_owner_list = []
    for i in range(num_of_rows):
        contact_owner_list.append(recruiter)
    contact_owner_column = pd.Series(contact_owner_list)
    return(contact_owner_column)


def contact_owner_col_move(df_sheet, recruiter):
    try:
        contact_owner_column = df_sheet.pop('Contact Owner')
    except KeyError:
        # if no Contact Owner column is in the dataframe
        num_of_df_rows = len(df_sheet.index)
        contact_owner = specialty.recruiter_dictionary.get(recruiter)
        contact_owner_column = add_contact_owner_column(num_of_df_rows, contact_owner)
    finally:
        df_sheet.insert(0, "Contact Owner", contact_owner_column)
        df_sheet = df_sheet.sort_values(by=['Email address'])
        return df_sheet


def df_drop_columns(report_df, list_of_unwanted_columns):
    for item in list_of_unwanted_columns:
        try:
            report_df = report_df.drop(columns=item)
        except KeyError:
            #print(f'{item} not found in df.')
            continue
    return report_df


def create_df(file_path, recruiter, df_type):
    if is_file_available(file_path):
        file = os.listdir(file_path)[0]
        df_new = pd.read_csv(file_path + file)
        df_new = df_drop_columns(df_new, unwanted_columns)
        df_new = contact_owner_col_move(df_new, recruiter)
        return df_new
    else:
        if df_type == 'opens':
            message = {'Details': ['You', 'need', 'to', 'download', 'files', 'from',
                                   'Constant Contact', 'and place', 'them', 'in', 'folders',
                                   'open, clicks, bounces, unsubscribed']}
            df_new = pd.DataFrame.from_dict(message)
            return df_new
        else:
            message = {'Details': ['There', 'are', 'no', f'{df_type}', 'for', 'this', 'blast.']}
            df_new = pd.DataFrame.from_dict(message)
            return df_new


def create_single_stat_report(directory, recruiter_name='none'):

    #Getting opens path and create opens dataframe for Excel sheet import
    opens_path = directory + "/1opens/"
    df_opens = create_df(opens_path, recruiter_name, df_type="opens")

    #Getting clicks path and create clicks dataframe for Excel sheet import
    clicks_path = directory + "/2clicks/"
    df_clicks = create_df(clicks_path, recruiter_name, df_type="clicks")

    #Getting bounces path and creating bounces dataframe for import into Excel
    bounces_path = directory + "/3bounces/"
    df_bounces = create_df(bounces_path, recruiter_name, df_type="bounces")

    #Getting unsubscribed path and creating unsubscribed dataframe for import into Excel
    if specialty.name == "Neurosurgery":
        unsubscribed_path = directory + "/4unsubscribed/"
        df_unsubscribed = create_df(unsubscribed_path, recruiter_name, df_type="unsubscribeds")

    #Creating the Excel Report
    report_prefix = values['-REPORT PREFIX-']
    export_file_path = directory + "/" + report_prefix + " Blast Report.xlsx"
    recruiter_name = " " + recruiter_name.capitalize()
    if recruiter_name != ' None':
        directory = os.path.dirname(directory)
        report_prefix = values['-NS REPORT PREFIX-']
        export_file_path = directory + "/" + report_prefix + recruiter_name + " Blast Report.xlsx"
    writer = pd.ExcelWriter(path=export_file_path, engine='xlsxwriter')
    df_opens.to_excel(writer, sheet_name='Opens', index=False)
    df_clicks.to_excel(writer, sheet_name='Clicks', index=False)
    df_bounces.to_excel(writer, sheet_name='Bounces', index=False)
    if specialty.name == "Neurosurgery":
        df_unsubscribed.to_excel(writer, sheet_name='Unsubscribed', index=False)


    #Create worksheet objects and set column width and formatting options
    workbook = writer.book
    border_format = workbook.add_format()
    border_format.set_border()

    worksheet_opens = writer.sheets['Opens']
    worksheet_opens.set_column(0, 40, 35, border_format)
    worksheet_clicks = writer.sheets['Clicks']
    worksheet_clicks.set_column(0, 40, 35, border_format)
    worksheet_bounces = writer.sheets['Bounces']
    worksheet_bounces.set_column(0, 40, 35, border_format)
    if specialty.name == "Neurosurgery":
        worksheet_unsubscribed = writer.sheets['Unsubscribed']
        worksheet_unsubscribed.set_column(0, 40, 35, border_format)

    writer.save()
    if recruiter_name != ' None':
        print(export_file_path + " created.")

def create_all_reports(directory, specialty):
    for recruiter in specialty.recruiter_folders:
        recruiter_directory = directory + f'/{recruiter}'
        create_single_stat_report(recruiter_directory, recruiter)


current_working_directory = os.getcwd()
stats_path = os.getcwd()

#Extra Fields that can be dropped from the Excel Sheet
unwanted_columns = ['Email status', 'Email permission status', 'Email update source', 'Street address line 1 - Home',
                        'Email Lists', 'Created At', 'Updated At', 'Confirmed Opt-Out Date', 'Phone - fax',
                        'City - Home', 'Phone - work', 'Zip/Postal Code - Home', 'Country - Home', 'State/Province - Other',
                        'Custom Field 1', 'Job title', 'Confirmed Opt-Out Source', 'Confirmed Opt-Out Reason', 'Street address line 1 - Work',
                        'City - Work', 'State/Province - Work', 'Zip/Postal Code - Work', 'Country - Work', 'off lim', 'Source Name', 'Phone - home']

todays_date = date.today()
current_month = todays_date.strftime("%Y-%m")

#Creating GUI with PySimpleGUI
sg.theme('DarkTeal9')
layout = [
    [sg.Text("OPEN REPORT HELPER", font=("Helvetica", 25), text_color="white")],
    [
        sg.Text("CHOOSE SPECIALTY: ", font=("Helvetica", 18)),
        sg.Radio("Neurosurgery", "group1", default=True, key='-SUB-NS-'),
        sg.Radio("Urology", "group1", default=False, key='-SUB-UL-'),
        sg.Radio("Neurology", "group1", default=False, key='-SUB-NL-'),
        sg.Radio("GI", "group1", default=False, key='-SUB-GI-')
    ],
    [sg.Text("Choose from any of the following options below", font=("Helvetica", 13), text_color="white")],
    [sg.Text('CREATE FOLDERS: 1opens, 2clicks, 3bounces for all recruiters')],
    [sg.Text('Choose destination:'), sg.Input(f'{current_working_directory}', key='-NEW-FOLDER-'), sg.FolderBrowse(target='-NEW-FOLDER-')],
    [sg.Button('Create Folders')],
    [sg.Text('CREATE SINGLE REPORT FOR 1 RECRUITER -- choose directory containing subfolders 1opens, 2clicks, 3bounces')],
    [sg.Text('Report Location:'), sg.Input(key='-FOLDER-'), sg.FolderBrowse(target='-FOLDER-')],
    [sg.Text('Report Name Prefix: '), sg.Input(f'{current_month}', key='-REPORT PREFIX-'), sg.Text('Blast Report.xlsx')],
    [sg.Button('Create Single Report')],
    [sg.Text('CREATE REPORTS FOR ALL RECRUITERS -- choose directory containing recruiter folders filled with cc folder/files')],
    [sg.Text('Select Folder'), sg.Input(key='-NS FOLDER-'), sg.FolderBrowse(target='-NS FOLDER-')],
    [sg.Text('Report Name Prefix: '), sg.Input(f'{current_month}', key='-NS REPORT PREFIX-'), sg.Text('[Recruiter Name] Blast Report.xlsx')],
    [sg.Button('Create All Reports')],
    #[sg.Text('Clear Subfolders of CSVs'), sg.Input(key='-CLEAR-FOLDER-'), sg.FolderBrowse(target='-CLEAR-FOLDER-')],
    #[sg.Button('Clear_Subfolders')],
    [sg.Button('Exit')]
]

window = sg.Window('BLAST REPORT HELPER', layout)

#Below loop is for the GUI and it is continuous until the GI is closed or 'Exit' is clicked

while True or event == 'OK':
    event, values = window.read()

    #Check radio buttons for specialty selected
    spec = ""
    if values['-SUB-NS-'] == True:
        spec = "Neurosurgery"
    elif values['-SUB-NL-'] == True:
        spec = "Neurology"
    elif values['-SUB-UL-'] == True:
        spec = "Urology"
    elif values['-SUB-GI-'] == True:
        spec = "Gastro"

    specialty = Specialty(spec)


    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    elif event == 'Create Folders':
        stats_path = values['-NEW-FOLDER-']
        path = make_recruiter_folders(stats_path, specialty)

    elif event == 'Create All Reports':
        subspecialty_folder = values['-NS FOLDER-']
        create_all_reports(subspecialty_folder, specialty)


    # elif event == 'Clear_Subfolders':
    #     chosen_path = values['-CLEAR-FOLDER-']
    #     print(chosen_path)
    #     clear_subfolders(chosen_path)

    elif event == 'Create Single Report':
        #Getting stat folder with folders opens, clicks, bounces and unsubscribed from the GUI
        stats_path = values['-FOLDER-']
        print(stats_path)
        create_single_stat_report(directory=stats_path)

window.close()

