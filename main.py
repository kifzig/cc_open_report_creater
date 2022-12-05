import pandas as pd
import PySimpleGUI as sg
import xlsxwriter
import os
from datetime import date


def is_file_available(path):
    try:
        file = os.listdir(path)[0]
    except IndexError:
        return False
    else:
        return True


def make_recruiter_folders(path):
    try:
        for folder in ns_recruiter_folders:
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
        for folder in subfolders:
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
    for folder in subfolders:
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
        contact_owner = ns_recruiter_dict.get(recruiter)
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
        df_new = df_drop_columns(df_new, unwanted_columns_opens)
        df_new = contact_owner_col_move(df_new, recruiter)
        return df_new
    else:
        if df_type == 'opens':
            message = {'Details': ['You', 'need', 'to', 'download', 'files', 'from',
                                   'Constant Contact', 'and place', 'them', 'in', 'folders',
                                   'open, clicks, bounces, unsubscribed']}
        else:
            message = {'Details': ['There', 'are', 'no', f'{df_type}', 'for', 'this', 'blast.']}
            df_new = pd.DataFrame.from_dict(message)
            return df_new


def create_opens_df(opens_path, recruiter):
    if is_file_available(opens_path):
        opens_file = os.listdir(opens_path)[0]
        df_opens = pd.read_csv(opens_path + opens_file)
        df_opens = df_drop_columns(df_opens, unwanted_columns_opens)
        df_opens = contact_owner_col_move(df_opens, recruiter)
        return df_opens
    else:
        message = {'Details': ['You', 'need', 'to', 'download', 'files', 'from',
                               'Constant Contact', 'and place', 'them', 'in', 'folders',
                               'open, clicks, bounces, unsubscribed']}
        df_opens = pd.DataFrame.from_dict(message)
        return df_opens


def create_clicks_df(clicks_path, recruiter):
    if is_file_available(clicks_path):
        clicks_file = os.listdir(clicks_path)[0]
        df_clicks = pd.read_csv(clicks_path + clicks_file)
        df_clicks = df_drop_columns(df_clicks, unwanted_columns_clicks)
        df_clicks = contact_owner_col_move(df_clicks, recruiter)
        return df_clicks
    else:
        message = {'Details': ['There', 'are', 'no', 'clicks', 'for', 'this', 'blast.']}
        df_clicks = pd.DataFrame.from_dict(message)
        return df_clicks


def create_bounces_df(bounces_path, recruiter):
    if is_file_available(bounces_path):
        bounces_file = os.listdir(bounces_path)[0]
        df_bounces = pd.read_csv(bounces_path + bounces_file)
        df_bounces = df_drop_columns(df_bounces, unwanted_columns_bounces)
        df_bounces = contact_owner_col_move(df_bounces, recruiter)
        return df_bounces
    else:
        message = {'Details': ['There', 'are', 'no', 'bounces', 'for', 'this', 'blast.']}
        df_bounces = pd.DataFrame.from_dict(message)
        return df_bounces


def create_unsubscribed_df(unsubscribed_path, recruiter):
    if is_file_available(unsubscribed_path):
        unsubscribed_file = os.listdir(unsubscribed_path)[0]
        df_unsubscribed = pd.read_csv(unsubscribed_path + unsubscribed_file)
        df_unsubscribed = df_drop_columns(df_unsubscribed, unwanted_columns_bounces)
        df_unsubscribed = contact_owner_col_move(df_unsubscribed, recruiter)
        return df_unsubscribed
    else:
        message = {'Details': ['There', 'are', 'no', 'unsubscribed', 'contacts', 'for', 'this', 'blast.']}
        df_unsubscribed = pd.DataFrame.from_dict(message)
        return df_unsubscribed


def create_single_stat_report(directory, recruiter_name='none'):

    #Getting opens path and create opens dataframe for Excel sheet import
    opens_path = directory + "/1opens/"
    df_opens = create_df(opens_path, recruiter_name, df_type="opens")

    #Getting clicks path and create clicks dataframe for Excel sheet import
    clicks_path = directory + "/2clicks/"
    #df_clicks = create_clicks_df(clicks_path, recruiter_name)
    df_clicks = create_df(clicks_path, recruiter_name, df_type="clicks")

    #Getting bounces path and creating bounces dataframe for import into Excel
    bounces_path = directory + "/3bounces/"
    df_bounces = create_df(bounces_path, recruiter_name, df_type="bounces")

    #Getting unsubscribed path and creating unsubscribed dataframe for import into Excel
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
    worksheet_unsubscribed = writer.sheets['Unsubscribed']
    worksheet_unsubscribed.set_column(0, 40, 35, border_format)

    writer.save()
    if recruiter_name != ' None':
        print(export_file_path + " created.")

def create_all_ns_reports(directory):
    for recruiter in ns_recruiter_folders:
        recruiter_directory = directory + f'/{recruiter}'
        create_single_stat_report(recruiter_directory, recruiter)


current_working_directory = os.getcwd()
ns_recruiter_folders = ['andrea', 'jonathan', 'rachel', 'nancy']
ns_recruiter_dict = {'andrea': 'Andrea Winslow', 'jonathan': 'Jonathan Haines', 'nancy': 'Nancy Cusick', 'rachel': 'Rachel Prero'}
subfolders = ['1opens', '2clicks', '3bounces', '4unsubscribed']
stats_path = os.getcwd()

#Extra Fields that can be dropped from the Excel Sheet
unwanted_columns_opens = ['Email status', 'Email permission status', 'Email update source', 'Street address line 1 - Home',
                        'Email Lists', 'Created At', 'Updated At', 'Confirmed Opt-Out Date', 'Phone - fax',
                        'City - Home', 'Phone - work', 'Zip/Postal Code - Home', 'Country - Home', 'State/Province - Other',
                        'Custom Field 1', 'Job title', 'Confirmed Opt-Out Source', 'Confirmed Opt-Out Reason', 'Street address line 1 - Work',
                        'City - Work', 'State/Province - Work', 'Zip/Postal Code - Work', 'Country - Work', 'off lim', 'Source Name', 'Phone - home']

unwanted_columns_clicks = ['Email status', 'Email Lists', 'Email permission status',
                           'Email update source', 'Created At', 'Updated At', 'Confirmed Opt-Out Date', 'Source Name', 'Job title']

unwanted_columns_bounces = ['Email status', 'Email permission status', 'Email update source',
                            'Source Name', 'Created At', 'Job title']

todays_date = date.today()
current_month = todays_date.strftime("%Y-%m")

#Creating GUI with PySimpleGUI

layout = [
    [sg.Text('CREATE FOLDERS: opens, clicks, bounces, unsubscribed')],
    [sg.Text('Choose directory'), sg.Input(f'{current_working_directory}', key='-NEW-FOLDER-'), sg.FolderBrowse(target='-NEW-FOLDER-')],
    [sg.Button('Create NS Folder')],
    [sg.Text('CREATE SINGLE REPORT -- choose directory containing subfolders opens, clicks, bounces, unsubscribed')],
    [sg.Text('Report Location:'), sg.Input(key='-FOLDER-'), sg.FolderBrowse(target='-FOLDER-')],
    [sg.Text('Report Name Prefix: '), sg.Input(f'2022-11 NS', key='-REPORT PREFIX-'), sg.Text('Blast Report.xlsx')],
    [sg.Button('Create Single Report')],
    [sg.Text('CREATE ALL REPORTS -- choose directory containing andrea, jonathan, nancy, rachel folders with populated folders')],
    [sg.Text('Select NS Folder'), sg.Input(key='-NS FOLDER-'), sg.FolderBrowse(target='-NS FOLDER-')],
    [sg.Text('Report Name Prefix: '), sg.Input(f'{current_month} NS', key='-NS REPORT PREFIX-'), sg.Text('[Recruiter Name] Blast Report.xlsx')],
    [sg.Button('Create All Reports')],
    #[sg.Text('Clear Subfolders of CSVs'), sg.Input(key='-CLEAR-FOLDER-'), sg.FolderBrowse(target='-CLEAR-FOLDER-')],
    #[sg.Button('Clear_Subfolders')],
    [sg.Button('Exit')]
]

window = sg.Window('BLAST REPORT HELPER', layout)

#Loop is for the GUI and it is continuous until the GI is closed or 'Exit' is clicked

while True or event == 'OK':
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    elif event == 'Create NS Folder':

        stats_path = values['-NEW-FOLDER-']
        path = make_recruiter_folders(stats_path)

    elif event == 'Create All Reports':
        ns_folder = values['-NS FOLDER-']
        create_all_ns_reports(ns_folder)


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

