#Seong Ma

import os, openpyxl, csv


def get_directory():
    while True:
        directory = input("Input directory: ")
        if os.path.isdir(directory):
            return directory
        print('INVALID PATH\n')
        
def create_csv(csv_file_to_write, excel_sheet):
    for row in excel_sheet.rows:
        pass

def convert_to_csv(excel_filename):
    excel_workbook = openpyxl.load_workbook(excel_filename)
    for sheet in excel_workbook.get_sheet_names():
        current_sheet_to_copy = excel_workbook.get_sheet_by_name(sheet)
        with open(_new_csv_filename(excel_filename, sheet), 'w', newline = '') as csv_file:
            pass

def search_directory(directory):
    for dir_name, subdir_list, file_list in os.walk(directory):
        for file in file_list:
            if _is_excel_file(file):
                absolute_directory = os.path.abspath(dir_name)
                absolute_filename  = os.path.join(absolute_directory, file)
                os.chdir(absolute_directory)
                convert_to_csv(file)
                
                
def _is_excel_file(filename):
    return filename.endswith('.xlsx')

def _new_csv_filename(excel_filename, sheet_name):
    return excel_filename[:-5] + '_' + sheet_name + '.csv'