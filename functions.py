#Seong Ma

import os, openpyxl, csv


def get_directory():
    while True:
        directory = input("Input directory: ")
        if os.path.isdir(directory):
            return directory
        print('INVALID PATH\n')
        
def read_excel():
    pass

def convert_to_csv(excel_file):
    pass

