#!/usr/bin/env python3

"""
Purpose:
  Reads the given csv file and formats the contents and saves a
  copy for use within the excel spreadsheet.
"""


import os
import sys
import time
import datetime

import tkinter

from datetime import datetime
from tkinter import filedialog
from handler import Excel, Analysis


def create_folder(name):
    """ Creates folder with a given name and optional return of new path. """
    if not os.path.exists(name):
        os.mkdir(name)


def create_path(path, file):
    return "{}\{}".format(path, file)


def current_time():
    """ Returns the local time in ISO format. Without microseconds. """
    return datetime.now().strftime("%m-%d-%Y %H'%M'%S")
    

def choose_file():
    tk = tkinter.Tk()
    # hide tkinter window when opening file browser
    tk.withdraw()
    filename = filedialog.askopenfilename()
    return filename


if __name__ == '__main__':
    reports_folder_path = create_path(os.getcwd(), "Reports")
    create_folder(reports_folder_path)

    try:

        data_file = choose_file()
        
        # no file was selected; cancel button most likely hit
        if not data_file:
            sys.exit()

        # open file inside the folder
        with open(data_file, 'r') as f:
            stow_rate_data = Analysis(f)
            stow_rate_data.analyze()

            # csv name shoud be date and time at generation
            creation_date = current_time()
            report_name = "{}_stow rates".format(creation_date)
            report_path = (create_path(reports_folder_path, report_name))

            Excel.append_data_to_template(
                Analysis.data_entries, save_as=report_path)

        print("\nConversion Complete!")
        print("Opening Excel Document...")
        time.sleep(2)

        os.startfile(report_path + '.xlsx')

    except FileNotFoundError as fnfe:
        print(fnfe)
        input("Press return to exit...")
    except NameError as ne:
        print(ne)
        input("Press return to exit...")
    except Exception as e:
        print(e)
        input("Press return to exit...")
