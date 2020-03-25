#!/usr/bin/env python3

"""
Purpose:
  Reads the given csv file and formats the contents and saves a
  copy for use within the excel spreadsheet.
"""

import os
import time
import datetime

from handler import Excel, Analysis


def create_folder(name):
    """ Creates folder with a given name and optional return of new path. """
    if not os.path.exists(name):
        os.makekdirs(name)


def create_path(path, file):
    return "{}\{}".format(path, file)


def current_time():
    """ Returns the local time in ISO format. Without microseconds. """
    return datetime.datetime.now().replace(microsecond=0).isoformat().replace(':', "'")


if __name__ == '__main__':

    # create directories for files if not done already
    cwd = os.getcwd()
    path_original_files = create_path(cwd, "Original Files")
    path_formatted_files = create_path(cwd, "Formatted Files")

    create_folder(path_original_files)
    create_folder(path_formatted_files)

    original_files = os.listdir(path_original_files)
    
    if not original_files:
        print("There is are no files to process in 'Original Files'")
        input("Press return to exit...")
        exit()

    for file_name in os.listdir(path_original_files):
        # open file inside the folder
        file_path = create_path(path_original_files, file_name)

        try:
            with open(file_path, 'r') as f:
                csv_file = Analysis(f)
                csv_file.analyze()

                # csv name shoud be date and time at generation
                generated_file_name = "{}_stow rates".format(current_time())
                edited_file_path = create_path(
                    path_formatted_files, generated_file_name)
                csv_file.save(edited_file_path)

            # convert csv to excel
            Excel.csv_to_excel(edited_file_path)
            Excel.append_data_to_sheet(Analysis.data_entries)
        except FileNotFoundError() as fnf:
            print("File not found:\n", fnf)
        except Exception() as e:
            print(e)

    #os.remove(path_original_files)
    #print("Conversion completed!")
    time.sleep(1.5)
