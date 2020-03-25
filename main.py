#!/usr/bin/env python3

"""
Purpose:
  Reads the given csv file and formats the contents and saves a
  copy for use within the excel spreadsheet.
"""


import os
import time
import datetime

import glob
from handler import Excel, Analysis


def create_folder(name):
    """ Creates folder with a given name and optional return of new path. """
    if not os.path.exists(name):
        os.mkdir(name)


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

    def clear_console(): return os.system('cls')

    if not original_files:
        print("There is are no files to process in 'Original Files'")
        input("Press return to exit...")
        exit()

    chosen_file = original_files[0]

    if len(original_files) > 1:
        chosen_file = None

        while chosen_file is None:
            clear_console()
            print("There are more than 1 file inside the folder!")
            print("\nPlease choose one of the following:")

            for number, f in enumerate(os.listdir(path_original_files)):
                print("\t[{}] {}".format(number, f))
            x = int(input("\nEnter a number: "))

            if x not in range(len(original_files)):
                clear_console()
                print("Invalid Selection!")
                time.sleep(1)
                continue

            chosen_file = original_files[x]
            print(chosen_file)

        clear_console()

    for count, file_name in enumerate(os.listdir(path_original_files)):

        if file_name != chosen_file:
            continue

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

            Excel.append_data_to_template(Analysis.data_entries, save_as=edited_file_path)
            break

        except FileNotFoundError() as fnf:
            print("File not found:\n", fnf)
        except Exception() as e:
            print(e)

    # os.remove(path_original_files)
    #print("Conversion completed!")
    print("\nConversion Complete!")
    print("Opening Excel Document...")
    time.sleep(2)

    os.startfile(edited_file_path + '.xlsx')
