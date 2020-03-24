#!/usr/bin/env python3

"""
Purpose:
  Reads the given csv file and formats the contents and saves a
  copy for use within the excel spreadsheet.
"""

import os
import datetime

from handler import Excel, Analysis


def create_folder(name, return_path=False):
    """ Creates folder with a given name and optional return of new path. """
    path = os.getcwd()
    folder = create_path(path, name)
    if not os.path.exists(folder):
        os.mkdir(folder)
    if return_path:
        return folder


def create_path(path, file):
    return "{}\{}".format(path, file)


def current_time():
    """ Returns the local time in ISO format. Without microseconds. """
    return datetime.datetime.now().replace(microsecond=0).isoformat().replace(':', "'")


if __name__ == '__main__':

    # create directories for files if not done already
    path_original_files = create_folder("Original Files", return_path=True)
    path_formatted_files = create_folder("Formatted Files", return_path=True)

    try:
        for file_name in os.listdir(path_original_files):
            # open file inside the folder
            file_path = create_path(path_original_files, file_name)

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

        pass
    except Exception as e:
        print(e)

    pass
