#!/usr/bin/env python3

"""
Purpose:
  Reads the given csv file and formats the contents and saves a
  copy for use within the excel spreadsheet.
"""

import os
import datetime

import csv


def create_folder(name, return_path=False):
    """ Creates folder with a given name and optional return of new path. """
    path = os.getcwd()
    folder = f"{path}\{name}"
    if not os.path.exists(folder):
        os.mkdir(folder)
    if return_path:
        return folder


def current_time():
    """ Returns the local time in ISO format. Without microseconds. """
    return datetime.datetime.now().replace(microsecond=0).isoformat().replace(':', "'")


def sort_by_rate(entries, increasing_order=False):
    entries.sort(key=lambda x: x[len(x)-1],
                 reverse=(increasing_order == False))
    return entries


def get_column_titles(file):
    column_header = str(file.readline()).split(',')

    # read first line (header) and store the indexes for "Employee Name", "Units", "UPH"
    columns = {
        'employee': column_header.index("Employee Name"),
        'units': column_header.index("Units"),
        'uph': column_header.index("UPH")
    }
    return columns


if __name__ == '__main__':

    # create directories for files
    path_original_files = create_folder("Original Files", return_path=True)
    path_formatted_files = create_folder("Formatted Files", return_path=True)

    try:
        for file_name in os.listdir(path_original_files):
            # open file inside the folder
            file_path = f"{path_original_files}\{file_name}"

            with open(file_path, 'r') as f:
                columns = get_column_titles(f)
                csv_file = csv.reader(f)

                data_entries = []

                # loop through file line-by-line
                for row in csv_file:
                    # if "UPH" != 0:
                    if row[columns['uph']] != '0':

                        # store data found in the stored index numbers obtained earlier
                        employee = row[columns['employee']]
                        uph = float(row[columns['uph']])
                        units = int(row[columns['units']])

                        # calculate the stow rate (UPH/60) to one decimal
                        stow_rate = float('{:.1f}'.format(uph / 60))

                        # append the stow rate to the data obtained
                        data_entries.append([employee, units, stow_rate])

            # csv name shoud be date and time at generation
            generated_file_name = f"{current_time()}_stow rates"
            edited_file_path = f"{path_formatted_files}\{generated_file_name}"

            # save new csv inside copy folder
            with open(edited_file_path, 'w', newline='') as edited_csv:
                csv.writer(edited_csv).writerows(sort_by_rate(data_entries))
        pass
    except Exception as e:
        print(e)

    pass
