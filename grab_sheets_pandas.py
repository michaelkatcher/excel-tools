"""
Name: Michael Katcher
Date: 07/31/2017
Desc: A script to solve the following problem:

    Background:     A folder with an .xlsm for each trading day
                    In each file is a data tab with the same name

    Requirement:    Collect all the tabs for a given date range
                    and save them in a single csv.

    Notes:          This version of grab_sheets uses pandas to filter
                    and group the data before it is saved.
                    
                    '~~~~' used in lieu of actual Strategy/Folder names
"""


from os import listdir
from datetime import datetime

import pandas as pd


def load_csv(save_path):
    """Return a DataFrame with existing csv data and a list of dates.

    Args:
        save_path (str):    The path of the existing csv data

    Returns:
        DataFrame:          A pandas DataFrame loaded with the csv data
        list of datetimes:  A list of unique dates from the DataFrame
    """

    existing_data = pd.read_csv(save_path)
    existing_data.dropna(inplace=True)

    existing_dates = [datetime.strptime(date_string, '%m/%d/%Y') for date_string in existing_data['Date'].unique()]

    return existing_data, existing_dates


def load_filenames(data_folder, filter=(lambda name: True)):
    """Return a list of file names in a given directory that pass a filter check

    Args:
        data_folder (str):              The folder that we're searching
        filter (callable, optional):    A callable that takes a string and returns a boolean of
                                        whether or not to include a filename.

    Returns:
        list of str:                    A list of file names (w/o full path) as strings.
    """

    return [name for name in listdir(data_folder) if filter(name)]


def date_from_file_name(file_name, return_datetime=False):
    """Return full date as a string or datetime from a date in a filename.

    Args:
        file_name (str): Formatted as: Trade E-mail Archive - yyyy.mm.dd.xlsm
        return_datetime (bool, optional): TRUE if return should be datetime
                                      FALSE if return should be string

    Returns:
        str or datetime: The return value. Date from filename as yyyy.mm.dd
    """
    if return_datetime:
        year = int(file_name[-15:-11])
        month = int(file_name[-10:-8])
        day = int(file_name[-7:-5])
        return datetime(year, month, day)
    else:
        return file_name[-15:-5]


def main():
    #####################################
    #####################################
    # Initial setup variables

    data_folder = 'N:\\Shares\\TA\\~~~~\\Trade Email\\'  # Folder containing the input files
    sheet_name = '~~~~'    # The sheet name we're looking for in each input file

    start_date = datetime(2017,7,1)  # The date of the first file you want to import
    end_date = datetime.now()   # The date of the last file you want to import

    filter_column = ''  # The column of data to filter
    filter_value = ''  # The value to filter

    group_column = ''  # The column to groupby
    sum_columns = []  # The columns to sum during groupby

    error_list = [] # Captures any errors in processing a given file

    save_path = 'H:\\~~~~\\Data\\{data_name}.csv'.format(data_name=sheet_name)  # The path of the output file

    overwrite_existing = False # Whether or not to skip files that have already been processed

    #####################################

    # Setup custom filter/group values for each sheet name
    if sheet_name == '~~~~':
        filter_column = 'Strategy'
        filter_value = '~~~~'

        group_column = 'SubStrategy'
        sum_columns = ['Day P&L', 'Qty']

    elif sheet_name == '~~~~':
        filter_column = 'Strategy'
        filter_value = '~~~~'

        group_column = 'SubStrategy'
        sum_columns = ['Quantity']

    else:
        print '{sheet_name} is not recognized as a valid sheet.'.format(sheet_name=sheet_name)

    #####################################
    #####################################

    # Load existing data if not overwriting
    if overwrite_existing:
        all_data = pd.DataFrame()
        existing_dates = []
    else:
        all_data, existing_dates = load_csv(save_path)

    # Setup the boolean function for whether to process a file or not
    def include_file(name): return name.endswith('.xlsm') and \
                                   (start_date <= date_from_file_name(name,True) <= end_date) and \
                                   date_from_file_name(name,True) not in existing_dates

    # Generate list of filenames
    file_names = load_filenames(data_folder, include_file)

    # If there are no new files to process, exit
    if not file_names:
        print 'All Data Previously Processed.'
        print 'Last Date: {date}'.format(date=max(existing_dates).strftime('%m/%d/%Y'))
        exit()

    # Give the user a chance to stop before overwriting
    if overwrite_existing:
        ans = raw_input('Are you sure you want to overwrite existing data (Y/N)')
        if ans.lower() != 'y':
            print 'Exiting without processing...'
            exit()

    print '{count} new files found to process.'.format(count=len(file_names))

    # Loop through each file name to populate all_data DataFrame
    for idx, file_name in enumerate(file_names):
        print 'Processing File {num} of {total}'.format(num=idx + 1, total=len(file_names))

        file_date = date_from_file_name(file_name, return_datetime=True).strftime('%m/%d/%Y')

        try:
            # Load Excel worksheet into a DataFrame
            ws = pd.read_excel(data_folder + file_name, sheetname=sheet_name)

            # Filter for the desired Strategy
            ws = ws[ws[filter_column]==filter_value]

            # Group data by SubStrategy, totaling Daily PNL and Quantity to new DataFrame
            grouped_data = ws.groupby(group_column, as_index=False)[sum_columns].sum()
            final_data = pd.DataFrame(grouped_data)

            # Add snapshot date to the grouped data
            final_data['Date'] = file_date

            # Append new data to existing data
            all_data = all_data.append(final_data, ignore_index=True)

        except Exception as exc:
            # Save the error and move onto next file
            error_msg = 'Error Processing {date} - ({type}): {msg}'.format(date=file_date, type=type(exc).__name__, msg=exc.message)
            error_list.append(error_msg)

    print 'Processing Complete\n'

    # Display any errors that occured during processing
    if error_list:
        print 'The following errors occurred:'
        for err in error_list:
            print '* {msg}'.format(msg=err)

        # Confirm the user still wants to save the data
        if raw_input('\nWould you still like to save the data (Y/N)?').lower() != 'y':
            print 'Exiting without saving...'
            exit()

    # Save the data
    all_data.to_csv(save_path, index=False)
    print 'Output Saved'


if __name__ == '__main__':
    main()
