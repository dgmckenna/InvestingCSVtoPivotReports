#! python3

# This script walks through the current working directory
# looking for all csv files. It will take TD Waterhouse
# downloaded CSV files, and extract the date, account information
# and list of securities from each file. The script consolidates
# all of the information together, and writes it into an
# excel file.

# This data can then be loaded into an excel data model
# which allows you to designate certain securities as
# certain asset types (Canadian equity, fixed income, etc.)
# as well as to list which account numbers are which account types.

import pprint
import os
import sys
import datetime
import re
import string

import csv


def write_csv (output_csv_filename, output_dict, overwrite):
# This routine creates an empty csv file for offline entry of
# assets that you do not have Waterhouse csv files for

    if overwrite or not(os.path.exists(output_csv_filename)):
        #Create workbook with the required tabs
        csv_file = open(output_csv_filename, 'w', newline='')
        csv_writer = csv.writer(csv_file)
        for row in output_dict:
        	csv_writer.writerow(row)
        csv_file.close()


def is_number(s):
# Routine to check if some values from the CSV file are numeric or not
# If they are numeric we need to format them differently when written
# to an excel file.

    try:
        float(s)
        return True
    except ValueError:
        return False


def initialize_pvt_data(pvt_data, header_row):
# This routine sets the columns that will appear in the pivot
# CSV file

    pvt_data.append(header_row)


def read_CSV_file(CSVFilePath, security_dict, pvt_data):
# This routine opens a CSV file (that has already been identified as
# a TD waterhouse file) and extracts the date, account information
# and security information. The extracted information is
# appended to the pvt_data list.

    csv_reader_file = open(CSVFilePath)
    csv_reader = csv.reader(csv_reader_file)
    csv_data = list (csv_reader)

    # The CSV file structure is:
    # [As of Date,2016-10-21 11:17:48],
    # [Account,TD Direct Investing - 538R77A],   (Note: CDN Cash might be TFSA, or RRSP)
    # ['Cash Balance (after settlement)', '[numeric_value]', '', '', '', ''],
    # ['Securities Market Value', '[numeric_value]', '', '', '', ''],
    # ['Total Account Value', '[numeric_value]', '', '', '', ''],
    # ['Margin Available (as of yesterday)', 'N/A', '', '', '', ''],
    # [],
    # ['Symbol', 'Security',  'Quantity',  'Price',  'Book Value',  'Market Value',  'Unrealized $',  'Gain/Loss %',  '% of Holdings'],
    # <Security data rows follow>

    # Extract the date from the file, and re-format it to a date time object
    # in YYYY-MM-DD

    file_date_string = csv_data[0][1].split()[0]
    print (file_date_string)
    file_date = datetime.datetime.strptime(file_date_string, '%Y-%m-%d').strftime('%Y-%m-%d')

    # Extract the account number information and split it into a list
    # Account info list:
    # 0=account number, 2=Currency, 3=account type SDRSP/TFSA/ETC. XXXXXXX - CDN TFSA TD Direct Investing

    account_string = csv_data[1][1]
    account_info_list = account_string.split(' ')

    # Loop through all of the securities information in the CSV file.
    # We skip the first 8 rows to get directly to the securities information.
    # For each item, build a record to insert into our pvt data, by taking
    # the date and account information (already extracted from the start
    # of the file) and the rest of the security information from the row.
    for i in range(8, len(csv_data)):
        # Delete the % holding info, which was copied from the csv file. We do not use this
        # item from the export file, and it simplifies building the row data.
        del csv_data[i][len(csv_data[i])-1]

        row = []
        row.append (file_date) #Date
        row.append (account_info_list[4]) #account number
    
        # Copy all items from the exported CSV file, (except %holding which deleted),
        # into our row. This provides data like name, description, quantity, etc.
        row.append(csv_data[i][0]) #symbol
        row.append(csv_data[i][1]) #market
        row.append(csv_data[i][2]) #desc
        row.append(csv_data[i][3]) #qty
        row.append(csv_data[i][5]) #price
        row.append(csv_data[i][6]) #book value
        row.append(csv_data[i][7]) #market value
        row.append(csv_data[i][8]) #unrealized $
        row.append(csv_data[i][9]) #unrealized percentage

        pvt_data.append (row)
        print (row)

        # Extract the security description, and add it to a dictionary.
        # We need to build a list of all the securities that are in the files
        # so we can generate the asset allocation tab, where there is one
        # row, per security and it's category is listed (ie. US equity, Canadian equity, etc.)
        # The description is used instead of the symbol, to account for the export
        # files not having the market (ie. Canadian Market, or US market, or other). There
        # may be issues where two different securities would have the same symbol on different
        # markets, however it is highly unlikely they would have the same description.

        security_desc = csv_data[i][2]
        security_symbol = csv_data[i][0]

        #make a dictionary with description as the key, then symbol as the value
        security_dict[security_desc] = security_symbol



def collect_offline_data(offline_filename, pvt_data_list):
# This routine copies data from an offline csv file
# and merges it with the current data list

    csv_file = open(offline_filename)
    csv_reader = csv.reader(csv_file)
    offline_data = list(csv_reader)
    new_data_list = pvt_data_list + offline_data[1:]
    csv_file.close()
    return new_data_list
    

# Start of main program

source_path = ''
output_csv_filename = 'consolidated.csv'
offline_filename = 'offline.csv'

print ('Start of program!!!')
print ('***')
print (source_path)
print ('***')

header_row = [
        'Date', 'Account' ,  
        'Symbol', 'Market', 'Security',  'Quantity',  'Price',  'Book Value',
        'Market Value',  'Unrealized $',  'Gain/Loss %'
        ]



pvt_data_list = [header_row]

# Create the initial excel file if not present
write_csv(offline_filename, pvt_data_list, False)

security_dict = {}
found_one_file = False

for foldername, subfolders, filenames in os.walk(os.getcwd()):
    for filename in filenames:
        #search for CSV export files from waterhouse
        #file must be numbers (typically 3, but not always), followed by a letter, followed by 2 numbers, followed by A,S,J or B
        #export CSV files from waterhouse are named (account number)-30-Jun-2015.csv
        matchWaterhouseFile = re.search(r'^\d+\w\d+[ASJB]', filename)
        if filename.endswith('.csv') and filename != offline_filename and filename != output_csv_filename:
            print ('Processing ' + os.path.join(foldername, filename))
            read_CSV_file (os.path.join(foldername, filename), security_dict, pvt_data_list)
            found_one_file = True

print (len(pvt_data_list))
# Merge any data in an offline file with the data currently
# read from all the csv files (pvt_data_list)
merge_data_list = collect_offline_data(offline_filename, pvt_data_list)

print (len(pvt_data_list))
# After parsing all of the csv files write the source data tab with
# all the csv file data
write_csv(output_csv_filename, merge_data_list, True)


if found_one_file == False:
    print(
        'No Waterhouse CSV files found.\n'
        'Put Waterhouse CSV files in this folder, or subfolders of this folder.\n'
        )


print('Press Enter to Continue')
input()
