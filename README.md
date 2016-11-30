# InvestingCSVtoPivotReports
Python application to convert TD Waterhouse (Canada) CSV export files to useful pivot table friendly data and reports

Note: this program is not associated with TD Canada Trust in any way shape or form.
Please read the licence agreement included with this project prior to using this software. The software is provided AS-IS.

Requirements:

openpyxl module - installed via pip install openpyxsl

Python 3.4

Quick Start:

INITIAL USAGE:
1. Download the script

2. Run the script

3. Observe the PythonExcel.xlsx file has been created.
 
4. Make a subdirectory with current date, in the script folder

5. Login to TD Webbroker, and download CSV files for each account you have. Also download CSV files for other accounts (spouse, etc.) you wish to be compared together.

6. Move all the CSV files into the folder with the current date, in the script folder.

7. If you have any assets at other instutitons, open PythonExcel.xlsx and type the date, and market value into the "Offline" tab

8. Close PythonExcel.xlsx

9. Run the script.

10. Open PythonExcel_Pivot.xlsx

11. Open PythonExcel.xlsx

12. Go to the "AssetAllocationLookup" tab and enter in the category that you want each security recorded under (Canada, USA, International, Fixed Income)

12. Click each chart/graph in PythonExcel_Pivot.xlsx and select "refresh data". If the graphs do not work, click "Source Data" in the ribbon, and then switch to the PythonExcel.xlsx file, and select all the available source data. When you select this it should convert the selection from a range (ie. A1:F333) to a named range "PivotData".

SUBSEQUENT USAGE:
1. Make a new folder with the current date.

2. Download all CSV files.

3. If you have any assets at other instutitons, open PythonExcel.xlsx and type the date, and market value into the "Offline" tab

4. Close PythonExcel.xlsx

5. Run the script.

6. Open PythonExcel.xlsx - if any new assets have been added, update the "AssetAllocationLookup" tab to replace "undefined" with the desired category.

7. Open PythonExcel_Pivot.xlsx

8. Enjoy the reports

--------------------------------------------------------

This script walks through the current working directory
looking for all csv files. It will take TD Waterhouse
downloaded CSV files, and extract the date, account information
and list of securities from each file. The script consolidates
all of the information together, and writes it into an
excel file. The goal of this is all of the data is in one file
with columns for date, account and a vlookup to create a column
for "category" (ie. US equity, fixed income). This allows the use of
pivot tables and pivot charts to easily look at such things as:
current asset allocation, asset allocation by account (how much
international equity in my RRSP, etc.

To support financial information that may be not available
in TD waterhouse csv format, the excel file has an "offline" tab.
Any entries made by hand into the "offline" tab are copied into
the source data tab each time the script is run. This is useful
if you have GIC's, assets at another institution or other items
you wish to track.

Note: you cannot have the pivot table, in the excel file that
this script builds. This script uses a python library which
cannot handle pivot tables. Instead, you have to have a separate
file with a pivot table, but it reads it's source data
from the file built by this Python script.

The python_sample folder shows .csv downloaded files from Waterhouse
in a folder structure, and includes sample results from the script,
and a PythonExcel_Pivot.xlsx file with pivot table reports.

The PythonExcel_Pivot.xlsx is also in the root of this project. To
use, open the PythonExcel.xlsx file created when you run the script,
and then open the PythonExcel_Pivot file. For the pivot table reports
you want, click them, then select "select source data" and repoint
them at your specific excel file on your computer, highlighting the
PivotData named range on the "SourceData" tab.

