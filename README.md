# InvestingCSVtoPivotReports
Python application to consolidate TD Waterhouse (Canada) CSV export files into one file, including pulling in data from one, non TD Waterhouse File.

Note: this program is not associated with TD Canada Trust in any way shape or form.
Please read the licence agreement included with this project prior to using this software. The software is provided AS-IS.

Once you have this one single CSV file, you can do your own PivotTable reporting and analysis, or to make it a lot more powerful, use Excel's PowerPivot features combined with the data model to add information about each asset. The wiki shows screenshots of what this means and what you can accomplish. In short, if you're using lots of vlookups, you should probably consider the data model instead. If you have Excel and PowerPivot setup properly, you do the following:

1. Download new CSV files from waterhouse
2. Update your offline.csv file with any assets you hold outside of TD waterhouse
3. Open your excel file with PowerPivot
4. Refresh all data sources -> The data model in the Excel updates from the CSV files
5. Look at your reports

This can allow you analyze your holdings with several minutes of work, and also allows you to analyze your, as well as your spouses assets all together.


Requirements:

Python 3


///////////////////////////

INITIAL USAGE:
1. Download the script

2. Run the script
 
3. Make a CSV subdirectory, in that make a directory with the current date, in the CSV folder

4. Login to TD Webbroker, and download CSV files for each account you have. Also download CSV files for other accounts (spouse, etc.) you wish to be compared together.

5 Move all the CSV files into the folder with the current date, in the script folder.

6. If you have any assets at other instutitons, the offline.csv file in the script folder and type in the details of these assets.

7. Run the script.

8. All assets are listed together in the consolidated.csv file

9. REFER to Wiki for instructions on using PowerPivot in Excel

--------------------------------------------------------

This script walks through the current working directory
looking for all csv files. It will take TD Waterhouse
downloaded CSV files, and extract the date, account information
and list of securities from each file. The script consolidates
all of the information together, and writes it into an
excel file. The goal of this is all of the data is in one file.
The data can then be linked with relationships to other tables,
in the Excel Data Model (PowerPivot), to allow you to track:

What percentage of my assets are in each asset class
(fixed income, international, Canadian, etc.)?

What percentage of my assets are in which accounts?

Since all reporting is via pivot tables, you can filter and drill
down to see the makeup. If for example you saw you had a lot
of fixed income assets outside of a tax sheltered account, you
could filter on fixed income and then view all the individual
securities to help identify them.


To support financial information that may be not available
in TD waterhouse csv format, the script creates an offline.csv
file. Any entries made by hand into the "offline" file are copied into
the consolidated.csv file  each time the script is run. This is useful
if you have GIC's, assets at another institution or other items
you wish to track.

