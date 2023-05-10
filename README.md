## Write Country Spreadsheets

### Dependencies

Requires running geotrellis annual update analysis on GADM boundaries.

You will need an updated metadata Excel sheet. Contact GFW data team to create this for you.

Currently only works on Mac OS X. Requires Excel 2010+ installed.

### Instructions

In your results folder, under each admin level (iso/adm1/adm2) there should be a "download" folder. Copy all folders locally.

Run country_stats.py with those folders as inputs to generate global and country .xlsx files:

`python country_stats.py <ISO download path> <ADM1 download path> <ADM2 download path>`

Once generated, copy the .xlsx files to a folder in ~/Library/Containers/com.microsoft.Excel/Data. This allows the script to access the files through the Excel API without prompting to grant permission for each one.

Now run the insert_info.py script*:

`python insert_info.py <metadata file path> <country spreads folder path>`

This script will prepend the metadata sheet to all the country workbooks. 

*NOTE if running from PyCharm: for the Excel API to work, you'll be prompted to grant permission to use Excel from the script. For some reason, running the script in PyCharm sometimes won't prompt you, causing you to get denied access. A workaround is to just run from a terminal outside of PyCharm.
 

