# excel-tools

Collection of scripts for Excel-related data analysis.

## grab_sheets.py

Background:     
- A folder with an .xlsm for each trading day
- In each file is a data tab with the same name

Requirement:    
- Copy all the tabs for a given date range into a single workbook with 
one tab for each date

## grab_sheets_pandas.py

Background:
- A folder with an .xlsm for each trading day
- In each file are various data tabs, identical in each file
- Goal was to extend functionality of grab_sheets.py and practice using pandas

Requirement:
- Copy a filtered, grouped slice of data for a given date range into a single csv file
- Have the option of merely updating the existing data instead of replacing all each time


