# excel-tools

Collection of scripts for Excel-related data analysis.

## grab_sheets.py

Background:     
- A folder with an .xlsm for each trading day
- In each file is a tab called 'SOTER'

Requirement:    
- Copy all the tabs for a given date range into a single workbook with 
one tab for each date.

Functions:
- def date_from_file_name: Return full date as a string or datetime from a date in a filename.
- def copy_worksheet: Copy values from one openpyxl Worksheet to another.

Classes:
- class OutputWorkbook(Workbook): An extension of the openpyxl Workbook class to add a context manager.



