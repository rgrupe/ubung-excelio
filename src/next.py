# Program: next.py - TBD.
# Author: Robert Grupe
# Updated: 2019-09-01

from openpyxl import Workbook

filename = "./data/next.xlsx"
workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "world!"
sheet["C10"] = "test"

workbook.save(filename)  # latest values all saved to file

def print_rows():                                   # method suggested as useful later
   print('\nThe following are the tuple values added to the', filename, 'spreadsheet file:')
   for row in sheet.iter_rows(values_only=True):
      print(row)

print_rows() # provide terminal summary output so don't need to open Excel

# METHODS
# .insert_rows(<idx>, <amount>)

#  .delete_rows(<idx>, <amount>)

# .insert_cols(<idx>, <amount>)
## sheet.insert_cols(idx=3, amount=5)    # Insert 5 columns before current "C"

# .delete_cols(<idx>, <amount>)
## sheet.delete_cols(idx=3, amount=5)    # Delete 5 columns starting with "C"