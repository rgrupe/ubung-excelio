# Program: helloExcel.py - Hello Word app to create a new Microsoft Excel workbook.
# Author: Robert Grupe
# Updated: 2019-09-01

from openpyxl import Workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "world!"

workbook.save(filename="hello_world.xlsx")