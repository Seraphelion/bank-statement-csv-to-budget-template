from linked_list import LinkedList
from openpyxl import Workbook, load_workbook
import csv
import os

filename = "statement_2025_Jan.csv"

fields = LinkedList()
rows = LinkedList()

with open(filename, "r", encoding="utf-8") as csvfile:
    csvreader = csv.reader(csvfile)
    for field in next(csvreader):
        if field is not None:
            fields.append(field)
    for row in csvreader:
        rows.append(row)

    print("Total no. of rows: %d" % (csvreader.line_num))

print(f"LinkedList fields has {fields.size()} elements")
for element in fields:
    print(element.value, end=" /// ")

print(f"LinkedList rows has {rows.size()} elements")

print('\nFirst 5 rows are:\n')
for row in rows[:5]:
    print(row)

if os.path.exists("budget.xlsx"):
    workbook = load_workbook("budget.xlsx")
    
    if "budget" not in workbook.sheetnames:
        worksheet = workbook.create_sheet("budget")
    else:
        worksheet = workbook ["budget"]

else:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "budget"

workbook.save("budget.xlsx")


