from data_structures.LinkedList import LinkedList # My LinkedList data structure
from data_structures.Queue import Queue # My Queue data structure
import xlwings as xw # For Creating, Reading, and Writing Excel Files
import csv # For reading csv files
import os # For file checking

filename = "tests\sample2_statement.csv"

# Reading the file and storing the values
row_queue = Queue()
with open(filename, "r", encoding="utf-8") as csvfile:
    csvreader = csv.reader(csvfile)

    # Skipping header fields
    next(csvreader)

    # Extracting the rest of the csv
    for row in csvreader:
        if any(row):
            ll_row = LinkedList()
            for field in row:
                ll_row.append(field)
            row_queue.enqueue(ll_row)

# If workbook exists open it, else create it
if os.path.exists("budget.xlsx"):
    wb = xw.Book("budget.xlsx")
    
    # If sheet called "budget" exist open it, else create it
    if "budget" in [sheet.name for sheet in wb.sheets]:
        ws = wb.sheets["budget"]
    else:
        ws = wb.sheets.add("budget")
        ws.range("A1").value = ["Year", "Month", "Date", "Description", "Category", "Income", "Debits", "Balance"]
        ws.range("D:D").column_width = 30
        
else:
    wb = xw.Book()
    wb.save("budget.xlsx")
    ws = wb.sheets.add("budget")
    ws.range("A1").value = ["Year", "Month", "Date", "Description", "Category", "Income", "Debits", "Balance"]
    ws.range("D:D").column_width = 30

existing_row_count = ws.used_range.last_cell.row if ws.used_range.value else 0

# Remove balance row after wb creation
if existing_row_count > 1:
    row_queue.dequeue()

# New row count without the last values
new_row_count = len(row_queue) - 3
row_count = existing_row_count + new_row_count

for i in range(existing_row_count + 1, row_count + 1):
    ll_row = row_queue.dequeue()

    # Date and format
    date = LinkedList()
    for element in ll_row[2].split("."):
        date.append(element)
    
    formated_date = f"{date[1]}/{date[0]}/{date[2]}"

    ws.range(f"C{i}").value = formated_date #ll_row[2].replace(".", "/")
    ws.range(f"C{i}").number_format = "DD/MM/YYYY"

    # Description
    ws.range(f"D{i}").value = ll_row[4]

    # Income / Debit
    if ll_row[7] == "K":  # Income
        ws.range(f"F{i}").value = ll_row[5]
    elif ll_row[7] == "D":  # Debits
        ws.range(f"G{i}").value = ll_row[5]

# Formulas for extracting date and month
for i in range(existing_row_count + 1, row_count + 1):
    ws.range(f"A{i}").formula = f"=YEAR(C{i})"
    ws.range(f"B{i}").formula = f"=TEXT(DATE(2011,MONTH(C{i}),1),\"MMMM\")"

    # Calculating balance
for i in range(existing_row_count + 1, row_count + 1):
    if i == 2:
        ws.range(f"H{i}").value = ws.range(f"F{i}").value or 0  # First row balance
    else:
        prev_balance = ws.range(f"H{i - 1}").value or 0
        income = ws.range(f"F{i}").value or 0
        debit = ws.range(f"G{i}").value or 0
        ws.range(f"H{i}").value = float(prev_balance) + float(income) - float(debit)

dropdown_values = "Earnings, Other Income, Investments, Groceries, Transportation, Clothes, Self-Improvment, Leisure, Utilities, Rent, Luxury, Other Expenses"
validation_range = ws.range(f"E2:E{row_count}")
validation_range.api.Validation.Delete()
validation_range.api.Validation.Add(
    Type = 3,
    AlertStyle = 1,
    Formula1 = dropdown_values
)

# If table exist, update it, else create it
table_name = "BudgetTable"
if table_name in [table.Name for table in ws.api.ListObjects]:
    # Find the table by name
    table = next(table for table in ws.api.ListObjects if table.Name == table_name)
    
    # Update the table range to include new rows
    table_range = ws.range(f"A1:H{row_count}")
    table.Resize(table_range.api)
else:
    # Create the table for the first time
    table_range = ws.range(f"A1:H{row_count}")
    ws.api.ListObjects.Add(1, table_range.api, None, 1).Name = table_name

# Save and close the workbook
wb.save("budget.xlsx")