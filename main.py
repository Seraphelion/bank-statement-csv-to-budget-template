from data_structures.LinkedList import LinkedList # My LinkedList data structure
from data_structures.Queue import Queue # My Queue data structure
import xlwings as xw # For Creating, Reading, and Writing Excel Files
import csv # For reading csv files
import os # For file checking

# This 
files = Queue()
for file in os.listdir("source"):
    if file.endswith(".csv"):
        files.enqueue(file)

while len(files) > 0:
    filename = "source/" + files.dequeue()

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
    if os.path.exists(r'results/budget.xlsx'):
        wb = xw.Book(r'results/budget.xlsx')
        
        # If sheet called "budget" exist open it, else create it
        if "budget" in [sheet.name for sheet in wb.sheets]:
            ws = wb.sheets["budget"]
        else:
            ws = wb.sheets.add("budget")
            ws.range("A1").value = ["Year", "Month", "Date", "Description", "Category", "Income", "Debits", "Balance"]
            ws.range("D:D").column_width = 30
            ws.range("V1").value = "First itter done: "
            ws.range("V:V").column_width = 15
            ws.range("W1").value = 0
            
    else:
        wb = xw.Book()
        wb.save(r'results/budget.xlsx')
        ws = wb.sheets.add("budget")
        ws.range("A1").value = ["Year", "Month", "Date", "Description", "Category", "Income", "Debits", "Balance"]
        ws.range("D:D").column_width = 30
        ws.range("V1").value = "First itter done: "
        ws.range("V:V").column_width = 15
        ws.range("W1").value = 0

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
        table = ws.api.ListObjects(table_name)
        table.TableStyle = "TableStyleMedium7"

    # Statistics tables, if they dont exist create them, else skip
    if int(ws.range("W1").value) == 0:
        ws.range("W1").value = 1

        # Headers for statistics 2 months ago
        ws.range("J1").value = r'=TEXT(EDATE(TODAY(), -2), "MMMM")' # Two months ago
        ws.range("K1").value = "Total"

        # Income
        ws.range("J2").value = "Earnings"
        ws.range("J:J").column_width = 15
        ws.range("K2").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],J1,BudgetTable[Category],J2)'
        ws.range("J3").value = "Other Income"
        ws.range("K3").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],J1,BudgetTable[Category],J3)'

        # Expenses
        ws.range("J5").value = "Investments"
        ws.range("K5").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J5)'
        ws.range("J6").value = "Groceries"
        ws.range("K6").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J6)'
        ws.range("J7").value = "Transportation"
        ws.range("K7").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J7)'
        ws.range("J8").value = "Clothes"
        ws.range("K8").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J8)'
        ws.range("J9").value = "Self-Improvment"
        ws.range("K9").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J9)'
        ws.range("J10").value = "Leisure"
        ws.range("K10").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J10)'
        ws.range("J11").value = "Utilities"
        ws.range("K11").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J11)'
        ws.range("J12").value = "Rent"
        ws.range("K12").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J12)'
        ws.range("J13").value = "Luxury"
        ws.range("K13").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J13)'
        ws.range("J14").value = "Other Expenses"
        ws.range("K14").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],J1,BudgetTable[Category],J14)'

        # Balance
        ws.range("J16").value = "Balance"
        ws.range("K16").formula = r'=SUM(K2:K3)-SUM(K5:K14)'

        # Formatting
        ws.range("J1:K1").color = (94, 165, 69)
        ws.range("J1:K1").font.bold = True
        ws.range("J1:K1").font.color = (255, 255, 255)

        for i in range(2, 17): # One row green, one row white
            if i % 2 == 0:
                ws.range(f"J{i}:K{i}").color = (221, 237, 214)
        
        ws.range("J2:J16").font.bold = True
        ws.range("J2:J16").font.color = (94, 165, 69)

        # Headers for statistics 1 month ago
        ws.range("M1").value = r'=TEXT(EDATE(TODAY(), -1), "MMMM")' # One month ago
        ws.range("N1").value = "Total"
        ws.range("M1:M1").column_width = 15

        # Income
        ws.range("M2").value = "Earnings"
        ws.range("N2").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],M1,BudgetTable[Category],M2)'
        ws.range("M3").value = "Other Income"
        ws.range("N3").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],M1,BudgetTable[Category],M3)'

        # Expenses
        ws.range("M5").value = "Investments"
        ws.range("N5").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M5)'
        ws.range("M6").value = "Groceries"
        ws.range("N6").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M6)'
        ws.range("M7").value = "Transportation"
        ws.range("N7").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M7)'
        ws.range("M8").value = "Clothes"
        ws.range("N8").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M8)'
        ws.range("M9").value = "Self-Improvment"
        ws.range("N9").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M9)'
        ws.range("M10").value = "Leisure"
        ws.range("N10").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M10)'
        ws.range("M11").value = "Utilities"
        ws.range("N11").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M11)'
        ws.range("M12").value = "Rent"
        ws.range("N12").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M12)'
        ws.range("M13").value = "Luxury"
        ws.range("N13").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M13)'
        ws.range("M14").value = "Other Expenses"
        ws.range("N14").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],M1,BudgetTable[Category],M14)'

        # Balance
        ws.range("M16").value = "Balance"
        ws.range("N16").formula = r'=SUM(N2:N3)-SUM(N5:N14)'

        # Formatting
        ws.range("M1:N1").color = (94, 165, 69)
        ws.range("M1:N1").font.bold = True
        ws.range("M1:N1").font.color = (255, 255, 255)

        for i in range(2, 17): # One row green, one row white
            if i % 2 == 0:
                ws.range(f"M{i}:N{i}").color = (221, 237, 214)
        
        ws.range("M2:M16").font.bold = True
        ws.range("M2:M16").font.color = (94, 165, 69)


        # Headers for statistics of current month
        ws.range("P1").value = r'=TEXT(EDATE(TODAY(), 0), "MMMM")' # Current month
        ws.range("Q1").value = "Total"
        ws.range("P1:P1").column_width = 15

        # Income
        ws.range("P2").value = "Earnings"
        ws.range("Q2").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],P1,BudgetTable[Category],P2)'
        ws.range("P3").value = "Other Income"
        ws.range("Q3").formula = r'=SUMIFS(BudgetTable[Income],BudgetTable[Month],P1,BudgetTable[Category],P3)'

        # Expenses
        ws.range("P5").value = "Investments"
        ws.range("Q5").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P5)'
        ws.range("P6").value = "Groceries"
        ws.range("Q6").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P6)'
        ws.range("P7").value = "Transportation"
        ws.range("Q7").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P7)'
        ws.range("P8").value = "Clothes"
        ws.range("Q8").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P8)'
        ws.range("P9").value = "Self-Improvment"
        ws.range("Q9").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P9)'
        ws.range("P10").value = "Leisure"
        ws.range("Q10").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P10)'
        ws.range("P11").value = "Utilities"
        ws.range("Q11").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P11)'
        ws.range("P12").value = "Rent"
        ws.range("Q12").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P12)'
        ws.range("P13").value = "Luxury"
        ws.range("Q13").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P13)'
        ws.range("P14").value = "Other Expenses"
        ws.range("Q14").formula = r'=SUMIFS(BudgetTable[Debits],BudgetTable[Month],P1,BudgetTable[Category],P14)'

        # Balance
        ws.range("P16").value = "Balance"
        ws.range("Q16").formula = r'=SUM(Q2:Q3)-SUM(Q5:Q14)'

        # Formatting
        ws.range("P1:Q1").color = (94, 165, 69)
        ws.range("P1:Q1").font.bold = True
        ws.range("P1:Q1").font.color = (255, 255, 255)

        for i in range(2, 17): # One row green, one row white
            if i % 2 == 0:
                ws.range(f"P{i}:Q{i}").color = (221, 237, 214)
        
        ws.range("P2:P16").font.bold = True
        ws.range("P2:P16").font.color = (94, 165, 69)


    # Save and close the workbook
    wb.save(r'results/budget.xlsx')
    print("Done!")