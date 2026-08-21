import openpyxl

wb = openpyxl.load_workbook('/Users/davidmichalove/Desktop/automate/PID OR (DATE)_TEMPLATE (2).xlsx')
sheet = wb.active
for row in sheet.iter_rows():
    for cell in row:
        if cell.value:
            print(f"Cell {cell.coordinate}: {cell.value}")
