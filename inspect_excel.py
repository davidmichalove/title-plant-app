import openpyxl
wb = openpyxl.load_workbook('/Users/davidmichalove/Desktop/automate/PID OR (DATE)_TEMPLATE (2).xlsx', data_only=True)
for sheet in wb.worksheets:
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str) and cell.strip():
                print(repr(cell))
