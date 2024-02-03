import openpyxl
import csv

# Path to the text file
txt_file = 'C:/Users/USER/Documents/Github/email-automation/files/report.txt'

# Path to the new Excel file
xlsx_file = 'C:/Users/USER/Documents/Github/email-automation/files/result.xlsx'

# Create a new workbook and select the active worksheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Open the text file
with open(txt_file, 'r', newline='', encoding='utf-8') as file:
    # Use csv.reader to handle the CSV parsing
    reader = csv.reader(file, delimiter=',', quotechar='"')
    
    for row_index, row in enumerate(reader, start=1):
        for column_index, cell_value in enumerate(row, start=1):
            # Write each cell value to the corresponding cell in the worksheet
            sheet.cell(row=row_index, column=column_index).value = cell_value

# Save the workbook to the new Excel file
workbook.save(xlsx_file)

# No need to explicitly close the workbook; it's automatically closed after saving
