import win32com.client

def get_column_from_excel(file_path, sheet_name, column_number):
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Keep Excel hidden

    # Open the workbook
    workbook = excel.Workbooks.Open(file_path)
    sheet = workbook.Sheets(sheet_name)

    # Get the data from the specified column
    column_data = []
    for row in range(1, sheet.UsedRange.Rows.Count + 1):
        cell_value = sheet.Cells(row, column_number).Value
        column_data.append(cell_value)

    # Close the workbook
    workbook.Close()

    # Quit Excel application
    excel.Quit()

    return column_data

# Example usage
file_path = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
sheet_name= input(r'Bitte gebe den sheet Name ein: ')
column_number = input('welche spaltennummer ist die mit den Drücken? ')  # Replace with the column number you want to extract

column_data = get_column_from_excel(file_path, sheet_name, column_number)
print(column_data)
