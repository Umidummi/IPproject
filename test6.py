import win32api
import win32com.client

# Create an instance of Excel
excel = win32com.client.Dispatch('Excel.Application')

# Make Excel visible
excel.Visible = True

# Add a new workbook
workbook = excel.Workbooks.Add()

# Access the first sheet
sheet = workbook.Sheets(1)

# Write data to the sheet
sheet.Cells(1, 1).Value = 'Hello, Pywin32!'
path = win32api.GetEnvironmentVariable('PATH')
print(path)
#programm_pfad = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# Programm öffnen
#handle = win32process.CreateProcess(None, programm_pfad, None, None, 0, win32process.CREATE_NO_WINDOW, None, None, win32process.STARTUPINFO())
