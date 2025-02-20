import pandas as pd
import os

def excelVectorGenerator0():
    try:
        excelfile = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
        #excelfile = r'D:\WIN7\Kirchwehm\dt.csv'
        # Überprüfen, ob die Datei existiert
        if not os.path.isfile(excelfile):
            print("Die Datei wurde nicht gefunden.")
            return excelVectorGenerator0()

        # Erkennen des Dateiformats anhand der Dateiendung
        file_extension = excelfile.lower().split('.')[-1]

        if file_extension == 'xlsx' or file_extension == 'xls':
            # Excel-Datei einlesen
            df = pd.read_excel(excelfile)
            print(f'df: {df}')
            print("Excel-Datei erfolgreich geladen.")
        elif file_extension == 'csv': # CSV-Datei einlesen
            df = pd.read_csv(excelfile)
            print(f'df: {df}')
            print("CSV-Datei erfolgreich geladen.")
        else:
            print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
            return excelVectorGenerator0()
        print("Verfügbare Spalten:", df.columns)
        print('df.header = ', df.header)
        # Überprüfen, ob eine der Spalten 'Drücke:' enthält, unabhängig von zusätzlichen Zeichen
        columns = df.header.split(';')
        spaltennr=None
        for i, column in enumerate(columns):
            print(f'Index: {i}, Spaltenname: {column}')
            if 'Drücke' in column:
              spaltennr=i
        if spaltennr is None:
            print("Spalte 'Drücke:' wurde nicht gefunden.")
            #return excelVectorGenerator0()
        print(f"Gefundene Spaltennummer: {spaltennr+1}")
        print(f'df[index] = {df[spaltennr]}')
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        return excelVectorGenerator0()


excelVectorGenerator0()