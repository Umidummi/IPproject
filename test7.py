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
            df = pd.read_excel(excelfile, header=None)
            print(f'df:\n {df}')
            print("Excel-Datei erfolgreich geladen.")
            rownum = df.shape[0]
            print('Anzahl der Zeilen: ',rownum)
            for i in range(rownum):
                row=df.iloc[i].tolist()
                if 'Drücke' in row:
                    header = row
            df.columns = header
            print(f"Erste Zeilen der Datei:\n{df.head()}")
        elif file_extension == 'csv': # CSV-Datei einlesen
            trennzeichen=input('Bitte gebe das Trennsymbol der CSV datei ein.')
            df = pd.read_csv(excelfile, sep=trennzeichen, encoding='latin1')
            print(f'df: {df}')
            print("CSV-Datei erfolgreich geladen.")
            header = df.columns.tolist()
        else:
            print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
            return excelVectorGenerator0()

        #print(f"Erste Zeilen der Datei:\n{df.head()}")


        print("Spaltennamen:", header)
        print('header length: ', len(header))
        spaltennr=None
        for i, column in enumerate(header):
            print(f'Index: {i}, Spaltenname: {column}')
            if 'Drücke' in column:
                spaltennr=i
        if spaltennr is None:
            print("Spalte 'Drücke:' wurde nicht gefunden.")
            #return excelVectorGenerator0()
        print(f"Gefundene Spaltennummer: {spaltennr + 1}")
        pressures = df.iloc[1:, spaltennr].dropna().values
        print(f"Extrahierte Drücke: {pressures}")
        return pressures
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        return excelVectorGenerator0()


excelVectorGenerator0()
# C:\Users\ge96cax\exceldruck\dt.csv
# C:\Users\ge96cax\exceldruck\drucktabelle1.xlsx
#