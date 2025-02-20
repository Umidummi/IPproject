import pandas as pd
import os

from pywin.framework.interact import valueFormatOutputError


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
            #df.columns = header
            #print(f"Erste Zeilen der Datei:\n{df.head()}")
        elif file_extension == 'csv': # CSV-Datei einlesen
            trennzeichen=input('Bitte gebe das Trennsymbol der CSV datei ein.')
            df = pd.read_csv(excelfile, sep=trennzeichen, encoding='latin1')
            print(f'df: {df}')
            print("CSV-Datei erfolgreich geladen.")
        else:
            print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
            return excelVectorGenerator0()

        #print(f"Erste Zeilen der Datei:\n{df.head()}")

        rownum = df.shape[0]
        print('Anzahl der Zeilen: ', rownum)
        zeilenstart = None
        spaltennr = None
        for i in range(rownum):
            print(f"Row {i}: {df.iloc[i]}")
            row = df.iloc[i].tolist()
            print('Reihenelemente: ', row)
            for index,j in enumerate(row):
                if type(j)==int:
                    continue
                elif type(j)==str:
                    if 'Drücke' in j:
                        header = row
                        zeilenstart = i
                        spaltennr= index
                        break
            if spaltennr == None and zeilenstart == None:
                    print(f'in der Zeile {i} wurde es nicht gefunden')
            else:
                break
        if spaltennr == None and zeilenstart == None:
            print("Spalte 'Drücke:' wurde nicht gefunden.")
            return excelVectorGenerator0()
        print('Zeilennummer mit Drücke: ', zeilenstart)
        print('Spaltennummer mit Drücke: ', spaltennr+1)
        print("Spaltennamen:", header)
        print('header length: ', len(header))
        pressures = df.iloc[zeilenstart:, spaltennr].dropna().values
        numpressure=[]
        for i in pressures:
            try:
                numvalue=float(i)
                numpressure.append(numvalue)
            except ValueError:
                continue

        print(f"Extrahierte Drücke: {numpressure}")
        return numpressure
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        return excelVectorGenerator0()


excelVectorGenerator0()
# C:\Users\ge96cax\exceldruck\dt.csv
# C:\Users\ge96cax\exceldruck\drucktabelle1.xlsx
#