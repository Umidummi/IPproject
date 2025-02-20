
#import win32api
#import win32com.client
#import time
#AcquisitionInstance = win32com.client.Dispatch('PSV.AcquisitionInstance')
#print(app)
#print(dir(app))
#app=AcquisitionInstance.GetApplication(True, 10000)
#print(app)
#print(dir(app))
#app.Application.Activate()
#eigenschaften=app.Application.Acquisition.ActiveProperties(1)
#print(eigenschaften)
#print(type(eigenschaften))

import pandas as pd
import os

def excelVectorGenerator1():
    while True:
        try:
            excelfile = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
            #excelfile = r'D:\WIN7\Kirchwehm\dt.csv'
            # Überprüfen, ob die Datei existiert
            if not os.path.isfile(excelfile):
                print("Die Datei wurde nicht gefunden.")
                continue

            # Erkennen des Dateiformats anhand der Dateiendung
            file_extension = excelfile.lower().split('.')[-1]

            if file_extension == 'xlsx' or file_extension == 'xls':
                # Excel-Datei einlesen
                df = pd.read_excel(excelfile)
                print("Excel-Datei erfolgreich geladen.")
                matching_column=list(df['Druecke:'])
                print('mathing_colum:  '+matching_column)
            elif file_extension == 'csv': # CSV-Datei einlesen
                df = pd.read_csv(excelfile)
                print("CSV-Datei erfolgreich geladen.")
                # Entfernen von führenden/nachfolgenden Leerzeichen in den Spaltenüberschriften
                df.columns = df.columns.str.strip()
                print("Verfügbare Spalten:", df.columns)
                # Überprüfen, ob eine der Spalten 'Drücke:' enthält, unabhängig von zusätzlichen Zeichen
                for column in df.columns:
                    if '1' in column:  # Suche nach 'Drücke' in der Spaltenüberschrift
                        matching_column = column
                        break

                if matching_column is None:
                    print("Spalte 'Drücke:' wurde nicht gefunden.")
                    continue
                matching_column = None
            else:
                print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
                excelVectorGenerator1()


            print(f"Gefundene Spalte: {matching_column}")

            # Werte der 'Drücke:;;Zeitsabstände:'-Spalte aufteilen und nur die Druckwerte extrahieren
            druckwerte = []
            for value in df[matching_column]:
                # Zelle nach ';;' aufteilen und nur den ersten Teil (Druckwert) nehmen
                split_values = value.split(';;')
                druckwert = split_values[0]  # Der Druckwert (erste Zahl)

                # Überprüfen, ob der Wert numerisch ist
                try:
                    # Versuche, den Wert in eine Zahl zu konvertieren
                    float(druckwert)  # Wenn erfolgreich, ist es ein gültiger Druckwert
                    druckwerte.append(druckwert)  # Füge den Druckwert hinzu
                except ValueError:
                    # Wenn der Wert keine Zahl ist, überspringe ihn
                    continue

            print(f"Druckwerte: {druckwerte}")
            return druckwerte  # Rückgabe der Liste mit den Druckwerten

        except Exception as e:
            print(f"Ein Fehler ist aufgetreten: {e}")
            continue
def excelVectorGenerator0():
    while True:
        try:
            excelfile = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
            #excelfile = r'D:\WIN7\Kirchwehm\dt.csv'
            # Überprüfen, ob die Datei existiert
            if not os.path.isfile(excelfile):
                print("Die Datei wurde nicht gefunden.")
                continue

            # Erkennen des Dateiformats anhand der Dateiendung
            file_extension = excelfile.lower().split('.')[-1]

            if file_extension == 'xlsx' or file_extension == 'xls':
                # Excel-Datei einlesen
                df = pd.read_excel(excelfile)
                print(f'df: {df}')
                print("Excel-Datei erfolgreich geladen.")
            elif file_extension == 'csv': # CSV-Datei einlesen
                df = pd.read_csv(excelfile)
                print("CSV-Datei erfolgreich geladen.")
            else:
                print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
                excelVectorGenerator1()
            print("Verfügbare Spalten:", df.columns)
            # Entfernen von führenden/nachfolgenden Leerzeichen in den Spaltenüberschriften
            #df.columns = df.columns.str.strip()
            # Überprüfen, ob eine der Spalten 'Drücke:' enthält, unabhängig von zusätzlichen Zeichen
            matching_column = None
            for column in df.columns:
                if 'Drücke' in column:  # Suche nach 'Drücke' in der Spaltenüberschrift
                    matching_column = column
                    break

            if matching_column is None:
                print("Spalte 'Drücke:' wurde nicht gefunden.")
                continue

            print(f"Gefundene Spalte: {matching_column}")

            # Werte der 'Drücke:;;Zeitsabstände:'-Spalte aufteilen und nur die Druckwerte extrahieren
            druckwerte = []
            for value in df[matching_column]:
                # Zelle nach ';;' aufteilen und nur den ersten Teil (Druckwert) nehmen
                split_values = value.split(',,')
                druckwert = split_values[0]  # Der Druckwert (erste Zahl)

                # Überprüfen, ob der Wert numerisch ist
                try:
                    # Versuche, den Wert in eine Zahl zu konvertieren
                    float(druckwert)  # Wenn erfolgreich, ist es ein gültiger Druckwert
                    druckwerte.append(druckwert)  # Füge den Druckwert hinzu
                except ValueError:
                    # Wenn der Wert keine Zahl ist, überspringe ihn
                    continue

            print(f"Druckwerte: {druckwerte}")
            return druckwerte  # Rückgabe der Liste mit den Druckwerten

        except Exception as e:
            print(f"Ein Fehler ist aufgetreten: {e}")
            continue
excelVectorGenerator0()