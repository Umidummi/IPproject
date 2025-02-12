import win32api
import win32com.client
import time
import os
import serial
import serial.tools.list_ports
import pandas as pd

def copy_binary_file(reference_file, new_file):
    # Check if the reference file exists
    if not os.path.exists(reference_file):
        print(f"Reference file '{reference_file}' does not exist.")
        return

    # Open the reference file in binary mode for reading
    with open(reference_file, 'rb') as ref_file:
        # Read the binary content
        binary_content = ref_file.read()

    # Open the new file in binary mode for writing
    with open(new_file, 'wb') as new_file_obj:
        # Write the binary content to the new file
        new_file_obj.write(binary_content)

    print(f"Binary content copied from '{reference_file}' to '{new_file}'.")
def createFile():
    #pfad=input('bitte gebe den pfad ein wo ein neuer ordner erstellt werden soll: ')
    pfad=r'D:\WIN7\Kirchwehm'
    name = input('Name des Ordners: ')
    folder_path = os.path.join(pfad, name)

    # Überprüfen, ob der Ordner bereits existiert, und ihn erstellen
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Ordner '{folder_path}' wurde erfolgreich erstellt.")
        return folder_path
    else:
        print(f"Der Ordner '{folder_path}' existiert bereits.")

def statusAbfrage():
    status = app.Application.Acquisition.State
    if status==3: #wenn gescanned wird ist status 3, wenn fertig 0 und wenn abgebrochen 5
        print(type(status))
        print(status)
        time.sleep(2)
        statusAbfrage()
    elif status==0:
        print(type(status))
        print(status)
        return 0
    elif status ==5:
        print(type(status))
        print(status)
        return 5
    else :
        print(f'neuer status: ', status)

#diese Funkton ließt eine excel Tabelle ein wird die Drücke in mBar abfahren. Diese Funktion wird in anderren skripten aufgerufen, daher bitte nicht löschen auch wenn sie hier nicht direkt benutzt wird
def psvDruckKontrolle(i):
    try:
        sp=portsuche()
        ser = serial.Serial(port=sp, baudrate=br, timeout=to)
        print(f'Verbindung hergestellt mit {sp}')
        druckBefehl = f'SP{i}\r'
        for char in druckBefehl:
            ser.write(char.encode('utf-8'))
            print(f'Command gesendet: {char.strip()}')
            time.sleep(0.2)
        time.sleep(0.5)
        response = ord(ser.readline().decode('utf-8').strip())
        if response:
            print(f'ACK=6 or NAK=21 : {response}')
            antwort = druckabfrage(ser, i)
            print(f'typ von druckaktuell: {type(antwort)}')
            print(f'Druckaktuell: {(antwort)}')
        else:
             print('keine Antwort. ')
    except serial.SerialException as e:
        print(f'Fehler: {repr(e)}')
    except UnicodeDecodeError as e:
        print(f'Fehler bei der Dekodierung: {repr(e)}')

    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print('Verbindung closed. ')

def excelVectorGenerator1():
    while True:
        try:
            #excelfile = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
            excelfile=r'D:\WIN7\Kirchwehm\dt.csv'
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
            elif file_extension == 'csv': # CSV-Datei einlesen
                df = pd.read_csv(excelfile)
                print("CSV-Datei erfolgreich geladen.")
            else:
                print("Nur Excel- (.xlsx, .xls) und CSV-Dateien (.csv) werden unterstützt.")
                continue
            print("Verfügbare Spalten:", df.columns)
            # Entfernen von führenden/nachfolgenden Leerzeichen in den Spaltenüberschriften
            df.columns = df.columns.str.strip()
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


def excelVectorGenerator(): # hier wird die excel datei eingelesen und ein vektor generiert sobald der Pfad zu der Datei korrekt eingegeben wurde
    try:
        excelfile = input(r'bitte gebe den Pfad Exceldatei mit den Drücken ein: ')
        df = pd.read_excel(excelfile)
        print("Verfügbare Spalten:", df.columns)
        # Entfernen von führenden/nachfolgenden Leerzeichen
        df.columns = df.columns.str.strip()
        # Überprüfen, ob die Spalte 'Drücke:' existiert
        #if 'Drücke:' not in df.columns:
        #    print("Spalte 'Drücke:' wurde nicht gefunden.")

        points = list(df['Drücke:'])
        print(points)
        return points
    except FileNotFoundError:
        print("Die Datei wurde nicht gefunden.")
        excelVectorGenerator()
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
        excelVectorGenerator()

def druckabfrage(ser, counter): #in dieser funktion pendelt es den Druck auf die neue Vorgabe ein
    gp = 'GP\r'
    ser.write(gp.encode('utf-8'))
    time.sleep(0.3)
    #ließt es den ersten Druckwert ein
    response1 = ser.readline().decode('utf-8').strip()
    print(f'Antwort 1: {response1}')
    ser.write(gp.encode('utf-8'))
    time.sleep(0.3)
    # ließt es den zweiten Druckwert ein
    response2 = ser.readline().decode('utf-8').strip()
    print(f'Antwort 2: {response2}')
    druckdavor=float(response1)
    druckaktuell=float(response2)
    #berechnet die sekante der zwei Drücke im Betrag
    steigung=abs(druckdavor-druckaktuell)/druckdavor
    print(f'Steigung= {steigung}')
    #print(f'typ von druckaktuell: {type(druckaktuell)}')
    #print(f'typ von counter: {type(counter)}')
    #berechnet relativer Fehler zwischen zweiten Druckwert und Sollwert
    relerror=abs(druckaktuell-counter)/counter
    print(f'relativer Fehler= {relerror}')
    if relerror<0.01 and steigung<0.05: # wenn rel. Fehler unter 1% und  sekante in Betrag unter 5% dann übergibt die funktion den zweiten Druck wert
        print(f'Druck eingestellt bei: {druckaktuell}mBar')
        return 1
    else:
        print("Bedingungen nicht erfüllt, erneuter Versuch...") #rekursiver aufruf der funktion solange die Bedingung nicht erfüllt ist
        return druckabfrage(ser, counter)

def portsuche():
    #der folgende Absatz sucht den usb port aus an dem sie die Vakuumpumpe(VP) aneschlossen haben
    ports = list(serial.tools.list_ports.comports()) #ruft eine Liste mit allen existierenden Anschlüssen an Ihrem Computer ab
    sp=None
    #durch Vergleichen der Namen von allen Anschlüssen mit dem Namen vom Adapter RS232 zu usb wählt es den richtigen Port aus.
    for p in ports:
        print(p)
        if 'ATEN'in p.description:
                print(f'this is the Device: {p.device}')
                sp=p.device
                return sp
        if sp is None:
            print('Das Gerät wurde nicht gefunden.')

br = 38400
to = 1

AcquisitionInstance = win32com.client.Dispatch('PSV.AcquisitionInstance')
app=AcquisitionInstance.GetApplication(True, 10000)
print(app)
app.Application.Activate()


print(app.ActiveDocument.Name)
#referenceFile=input("Bitte gebe den Pfad der ersten Referenz messung ein: ")
referenceFile=r'D:\WIN7\Kirchwehm\OG.svd'
pressureVector=excelVectorGenerator1()
print('app.ActiveDocument.Name', app.ActiveDocument.Name)
print('app.Application.Acquisition.ScanFileName', app.Application.Acquisition.ScanFileName)
ordnerNeu=createFile()
for i in pressureVector:
    floatvoni=float(i)
    psvDruckKontrolle(floatvoni)
    newFile=rf"{ordnerNeu}\Scan_{i}.svd"
    print('Name der neuen Scandatei: ',newFile)
    copy_binary_file(referenceFile, newFile)
    app.Application.Acquisition.ScanFileName = newFile
    print('app.Application.Acquisition.ScanFileName: ', app.Application.Acquisition.ScanFileName)
    app.Application.Acquisition.Scan(0)
    print('app.Application.ActiveDocument.Name: ', app.Application.ActiveDocument.Name)
    if statusAbfrage()==1:
        continue
    elif statusAbfrage()==0:
        print('Messung wurde abgebrochen. ') #hier kann noch erweitert werden in dem erneut ein scan bei dem fehlgeschlagenen Druck wiederholt wird aber tbcl
print('Messung erfolgreich abgeschlossen.')