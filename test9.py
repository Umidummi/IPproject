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

    # Überprüfen, ob der Ordner bereits existiert, und ihn zu erstellen
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Ordner '{folder_path}' wurde erfolgreich erstellt.")
        return folder_path
    else:
        print(f"Der Ordner '{folder_path}' existiert bereits.")
        return createFile()

def statusAbfrage():
    status = app.Application.Acquisition.State
    print(type(status))
    print(status)
    if status==3: #wenn gescanned wird ist status 3, wenn fertig 0 und wenn abgebrochen 5
        time.sleep(2)
        return statusAbfrage()
    elif status==0:
        print('Scan war erfolgreich.')
        return 0
    else:
        print(f'neuer status: ', status)
        return status
#diese Funkton ließt eine excel Tabelle ein wird die Drücke in mBar abfahren. Diese Funktion wird in anderren skripten aufgerufen, daher bitte nicht löschen auch wenn sie hier nicht direkt benutzt wird
def psvDruckKontrolle(i):
    try:
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
            return antwort
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
    try:
        excelfile = input(r'Bitte gebe den Pfad der Excel- oder CSV-Datei mit den Drücken ein: ')
        #excelfile = r'D:\WIN7\Kirchwehm\dt.csv'
        # Überprüfen, ob die Datei existiert
        if not os.path.isfile(excelfile):
            print("Die Datei wurde nicht gefunden.")
            return excelVectorGenerator1()

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
            return excelVectorGenerator1()

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
            return excelVectorGenerator1()
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
        return excelVectorGenerator1()

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
        return druckaktuell
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

def check_referenceFile():
    # referenceFile=input("Bitte gebe den Pfad der ersten Referenz messung ein: ")
    referenceFile = r'D:\WIN7\Kirchwehm\OG.svd'
    if os.path.isfile(referenceFile):
        print(f'Datei existiert: {referenceFile}')
        return referenceFile
    else:
        print(f'die Datei existiert nicht: {referenceFile}')
        check_referenceFile()

br = 38400
to = 1
sp = portsuche()

AcquisitionInstance = win32com.client.Dispatch('PSV.AcquisitionInstance')
app=AcquisitionInstance.GetApplication(True, 10000)
print(app)
app.Application.Activate()


print(app.ActiveDocument.Name)
referenceFile=check_referenceFile()
pressureVector=excelVectorGenerator1()
print('app.ActiveDocument.Name', app.ActiveDocument.Name)
print('app.Application.Acquisition.ScanFileName', app.Application.Acquisition.ScanFileName)
ordnerNeu=createFile()
for i in pressureVector:
    floatvoni=float(i)
    DruckAktuell = psvDruckKontrolle(floatvoni)
    print('DruckAktuell: ', DruckAktuell)
    newFile=rf"{ordnerNeu}\Scan_{i}.svd"
    print('Name der neuen Scandatei: ',newFile)
    copy_binary_file(referenceFile, newFile)
    app.Application.Acquisition.ScanFileName = newFile
    print('app.Application.Acquisition.ScanFileName: ', app.Application.Acquisition.ScanFileName)
    app.Application.Acquisition.Scan(0)
    print('app.Application.ActiveDocument.Name: ', app.Application.ActiveDocument.Name)
    status2 = statusAbfrage()
    print('typ von Status: ', type(status2))
    print(f'was ist der Status? ', status2)
    if not status2 == 0:
        print('Messung wurde abgebrochen. ') #hier kann noch erweitert werden in dem erneut ein scan bei dem fehlgeschlagenen Druck wiederholt wird aber tbcl
        break

print('Messung erfolgreich abgeschlossen.')