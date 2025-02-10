#alle wichtigen libarys erstmal ein Binden.
import serial
import time
import serial.tools.list_ports
import pandas as pd
#import numpy as np

#der folgende Absatz sucht den usb port aus an dem sie die Vakuumpumpe(VP) aneschlossen haben
ports = list(serial.tools.list_ports.comports()) #ruft eine Liste mit allen existierenden Anschlüssen an Ihrem Computer ab
sp=None
#durch Vergleichen der Namen von allen Anschlüssen mit dem Namen vom Adapter RS232 zu usb wählt es den richtigen Port aus.
for p in ports:
    print(p)
    if 'ATEN'in p.description:
            print(f'this is the Device: {p.device}')
            sp=p.device
    if sp is None:
        print('Das Gerät wurde nicht gefunden.')

br = 38400
to = 1
def getpressure(): #"Druckauslesebefehl" als Symbolkette erstellen
    befehl=('GP')
    return befehl

def setpressure(): #"Druckeinstellungsbefehl" als Symbolkette erstellen
    druck=input('Druck in mbar: ')
    befehl = (f'SP{druck}')
    return befehl

def main(befehl):
    try:
        ser = serial.Serial(port=sp, baudrate=br, timeout=to) #stellt Verbindung mit der Vakuumpumpe(VP) her (öffnet Chanel)
        print(f'Verbindung herestellt mit {sp}')
        command=f'{befehl}\r'
        for char in command: #wandelt die Befehlsymbolkette einzeln in ASCII code um und schick diese mit 0.2Sekunden Abstand zum Gerät ab
            ser.write(char.encode('utf-8'))
            print(f'Command gesendet: {char.strip()}')
            time.sleep(0.2)
        time.sleep(0.3) #wartet insgesamt 0.5Sekunden ab bevor Empfangen der Antwort
        response = ser.readline().decode('utf-8').strip() #liest die Antwort von der VP
# nur für Überprüfung ob Response gleich ACK(6) ist      print(f'Response: {ord(response)}')
        if response:
            print(f'Antwort: {response}')
        else:
            print('keine Antwort. ')
    # der Teil der Funktion bewahrt lediglich das ganze Programmm vor Absturz/Error, da bei serieller Kommunikation durchaus unverständliche Signale empfangen werden.
    except serial.SerialException as e:
        print(f'Fehler: {e}')
    except UnicodeDecodeError as e:
        print(f'Fehler bei der Dekodierung: {e}')
    finally: #hier schließt es die Verbindung zur VP, damit auch andere Programme Zugriff auf die VP bekommen können
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print('Verbindung closed. ')

def choice(): #drei Auswahlmöglichkeiten die ausgeführt werden können, je nach wofür sich der Nutzer nach Öffnung des Programms entscheidet
    wahl=input('(1) Druckabfrage, (2) Druckeinstellung oder (3) Druckabfahren? ')
    if wahl=='1':
        main(getpressure())
    elif wahl=='2':
        main(setpressure())
    elif wahl=='3':
        stufen()
    else:
        print('Die Eingabe ist ungültig. ')
        choice()

def end(): #Beendigung des Programms
    beenden=input('Möchten Sie das Programm beenden drücken Sie bitte die 0.\n Möchten Sie noch ein Befehl eingeben drücken Sie die 1. \n ')
    w2=int(beenden)
    if w2==1:
        choice()
        end()
    elif w2==0:
        input('Drücke irgendeine Taste um Programm zu schließen. ')

def stufen(): #diese Funkton ließt eine excel Tabelle ein wird die Drücke in mBar abfahren
    try:
        ser = serial.Serial(port=sp, baudrate=br, timeout=to)
        print(f'Verbindung hergestellt mit {sp}')
#folgender auskommentierter Block ist eine alternative Möglichkeit ein vektor mit verschiedenen Drücken abzufahren
#        SW = float(input('Startwert[mBar]: '))
#        EW = float(input('Endwert[mBar]: '))
#        schritte = int(input('Anzahl der Abtastungspunkte: '))
#        druckhalten = float(input('Anzahl Sekunden, bei dem der Druck gehalten werden soll: '))  # zeit in sekunden, bei der druck gehalten werden soll
#        points = np.linspace(SW, EW, schritte)
        try:
            #hier wird die excel datei eingelesen und ein vektor generiert sobald der Pfad zu der Datei korrekt eingegeben wurde
            excelfile = input(r'bitte gebe den Pfad ein: ')
            df = pd.read_excel(excelfile)
            #hier wird auch Zeit eingelsen, aber zukünftig soll die Zeit durch die kommunikation mit dem PVS programm bestimmt werden
            druckhalten=df.at[0, 'Zeitsabstand[s]: ']
            points = list(df['Druck[mBar]:'])
            print(points)
        except FileNotFoundError:
            print("Die Datei 'book2.xlsx' wurde nicht gefunden.")
        except Exception as e:
            print(f"Ein Fehler ist aufgetreten: {e}")


        for counter in points: #hier werden die Ascii Symbole wie in Main funktion der VP gesendet
            command = f'SP{counter}\r'
            for char in command:
                ser.write(char.encode('utf-8'))
                print(f'Command gesendet: {char.strip()}')
                time.sleep(0.2)
            time.sleep(0.5)
            response = ord(ser.readline().decode('utf-8').strip())
            if response:
                print(f'ACK=6 or NAK=21 : {response}')
                antwort = druckabfrage(ser, counter)
                print(f'typ von druckaktuell: {type(antwort)}')
                print(f'Antwort Druck: {antwort}')
            else:
                print('keine Antwort. ')
            time.sleep(druckhalten)


    except serial.SerialException as e:
        print(f'Fehler: {repr(e)}')
    except UnicodeDecodeError as e:
        print(f'Fehler bei der Dekodierung: {repr(e)}')
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print('Verbindung closed. ')

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
    #berechnet relativer Fehler zwischen zweiten Druckwert und Sollwert
    relerror=abs(druckaktuell-counter)/counter
    if relerror<0.01 and steigung<0.05: # wenn rel. Fehler unter 1% und  sekante in Betrag unter 5% dann übergibt die funktion den zweiten Druck wert
        return druckaktuell
    else:
        print("Bedingungen nicht erfüllt, erneuter Versuch...") #rekursiver aufruf der funktion solange die Bedingung nicht erfüllt ist
        return druckabfrage(ser, counter)

def psvDruckKontrolle(): #diese Funkton ließt eine excel Tabelle ein wird die Drücke in mBar abfahren. Diese Funktion wird in anderren skripten aufgerufen, daher bitte nicht löschen auch wenn sie hier nicht direkt benutzt wird
    try:
        ser = serial.Serial(port=sp, baudrate=br, timeout=to)
        print(f'Verbindung hergestellt mit {sp}')
        try:
            #hier wird die excel datei eingelesen und ein vektor generiert sobald der Pfad zu der Datei korrekt eingegeben wurde
            excelfile = input(r'bitte gebe den Pfad ein: ')
            df = pd.read_excel(excelfile)
            #hier wird auch Zeit eingelsen, aber zukünftig soll die Zeit durch die kommunikation mit dem PVS programm bestimmt werden
            points = list(df['Druck[mBar]:'])
            print(points)
        except FileNotFoundError:
            print("Die Datei wurde nicht gefunden.")
        except Exception as e:
            print(f"Ein Fehler ist aufgetreten: {e}")


        for counter in points: #hier werden die Ascii Symbole wie in Main funktion der VP gesendet
            command = f'SP{counter}\r'
            for char in command:
                ser.write(char.encode('utf-8'))
                print(f'Command gesendet: {char.strip()}')
                time.sleep(0.2)
            time.sleep(0.5)
            response = ord(ser.readline().decode('utf-8').strip())
            if response:
                print(f'ACK=6 or NAK=21 : {response}')
                antwort = druckabfrage(ser, counter)
                print(f'typ von druckaktuell: {type(antwort)}')
                print(f'Antwort Druck: {antwort}')
            else:
                print('keine Antwort. ')


choice()
end()

#neue änderung
#123
#lARLarlAR