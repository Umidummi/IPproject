import serial
import time
import serial.tools.list_ports
import numpy as np
import pandas as pd

def druckabfrage(ser, counter):
#    try:
        gp = 'GP\r'
        ser.write(gp.encode('utf-8'))
        time.sleep(0.3)
        response1 = ser.readline().decode('utf-8').strip()
        print(f'Antwort 1: {response1}')
        ser.write(gp.encode('utf-8'))
        time.sleep(0.3)
        response2 = ser.readline().decode('utf-8').strip()
        print(f'Antwort 2: {response2}')
        druckdavor=float(response1)
        druckaktuell=float(response2)
        steigung=abs(druckdavor-druckaktuell)/druckdavor
        relerror=abs(druckaktuell-counter)/counter
        if relerror<0.01 and steigung<0.05:
            return druckaktuell
        else:
            print("Bedingungen nicht erfüllt, erneuter Versuch...")
            return druckabfrage(ser, counter)
#    except ValueError as e:
#        print(f'Error parsing response: {e}')
#        return None

def main():
    try:
        br = 38400
        to = 1
        sp = 'COM17'
        print(br, to, sp)
        ser = serial.Serial(port=sp, baudrate=br, timeout=to)
        print(f'Verbindung hergestellt mit {sp}')
        try:
            excelfile = input(r'bitte gebe den Pfad ein: ')
            df = pd.read_excel(excelfile)
            druckhalten=df.at[0, 'Zeitsabstand[s]: ']
            points = list(df['Druck[mBar]:'])
            print(points)
        except FileNotFoundError:
            print("Die Datei 'book2.xlsx' wurde nicht gefunden.")
        except Exception as e:
            print(f"Ein Fehler ist aufgetreten: {e}")

        for counter in points:
            command=f'SP{counter}\r'
            for char in command:
                ser.write(char.encode('utf-8'))
                print(f'Command gesendet: {char.strip()}')
                time.sleep(0.2)
            time.sleep(0.5)
            response = ord(ser.readline().decode('utf-8').strip())
            if response:
                print(f'ACK=6 or NAK=21 : {response}')
                antwort=druckabfrage(ser,counter)
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
main()

""" antwort=druckabfrage(ser)
            print('Antwort: '+antwort)
            relError= abs(antwort-counter)/counter
            print({'relativer Fehler: '}, +relError)
            time.sleep(4)"""

