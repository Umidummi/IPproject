import serial
import time
import serial.tools.list_ports

sp='COM17'
br = 38400
to = 1
def getpressure():
    befehl=('GP')
    return befehl

def setpressure():
    druck=input('Druck in mbar: ')
    befehl = (f'SP{druck}')
    return befehl

def main(befehl):
    try:
        ser = serial.Serial(port=sp, baudrate=br, timeout=to)
        print(f'Verbindung herestellt mit {sp}')
        command=f'{befehl}\r'
        for char in command:
            ser.write(char.encode('utf-8'))
            print(f'Command gesendet: {char.strip()}')
            time.sleep(0.2)
        time.sleep(0.5)
        response = ser.readline().decode('utf-8').strip()
        print(f'Response: {response}')
        if response:
            print(f'Antwort: {response}')
        else:
            print('keine Antwort. ')
    except serial.SerialException as e:
        print(f'Fehler: {e}')
    except UnicodeDecodeError as e:
        print(f'Fehler bei der Dekodierung: {e}')
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print('Verbindung closed. ')
def choice():
    wahl=input('(1) Druckabfrage oder (2) Druckeinstellung? ')
    if wahl=='1':
        main(getpressure())
    elif wahl=='2':
        main(setpressure())
    else:
        print('Die Eingabe ist ungültig. ')
        choice()

def end():
    beenden=input('Möchten Sie das Programm beenden drücken Sie bitte die 0.\n Möchten Sie noch ein Befehl eingeben drücken Sie die 1. \n ')
    w2=int(beenden)
    if w2==1:
        choice()
        end()
    elif w2==0:
        input('Drücke irgendeine Taste um Programm zu schließen. ')

choice()
end()