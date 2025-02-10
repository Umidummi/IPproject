import win32api
import win32com.client
import time
import os


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
    pfad=input('bitte gebe den pfad ein wo ein neuer ordner erstellt werden soll: ')
    name = input('Name des Ordners: ')
    folder_path = os.path.join(pfad, name)

    # Überprüfen, ob der Ordner bereits existiert, und ihn erstellen
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Ordner '{folder_path}' wurde erfolgreich erstellt.")
        return folder_path
    else:
        print(f"Der Ordner '{folder_path}' existiert bereits.")


AcquisitionInstance = win32com.client.Dispatch('PSV.AcquisitionInstance')
app=AcquisitionInstance.GetApplication(True, 10000)
print(app)
#print(dir(app))
app.Application.Activate()
app.ActiveDocument=r''
print(app.ActiveDocument.Name)
referenceFile=input("Bitte gebe den Pfad der ersten Referenz messung ein: ")
app.ActiveDocument=referenceFile
print(app.ActiveDocument.Name)
#ScanName = app.Application.Acquisition.ScanFileName.title().replace(app.ActiveDocument.Name, 'Scan9')
print( app.Application.Acquisition.ScanFileName)
ordner=createFile()

for i in range(10):
    newFile=rf"{ordner}\Scan_{i+1}.svd"
    print(newFile)
    #if os.path.exists(newFile):
    #    os.remove(newFile)
    #    print(f"Die Datei {newFile} wurde gelöscht, um sie zu überschreiben.")
    copy_binary_file(referenceFile, newFile)
    app.Application.Acquisition.ScanFileName = newFile
    app.Application.Acquisition.Scan(0)
    print(app.Application.Acquisition.ScanFileName)
    #app.Application.Acquisition.Start
    status=app.Application.Acquisition.State
    while status==3: #wenn gescanned wird ist status=3
        status = app.Application.Acquisition.State
        print(type(status))
        print(status)
        time.sleep(2)
    print(app.Application.ActiveDocument.Name)




#print(dokument)
#app.Application.PrintOut()  #ausdruck von messung graphic
#app.Application.NewWindow()  #öffnet neues fenster von der Messungsgraphic
#app.Acquisition.Start(1)
#app.Application.Acquisition.GeneratorsOn=False #generator an oder aus machen mit True oder False
#vllt am ende wir es gebraucht um laser auszuschalten






