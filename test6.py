import Druckregulierung
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
