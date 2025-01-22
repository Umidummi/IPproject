import pyautogui
import time

# Warten, um sicherzustellen, dass das Programm geöffnet ist
time.sleep(5)

# Koordinaten des Eingabefeldes und des Buttons
input_field_coords = (x, y)
button_coords = (x, y)

# Zahl in das Eingabefeld eingeben
pyautogui.click(input_field_coords)
pyautogui.typewrite('12345')

# Button klicken
pyautogui.click(button_coords)
