import win32api
import win32con

# Open a registry key
key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 'Software\MyApp', 0, win32con.KEY_ALL_ACCESS)

# Read a value
value, regtype = win32api.RegQueryValueEx(key, 'MyValue')
print(value)

# Write a new value
win32api.RegSetValueEx(key, 'MyNewValue', 0, win32con.REG_SZ, 'Hello, Registry!')

# Close the key
win32api.RegCloseKey(key)