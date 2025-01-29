
import win32api
import win32com.client
import time
AcquisitionInstance = win32com.client.Dispatch('PSV.AcquisitionInstance')
#print(app)
#print(dir(app))
app=AcquisitionInstance.GetApplication(True, 10000)
print(app)
print(dir(app))
app.Application.Activate()
eigenschaften=app.Application.Acquisition.ActiveProperties(1)
print(eigenschaften)
print(type(eigenschaften))