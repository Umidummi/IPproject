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
#start acquisition
#app.Acquisition.Start(1)
#app.Application.Acquisition.GeneratorsOn=False #generator an oder aus machen mit True oder False
#vllt am ende wir es gebraucht um laser auszuschalten
app.Application.Acquisition.Scan(0)
status=app.Application.Acquisition.State
while status==3: #wenn gescanned wird ist status=3
    status = app.Application.Acquisition.State
    print(type(status))
    print(status)
    time.sleep(2)
