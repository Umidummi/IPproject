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
#app.Application.PrintOut()  #ausdruck von messung graphic
#app.Application.NewWindow()  #öffnet neues fenster von der Messungsgraphic
#start acquisition
#app.Acquisition.Start(1)
#app.Application.Acquisition.GeneratorsOn=False #generator an oder aus machen mit True oder False
#vllt am ende wir es gebraucht um laser auszuschalten
#app.Application.Acquisition.Scan(0)
app.Application.Acquisition.ScanFileName.title()
print(app.Application.Acquisition.ScanFileName.title())
#app.Application.Acquisition.Start
status=app.Application.Acquisition.State
#while status==3: #wenn gescanned wird ist status=3
#    status = app.Application.Acquisition.State
#    print(type(status))
#    print(status)
#    time.sleep(2)
print(dir(app.Application.Acquisition.Infos))
print(type(app.Application.Acquisition.Infos.GetIDsOfNames))
print(app.Application.Acquisition.Infos.GetIDsOfNames)
print(dir(app.Application.Acquisition.ScanFileName.title().replace))
dokument=app.Application.Acquisition.ScanFileName.title().replace(app.Application.ActiveDocument.Name, app.Application.Acquisition.ScanFileName.title())
#app.Application.Acquisition.Document.Save.Title
print(app.Application.ActiveDocument.Name)
print(dokument)
