
from pywinauto import application
from pywinauto.application import Application
import time
import pyautogui
import xlwings as xw
from xlwings import Range, constants
import pydirectinput
from datetime import date

#Extragere date din excel

wb = xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Foaie1']
today = date.today()
formatted_date = today.strftime('%m.%d.%Y')
if xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("I9").value != formatted_date:
    #xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F7").value = "NU"
    #xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F9").value = "NU"
    xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("I9").value = formatted_date
    xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("G3").value = ""

lastCell = wb.range('E' + str(wb.cells.last_cell.row)).end('up').row
nrReceptie = 0
if xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("G3").value is None:
    nrReceptie = 9
else:
    nrReceptie = int(xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("G3").value)
    nrReceptie = nrReceptie + 8

masina = wb.range("K" + str(nrReceptie) + ":K" + str(lastCell)).value
propietar = wb.range("H" + str(nrReceptie) + ":H" + str(lastCell)).value
codExploatatie = wb.range("J" + str(nrReceptie) + ":J" + str(lastCell)).value
localitate = wb.range("I" + str(nrReceptie) + ":I" + str(lastCell)).value
nrCrotal = wb.range("E" + str(nrReceptie) + ":E" + str(lastCell)).value
varsta = wb.range("G" + str(nrReceptie) + ":G" + str(lastCell)).value
doctor = int(xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("E2").value)
#organeCapre = False
#organeOi = False

listaOF = [
    
]
nrArticole = 0


# Introducere in p01

app = Application(backend="uia").start("C:\\vifout\\Putty\\putty.exe")
pid = application.process_from_module(module = "C:\\vifout\\Putty\\putty.exe")
#app.PuTTYConfiguration.print_control_identifiers()
btnOpen = app.PuTTYConfiguration.child_window(title="Open", auto_id="1009", control_type="Button")
srvOption = app.PuTTYConfiguration.child_window(title="srvvif57", control_type="ListItem")
srvOption.select()
btnOpen.click()
time.sleep(1)
app = Application(backend="uia").connect(process=pid)
#pp.VIF5_7.print_control_identifiers()
pyautogui.typewrite("p01")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter", presses=4)
pyautogui.press('r')
pyautogui.press('e')
pyautogui.press('enter', presses=2)
pyautogui.press("f3")
pyautogui.press("enter")
pyautogui.typewrite("00")
pyautogui.press("enter")
pyautogui.typewrite(masina[0])
pyautogui.press("enter")
pyautogui.typewrite(propietar[0])
pyautogui.press("enter")
pyautogui.typewrite(codExploatatie[0])
pyautogui.press("enter")
pyautogui.typewrite(propietar[0])
pyautogui.press("enter")
pyautogui.typewrite(localitate[0])
pyautogui.press("enter")
pyautogui.typewrite(str(doctor))
pyautogui.press("enter")
pyautogui.press("f2")
time.sleep(1)
if nrCrotal[0][:3] == "RO2":
    pyautogui.typewrite("10401")
    #organeCapre = True
else:
    pyautogui.typewrite("10201")
    #organeOi = True
pyautogui.press("enter")
pyautogui.typewrite(nrCrotal[0])
pyautogui.press("enter")
pyautogui.press("enter")
if nrCrotal[0][:3] == "RO2":
    pyautogui.typewrite("F")
    pyautogui.press("enter")
if varsta[0] == ">18LUNI":
    pyautogui.typewrite("18+")
elif varsta[0] == "<18LUNI":
    pyautogui.typewrite("12-18")
else:
    pyautogui.typewrite("<12")
pyautogui.press("enter")
pyautogui.press("f2")
time.sleep(1)
nrArticole = nrArticole + 1

propietarAnterior = wb.range("H" + str(nrReceptie)).value

for i in range(1,len(propietar)):
    if propietarAnterior != propietar[i]:
        listaOF.append(nrArticole)
        nrArticole = 0
        pyautogui.press("f4")
        pyautogui.press("d")
        time.sleep(1)
        pyautogui.press("f3")
        pyautogui.press("enter")
        pyautogui.typewrite("00")
        pyautogui.press("enter")
        pyautogui.typewrite(masina[i])
        pyautogui.press("enter")
        pyautogui.typewrite(propietar[i])
        pyautogui.press("enter")
        pyautogui.typewrite(codExploatatie[i])
        pyautogui.press("enter")
        pyautogui.typewrite(propietar[i])
        pyautogui.press("enter")
        pyautogui.typewrite(localitate[i])
        pyautogui.press("enter")
        pyautogui.typewrite(str(doctor))
        pyautogui.press("enter")
        pyautogui.press("f2")
        time.sleep(1)
        if nrCrotal[i][:3] == "RO2":
            pyautogui.typewrite("10401")
            # organeCapre = True
        else:
            pyautogui.typewrite("10201")
            # organeOi = True
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i])
        pyautogui.press("enter")
        pyautogui.press("enter")
        if nrCrotal[i][:3] == "RO2":
            pyautogui.typewrite("F")
            pyautogui.press("enter")
        if varsta[i] == ">18LUNI":
            pyautogui.typewrite("18+")
        elif varsta[i] == "<18LUNI":
            pyautogui.typewrite("12-18")
        else:
            pyautogui.typewrite("<12")
        pyautogui.press("enter")
        pyautogui.press("f2")
        time.sleep(1)
        propietarAnterior = propietar[i]
        nrArticole = nrArticole + 1
    else:
        if nrCrotal[i][:3] == "RO2":
            pyautogui.typewrite("10401")
            # organeCapre = True
        else:
            pyautogui.typewrite("10201")
            # organeOi = True
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i])
        pyautogui.press("enter")
        pyautogui.press("enter")
        if nrCrotal[i][:3] == "RO2":
            pyautogui.typewrite("F")
            pyautogui.press("enter")
        if varsta[i] == ">18LUNI":
            pyautogui.typewrite("18+")
        elif varsta[i] == "<18LUNI":
            pyautogui.typewrite("12-18")
        else:
            pyautogui.typewrite("<12")
        pyautogui.press("enter")
        pyautogui.press("f2")
        time.sleep(1)
        propietarAnterior = propietar[i]
        nrArticole = nrArticole + 1
listaOF.append(nrArticole)
pyautogui.press("f4")
pyautogui.press("d")
time.sleep(1)

xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("G3").value = lastCell - 7


# Validare in p02

pydirectinput.PAUSE = 0.05


crotaleSortate = sorted(nrCrotal)

app = Application(backend="uia").start("C:\\vifout\\Putty\\putty.exe")
pid = application.process_from_module(module = "C:\\vifout\\Putty\\putty.exe")
#app.PuTTYConfiguration.print_control_identifiers()
btnOpen = app.PuTTYConfiguration.child_window(title="Open", auto_id="1009", control_type="Button")
srvOption = app.PuTTYConfiguration.child_window(title="srvvif57", control_type="ListItem")
srvOption.select()
btnOpen.click()
time.sleep(1)
app = Application(backend="uia").connect(process=pid)
#pp.VIF5_7.print_control_identifiers()
pyautogui.typewrite("p03")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter", presses=4)
pyautogui.press('b')
pyautogui.press('e')
time.sleep(1)


"""
if xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F7").value == "NU" and organeOi == True:
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10201")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite("10201_ORGANE")
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F7").value = "DA"
if xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F9").value == "NU" and organeCapre == True:
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10401")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite("10401_ORGANE")
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("F9").value = "DA"
"""

propietarAnterior = wb.range("H" + str(nrReceptie)).value

pyautogui.hotkey('ctrl', 'o')
if nrCrotal[0][:3] == "RO2":
    pyautogui.typewrite("10401")
else:
    pyautogui.typewrite("10201")
pyautogui.press("enter")
pyautogui.typewrite(str(listaOF[0]))
pyautogui.press("enter")
pyautogui.press("f2")
pyautogui.press("enter")
pyautogui.press("f5")
time.sleep(1)
for j in range(len(nrCrotal)):
    if nrCrotal[0] != crotaleSortate[j]:
        pydirectinput.press("down")
    else:
        pydirectinput.press("enter")
        crotaleSortate.pop(j)
        pydirectinput.press("f2")
        time.sleep(1)
        break

for i in range(1,len(propietar)):
    if propietarAnterior != propietar[i]:
        listaOF.pop(0)
        pyautogui.hotkey('ctrl', 'o')
        if nrCrotal[i][:3] == "RO2":
            pyautogui.typewrite("10401")
        else:
            pyautogui.typewrite("10201")
        pyautogui.press("enter")
        pyautogui.typewrite(str(listaOF[0]))
        pyautogui.press("enter")
        pyautogui.press("f2")
        pyautogui.press("enter")
        pyautogui.press("f5")
        time.sleep(1)
        for j in range(len(nrCrotal)):
            if nrCrotal[i] != crotaleSortate[j]:
                pydirectinput.press("down")
            else:
                pydirectinput.press("enter")
                crotaleSortate.pop(j)
                pydirectinput.press("f2")
                time.sleep(1)
                break
        propietarAnterior = propietar[i]
    else:
        pyautogui.press("enter")
        pyautogui.press("f5")
        time.sleep(1)
        for j in range(len(nrCrotal)):
            if nrCrotal[i] != crotaleSortate[j]:
                pydirectinput.press("down")
            else:
                pydirectinput.press("enter")
                crotaleSortate.pop(j)
                pydirectinput.press("f2")
                time.sleep(1)
                break
        propietarAnterior = propietar[i]
