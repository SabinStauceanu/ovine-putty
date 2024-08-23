
from pywinauto import application
from pywinauto.application import Application
import time
import pyautogui
import xlwings as xw
from xlwings import Range, constants
import pydirectinput
from datetime import date
import win32api, win32con

pyautogui.FAILSAFE = False
pydirectinput.FAILSAFE = False

# Selectare sheet Foaie1 inainte de rularea programului
xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Foaie1'].select()

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

if nrReceptie == lastCell:
    nrCrotal = wb.range("E" + str(nrReceptie)).value
    if wb.range("K" + str(nrReceptie)).value is None:
        wb.range("K" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste masina la pozitia:" + str(nrReceptie - 8),"Masina lipsa!", 0)
        sys.exit()
    else:
        masina = wb.range("K" + str(nrReceptie)).value
    if wb.range("H" + str(nrReceptie)).value is None:
        wb.range("H" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste propietarul la pozitia:" + str(nrReceptie - 8),"Propietar lipsa!", 0)
        sys.exit()
    else:
        propietar = wb.range("H" + str(nrReceptie)).value
    if wb.range("J" + str(nrReceptie)).value is None:
        wb.range("J" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste codul de exploatatie la pozitia:" + str(nrReceptie - 8),"Cod exp lipsa!", 0)
        sys.exit()
    else:
        codExploatatie = wb.range("J" + str(nrReceptie)).value
    if wb.range("I" + str(nrReceptie)).value is None:
        wb.range("I" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste localitatea la pozitia:" + str(nrReceptie - 8),"Localitate lipsa!", 0)
        sys.exit()
    else:
        localitate = wb.range("I" + str(nrReceptie)).value
    if wb.range("G" + str(nrReceptie)).value is None:
        wb.range("G" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste varsta la pozitia:" + str(nrReceptie - 8),"Varsta lipsa!", 0)
        sys.exit()
    else:
        varsta = wb.range("G" + str(nrReceptie)).value
else:
    masina = wb.range("K" + str(nrReceptie) + ":K" + str(lastCell)).value
    propietar = wb.range("H" + str(nrReceptie) + ":H" + str(lastCell)).value
    codExploatatie = wb.range("J" + str(nrReceptie) + ":J" + str(lastCell)).value
    localitate = wb.range("I" + str(nrReceptie) + ":I" + str(lastCell)).value
    nrCrotal = wb.range("E" + str(nrReceptie) + ":E" + str(lastCell)).value
    varsta = wb.range("G" + str(nrReceptie) + ":G" + str(lastCell)).value
    # Se verifica daca celulele sunt goale
    for i in range(len(masina)):
        if masina[i] is None:
            wb.range("K" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste masina la pozitia:" + str(i + nrReceptie - 8),"Masina lipsa!", 0)
            sys.exit()
    for i in range(len(propietar)):
        if propietar[i] is None:
            wb.range("H" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste propietarul la pozitia:" + str(i + nrReceptie - 8),"Propietar lipsa!", 0)
            sys.exit()
    for i in range(len(codExploatatie)):
        if codExploatatie[i] is None:
            wb.range("J" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste codul de exploatatie la pozitia:" + str(i + nrReceptie - 8),"Cod exp lipsa!", 0)
            sys.exit()
    for i in range(len(localitate)):
        if localitate[i] is None:
            wb.range("I" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste localitatea la pozitia:" + str(i + nrReceptie - 8),"Localitate lipsa!", 0)
            sys.exit()
    for i in range(len(varsta)):
        if varsta[i] is None:
            wb.range("G" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste varsta la pozitia:" + str(i + nrReceptie - 8),"Varsta lipsa!", 0)
            sys.exit()
doctor = int(xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("E2").value)
#organeCapre = False
#organeOi = False

# Se selecteaza o celula goala pentru a evita o eroare
wb.range("E1").select()

listaOF = [
    
]
nrArticole = 0

#Se va apasa tasta capslock daca este on

caps_status = win32api.GetKeyState(win32con.VK_CAPITAL)

if caps_status==1:
    pyautogui.press("capslock")


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
#pp.VIF5_7.print_control_identifie
# rs()
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

if nrReceptie == lastCell:
    pyautogui.typewrite(masina)
    pyautogui.press("enter")
    pyautogui.typewrite(propietar)
    pyautogui.press("enter")
    pyautogui.typewrite(codExploatatie)
    pyautogui.press("enter")
    pyautogui.typewrite(propietar)
    pyautogui.press("enter")
    pyautogui.typewrite(localitate)
    pyautogui.press("enter")
    pyautogui.typewrite(str(doctor))
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    if nrCrotal[:3] == "RO2":
        pyautogui.typewrite("10401")
        # organeCapre = True
    else:
        pyautogui.typewrite("10201")
        # organeOi = True
    pyautogui.press("enter")
    pyautogui.typewrite(nrCrotal)
    pyautogui.press("enter")
    pyautogui.press("enter")
    if nrCrotal[:3] == "RO2":
        pyautogui.typewrite("F")
        pyautogui.press("enter")
    if varsta == ">18LUNI":
        pyautogui.typewrite("18+")
    elif varsta == "<18LUNI":
        pyautogui.typewrite("12-18")
    else:
        pyautogui.typewrite("<12")
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    pyautogui.press("f4")
    pyautogui.press("d")
    time.sleep(1)
else:
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
            time.sleep(2)
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
            time.sleep(2)
            propietarAnterior = propietar[i]
            nrArticole = nrArticole + 1
    listaOF.append(nrArticole)
    pyautogui.press("f4")
    pyautogui.press("d")
    time.sleep(1)

xw.Book("C:\\Users\\CALITATE\\Desktop\\OVINA PUTTY.xls").sheets['Date'].range("G3").value = lastCell - 7

app.VIF5_7.child_window(title="Închidere", control_type="Button").click()
pyautogui.press("enter")


# Validare in p02

pydirectinput.PAUSE = 0.03


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


propietarAnterior = wb.range("H" + str(nrReceptie)).value

if nrReceptie == lastCell:
    pyautogui.hotkey('ctrl', 'o')
    if nrCrotal[:3] == "RO2":
        pyautogui.typewrite("10401")
    else:
        pyautogui.typewrite("10201")
    pyautogui.press("enter")
    pyautogui.typewrite(str(1))
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.press("enter")
    pyautogui.press("f5")
    time.sleep(1)
    pydirectinput.press("enter")
    pydirectinput.press("f2")
    time.sleep(1)

else:
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
    time.sleep(2)
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
            time.sleep(2)
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

app.VIF5_7.child_window(title="Închidere", control_type="Button").click()
pyautogui.press("enter")

