from pywinauto import application
from pywinauto.application import Application
import time
import pyautogui
import xlwings as xw
from xlwings import Range, constants
import pydirectinput
from datetime import date
import win32api, win32con

caleExcel = "C:\\Users\\CALITATE\\Desktop\\BOVINA PUTTY.xls"
calePutty = "C:\\vifout\\Putty\\putty.exe"
foaieCalculReceptii = 'Foaie1'
foaieCalculAutomat = 'Date'
denumireButonInchidere = 'Close'

# Functie deschidere consola putty

def deschidereConsola(postLucru):
    global app
    app = Application(backend="uia").start(calePutty)
    pid = application.process_from_module(module=calePutty)
    # app.PuTTYConfiguration.print_control_identifiers()
    btnOpen = app.PuTTYConfiguration.child_window(title="Open", auto_id="1009", control_type="Button")
    srvOption = app.PuTTYConfiguration.child_window(title="srvvif57", control_type="ListItem")
    srvOption.select()
    btnOpen.click()
    time.sleep(1)
    app = Application(backend="uia").connect(process=pid)
    pyautogui.typewrite(postLucru)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter", presses=4)

pyautogui.FAILSAFE = False
pydirectinput.FAILSAFE = False

# Selectare sheet Foaie1 inainte de rularea programului
try:
    xw.Book(caleExcel).sheets[foaieCalculReceptii].select()
except:
    ctypes.windll.user32.MessageBoxW(0, "Te rog selecteaza sheet-ul Foaie1", "Eroare selectie sheet!", 0)
    sys.exit()

#Extragere date din excelul de bovine

wb = xw.Book(caleExcel).sheets[foaieCalculReceptii]
today = date.today()
formatted_date = today.strftime('%m.%d.%Y')
if xw.Book(caleExcel).sheets[foaieCalculAutomat].range("I9").value != formatted_date:
    xw.Book(caleExcel).sheets[foaieCalculAutomat].range("I9").value = formatted_date
    xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value = ""

lastCell = wb.range('E' + str(wb.cells.last_cell.row)).end('up').row
nrReceptie = 0

if xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value is None:
    nrReceptie = 9
else:
    nrReceptie = int(xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value)
    nrReceptie = nrReceptie + 8

if nrReceptie == lastCell:
    verificareCrotal = wb.range("E9" + ":E" + str(lastCell)).value
    nrCrotal = wb.range("E" + str(nrReceptie)).value
    # Se verifica daca celulele sunt goale
    if wb.range("C" + str(nrReceptie)).value is None:
        wb.range("C" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste numarul de criteriu:" + str(nrReceptie - 7),"Nr criteriu lipsa!", 0)
        sys.exit()
    else:
        nrCriteriu = wb.range("C" + str(nrReceptie)).value
    if wb.range("G" + str(nrReceptie)).value is None:
        wb.range("G" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste varsta la pozitia:" + str(nrReceptie - 8),"Varsta lipsa!", 0)
        sys.exit()
    else:
        varsta = wb.range("G" + str(nrReceptie)).value
    if wb.range("H" + str(nrReceptie)).value is None:
        wb.range("H" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste sexul la pozitia:" + str(nrReceptie - 8),"Sex lipsa!", 0)
        sys.exit()
    else:
        sex = wb.range("H" + str(nrReceptie)).value
    if wb.range("I" + str(nrReceptie)).value is None:
        wb.range("I" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste rasa la pozitia:" + str(nrReceptie - 8),"Rasa lipsa!", 0)
        sys.exit()
    else:
        rasa = wb.range("I" + str(nrReceptie)).value
    if wb.range("J" + str(nrReceptie)).value is None:
        wb.range("J" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste propietarul la pozitia:" + str(nrReceptie - 8),"Propietar lipsa!", 0)
        sys.exit()
    else:
        propietar = wb.range("J" + str(nrReceptie)).value
    if wb.range("K" + str(nrReceptie)).value is None:
        wb.range("K" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste localitatea la pozitia:" + str(nrReceptie - 8),"Localitate lipsa!", 0)
        sys.exit()
    else:
        localitate = wb.range("K" + str(nrReceptie)).value
    if wb.range("L" + str(nrReceptie)).value is None:
        wb.range("L" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste cod exploatatie la pozitia:" + str(nrReceptie - 8),"Cod exp lipsa!", 0)
        sys.exit()
    else:
        codExploatatie = wb.range("L" + str(nrReceptie)).value
    if wb.range("M" + str(nrReceptie)).value is None:
        wb.range("M" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste numar pasaport la pozitia:" + str(nrReceptie - 8),"Nr. pasaport lipsa!", 0)
        sys.exit()
    else:
        nrPasaport = wb.range("M" + str(nrReceptie)).value
    if wb.range("N" + str(nrReceptie)).value is None:
        wb.range("N" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Lipseste masina la pozitia:" + str(nrReceptie - 8),"Masina lipsa!", 0)
        sys.exit()
    else:
        masina = wb.range("N" + str(nrReceptie)).value
    nrCriteriu = int(nrCriteriu)
    varsta = int(varsta)
else:
    nrCriteriu = wb.range("C" + str(nrReceptie) + ":C" + str(lastCell)).value
    nrCrotal = wb.range("E" + str(nrReceptie) + ":E" + str(lastCell)).value
    verificareCrotal = wb.range("E9" + ":E" + str(lastCell)).value
    varsta = wb.range("G" + str(nrReceptie) + ":G" + str(lastCell)).value
    sex = wb.range("H" + str(nrReceptie) + ":H" + str(lastCell)).value
    rasa = wb.range("I" + str(nrReceptie) + ":I" + str(lastCell)).value
    propietar = wb.range("J" + str(nrReceptie) + ":J" + str(lastCell)).value
    localitate = wb.range("K" + str(nrReceptie) + ":K" + str(lastCell)).value
    codExploatatie = wb.range("L" + str(nrReceptie) + ":L" + str(lastCell)).value
    nrPasaport = wb.range("M" + str(nrReceptie) + ":M" + str(lastCell)).value
    masina = wb.range("N" + str(nrReceptie) + ":N" + str(lastCell)).value

    # Mesaje de eraore in cazul in care o celula este goala
    for i in range(len(nrCriteriu)):
        if nrCriteriu[i] is None:
            wb.range("C" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste numarul de criteriu:" + str(i + nrReceptie - 8),"Nr criteriu lipsa!", 0)
            sys.exit()
    for i in range(len(varsta)):
        if varsta[i] is None:
            ctypes.windll.user32.MessageBoxW(0, "Lipseste varsta la pozitia:" + str(i + nrReceptie - 8),"Varsta lipsa!", 0)
            sys.exit()
    for i in range(len(sex)):
        if sex[i] is None:
            wb.range("H" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste sexul la pozitia:" + str(i + nrReceptie - 8),"Sex lipsa!", 0)
            sys.exit()
    for i in range(len(rasa)):
        if rasa[i] is None:
            wb.range("I" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste rasa la pozitia:" + str(i + nrReceptie - 8),"Rasa lipsa!", 0)
            sys.exit()
    for i in range(len(propietar)):
        if propietar[i] is None:
            wb.range("J" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste propietarul la pozitia:" + str(i + nrReceptie - 8),"Propietar lipsa!", 0)
            sys.exit()
    for i in range(len(localitate)):
        if localitate[i] is None:
            wb.range("K" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste localitatea la pozitia:" + str(i + nrReceptie - 8),"Localitate lipsa!", 0)
            sys.exit()
    for i in range(len(codExploatatie)):
        if codExploatatie[i] is None:
            wb.range("L" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste codul de exploatatie la pozitia:" + str(i + nrReceptie - 8),"Cod exp lipsa!", 0)
            sys.exit()
    for i in range(len(nrPasaport)):
        if nrPasaport[i] is None:
            wb.range("M" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste numarul de pasaport la pozitia:" + str(i + nrReceptie - 8),"Nr pasaport lipsa!", 0)
            sys.exit()
    for i in range(len(masina)):
        if masina[i] is None:
            wb.range("N" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Lipseste masina la pozitia:" + str(i + nrReceptie - 8),"Masina lipsa!", 0)
            sys.exit()
    nrCriteriu = [int(nrCriteriu) for nrCriteriu in nrCriteriu]
    varsta = [int(varsta) for varsta in varsta]
    rasa = [rasa.strip(' ') for rasa in rasa]
doctor = int(xw.Book(caleExcel).sheets[foaieCalculAutomat].range("E2").value)

# Verificare rasa inainte de lansare program

listRasa = ["AB ANGUS", "AYRS", "BIVOL", "BU", "BB", "BMM", "BN", "BNR", "BR", "BRAUN", "BRUNA", "CHAROL", "FLECK", "FRIZA", "HER", "HOLL", "JER", "LYM", "MET", "MONTB", "PINZG", "RED HOLL", "RED HOOL", "SIMENT", "SURA", "AUBRAC", "HG"
                                                                                                                                                                                                                             ""]
contineRasa = False

if nrReceptie == lastCell:
    for i in range(len(listRasa)):
        if listRasa[i] in rasa:
            contineRasa = True
            break
    if contineRasa == False:
        wb.range("I" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Reintrodu rasa de la pozitia:" + str(nrReceptie - 8), "Rasa incorecta", 0)
        sys.exit()
else:
    for i in range(len(rasa)):
        for j in range(len(listRasa)):
            if listRasa[j] in rasa[i]:
                contineRasa = True
                break

        if contineRasa == False:
            wb.range("I" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Reintrodu rasa de la pozitia:" + str(i + nrReceptie - 8), "Rasa incorecta", 0)
            sys.exit()
        contineRasa = False

# Verificare sex animal

if nrReceptie == lastCell:
    if "F" in sex or "M" in sex:
       print("")
    else:
        wb.range("H" + str(nrReceptie)).color = (235, 52, 52)
        ctypes.windll.user32.MessageBoxW(0, "Reintrodu sexul de la pozitia:" + str(nrReceptie - 8), "Sex incorect", 0)
        sys.exit()
else:
    for i in range(len(rasa)):
        if "F" in sex[i] or "M" in sex[i]:
            continue
        else:
            wb.range("H" + str(nrReceptie + i)).color = (235, 52, 52)
            ctypes.windll.user32.MessageBoxW(0, "Reintrodu sexul de la pozitia:" + str(i + nrReceptie - 8), "Sex incorect", 0)
            sys.exit()

# Verificare crotale duplicate

for i in range(lastCell - 8):
    for j in range(lastCell - 8):

        if i == j:
            pass
        elif verificareCrotal[i] == verificareCrotal[j]:
            ctypes.windll.user32.MessageBoxW(0, "Crotalul " + verificareCrotal[i] + " este duplicat la pozitia " + str(i+1) + " si pozitia " + str(j+1),
                                             "Crotal duplicat", 0)
            sys.exit()

# Se selecteaza o celula goala pentru a evita o eroare
wb.range("E1").select()

#Se va apasa tasta capslock daca este on

caps_status = win32api.GetKeyState(win32con.VK_CAPITAL)

if caps_status==1:
    pyautogui.press("capslock")

# Creare lista de crotale pentru furnizori diferiti

listaCrotale= [

]

#Introducere date in p01

deschidereConsola("p01")
pyautogui.press('r')
pyautogui.press('e')
pyautogui.press('enter', presses=2)

###################################


propietarAnterior = wb.range("J" + str(nrReceptie)).value
#print(propietarAnterior)

# In cazul in care exita doar o singura receptie avem scriptul asta
if nrReceptie == lastCell:
    pyautogui.press("f3")
    pyautogui.press("enter")
    pyautogui.typewrite("0")
    pyautogui.press("enter")
    pyautogui.typewrite(masina)
    pyautogui.press("enter")
    pyautogui.typewrite(str(doctor))
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.typewrite(nrCrotal)
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite("UE")
    pyautogui.press("enter")
    pyautogui.typewrite(nrPasaport)
    pyautogui.press("enter")
    if "RED HOOL" in rasa[0]:
        rasa[0] = "RED HOLL"
    pyautogui.typewrite(rasa)
    pyautogui.press("enter")
    pyautogui.typewrite(sex)
    pyautogui.press("enter")
    pyautogui.typewrite(str(varsta))
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite(propietar)
    pyautogui.press("enter")
    pyautogui.typewrite(propietar)
    pyautogui.press("enter")
    pyautogui.typewrite(codExploatatie)
    pyautogui.press("enter")
    pyautogui.typewrite(localitate)
    pyautogui.press("enter")
    pyautogui.typewrite(str(nrCriteriu))
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.press("f2")
    time.sleep(1)
    pyautogui.press("f4")
    pyautogui.press("d")

    # Salvare nr criteriu in sheet-ul de date in cazul in care conexiunea la server este intrerupta
    xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value = nrCriteriu + 1
else:
    # In cazul in care exita mai multe receptii avem scrptul asta
    pyautogui.press("f3")
    pyautogui.press("enter")
    pyautogui.typewrite("0")
    pyautogui.press("enter")


    pyautogui.typewrite(masina[0])
    pyautogui.press("enter")

    pyautogui.typewrite(str(doctor))
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.typewrite("10301")
    pyautogui.press("enter")

    pyautogui.typewrite(nrCrotal[0])
    pyautogui.press("enter")
    pyautogui.press("enter")

    pyautogui.typewrite("UE")
    pyautogui.press("enter")

    pyautogui.typewrite(nrPasaport[0])
    pyautogui.press("enter")
    if "RED HOOL" in rasa[0]:
        rasa[0] = "RED HOLL"

    pyautogui.typewrite(rasa[0])
    pyautogui.press("enter")

    pyautogui.typewrite(sex[0])
    pyautogui.press("enter")

    pyautogui.typewrite(str(varsta[0]))
    pyautogui.press("enter")
    pyautogui.press("enter")

    pyautogui.typewrite(propietar[0])
    pyautogui.press("enter")

    pyautogui.typewrite(propietar[0])
    pyautogui.press("enter")

    pyautogui.typewrite(codExploatatie[0])
    pyautogui.press("enter")

    pyautogui.typewrite(localitate[0])
    pyautogui.press("enter")

    pyautogui.typewrite(str(nrCriteriu[0]))
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.press("f2")
    time.sleep(3)

    listaCrotale.append(nrCrotal[0])

    for i in range(1,len(propietar)):
        if propietarAnterior != propietar[i]:
            pyautogui.press("f4")
            pyautogui.press("d")
            time.sleep(2)
            app.VIF5_7.child_window(title=denumireButonInchidere, control_type="Button").click()
            pyautogui.press("enter")

            #Salvare nr criteriu in sheet-ul de date in cazul in care conexiunea la server este intrerupta
            xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value = nrCriteriu[i] + 1

            #Deschidere post 2

            pydirectinput.PAUSE = 0.03

            crotaleSortate = sorted(listaCrotale)

            deschidereConsola("p02")
            # pp.VIF5_7.print_control_identifiers()
            pyautogui.press('b')
            pyautogui.press('e')
            time.sleep(1)

            for j in range(len(listaCrotale)):
                pyautogui.hotkey('ctrl', 'o')
                pyautogui.typewrite("10301")
                pyautogui.press("enter")
                pyautogui.press("f2")
                pyautogui.press("enter")
                pyautogui.typewrite(listaCrotale[j])
                pydirectinput.press("enter")
                pydirectinput.press("f2")
                time.sleep(2.5)

            listaCrotale.clear()

            app.VIF5_7.child_window(title=denumireButonInchidere, control_type="Button").click()
            pyautogui.press("enter")

            deschidereConsola("p01")
            # app.VIF5_7.print_control_identifiers()
            pyautogui.press('r')
            pyautogui.press('e')
            pyautogui.press('enter', presses=2)

            pyautogui.press("f3")
            pyautogui.press("enter")
            pyautogui.typewrite("0")
            pyautogui.press("enter")

            pyautogui.typewrite(masina[i])
            pyautogui.press("enter")

            pyautogui.typewrite(str(doctor))
            pyautogui.press("enter")
            pyautogui.press("f2")
            pyautogui.typewrite("10301")
            pyautogui.press("enter")

            pyautogui.typewrite(nrCrotal[i])
            pyautogui.press("enter")
            pyautogui.press("enter")

            pyautogui.typewrite("UE")
            pyautogui.press("enter")

            pyautogui.typewrite(nrPasaport[i])
            pyautogui.press("enter")
            if "RED HOOL" in rasa[i]:
                rasa[i] = "RED HOLL"

            pyautogui.typewrite(rasa[i])
            pyautogui.press("enter")

            pyautogui.typewrite(sex[i])
            pyautogui.press("enter")

            pyautogui.typewrite(str(varsta[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")

            pyautogui.typewrite(propietar[i])
            pyautogui.press("enter")

            pyautogui.typewrite(propietar[i])
            pyautogui.press("enter")

            pyautogui.typewrite(codExploatatie[i])
            pyautogui.press("enter")

            pyautogui.typewrite(localitate[i])
            pyautogui.press("enter")

            pyautogui.typewrite(str(nrCriteriu[i]))
            pyautogui.press("enter")
            pyautogui.press("f2")
            pyautogui.press("f2")
            time.sleep(3)
            propietarAnterior = propietar[i]
            listaCrotale.append(nrCrotal[i])

        else:
            pyautogui.press("enter")

            pyautogui.typewrite(nrCrotal[i])
            time.sleep(1)
            pyautogui.press("enter")
            pyautogui.press("enter")

            pyautogui.typewrite("UE")
            pyautogui.press("enter")

            pyautogui.typewrite(nrPasaport[i])
            pyautogui.press("enter")
            if "RED HOOL" in rasa[i]:
                rasa[i] = "RED HOLL"

            pyautogui.typewrite(rasa[i])
            pyautogui.press("enter")

            pyautogui.typewrite(sex[i])
            pyautogui.press("enter")

            pyautogui.typewrite(str(varsta[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")

            pyautogui.typewrite(str(nrCriteriu[i]))
            pyautogui.press("enter")
            pyautogui.press("f2")
            pyautogui.press("f2")
            time.sleep(3)
            propietarAnterior = propietar[i]
            listaCrotale.append(nrCrotal[i])
    pyautogui.press("f4")
    pyautogui.press("d")
    time.sleep(3)

#Memorare in excel urmatoarea introducere
xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value = lastCell - 7

# Inchidere post P01
app.VIF5_7.child_window(title=denumireButonInchidere, control_type="Button").click()
pyautogui.press("enter")

#Validare date in p02

pydirectinput.PAUSE = 0.02

crotaleSortate = sorted(nrCrotal)

deschidereConsola("p02")
# pp.VIF5_7.print_control_identifiers()
pyautogui.press('b')
pyautogui.press('e')
time.sleep(1)

if nrReceptie == lastCell:
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.press("enter")
    pyautogui.typewrite(nrCrotal)
    pydirectinput.press("enter")
    pydirectinput.press("f2")
else:
    for i in range(len(listaCrotale)):
        pyautogui.hotkey('ctrl', 'o')
        pyautogui.typewrite("10301")
        pyautogui.press("enter")
        pyautogui.press("f2")
        pyautogui.press("enter")
        pyautogui.typewrite(listaCrotale[i])
        pydirectinput.press("enter")
        pydirectinput.press("f2")
        time.sleep(2.5)

# Inchidere post P02
app.VIF5_7.child_window(title=denumireButonInchidere, control_type="Button").click()
pyautogui.press("enter")

#Memorare in excel urmatoarea introducere (asta este pentru cand se face doar P02)
xw.Book(caleExcel).sheets[foaieCalculAutomat].range("G3").value = lastCell - 7

# Salvare fisier excel

xw.Book(caleExcel).save()
