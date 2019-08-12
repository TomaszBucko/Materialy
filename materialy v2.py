# -*- coding: utf-8 -*-
import datetime
import os.path   # moduł udostępniający funkcję isfile()
import xlrd
import xlwt


def odczytajstany(sfile):
    slownik = {}
    lista = []
    if os.path.isfile(sfile):  # czy istnieje plik słownika?
        with open(sfile, "r", encoding='CP1250') as sTxt:  # otwórz plik do odczytu
            for i, line in enumerate(sTxt):  # type: (int, object) # przeglądamy kolejne linie
                lista = []
                if i == 2:
                    data = line
                if '|' in line and i > 5:
                    t = line
                    t = t.split("|")
                    indeks = str(t[3]).replace("\n", "").upper().strip()
                    stan = int(str(t[6].replace("\n", "")).strip()[:-4].replace(".", ""))
                    if indeks == "201076":
                        stan -= 5200
                    dziennie = int(str(t[9].replace("\n", "")).strip()[:-4].replace(".", ""))
                    #zapasdni = int(str(t[10].replace("\n", "")).strip()[:-4].replace(".", ""))
                    zamowione = int(str(t[11].replace("\n", "")).strip()[:-4].replace(".", ""))
                    dostawa = str(t[12].replace("\n", "")).strip()
                    lista.append(round(stan, 1))
                    lista.append(round(dziennie, 1))
                    #lista.append(round(zapasdni, 1))
                    lista.append(round(zamowione, 1))
                    if dostawa == "":
                        lista.append(0)
                    else:
                        lista.append(dostawa)
                    slownik[int(indeks)] = lista
                elif i > 6:
                    break
    else:
        print("Nie ma pliku ze stanami: " + sfile)
        print("Wygeneruj stany z SAP i zapisz w folderze, ze skryptem.")
    return slownik



def odczytajnormy(arkusz1):
    slownik = {}
    lista = []
    # tutaj musze odczytać liczbę wierszy.
    rows = arkusz1.nrows
    for lp in range(1, rows, 1):
        lista = []
        lista.append(arkusz1.row_values(lp)[0])
        lista.append(arkusz1.row_values(lp)[2])
        lista.append(arkusz1.row_values(lp)[3])
        lista.append(arkusz1.row_values(lp)[4])
        lista.append(arkusz1.row_values(lp)[5])
        lista.append(arkusz1.row_values(lp)[6])
        lista.append(arkusz1.row_values(lp)[7])
        lista.append(arkusz1.row_values(lp)[8])
        lista.append(arkusz1.row_values(lp)[9])
        lista.append(arkusz1.row_values(lp)[10])
        lista.append(arkusz1.row_values(lp)[11])
        lista.append(arkusz1.row_values(lp)[12])
        slownik[int(arkusz1.row_values(lp)[1])] = lista   # indeks materiału + lista informacji
    return slownik

def odczytajpromocje2(arkusz1):
    slownik = {}
    lista = []
    # tutaj musze odczytać liczbę wierszy.
    rows = arkusz1.nrows
    for lp in range(1, rows, 1):
        lista = []
        dni = arkusz1.row_values(lp)[1]
        palety = arkusz1.row_values(lp)[2]
        lista.append(dni)
        lista.append(palety)
        if arkusz1.row_values(lp)[0] not in slownik:
            slownik[arkusz1.row_values(lp)[0]] = []
        slownik[arkusz1.row_values(lp)[0]].append(lista)   # indeks materiału + lista informacji
    return slownik


def sort(arkusz1):
    kolejnosc = []
    rows = arkusz1.nrows
    for lp in range(1, rows, 1):
        kolejnosc.append(int(arkusz1.row_values(lp)[1]))
    return kolejnosc

def czyzamowic(pokrycie, czasrealizacji, dziennie, chodliwosc, stan, czypoddostawie, zamowione, minimalnezamowienie):
    tekst = []
    if pokrycie > (3 * czasrealizacji) and dziennie > 0.1:
        tekst.append(["Duży zapas", 0])
    else:
        naile = czasrealizacji * 3
        if naile > 90:
            naile = 90
        ilezamowic = int(round(naile * dziennie - stan, 0))
        if chodliwosc == 1 and czypoddostawie is True:
            if pokrycie < (czasrealizacji * 2):
                if minimalnezamowienie > 0:
                    procent = int(ilezamowic / minimalnezamowienie * 100)
                else:
                    procent = 0
                tekst.append(["Zamówić: " + str(ilezamowic) + "(" + str(procent) + "%)" + " na " + str(int(naile)) + " dni.", 0])
                if procent < 100:
                    nailepominimalnym = int((stan + minimalnezamowienie) / dziennie)
                    tekst.append(["(" + str(int(minimalnezamowienie)) + " min na " + str(nailepominimalnym) + "d)", 0])
                else:
                    tekst.append(["", 0])
            else:
                tekst.append(["", 0])
        else:
            tekst.append(["", 0])
    return tekst



def zestawienie(indeks, stan, normy, promocje):
    # liczba znaków na kolejne pozycje w tabeli
    a = 12  # cecha sortowania
    b = 6  # indeks
    c = 40  # opis
    d = 7  # zejscia/stan
    e = 12
    f = 30
    g = 10
    h = 10
    i = c - h
    j = 12
    k = 15
    l = 33
    m = 7
    linia = []
    spearatot = [" | ", 0]
    roznica = ["-> ", 0]
    if indeks == 1:
        linia.append(["Adnotacja", a])
        linia.append(spearatot)
        linia.append(["Indeks", b])
        linia.append(spearatot)
        linia.append(["Opis:", h + i])
        linia.append(spearatot)
        linia.append(["(Zejście)->Stan po", 2*d + len(roznica) + 1])
        linia.append(spearatot)
        linia.append(["(Palety)->Palety po", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        linia.append(["(Tony)->Tony po", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        linia.append(["(Dni)->Dni po", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        linia.append(["Uwagi:", f])
        linia.append(["\n", 0])
    elif indeks == 2:
        pomocnicza = "-" * 9
        for asd in range(0, a + b + h + i + 2 * d + 6 * m + f):
            pomocnicza += "-"
        linia.append([pomocnicza, 0])
        linia.append(["\n", 0])
    else:
        stansap = stan[0]
        dziennie = stan[1]
        zamowione = stan[2]
        dnidodostawy = stan[3]
        cecha = normy[0]
        opis = normy[1]
        norma = normy[2]
        chodliwosc = normy[3]
        grupamaterialowa = normy[4]
        wagapalety = normy[5]
        minimalnezamowienie = normy[6]
        czasrealizacji = normy[7]
        marka = normy[8]
        dniodinwentaryzacji = normy[9]
        indeksy = normy[10].upper().split("|")
        if grupamaterialowa != 2:
            straty = round(-1 * dniodinwentaryzacji * dziennie * 0.03, 1)
            dziennie = dziennie * 1.03
            norma = norma * 1.03
        else:
            straty = 0
            stratytony = round(straty / norma, 1)
            stratypalety = round(stratytony * 1000 / wagapalety, 1)
        stanpostratach = int(round(stansap + straty, 1))
        if stanpostratach < 0:
            stanpostratach = 0
        if norma != 0 and wagapalety != 0:
            tony = round(stanpostratach / norma, 1)
            palety = round(tony * 1000 / wagapalety, 1)
        else:
            palety = "-"
            tony = "-"
        if dziennie == 0:
            dziennie = 0.1
        if dziennie > 0.1:
            pokrycie = round(stanpostratach / dziennie, 1)
        else:
            pokrycie = 999
        linia.append([str(cecha), a])
        linia.append(spearatot)
        linia.append([str(indeks), b])
        linia.append(spearatot)
        linia.append([str(opis), h + i])
        linia.append(spearatot)
        linia.append([str(stanpostratach), 2 * d + len(roznica) + 1])
        if grupamaterialowa != 1:
            linia.append(["szt", 4])
        else:
            linia.append(["mb", 4])
        linia.append(spearatot)
        linia.append([str(palety) + "p", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        linia.append([str(tony) + "t", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        linia.append([str(pokrycie) + "d", 2 * m + len(roznica) + 1])
        linia.append(spearatot)
        if pokrycie < 18 and dziennie > 0.1:
            brak = round((25 - pokrycie) * int(dziennie), 0)
            if chodliwosc == 1:
                if zamowione > 0:
                    linia.append([("Wywołać: " + str(int(brak))), 0])
                else:
                    linia.append([" Niski stan i brak zamówienia !!!", 0])
        czypodostawie = False
        if zamowione == 0:
            czypodostawie = True
        czy = czyzamowic(pokrycie, czasrealizacji, dziennie, chodliwosc, stanpostratach, czypodostawie, zamowione, minimalnezamowienie)
        if czy:
            linia.extend(czy)
        linia.append(["\n", 0])
        listazmianmaterialow = []
        if zamowione > 0:
            listazmianmaterialow.append([int(dnidodostawy), zamowione, "Zamówione"])
        for indekssera in indeksy:
            if str(indekssera) in promocje:
                for numerprom in range(len(promocje[indekssera])):
                    dni = int(promocje[indekssera][numerprom][0])
                    palety = int(promocje[indekssera][numerprom][1])
                    potrzebnematerialy = int(round(palety * wagapalety / 1000 * norma, 2))
                    if promocje[indekssera][numerprom][0] > 0:
                        test = [dni, -potrzebnematerialy, indekssera]
                        if len(listazmianmaterialow) > 0:
                            for lp in range(0, len(listazmianmaterialow)):
                                if listazmianmaterialow[lp][0] > dni:
                                    listazmianmaterialow.insert(lp, test)
                                    break
                                else:
                                    if lp >= len(listazmianmaterialow) - 1:
                                        listazmianmaterialow.append(test)
                        else:
                            listazmianmaterialow.append(test)
                    else:
                        linia.append(["Usuń nieatkualną promocje z pliku. \n", 0])
        stanpopromocji = stanpostratach
        for lp in range(0, len(listazmianmaterialow)):
            dni = listazmianmaterialow[lp][0]
            potrzebnematerialy = listazmianmaterialow[lp][1]
            tekst = listazmianmaterialow[lp][2]
            tonypromocja = 0
            paletypromocja = 0
            indekspromocji = ""
            if zamowione == 0:
                czypodostawie = True
            if tekst == "Zamówione":
                czypodostawie = True
                indekspromocji = "Dostawa"
            else:
                indekspromocji = tekst
                tekst = "Promocja"
            stanpopromocji += potrzebnematerialy
            if norma != 0 and wagapalety != 0:
                tonypromocja = round(potrzebnematerialy / norma, 1)
                paletypromocja = round(tonypromocja * 1000 / wagapalety, 1)
                tony = round(stanpopromocji / norma, 1)
                palety = round(tony * 1000 / wagapalety, 1)
            if dziennie > 0.1:
                pokrycie = round(stanpopromocji / dziennie, 1)
                dnipromocja = round(potrzebnematerialy / dziennie, 0)
            else:
                pokrycie = 999
                dnipromocja = 999
            linia.append([str(tekst), a + b + 3])
            linia.append(spearatot)
            linia.append([indekspromocji + " - za " + str(dni) + "d", h + i])
            linia.append(spearatot)
            linia.append([str(potrzebnematerialy), d])
            linia.append(roznica)
            linia.append([str(stanpopromocji), d])
            if grupamaterialowa != 1:
                linia.append(["szt", 4])
            else:
                linia.append(["mb", 4])
            linia.append(spearatot)
            linia.append([str(paletypromocja) + "p", m])
            linia.append(roznica)
            linia.append([str(palety) + "p", m])
            linia.append(spearatot)
            linia.append([str(tonypromocja) + "t", m])
            linia.append(roznica)
            linia.append([str(tony) + "t", m])
            linia.append(spearatot)
            linia.append([str(int(dnipromocja)) + "d", m])
            linia.append(roznica)
            linia.append([str(pokrycie) + "d", m])
            linia.append(spearatot)
            czy = czyzamowic(pokrycie, czasrealizacji, dziennie, chodliwosc, stanpopromocji, czypodostawie, zamowione, minimalnezamowienie)
            if czy:
                linia.extend(czy)
            linia.append(["\n", 0])

    wydruk = ""
    for i in range(0, len(linia)):
        if indeks == 1:
            wydruk += str(linia[i][0]).center(linia[i][1])
        else:
            wydruk += str(linia[i][0]).rjust(linia[i][1])
    return wydruk



pliksap = "materialy.txt"  # plik ze stanami wyeksporotwane z sapa
stany = odczytajstany(pliksap)


normy = "normy.xls" # specjalnie przygotowany plik z normami
exel = xlrd.open_workbook(normy)
#arkusz1 = exel.sheet_by_name("NORMY")
arkusz1 = exel.sheet_by_index(1)
normy = odczytajnormy(arkusz1)

promocje = "promocje.xls" # specjalnie przygotowany plik z normami
exel2 = xlrd.open_workbook(promocje)
#arkusz1 = exel.sheet_by_name("NORMY")
arkusz2 = exel2.sheet_by_index(1)


prom = odczytajpromocje2(arkusz2)



kolejnosc = sort(arkusz1)
dodruku = "zbiorczy.txt"  # plik wyjsciowy po przetworzeniu
sfile = dodruku
file1 = open(sfile, "w", encoding="utf-8")  # otwieramy plik do zapisu, istniejący plik zostanie nadpisany(!)

linia = zestawienie(1, 1, 1, 1)
file1.write(linia)

listaKluczowe = {}
for i in normy:
    if normy[i][11] not in listaKluczowe and normy[i][11] != 0:
        adres = "Kluczowi/" + str(normy[i][11]) + ".txt"
        listaKluczowe[normy[i][11]] = open(adres, "w", encoding="utf-8")
        listaKluczowe[normy[i][11]].write(linia)

a1 = 1
a2 = 1
for i in kolejnosc:
    if i in normy and i in stany:
        linia = ""
        a1 = normy[i][0]
        if a1 != a2:
            linia = zestawienie(2, 1, 1, 1)
        linia += zestawienie(i, stany[i], normy[i], prom)
        if normy[i][11] in listaKluczowe:
            listaKluczowe[normy[i][11]].write(linia)
        file1.write(linia)
        file1.write("\n")
        a2 = a1
    else:
        print("coś nie tak")