#!/usr/bin/python3
# -*- utf8
# notepad++ ausführen: python -i "$(FULL_CURRENT_PATH)"

# Dieses Programm sucht in einem Verzeichnis nach PDF Dateien die im Dateinamen Matrikelgruppe, Versuchsname+Nummer und Gruppennummer enthalten:
# z.b: ABC15_SigSys2_Gr2.pdf (Andere Reihenfolge? -> in get_Protokoll_dateien() ändern
# nach der Gruppennummer kann ergänzt werden, ob die Datei kontrolliert wurde, nach Checkliste, etc.  --> 'n' ergänzen
# die letzte Ziffer vor '.pdf' kann dann als Note gewertet werden. So dass eine Datei im Ergebnis
# ABC15_SigSys2_Gr2n2.pdf -->

# YYY_XXXX_GrZ.pdf
# XXXX – Name + Ziffer 1, 2, 3 oder 4 entsprechend für Versuch 1 bis 4
# YYY – Seminargruppe ohne ‚d‘ oder ‚b‘ für Diplom/Bachelor, kein Bindestrich bei KIA, z.B. EAS15, KES14, …
# Z – Ziffer für Gruppennummer laut Praktikumsplan oder mündlicher Absprache


import os  # Dateizugriff

# Dateinamen einlesen

# dirname = "H:/python/2016"
# dirname = os.getcwd() # aktuelles Verzeichnis  --> Verzeichnis in dem aufgerufen wurde
dirname = os.path.dirname(
    os.path.realpath(__file__)
)  # http://stackoverflow.com/questions/5137497/find-current-directory-and-files-directory
print("Verzeichnis: ", dirname)

# Pfad mit Schrägstrich und einfachen Anführungszeichen
# evtl. später automatisch bestimmen
def get_Protokoll_dateien(dirname):

    if (
        os.path.exists(dirname) == False
        or os.path.isdir(dirname) == False
        or os.access(dirname, os.R_OK) == False
    ):
        print("Fehler mit dem Verzeichnis")
        return -1
    else:
        objects = os.listdir(dirname)  # von http://www.decocode.de/?323#2644
        objects.sort()
        files = (
            list()
        )  # Liste mit pro ZEile: Dateiname, versuchnummer, Matrikel, Gruppennummer
        for objectname in objects:
            # print(objectname)
            pos_appendix = -1
            pos_appendix = objectname.lower().find(".pdf")
            if pos_appendix > 0:
                ## Es ist eine PDF!
                # Versuchsnummer etc. abspalten
                # print(objectname)
                Teile = objectname[0:pos_appendix].split("_")
                # es werden die ersten drei Betandteile zwischen '_' genutzt
                # Versuch,Matrikel,Gruppe = Teile[1], Teile[0], Teile[2] # Reihenfolge im Dateinamen festlegen
                Versuch, Matrikel, Gruppe = Teile[0], Teile[1], Teile[2]
                Versuch = int(Versuch[-1:])
                Gruppe = int(Gruppe.lower().strip("grn.")[0])  # erste Ziffer
                files.append(
                    [objectname[0:pos_appendix], Versuch, Matrikel.upper(), Gruppe]
                )

    # print("PDFs::\n")
    return files


# Dateinamen zuergliedern  in
## Versuch MRT1-4
## Matrikel EA14 etc.
## Gruppennummer
## Merkmale
## n - Namen stehen im Versuch
## [1-6] Bewertung
## abweichende Dateinamen erkennen und zusammenfassen

# Intelligenz
## Welche Matrikel gibt es, wie viele Gruppen pro Matrikel?

### Matrikel suchen, zählen und auflisten
#### Liste der enthaltenen Matrikel erstellen
def lese_Matrikel(files):
    alle_Matrikel = list()
    for file in files:
        alle_Matrikel.append(file[2])

    alle_Matrikel.sort()  # sortieren, doppelte stehen jetzt hintereinander
    neue_liste = list()
    vorheriger_eintrag = ""
    for eintrag in alle_Matrikel:
        if eintrag == vorheriger_eintrag:
            continue
        else:
            neue_liste.append(eintrag)
            vorheriger_eintrag = eintrag

    alle_Matrikel = neue_liste  # die erstellte Liste enthällt alle Matrikel

    return neue_liste


#### Abfragen wieviele Gruppen es in dem Matrikel gab, welche Protokolle da sind.
def get_Anzahl_Gruppen_in_Matrikel(Matrikel, files):

    anzahl_Gruppen = -1
    for file in files:
        if file[2].upper() == Matrikel.upper():
            if anzahl_Gruppen < int(
                file[3]
            ):  ## bisherige gespeicherte Anzahl kleiner als Aktuelle Gruppennummer?
                anzahl_Gruppen = int(file[3])  ## --> Speichern

    return anzahl_Gruppen


def get_Einzelnote(file):
    note = 0
    pos_first = file[0].lower().find("_")  # erstes Vorkommen
    pos_second = pos_first + file[0].lower()[pos_first + 1 : -1].find("_")
    pos_Gruppenziffer = pos_second + 4
    if pos_Gruppenziffer + 3 == len(file[0]):
        if file[0][pos_Gruppenziffer + 2].isdigit() > 0:
            note = int(file[0][pos_Gruppenziffer + 2])
            return note
    else:
        return 999  # besser eine zu hohe Zahl zurückgeben als '-1' was zu einer Verbesserung führen würde


def get_Attribute(file):
    # Zwischen Gruppennummer und Ende des Dateinamen suchen
    # pos_appendix = -1
    # pos_appendix = file[0].lower().find('.pdf')
    # if (pos_appendix  > 0):
    # 	# Pos. von Gruppennummer suchen
    # 	pos_Gruppen_nr = 0
    pos_first = file[0].lower().find("_")  # erstes Vorkommen
    pos_second = pos_first + file[0].lower()[pos_first + 1 : -1].find("_")
    pos_Gruppenziffer = pos_second + 4
    # print(file)
    # print(file[0])
    # print('Position: ' + str(pos_Gruppenziffer) + ' Laenge: ' + str(len(file[0])))
    # print(file[0][pos_Gruppenziffer])
    if pos_Gruppenziffer + 1 == len(file[0]):
        return 1  # Protokoll vorhanden
        # (sonst wäre die Fkt nicht aufgerufen worden)

    if pos_Gruppenziffer + 2 == len(file[0]):
        if file[0][pos_Gruppenziffer + 1] == "n":
            return 2  # Namen eingetragen
        else:
            return -1

    if pos_Gruppenziffer + 3 == len(file[0]):
        if file[0][pos_Gruppenziffer + 2].isdigit() > 0:
            return 3  # Note eingetragen
        else:
            return -2

    return -3  # default = Fehler


# Tabelle mit Vollständigkeit
## zusammengefasst pro Gruppe
# abfragen, welche Protokolle schon da sind
def get_Vollstaendigkeit(Matrikel, Gruppe, files):
    Protokolle = [0] * 4  # für 4 Versuche
    for file in files:
        if file[2].upper() == Matrikel.upper():
            Versuch = file[1]
            NrGruppe = file[3]
            if Gruppe == NrGruppe:
                # print("File: " + str(file) + " Matrikel: " + str(Matrikel) + " Gruppe: " + str(Gruppe)) # for debug
                Protokolle[Versuch - 1] = 1  # 1 Protokoll vorhanden
                prot_Status = get_Attribute(file)
                # print("Status: " + str(prot_Status))
                if prot_Status >= 1 and prot_Status <= 3:
                    Protokolle[Versuch - 1] = prot_Status
                    # 2 Namen vorhanden
                    # 3 Bewertet, Note eingetragen

    return Protokolle


def get_note(Matrikel, Gruppe, files):
    # liest die Noten für eine Gruppe aus
    Note = 0.0
    Anz_Noten = 0.0
    for file in files:
        if file[2].upper() == Matrikel.upper():
            Versuch = file[1]
            NrGruppe = file[3]
            if Gruppe == NrGruppe:
                # print("File: " + str(file) + " Matrikel: " + str(Matrikel) + " Gruppe: " + str(Gruppe)) # for debug
                # Protokolle[Versuch-1] = 1 	# 1 Protokoll vorhanden
                prot_Status = get_Attribute(file)
                # print("Status: " + str(prot_Status))
                if prot_Status == 3:
                    # print(type(Note))
                    # print("Note")
                    # print("Note: " + str(get_Einzelnote(file)))
                    Note = Note + get_Einzelnote(file)
                    Anz_Noten = 1 + Anz_Noten

    if Anz_Noten > 0:
        return Note / Anz_Noten, Anz_Noten
    else:
        return 0, 0


### Hier beginnt das eigentliche Programm
files = get_Protokoll_dateien(dirname)
alle_Matrikel = lese_Matrikel(files)

print("Anzahl der Matrikel ist " + str(len(alle_Matrikel)) + "\n")
print(alle_Matrikel)

for Matrikel in alle_Matrikel:
    Gruppenanzahl = get_Anzahl_Gruppen_in_Matrikel(Matrikel, files)
    print(Matrikel + "\that " + str(Gruppenanzahl) + " Versuchsgruppen")
    for Gruppe in range(1, Gruppenanzahl + 1):
        v1, v2, v3, v4 = get_Vollstaendigkeit(Matrikel, Gruppe, files)
        print(
            "\tGruppe "
            + str(Gruppe)
            + ":\t"
            + str(v1)
            + "\t"
            + str(v2)
            + "\t"
            + str(v3)
            + "\t"
            + str(v4)
            + "\t"
        )

for Matrikel in alle_Matrikel:
    Gruppenanzahl = get_Anzahl_Gruppen_in_Matrikel(Matrikel, files)
    print("\n\n")  # leerzeilen
    print(Matrikel + " = " + str(Gruppenanzahl) + " Versuchsgruppen:")
    print("\t\t\tNote\tAnzahl Protokolle")
    for Gruppe in range(1, Gruppenanzahl + 1):
        note, anz_note = get_note(Matrikel, Gruppe, files)
        print("\tGruppe " + str(Gruppe) + ":\t" + str(note) + "\t" + str(anz_note))
