from pypxlib import Table
import xlsxwriter

#########################################################################################################
# Modul zum Exportieren von Pardadox Tabellen nach Excel
#########################################################################################################

def tabelle_nach_excel_auslagern(TabellenPfad,NameDerExcelTeabelle):

    workbook = xlsxwriter.Workbook(f'{NameDerExcelTeabelle}.xlsx')              #.xls File erstellen
    worksheet = workbook.add_worksheet()    # Tabellenseite erstellen

    tabelle = Table(f"{TabellenPfad}")          # Paradox.db einlesen
    x = tabelle.fields                          # Felder ermitteln

    #  tabelle.fields gibt ein dictionary zurück, daher die keys nehmen
    tabellenspalten = x.keys()

    # dictionary keys zum iterieren in liste füllen
    liste_fuer_tabellenspalten_namen = []
    for i in tabellenspalten:
        liste_fuer_tabellenspalten_namen.append(i)


    # nun .xls füllen worksheet.write braucht als parameter row und column, sowie den inhalt
    # beginne bei 0,0
    row1 = 0
    col1 = 0
    for i in liste_fuer_tabellenspalten_namen:      # namen der Tabellenspalten in xls schreiben
        worksheet.write(col1, row1, i)
        row1 += 1       # row hochzählen um von links nach rechts zu füllen

    col = 1     # dannach col auf 1 setzen, um den tabelleninhalt erst in zeile zwei zu beginnen


    #   range(len(tabelle)) = für jeden Datensatz in der Tabelle:
    for i in range(len(tabelle)):
        # datensatz einlesen
        Datensatz1 = tabelle[i]

        row = 0 # row auf 0 setzen, um von links anzufangen

        # range(len(liste_fuer_tabellenspalten_namen)) = schauen wieviele FElder die tabelle hat und für jedes Feld folgendes tun:
        for i in range(len(liste_fuer_tabellenspalten_namen)):  # i = zahl

            # da pypxlib nur Objekte erzeugt muss für jedes betrachtete  Element der Name des Feldes gesucht werden
            aktuelles_feld = f"{liste_fuer_tabellenspalten_namen[i]}" # z.B. aktuelles feld = str "Artikel"

            # getattr(Datensatz1, aktuelles_feld) fragt den Wert für ein bestimmtes Atribut eines Objektes ab.
            wert = (getattr(Datensatz1, aktuelles_feld))

            # wenn das Feld in der Datenbank nicht leer ist, schreib es in .xls und zähle die rows hoch
            if wert != None:
                #liste.append(aktuelles_feld_des_betrachteten_datensatzes)
                worksheet.write(col, row, wert)
                row += 1

            else:   # nichts adden, aber ins nächste feld springen
                row += 1
        # nach jedem Datensatz die columns hochzählen
        col += 1
    #speichern
    workbook.close()


