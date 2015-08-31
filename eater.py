#!/usr/bin/python
# -*- coding: utf-8 -*-

import datetime
import pyexcel
from pyexcel.ext import ods3

'''
Klasse für ein Kind, was am Mittagessen teilnimmt mit allen dazugehörigen Daten.
Beim Import werden die Daten aus der Tabelle in Objekte dieses Typs überführt.
'''
class Eater:
    def __init__(self):
        ''' Stammdaten '''
        self.id = None        # Eindeutige ID
        self.vorname = ''     # Vorname des Kindes
        self.nachname = ''    # Nachname des Kindes
        self.nachname2 = ''   # Wenn vorhanden: Familienname des zweiten Elternteils
        self.strasse = ''     # Straße für die Rechnung
        self.plz = ''         # Postleitzahl für Rechnungsadresse
        self.stadt = ''       # Stadt für Rechnungsadresse
        self.but = ''         # Auswahl aus [job, wg] = Jobcenter bzw. Wohngeld
        self.but_ende = None  # Ende des Bewilligungszeitraumes für BuT
        self.but_gutschein = '' # Gutscheinnummer für BuT (bei Jobcenter)
        ''' Abrechnungsweise '''
        self.dauer_last = False   # ja / nein
        self.dauer_last_datum = None # Tag des Monats an dem Dauerlastschrift gezogen wird
        self.lastschrift = False  # ja / nein
        self.rechnung = False     # ja / nein
        self.kontoinhaber = ''    # Kontoinhaber des Bankkontos von dem abgebucht wird
        self.IBAN = ''            # Bank: IBAN
        self.BIC = ''             # Bank: BIC
        self.mail1 = ''           # Erste Mailadresse für Zusendung der Rechnung
        self.mail2 = ''           # Zweite Mailadresse für die Zusendung der Rechnung
        ''' Anzahl an Essen an denen teilgenommen wurde '''
        self.anzahl = 0

'''
Wandelt ein gegebenes Sheet eines ODS-Files im in ein Dictionary um in dem die
id als Schlüssel für ein eater-Objekt verwendet wird.
'''
def getEaterFromSheet(sheet):
    # Header auf richtige Reihenfolge überprüfen
    header = ['id', 'Vorname', 'Nachname', 'Nachname2', 'Strasse',
        'PLZ', 'Stadt', 'BuT', 'BuT Frist', 'Gutschein Nummer', 'Dauerlastschrift',
        'Dauerlastschrift Datum', 'Lastschrift', 'Rechnung', 'Kontoinhaber',
        'IBAN', 'BIC', 'Mail1', 'Mail2']
    first_row = sheet.row[0]
    for i in range(0, len(header)):
        if header[i] != first_row[i]:
            print u'Falsche Überschrift in Spalte ' + str(i) + ' \'' + first_row[i] + '\''
            print u'Sollte sein: \'' + header[i] + '\'.' 

    # Daten einlesen
    eaters = {}
    records = sheet.to_array()
    for i, line in enumerate(records):
        id = line[header.index('id')]
        # Leere Zeilen und Header auslassen.
        if id != '' and id != 'id':
            eater = Eater()
            # Adressdaten einfach nur kopieren
            eater.vorname   = line[header.index('Vorname')]
            eater.nachname  = line[header.index('Nachname')]
            eater.nachname2 = line[header.index('Nachname2')]
            eater.strasse   = line[header.index('Strasse')]
            eater.plz       = line[header.index('PLZ')]
            eater.stadt     = line[header.index('Stadt')]
            eater.but       = line[header.index('BuT')]
            # Wenn Teilnehmer von Bildung und Teilhabe, hole Ablaufdatum und
            # Gutscheinnummer.
            if eater.but in ['job','wg']:
                eater.but_gutschein = line[header.index('Gutschein Nummer')]
                datum = line[header.index('BuT Frist')]
                if datum:
                    eater.but_ende = datum
                else:
                    print 'Bitte BuT Frist für \'' + eater.vorname + ' ' + eater.nachname + '\' angeben.'
            # Abrechnungsart
            eater.dauer_last = line[header.index('Dauerlastschrift')] == 'ja' and True
            eater.dauer_last_datum = line[header.index('Dauerlastschrift Datum')]
            # Meckern wenn Dauerlastschrift, aber kein Datum angegeben
            if eater.dauer_last:
                if eater.dauer_last_datum == '':
                    print 'Bitte Dauerlastschrift Datum angeben!'
                    print 'Teilnehmer: ' + eater.vorname + ' ' + eater.nachname
                else:
                    eater.dauer_last_datum = int(eater.dauer_last_datum)
            eater.lastschrift = line[header.index('Lastschrift')] == 'ja' and True
            eater.rechnung = line[header.index('Rechnung')] == 'ja' and True
            # Konto
            eater.kontoinhaber = line[header.index('Kontoinhaber')]
            if not ',' in eater.kontoinhaber:
                print 'Kontoinhaber bitte in der Form Nachname, Vorname eintragen.'
                print 'Fehler bei ' + eater.kontoinhaber
            eater.IBAN = line[header.index('IBAN')]
            eater.BIC = line[header.index('BIC')]
            eater.mail1 = line[header.index('Mail1')]
            if len(line) > header.index('Mail2'):
                eater.mail2 = line[header.index('Mail2')]
            # Teilnehmer zu Dictionary hinzufügen
            eaters.update({ id : eater })
    return eaters

'''
Ermittelt die Anzahl an Essen an denen teilgenommen wurde und ordnet sie den
Stammdaten zu und gibt ein Array zurück welches alle Menschen enthält, die
diesen Monat mindestens einmal mitgegessen habe.
    @param sheet: Worksheet, auf dem die Anzahl der Essen notiert sind
    @param eaters: Dictionary der mitessenden Personen
    @returns: Array der Mitessenden Personen mit Angabe, wie oft sie am Essen
              teilgenommen haben.
'''
def getAmountFromSheet(sheet, eaters):
    # Header auf richtige Reihenfolge überprüfen
    header = ['id', 'Vorname', 'Nachname', 'Anzahl']
    first_row = sheet.row[0]
    for i in range(0, len(header)):
        if header[i] != first_row[i]:
            print u'Falsche Überschrift in Spalte ' + str(i) + ' \'' + first_row[i] + '\''
            print u'Sollte sein: \'' + header[i] + '\'.' 
    # Daten einlesen
    eaters_amount = []
    records = sheet.to_array()
    for i, line in enumerate(records):
        id = line[header.index('id')]
        amount_as_string = line[header.index('Anzahl')]
        # Leere Zeilen und Erste Zeile überspringen
        if id != '' and id != 'id' and amount_as_string != '':
          amount = int(amount_as_string)
          # Teilnehmer mit gleicher id finden und Anzahl der teilgenommenen
          # Essen hinzufügen und an Array anhängen
          if amount > 0:
              eater = eaters.get(id)
              eater.amount = amount
              eaters_amount.append(eater)
    return eaters_amount
