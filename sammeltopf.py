#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import pyexcel
import pyexcel.ext.ods3
import datetime
from xhtml2pdf import pisa
import codecs
import locale

import eater


# Prüfen der Argumente
if len(sys.argv) < 2:
    print 'Bitte Datei mit Berechnungsgrundlagen angeben!'
    print 'USAGE: sammeltopf.py [Daten.ods] [Test]'
    sys.exit()
else:
    data_file = sys.argv[1]
if len(sys.argv) > 2:
    test = True

# Locale setzen
locale.setlocale(locale.LC_ALL, 'de_DE.utf8')

# Abrechnungsmonat ermitteln und bestätigen
today = datetime.datetime.today()
month = today.month - 1
# Wenn Dezember abgerechnet wird, müssen wir auch ein Jahr zurück gehen.
year = month == 12 and today.year - 1 or today.year
print 'Abrechnungsmonat: ' + str(month) + ' / ' + str(year)
ok = raw_input('OK? [Enter/n]')
if ok != '':
    month = int(input("Bitte Monat eingeben: "))
    year = int(input("Bitte Jahr eingeben: "))

# Öffnen der Datei
book = pyexcel.get_book(file_name=data_file)
try:
    eater_data = book['Stammdaten']
    amount_data = book['Anzahl']
    price_data = book['Preis']
except:
    print "Bitte überprüfen, ob alle Arbeitsmappen vorhanden sind: "
    print "Stammdaten / Anzahl / Preis\n"
    sys.exit(0)

# Importiere Stammdaten, Teilnahme am Essen und Essenspreis
eaters = eater.getEaterFromSheet(eater_data)
eaters_amount = eater.getAmountFromSheet(amount_data, eaters)
price = float(price_data.row[0][0])

# Arrays für Speicherung der Buchungsdaten
header = ['Name Kind', 'Kontoinhaber', 'Verwendungszweck', 'IBAN', 'BIC', 'Betrag']
header_but = ['Name Kind', 'Gutschein Nr.', 'Betrag']
lastschrift = [header]
dauerlastschrift = [header]
rechnung = [header]
but_job = [header_but]
but_wg = [header_but]

# Erzeuge Rechnungen
print 'Erzeuge Rechnungen.'
for i, eater in enumerate(eaters_amount):
    # Vorlage einlesen
    with codecs.open('rechnung.html','r', encoding='utf-8') as template_file:
        template = template_file.read()
    # print 'Kann Rechnungsvorlage \'rechnung.html\' nicht finden.'
    # sys.exit()
    ''' Daten vorbereiten '''
    # Name in Anschrift
    # Wenn Kontoinhaber gleich Teilnehmer => Pädagoge => Name statt Familie
    if eater.kontoinhaber and eater.vorname == eater.kontoinhaber.split(',')[1].strip(' '):
        familie = eater.vorname + ' ' + eater.nachname
    else:
        familie = 'Familie ' + eater.nachname
        familie += eater.nachname2 and ' / ' + eater.nachname2 or ''
    name = eater.vorname + ' ' + eater.nachname
    # Rechnungsnummer
    nummer = 'ES ' + str(year) + '-' + str(month) + '-' + str(i + 1)
    teilsumme = price * eater.amount
    verwendungszweck = nummer + ' ' + eater.nachname + ', ' + eater.vorname

    text = ''
    replacer = []
    # Ist BuT-Gutschein abgelaufen?
    if eater.but in ['wg','job']:
        # Nimm den ersten des Abrechnungsmonat als Referenz
        erster = datetime.date(year, month, 1)
        # Falls Frist noch nicht zuende ist, setze Summe
        if (eater.but_ende - erster) > datetime.timedelta(days = 1):
            but_summe = (price - 1.0) * eater.amount
            # Zeile aktivieren indem der Kommentar entfernt wird.
            replacer.append(['!--but ',''])
            replacer.append([' but--',''])
        else:
            text += u'Der Betrag für Bildung und Teilhabe kann nicht abgezogen '
            text += u'werden, weil kein aktueller Gutschein vorliegt. ' 
            but_summe = 0
    else:
        but_summe = 0

    # Summen als Strings im Währungsformat
    betrag = teilsumme - but_summe
    teilsumme_str = locale.currency(teilsumme).decode('utf-8')
    butsumme_str = locale.currency(but_summe).decode('utf-8')
    betrag_str = locale.currency(betrag).decode('utf-8')
    
    # Buchungstext und Sammeln der Buchungsdaten
    if eater.rechnung:
        text += u'''Bitte überweisen Sie den angefallenen Rechnungsbetrag
                   innerhalb von 14 Tagen!'''
        rechnung.append([name, eater.kontoinhaber, verwendungszweck, eater.IBAN,
            eater.BIC, betrag]) 
    elif eater.lastschrift:
        text += u'Der Einzug des Rechnungsbetrages erfolgt per Lastschrift.'
        lastschrift.append([name, eater.kontoinhaber, verwendungszweck, eater.IBAN,
            eater.BIC, betrag])
    elif eater.dauer_last:
        text += u'''Der Einzug des festgelegten Betrages erfolgt per Lastschrift
                 zum ''' + str(eater.dauer_last_datum) + '. des Monats.'
        dauerlastschrift.append([name, eater.kontoinhaber, verwendungszweck, eater.IBAN,
            eater.BIC, betrag])
   
    # Füge Buchung für BuT hinzu
    if but_summe != 0:
        if eater.but == 'wg':
            but_wg.append([name, ' ', but_summe])
        elif eater.but == 'job':
            but_job.append([name, eater.but_gutschein, but_summe])
    
    # Daten einsetzen
    replacer.append(['__familie__', familie])
    replacer.append(['__strasse__', eater.strasse])
    replacer.append(['__plz__', eater.plz])
    replacer.append(['__stadt__', eater.stadt])
    replacer.append(['__monat__', str(month)])
    replacer.append(['__jahr__', str(year)])
    replacer.append(['__datum__', today.strftime('%d.%m.%Y')])
    replacer.append(['__nummer__', nummer])
    replacer.append(['__name__', eater.nachname + ', ' + eater.vorname])
    replacer.append(['__anzahl__', str(eater.amount)])
    replacer.append(['__einzelpreis__', locale.currency(price).decode('utf-8')])
    replacer.append(['__teilsumme__', teilsumme_str])
    replacer.append(['__butsumme__', butsumme_str])
    # TODO: Rücklastschrift
    # template.replace('__ruecklastschrift__', 'nein')
    # template.replace('__ruecklastschriftsumme__', u'0,00 €')
    replacer.append(['__betrag__', betrag_str])
    replacer.append(['__buchungstext__', text])
    for item in replacer:
        template = template.replace(item[0], item[1])
    # HTML in PDF umwandeln
    pdf_file_name = eater.nachname + '-' + eater.vorname + '_' + str(year) + '-'
    pdf_file_name += str(month) + '.pdf'
    with open(pdf_file_name,'w+b') as pdf_file:
        pisa.CreatePDF(template, dest=pdf_file)
    # Verschicke Rechnungen per Mail
# Buchungsdaten speichern
print 'Speichere Buchungsdaten.'
buchungsdaten = {
        'Lastschrift' : lastschrift,
        'Rechnung' : rechnung,
        'Dauerlastschrift' : dauerlastschrift,
        'BuT Job' : but_job,
        'BuT WG' : but_wg,
        }
buchungsdaten_book = pyexcel.Book(buchungsdaten)
filename = str(year) + '-' + str(month) + '-Buchungsdaten.ods'
buchungsdaten_book.save_as(filename)
print 'Fertig.'
