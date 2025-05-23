import os
import pandas as pd
from datetime import datetime

# Funktion zur Validierung der Gegenkontonummer
# Erlaubt nur Zahlen im Bereich 1200 bis 1299 (SKR03 typische Bankkonten)
def validiere_gegenkonto(kontonr):
    if not kontonr.isdigit():
        return False
    nummer = int(kontonr)
    return 1200 <= nummer <= 1299

# Funktion zur Formatierung des Belegdatums im Format TTMM (z.B. 0105 für 1. Mai)
def belegdatum_fmt(date_str):
    if pd.isnull(date_str):
        return ""
    if isinstance(date_str, pd.Timestamp):
        return date_str.strftime("%d%m")
    s = str(date_str)
    for sep in ['/', '.']:
        parts = s.split(sep)
        if len(parts) >= 2:
            tag = parts[0].zfill(2)  # Tag 2-stellig
            monat = parts[1].zfill(2)  # Monat 2-stellig
            return f"{tag}{monat}"
    return ""

# Funktion zur Bereinigung und Umformatierung von Beträgen
# Wandelt z.B. 1234.56 in "1.234,56" um für das deutsche DATEV-Format
def clean_betrag(x):
    if pd.isnull(x):
        return ""
    val = str(x).replace('.', '').replace(',', '.')
    try:
        return "{:.2f}".format(float(val)).replace('.', ',')
    except:
        return x

# Konvertierungsfunktion für Amex Excel Dateien
# Liest die Datei ein, extrahiert relevante Spalten und speichert pro Monat DATEV-CSV
def konvertiere_amex(excel_path, gegenkonto, export_ordner):
    print(f"Lade Amex-Datei: {excel_path}")
    sheet_name = "Transaktionsdetails"

    # Suche Header-Zeile dynamisch (Zeile mit "Datum")
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    header_idx = None
    for idx in range(len(raw)):
        row = raw.iloc[idx]
        if 'Datum' in row.values:
            header_idx = idx
            break

    # Lade das Datenblatt mit gefundenem Header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_idx)

    # Konvertiere Datumsspalte in datetime-Objekte
    df["Datum"] = pd.to_datetime(df["Datum"], errors='coerce', dayfirst=True)

    # Extrahiere Monat aus dem Datum für spätere Gruppierung
    df["Monat"] = df["Datum"].apply(lambda x: x.month if not pd.isnull(x) else None)

    # Definiere Spalten für die DATEV-Ausgabe in der richtigen Reihenfolge
    zielspalten = [
        'Währung', 'VorzBetrag', 'RechNr', 'BelegDatum', 'Belegtext', 'UStSatz',
        'BU', 'Gegenkonto', 'Kost1', 'Kost2', 'Kostmenge', 'Skonto', 'Nachricht '
    ]

    # Pro Monat eine eigene CSV-Datei erstellen
    for monat in sorted(df["Monat"].dropna().unique()):
        monatsdf = df[df["Monat"] == monat].copy()
        out = pd.DataFrame()

        # Fülle die Spalten mit den passenden Werten
        out["Währung"] = "EUR"
        out["VorzBetrag"] = monatsdf["Betrag"].apply(clean_betrag)
        out["RechNr"] = ""
        out["BelegDatum"] = monatsdf["Datum"].apply(belegdatum_fmt)
        out["Belegtext"] = monatsdf["Beschreibung"].astype(str).str.strip()
        out["UStSatz"] = 0
        out["BU"] = 9  # Buchungsschlüssel
        out["Gegenkonto"] = gegenkonto
        out["Kost1"] = 1
        out["Kost2"] = ""
        out["Kostmenge"] = 1
        out["Skonto"] = 1
        out["Nachricht "] = out["Belegtext"]

        # Sortiere die Spalten passend zum DATEV-Format
        out = out[zielspalten]

        # Dateiname mit Jahr und Monat
        dateiname = f"AMEX_51003_{monat:02d}_{datetime.now().year}_DATEV.csv"
        pfad = os.path.join(export_ordner, dateiname)

        # Speichere CSV-Datei mit Semikolon als Trenner und UTF-8 Kodierung
        out.to_csv(pfad, sep=";", index=False, encoding="utf-8")
        print(f"Exportiert: {pfad}")

# Konvertierungsfunktion für Revolut CSV Dateien
# Funktioniert analog zu Amex, nur mit CSV Input und anderen Spaltennamen
def konvertiere_revolut(csv_path, gegenkonto, export_ordner):
    print(f"Lade Revolut-Datei: {csv_path}")
    df = pd.read_csv(csv_path)
    df["Datum"] = pd.to_datetime(df["Date completed (UTC)"], errors='coerce')
    df["Monat"] = df["Datum"].dt.month

    zielspalten = [
        'Währung', 'VorzBetrag', 'RechNr', 'BelegDatum', 'Belegtext', 'UStSatz',
        'BU', 'Gegenkonto', 'Kost1', 'Kost2', 'Kostmenge', 'Skonto', 'Nachricht '
    ]

    for monat in sorted(df["Monat"].dropna().unique()):
        monatsdf = df[df["Monat"] == monat].copy()
        out = pd.DataFrame()
        out["Währung"] = "EUR"
        # Betrag als String mit Komma statt Punkt
        out["VorzBetrag"] = monatsdf["Amount"].apply(lambda x: f"{x:.2f}".replace('.', ','))
        out["RechNr"] = ""
        out["BelegDatum"] = monatsdf["Datum"].apply(belegdatum_fmt)
        out["Belegtext"] = monatsdf["Description"].astype(str).str.strip()
        out["UStSatz"] = 0
        out["BU"] = 9
        out["Gegenkonto"] = gegenkonto
        out["Kost1"] = 1
        out["Kost2"] = ""
        out["Kostmenge"] = 1
        out["Skonto"] = 1
        out["Nachricht "] = out["Belegtext"]
        out = out[zielspalten]

        dateiname = f"REVOLUT_{datetime.now().year}_{monat:02d}_DATEV.csv"
        pfad = os.path.join(export_ordner, dateiname)
        out.to_csv(pfad, sep=";", index=False, encoding="utf-8")
        print(f"Exportiert: {pfad}")

# Hauptfunktion, die den Nutzer interaktiv führt
def main():
    print("Welches Konto möchtest du importieren? (Amex/Revolut)")
    konto = input().strip().lower()

    print("Gegenkonto (Standard 1250):")
    gegenkonto = input().strip()
    if gegenkonto == "":
        gegenkonto = "1250"
    while not validiere_gegenkonto(gegenkonto):
        print("Ungültiges Gegenkonto. Bitte eine Zahl zwischen 1200 und 1299 eingeben:")
        gegenkonto = input().strip()

    print("Gib den Ordner an, in den exportiert werden soll:")
    export_ordner = input().strip()
    if not os.path.isdir(export_ordner):
        print("Ordner existiert nicht. Erstelle Ordner...")
        os.makedirs(export_ordner)

    if konto == "amex":
        print("Gib den Pfad zur Amex Excel-Datei an (Standard: amex.xlsx):")
        datei = input().strip()
        if datei == "":
            datei = os.path.join(os.getcwd(), "amex.xlsx")
        konvertiere_amex(datei, gegenkonto, export_ordner)
    elif konto == "revolut":
        print("Gib den Pfad zur Revolut CSV-Datei an (Standard: revolut.csv):")
        datei = input().strip()
        if datei == "":
            datei = os.path.join(os.getcwd(), "revolut.csv")
        konvertiere_revolut(datei, gegenkonto, export_ordner)
    else:
        print("Unbekanntes Konto. Bitte 'Amex' oder 'Revolut' eingeben.")

if __name__ == "__main__":
    main()

