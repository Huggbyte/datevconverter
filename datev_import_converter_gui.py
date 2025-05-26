import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import webbrowser  # Für das Öffnen von Weblinks

# -------- Funktion zur Validierung der Kontonummer -----------
def validiere_gegenkonto(kontonr):
    """
    Prüft, ob die eingegebene Gegenkontonummer eine gültige Zahl zwischen 1200 und 1299 ist.
    Gibt True zurück, wenn ja, sonst False.
    """
    if not kontonr.isdigit():
        return False
    nummer = int(kontonr)
    return 1200 <= nummer <= 1299

# -------- Funktion zur Formatierung des Belegdatums -----------
def belegdatum_fmt(date_str):
    """
    Formatiert das Datum im Format TTMM, z.B. '0105' für den 1. Mai.
    Akzeptiert sowohl datetime-Objekte als auch Strings mit '/' oder '.' als Trenner.
    """
    if pd.isnull(date_str):
        return ""
    if isinstance(date_str, pd.Timestamp):
        return date_str.strftime("%d%m")
    s = str(date_str)
    for sep in ['/', '.']:
        parts = s.split(sep)
        if len(parts) >= 2:
            tag = parts[0].zfill(2)  # 2-stellig mit führender Null
            monat = parts[1].zfill(2)
            return f"{tag}{monat}"
    return ""

# -------- Funktion zur Bereinigung von Beträgen -----------
def clean_betrag(x):
    """
    Formatiert einen Betrag ins deutsche Format mit Komma als Dezimaltrennzeichen.
    Beispiel: 1234.56 -> "1234,56"
    """
    if pd.isnull(x):
        return ""
    val = str(x).replace('.', '').replace(',', '.')
    try:
        return "{:.2f}".format(float(val)).replace('.', ',')
    except:
        return x

# -------- Funktion zur Konvertierung von Amex Excel Dateien -----------
def konvertiere_amex(excel_path, gegenkonto, export_ordner, log_func):
    """
    Konvertiert eine Amex Excel-Datei in DATEV-kompatible CSV-Dateien,
    aufgeteilt nach Monaten.
    """
    log_func(f"Lade Amex-Datei: {excel_path}")
    sheet_name = "Transaktionsdetails"

    # Suche Header-Zeile dynamisch (enthält "Datum")
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    header_idx = None
    for idx in range(len(raw)):
        row = raw.iloc[idx]
        if 'Datum' in row.values:
            header_idx = idx
            break

    # Lade das Blatt mit korrektem Header
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_idx)

    # Datumsspalte als datetime konvertieren
    df["Datum"] = pd.to_datetime(df["Datum"], errors='coerce', dayfirst=True)
    # Monat extrahieren
    df["Monat"] = df["Datum"].apply(lambda x: x.month if not pd.isnull(x) else None)

    # Spalten für Export definieren
    zielspalten = [
        'Währung', 'VorzBetrag', 'RechNr', 'BelegDatum', 'Belegtext', 'UStSatz',
        'BU', 'Gegenkonto', 'Kost1', 'Kost2', 'Kostmenge', 'Skonto', 'Nachricht '
    ]

    # Schreibe für jeden Monat eine eigene CSV-Datei
    for monat in sorted(df["Monat"].dropna().unique()):
        monatsdf = df[df["Monat"] == monat].copy()
        out = pd.DataFrame()

        # Fülle Spalten mit passenden Werten
        out["Währung"] = "EUR"
        out["VorzBetrag"] = monatsdf["Betrag"].apply(clean_betrag)
        out["RechNr"] = ""
        out["BelegDatum"] = monatsdf["Datum"].apply(belegdatum_fmt)
        out["Belegtext"] = monatsdf["Beschreibung"].astype(str).str.strip()
        out["UStSatz"] = 0
        out["BU"] = 9
        out["Gegenkonto"] = gegenkonto
        out["Kost1"] = 1
        out["Kost2"] = ""
        out["Kostmenge"] = 1
        out["Skonto"] = 1
        out["Nachricht "] = out["Belegtext"]

        # Sortiere Spalten in der richtigen Reihenfolge
        out = out[zielspalten]

        # Erstelle Dateiname mit aktuellem Jahr und Monat
        dateiname = f"AMEX_51003_{monat:02d}_{datetime.now().year}_DATEV.csv"
        pfad = os.path.join(export_ordner, dateiname)

        # Speichere CSV-Datei
        out.to_csv(pfad, sep=";", index=False, encoding="utf-8")
        log_func(f"Exportiert: {pfad}")

# -------- Funktion zur Konvertierung von Revolut CSV Dateien -----------
def konvertiere_revolut(csv_path, gegenkonto, export_ordner, log_func):
    """
    Konvertiert Revolut CSV-Dateien in DATEV-kompatible CSV-Dateien,
    aufgeteilt nach Monaten.
    """
    log_func(f"Lade Revolut-Datei: {csv_path}")
    df = pd.read_csv(csv_path)

    # Datum konvertieren und Monat extrahieren
    df["Datum"] = pd.to_datetime(df["Date completed (UTC)"], errors='coerce')
    df["Monat"] = df["Datum"].dt.month

    # Definiere Spalten für Export
    zielspalten = [
        'Währung', 'VorzBetrag', 'RechNr', 'BelegDatum', 'Belegtext', 'UStSatz',
        'BU', 'Gegenkonto', 'Kost1', 'Kost2', 'Kostmenge', 'Skonto', 'Nachricht '
    ]

    # Schreibe CSV pro Monat
    for monat in sorted(df["Monat"].dropna().unique()):
        monatsdf = df[df["Monat"] == monat].copy()
        out = pd.DataFrame()

        out["Währung"] = "EUR"
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
        log_func(f"Exportiert: {pfad}")

# -------- Haupt-GUI-Funktion mit Tkinter --------
def gui_app():
    # Hauptfenster erstellen
    root = tk.Tk()
    root.title("DATEV Import Konverter")

    # Variablen für GUI-Widgets definieren
    konto_var = tk.StringVar(value="Amex")  # Standard Konto Amex
    gegenkonto_var = tk.StringVar(value="1250")  # Standard Gegenkonto
    export_dir_var = tk.StringVar()  # Exportordner (leer)
    input_file_var = tk.StringVar()  # Importdatei (leer)
    lizenz_var = tk.IntVar()  # Für Lizenz-Checkbox

    # Textfeld mit Scrollbar für Log-Ausgaben
    log_text = scrolledtext.ScrolledText(root, width=80, height=20)
    log_text.grid(row=7, column=0, columnspan=3, padx=10, pady=10)

    def log(msg):
        """Schreibt Log-Meldungen in das Textfeld"""
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)

    # Ordner-Auswahl Dialog
    def browse_export_dir():
        folder = filedialog.askdirectory()
        if folder:
            export_dir_var.set(folder)

    # Datei-Auswahl Dialog
    def browse_input_file():
        if konto_var.get() == "Amex":
            filetypes = [("Excel files", "*.xlsx *.xls")]
        else:
            filetypes = [("CSV files", "*.csv")]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            input_file_var.set(filename)

    # Start der Konvertierung mit Validierung und Fehlerbehandlung
    def start_conversion():
        # Lizenzbedingung prüfen
        if lizenz_var.get() != 1:
            messagebox.showwarning("Lizenz", "Bitte akzeptiere die Lizenzbedingungen, um fortzufahren.")
            return

        konto = konto_var.get()
        gegenkonto = gegenkonto_var.get()
        export_ordner = export_dir_var.get()
        input_datei = input_file_var.get()

        # Eingaben validieren
        if not validiere_gegenkonto(gegenkonto):
            messagebox.showerror("Fehler", "Ungültiges Gegenkonto! Bitte eine Zahl zwischen 1200 und 1299 eingeben.")
            return
        if not os.path.isdir(export_ordner):
            messagebox.showerror("Fehler", "Exportordner existiert nicht!")
            return
        if not os.path.isfile(input_datei):
            messagebox.showerror("Fehler", "Import-Datei nicht gefunden!")
            return

        try:
            # Je nach Konto Typ passende Konvertierungsfunktion aufrufen
            if konto == "Amex":
                konvertiere_amex(input_datei, gegenkonto, export_ordner, log)
            else:
                konvertiere_revolut(input_datei, gegenkonto, export_ordner, log)
            messagebox.showinfo("Erfolg", "Konvertierung abgeschlossen!")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}")

    # GUI-Elemente anordnen
    tk.Label(root, text="Konto wählen:").grid(row=0, column=0, sticky="w")
    tk.Radiobutton(root, text="Amex", variable=konto_var, value="Amex").grid(row=0, column=1)
    tk.Radiobutton(root, text="Revolut", variable=konto_var, value="Revolut").grid(row=0, column=2)

    tk.Label(root, text="Gegenkonto (1200-1299):").grid(row=1, column=0, sticky="w")
    tk.Entry(root, textvariable=gegenkonto_var).grid(row=1, column=1, columnspan=2, sticky="ew")

    tk.Label(root, text="Export-Ordner:").grid(row=2, column=0, sticky="w")
    tk.Entry(root, textvariable=export_dir_var, width=50).grid(row=2, column=1)
    tk.Button(root, text="Durchsuchen", command=browse_export_dir).grid(row=2, column=2)

    tk.Label(root, text="Import-Datei:").grid(row=3, column=0, sticky="w")
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=3, column=1)
    tk.Button(root, text="Datei wählen", command=browse_input_file).grid(row=3, column=2)

    # Lizenzbedingungen Checkbox
    tk.Checkbutton(root, text="Ich akzeptiere die Open Source Lizenzbedingungen", variable=lizenz_var).grid(row=4, column=0, columnspan=3, sticky="w")

    # BuyMeACoffee Button mit Link
    def open_buymeacoffee():
        url = "https://buymeacoffee.com/huggbyte"
        webbrowser.open(url)

    tk.Button(root, text="Buy me a coffee ☕", command=open_buymeacoffee).grid(row=5, column=1, pady=5)

    # Start-Button für Konvertierung
    tk.Button(root, text="Konvertierung starten", command=start_conversion).grid(row=6, column=1, pady=10)

    # GUI Hauptloop starten
    root.mainloop()

if __name__ == "__main__":
    gui_app()

