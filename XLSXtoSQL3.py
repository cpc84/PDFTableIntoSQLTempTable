# pip install openpyxl
# pip install pandas
# pip install tabula-py
# pip install tqdm
# pip install xlrd

import tabula
import pandas as df
import os
import io
import sys
import getopt
from datetime import datetime
from numpy import nan
import tkinter as tk
from tkinter import filedialog 

root = tk.Tk()

filenameAndPath = ""
filename = ""

def callback():
    initdir = os.getcwd() # "/"
    filenameAndPath = filedialog.askopenfilename(initialdir = initdir, title = "Datei öffnen",filetypes = (("PDF files","*.pdf"),("CSV files","*.csv"),("XLS files","*.xls"),("XLSX files","*.xlsx")))
    eingabefeld_wert1.set(filenameAndPath)
    print(filenameAndPath)

def aktionSF():
    # input
    bereich = str(eingabefeld_wert2.get()) #"10-15"
    inputFileName = str(eingabefeld_wert1.get()) #"I03-Fachlosspezifikation und Preisanker.pdf"#"I04-Kassenabwahl.pdf" #"Komplettdatei.xls" # "Umsatzdaten Apotheken 26.-30. September 2020.xls" #"Komplettdatei.xls" #"I04-Kassenabwahl.pdf"
    filename = inputFileName[::-1] # Reverse
    filename = inputFileName.replace(filename[filename.find("/"):len(inputFileName)][::-1], "")
    
    # globale Variablen
    filenameSQL = ""
    spalten = ""
    values = ""
    tempdf = df

    # zeitstempel generieren
    now = datetime.now()
    timestamp = str(now.strftime("%m%d%Y%H%M%S"))
    #print(timestamp)	

    # maximale spaltenlänge definieren
    df.set_option("display.max_colwidth", -1)

    # Tabelle aus PDF-Datei extrahieren   
    if ".pdf" in inputFileName:
        print("pdf " + filename)
        # csv filename generieren
        FilenameCsv = timestamp + "_" + filename.replace(".pdf", ".csv")
        # tabelle aus pdf-datei in csv konvertieren
        tabula.convert_into(inputFileName, FilenameCsv, output_format="csv", pages=bereich, lattice=True, stream=True)
        # csv-datei in dataframe einlesen 
        tempdf = df.read_csv(FilenameCsv, encoding = "ISO-8859-1", engine='python', header=1) # , usecols=[0, 1]
        # xlsx filename generieren
        filenameXls = FilenameCsv.replace(".csv", ".xlsx")
        # xlsx-datei generieren
        tempdf.to_excel(filenameXls, sheet_name='vergleich', index=True)
        # SQL filename generieren
        filenameSQL = filenameXls.replace(".xlsx", ".sql")
        # xlsx-datei in dataframe einlesen
        tempdf = df.read_excel(filenameXls) 

    # csv-datei einlesen
    if ".csv" in inputFileName:
        print("csv " + filename)
        # csv-datei in dataframe einlesen
        tempdf = df.read_csv(inputFileName, encoding = "ISO-8859-1", engine='python', header=1) # , usecols=[0, 1]
        # xlsx filename generieren
        filenameXls = timestamp + "_" + inputFileName.replace(".csv", ".xlsx")
        # xlsx-datei generieren
        tempdf.to_excel(filenameXls, sheet_name='vergleich', index=True)
        # SQL filename generieren
        filenameSQL = filenameXls.replace(".xls", ".sql")
        # xlsx-datei in dataframe einlesen
        tempdf = df.read_excel(filenameXls) 

    # Xlsx-datei einlesen
    if ".xls" in inputFileName:
        print("xls " + filename)
        # xlsx filename generieren
        filenameXls = filename
        # SQL filename generieren
        filenameSQL = timestamp + "_" + filename.replace(".xls", ".sql")
        # xlsx-datei in dataframe einlesen
        tempdf = df.read_excel(inputFileName) 

    if ".xlsx" in inputFileName:
        print("xlsx " + filename)
        # xlsx filename generieren
        filenameXls = filename
        # SQL filename generieren
        filenameSQL = timestamp + "_" + filename.replace(".xlsx", ".sql")
        # xlsx-datei in dataframe einlesen
        tempdf = df.read_excel(inputFileName) 

    # Wenn Spaltennamen nicht gefüllt sind
    if "Unnamed" in str(tempdf.columns) and not tempdf.iloc[0].isnull().values.any():
        # column names from 1 row
        tempdf.columns = tempdf.iloc[0]
        # erste zeile löschen
        tempdf = tempdf[1:]

    # NAN feldern mit NULL ersetzen
    tempdf = tempdf.fillna('NULL')

    # SQL-Datei löschen, wenn existiert
    if os.path.exists(filenameSQL):
          os.remove(filenameSQL)

    # replace new line
    tempdf.replace(to_replace=[r"\\t|\\n|\\r", "\t|\n|\r"], value=[""," "], regex=True, inplace=True)

    # replace '
    tempdf.replace(to_replace=["'", "'"], value=[""," "], regex=True, inplace=True)

    # SQL-Datei befüllen
    f = open(filenameSQL, "a", encoding="utf-8") 

    f.write("IF OBJECT_ID('tempdb..#Temp', 'U') IS NOT NULL DROP TABLE #Temp\n\r")

    # spalten mit werte befüllen
    for col in tempdf.columns:
        spalten += "[" + str(col).replace("\n","_") + "] VARCHAR(MAX),"
        values += "'" + tempdf[col].astype(str) + "',"

    # letzte char abschneiden
    spalten = spalten[:-1]
    values = "SELECT " + values.astype(str).str[:-1] + " UNION ALL"

    f.write("CREATE TABLE #Temp(" + spalten + ")\n\r")
    f.write("INSERT INTO #Temp\n\r")
    f.write("" + values.to_string(index=False).rstrip() + "SELECT * FROM #Temp")

    f.close()

    # SQL-Datei öffnen zum korrektur
    with open(filenameSQL, 'r') as f:
         txt = f.read().replace('UNION ALLSELECT *', '\n\rSELECT *').replace("'NULL'","NULL").replace("Unnamed: ","Col")

    # Korrektur in SQL-Datei Schreiben
    with open(filenameSQL, 'w') as f:
       f.write(txt)

    label3 = tk.Label(root, text="Aktion durchgeführt", bg="yellow")
    label3.pack()

# Textausgabe erzeugen
label1 = tk.Label(root, text="dateiname", fg="black")
# in GUI Elemente einbetten
label1.pack()

errmsg = 'Error!'
schaltf2 = tk.Button(text='File Open', command=callback)
schaltf2.pack()

eingabefeld_wert1=tk.StringVar()
eingabefeld_1=tk.Entry(root, textvariable=eingabefeld_wert1)
eingabefeld_1.pack()
eingabefeld_wert1.set('pfad')

# Textausgabe erzeugen
label2 = tk.Label(root, text="Bereich:", fg="black")
label2.pack()

eingabefeld_wert2=tk.StringVar()
eingabefeld=tk.Entry(root, textvariable=eingabefeld_wert2)
eingabefeld.pack()
eingabefeld_wert2.set('all')

schaltf1 = tk.Button(root, text="Aktion durchführen", command=aktionSF)
schaltf1.pack()

root.mainloop()