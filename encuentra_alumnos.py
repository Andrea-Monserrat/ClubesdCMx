import re
import openpyxl
from unicodedata import normalize

doc = openpyxl.load_workbook('participantes_totales.xlsx')
conacyt = openpyxl.load_workbook('Listas_CONACYT.xlsx')
hoja = doc.get_sheet_by_name('Hoja1')
Nombres= []
for fila in hoja.rows:
    for columna in fila:
        s=columna.value
        s = re.sub(
            r"([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+", r"\1",
            normalize( "NFD", s), 0, re.I
            )
        s = normalize( 'NFC', s)
        Nombres.append(s)

k17=[]
diecisiete = conacyt.get_sheet_by_name('2017')
for fila in diecisiete.rows:
    for columna in fila:
        w=columna.value
        if w in Nombres:
            k17.append(w)
k18=[]
dieciocho = conacyt.get_sheet_by_name('2018')
for fila in dieciocho.rows:
    for columna in fila:
        w=columna.value
        if w in Nombres:
            k18.append(w)
k19=[]
diecinueve = conacyt.get_sheet_by_name('2019')
for fila in diecinueve.rows:
    for columna in fila:
        w=columna.value
        if w in Nombres:
            k19.append(w)
