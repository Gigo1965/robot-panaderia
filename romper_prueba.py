"""Rompe A PROPOSITO las formulas de NODO INTEGRAL C19/C20 en la copia de
prueba, quitando el prefijo _xlfn. (igual que lo hizo Google Sheets).
Antes de romper, guarda un respaldo de la copia de prueba.
Nunca toca el Excel real de Drive.
"""
import shutil

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

ruta = 'datos_prueba/maestro_prueba.xlsx'
respaldo = 'datos_prueba/maestro_prueba_RESPALDO.xlsx'

shutil.copy2(ruta, respaldo)
print('Respaldo guardado en', respaldo)

wb = openpyxl.load_workbook(ruta)
ws = wb['NODO INTEGRAL']

rotas = 0
for celda in ('C19', 'C20'):
    valor = ws[celda].value
    if isinstance(valor, ArrayFormula):          # formula matricial
        antes = valor.text
        valor.text = antes.replace('_xlfn.', '')
        ahora = valor.text
    elif isinstance(valor, str) and valor.startswith('='):   # formula normal
        antes = valor
        ws[celda] = antes.replace('_xlfn.', '')
        ahora = ws[celda].value
    else:
        print(celda, 'no tiene formula, se salta')
        continue
    print(celda, 'antes:', antes[:60])
    print(celda, 'ahora:', ahora[:60])
    rotas += 1

wb.save(ruta)
print(f'{rotas} formula(s) rotas a proposito (igual que lo hizo Google Sheets)')
