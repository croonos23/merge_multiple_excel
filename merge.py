import pandas as pd
import glob
import os


ruta_actual = os.path.dirname(os.path.abspath(__file__))

archivos_excel = []

ruta_archivos = ruta_actual+"\\*.xlsx"
for archivo in glob.glob(ruta_archivos):
    archivos_excel.append(archivo)

datos_finales = {}
for archivo in archivos_excel:

    xls = pd.ExcelFile(archivo)
    for hoja in xls.sheet_names:
        
        datos_hoja = pd.read_excel(archivo, sheet_name=hoja)
        if hoja in datos_finales:
            datos_finales[hoja] = pd.concat([datos_finales[hoja], datos_hoja])
        else:
            datos_finales[hoja] = datos_hoja
        
archivo_final = pd.ExcelWriter(ruta_actual+'\\archivo_final.xlsx')

for hoja, datos in datos_finales.items():
    datos.to_excel(archivo_final, sheet_name=hoja, index=False)

archivo_final.save()
