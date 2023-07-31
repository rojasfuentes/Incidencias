import os
import pandas as pd
import openpyxl
import re

patron_mes = r'(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)'
carpeta = r"C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Coecillo-20230711T184742Z-001\Coecillo\GE"
salida = r"C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados\Coecillo\GE"

# Obtener lista de archivos en la carpeta
archivos = os.listdir(carpeta)
i = 1
for archivo in archivos:
    if archivo.endswith(".xlsx"):
        # Ruta completa del archivo de entrada
        path_entrada = os.path.join(carpeta, archivo)

        # Trabajar con el archivo de Excel
        archivo_excel = openpyxl.load_workbook(path_entrada)
        hoja_trabajo = archivo_excel.active

        fecha = hoja_trabajo['G5'].value
        coincidencias = re.findall(patron_mes, fecha.lower())
        if coincidencias:
            fecha = coincidencias[0].capitalize()
        else:
            fecha = 'Error'

        compañia = hoja_trabajo['C9'].value

        archivo_excel.close()

        # Crear dataframe con el archivo Excel
        df = pd.read_excel(path_entrada, skiprows=10, usecols="C:O")

        # Agregar columnas al dataframe
        df['Compañia'] = compañia
        df['Mes'] = fecha
        df['Cede'] = 'Coecillo'

        # Ruta completa del archivo de salida
        path_salida = os.path.join(salida, f"{fecha + '_' + compañia}.xlsx")

        # Guardar dataframe en un nuevo archivo de Excel
        df.to_excel(path_salida, index=False)

        # Imprimir información del archivo procesado
        print("Archivo " + str(i) + " de " + str(len(archivos)) + " "+ fecha + '_' + compañia)
        print("Procesando:", archivo)
        print()
        i += 1

print("Proceso completado.")
