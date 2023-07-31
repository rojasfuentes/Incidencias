import pandas as pd
from datetime import datetime


path = r'C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados\consolidado.xlsx'
df = pd.read_excel(path)

# Eliminar filas con fechas en blanco
df = df.dropna(subset=['Fecha'])

# Convertir la columna 'Fecha'
df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %I:%M:%S %p', errors='coerce')

#'Semana'
def obtener_semana(fecha):
    if pd.isnull(fecha):
        return None
    else:
        return fecha.isocalendar()[1]

df['Semana'] = df['Fecha'].apply(obtener_semana)


print(df.head())

df.to_excel(r'C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados\consolidado_semanas.xlsx', index=False)
print("Proceso completado.")
