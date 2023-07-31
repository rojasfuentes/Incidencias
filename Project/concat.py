import os
import pandas as pd

carpeta_raiz = r"C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados"
dataframes = []
contador_archivos = 0

# Recorrer la estructura de carpetas y subcarpetas
for carpeta_cede in os.listdir(carpeta_raiz):
    carpeta_cede_path = os.path.join(carpeta_raiz, carpeta_cede)
    if os.path.isdir(carpeta_cede_path):
        for carpeta_cliente in os.listdir(carpeta_cede_path):
            carpeta_cliente_path = os.path.join(carpeta_cede_path, carpeta_cliente)
            if os.path.isdir(carpeta_cliente_path):
                for archivo in os.listdir(carpeta_cliente_path):
                    if archivo.endswith(".xlsx"):
                        archivo_path = os.path.join(carpeta_cliente_path, archivo)
                        # Leer el archivo de Excel y agregarlo al listado de dataframes
                        df = pd.read_excel(archivo_path)
                        dataframes.append(df)
                        
                        
                        contador_archivos += 1

# Combinar todos
df_final = pd.concat(dataframes, ignore_index=True)
print(df_final.head())

df_final.to_excel(r"C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados\consolidado.xlsx", index=False)

print("Proceso completado.")
print("Cantidad de archivos procesados:", contador_archivos)

