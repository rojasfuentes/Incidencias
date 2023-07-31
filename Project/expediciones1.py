import os
import pandas as pd

def combine_excel_files(folder_path, output_file):
    file_list = os.listdir(folder_path)
    excel_files = [f for f in file_list if f.endswith('.xlsx') or f.endswith('.xls')]
    
    combined_data = pd.DataFrame()
    
    for file in excel_files:
        # Leer cada archivo de Excel en un DataFrame separado
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        
        # Combinar el DataFrame 
        combined_data = pd.concat([combined_data, df], ignore_index=True)
    
    
    output_path = os.path.join(folder_path, output_file)
    combined_data.to_excel(output_path, index=False)
    print("Completado. El resultado se ha guardado en:", output_path)


# Ejemplo de uso
folder_path = r'C:\Users\JFROJAS\Desktop\Consolidado Incidencias\Resultados'  # Reemplaza con la ruta de tu carpeta
output_file = '0IncidenciasSemana_30.xlsx'  # Nombre del archivo de salida
combine_excel_files(folder_path, output_file) 
