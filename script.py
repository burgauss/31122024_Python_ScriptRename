import pandas as pd
import os

excel_file = "C:/Users/JBURGOS/Documents/07_Tecnologias/16_MantenimientoDeEquipos/NombresArchivos.xlsx"
sheet_name = "Hoja1"
df = pd.read_excel(excel_file, sheet_name=sheet_name)

old_names_column = "NombreArchivo" 
new_names_column = "NumSerie"

directory = "C:/Users/JBURGOS/Documents/07_Tecnologias/16_MantenimientoDeEquipos/BitacoraIndividualWord"

for index, row in df.iterrows():
    old_path = os.path.join(directory, f"{row[old_names_column]}.docx")
    new_path = os.path.join(directory, f"{row[new_names_column]}.docx")
    
    try:
        os.rename(old_path, new_path)
        print(f"Renamed: {old_path} -> {new_path}")
    except FileNotFoundError:
        print(f"File not found: {old_path}")
    except Exception as e:
        print(f"Error renaming {old_path}: {e}")