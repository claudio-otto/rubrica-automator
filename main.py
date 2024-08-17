import zipfile
import os
import openpyxl

def obtener_nombres_alumnos(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        extraction_path = 'extracted_files'
        zip_ref.extractall(extraction_path)
        
        extracted_files = os.listdir(extraction_path)
        nombres_alumnos = [nombre.split('_')[0] for nombre in extracted_files]
        return nombres_alumnos

def copiar_hoja_excel(archivo_excel, hoja_origen, archivo_zip, archivo_salida):
    # Obtener la lista de alumnos desde el archivo ZIP
    alumnos = obtener_nombres_alumnos(archivo_zip)

    # Cargar el archivo original de Excel
    wb_original_v2 = openpyxl.load_workbook(archivo_excel)
    ws_original_v2 = wb_original_v2[hoja_origen]

    # Crear una copia de la hoja original para cada alumno dentro del mismo archivo
    for alumno in alumnos:
        ws_nueva_v2 = wb_original_v2.copy_worksheet(ws_original_v2)
        ws_nueva_v2.title = alumno[:31]  # Asegurarse de que el nombre no exceda 31 caracteres

    # Guardar el archivo con todas las copias
    wb_original_v2.save(archivo_salida)

if __name__ == "__main__":
    archivo_excel = 'Rubrica Evaluacion Final Precio Propiedades (1).xlsx'
    hoja_origen = 'Hoja1'
    archivo_zip = 'BOTIC-SOFOF-23-30-13-0006-M3 - EVALUACIÓN FINAL DEL MÓDULO-52388.zip'
    archivo_salida = 'Final_Rubrica_Alumnos_v3.xlsx'
    
    copiar_hoja_excel(archivo_excel, hoja_origen, archivo_zip, archivo_salida)
