import argparse
import shutil
import unicodedata
import zipfile
import os
import openpyxl
from pathlib import Path

def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def obtener_nombres_alumnos(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        extraction_path = 'extracted_files'

        # Eliminar el directorio si ya existe
        if os.path.exists(extraction_path):
            shutil.rmtree(extraction_path)

        # Crear el directorio nuevamente para extraer los archivos
        os.makedirs(extraction_path)

        zip_ref.extractall(extraction_path)
        extracted_files = os.listdir(extraction_path)
        nombres_alumnos = [nombre.split('_')[0] for nombre in extracted_files]
        nombres_alumnos.sort(key=remove_accents)
        return nombres_alumnos


def copiar_hoja_excel(ruta, archivo_excel, hoja_origen, archivo_zip):
    path = Path(ruta)
    # Obtener la lista de alumnos desde el archivo ZIP
    alumnos = obtener_nombres_alumnos(path / archivo_zip)

    # Cargar el archivo original de Excel
    wb_original_v2 = openpyxl.load_workbook(path / archivo_excel)
    ws_original_v2 = wb_original_v2[hoja_origen]

    # Crear una copia de la hoja original para cada alumno dentro del mismo archivo
    for alumno in alumnos:
        ws_nueva_v2 = wb_original_v2.copy_worksheet(ws_original_v2)
        ws_nueva_v2.title = alumno[:31]  # Asegurarse de que el nombre no exceda 31 caracteres

    # Guardar el archivo con todas las copias
    wb_original_v2.save(path / "evaluaciones.xlsx")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automatiza la creación de hojas de Excel personalizadas.")
    parser.add_argument('--ruta', type=str, required=True,
                        help="Ruta con los archivos a leer.")
    parser.add_argument('--zip', type=str, required=True,
                        help="Nombre del archivo ZIP que contiene los nombres de los estudiantes.")
    parser.add_argument('--rubrica', type=str, required=True,
                        help="Nombre del archivo Excel que contiene la hoja a copiar.")
    parser.add_argument('--hoja', type=str, required=True, help="Nombre de la hoja que será copiada.")

    args = parser.parse_args()

    copiar_hoja_excel(args.ruta, args.rubrica, args.hoja, args.zip)
