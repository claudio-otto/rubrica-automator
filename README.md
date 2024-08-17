# RubricaAutomator

**RubricaAutomator** es una herramienta que automatiza la creación de hojas personalizadas en Excel para cada estudiante, 
copiando una rúbrica predeterminada y facilitando la gestión de evaluaciones.

## Requisitos

- Python 3.x
- Librería openpyxl

## Uso

1. Coloca el archivo original de Excel en el mismo directorio del proyecto.
2. Coloca el archivo ZIP con los trabajos de los alumnos en el mismo directorio.
3. Abre el archivo `main.py` y ajusta los siguientes parámetros si es necesario:
   - `archivo_excel`: Nombre del archivo Excel que contiene la hoja a copiar.
   - `hoja_origen`: Nombre de la hoja que será copiada.
   - `archivo_zip`: Nombre del archivo ZIP que contiene los nombres de los estudiantes.
   - `archivo_salida`: Nombre del archivo Excel que se generará con las copias.
4. Ejecuta el script `main.py`.
5. Se generará un nuevo archivo Excel con las hojas copiadas, llamado según el valor de `archivo_salida`.

## Estructura del Proyecto

- `main.py`: Código principal que realiza la lectura del archivo ZIP, extrae los nombres de los alumnos, y copia la hoja del archivo Excel.
- `README.md`: Este archivo de documentación.
- `requirements.txt`: Lista de dependencias necesarias para ejecutar el proyecto.
