# Rubrica Proyecto

Este proyecto de Python permite copiar una hoja de un archivo Excel original a múltiples hojas dentro del mismo archivo, 
cada una con el nombre de un alumno, preservando el contenido y formato original.

Además, el proyecto incluye una función para leer un archivo ZIP, extraer los nombres de los alumnos sin sus identificadores, 
y utilizar estos nombres para crear las nuevas hojas en el archivo Excel.

## Requisitos

- Python 3.x
- Librería openpyxl

## Uso

1. Coloque el archivo original llamado `Rubrica Evaluacion Final Precio Propiedades (1).xlsx` en el mismo directorio.
2. Coloque el archivo ZIP con los trabajos de los alumnos llamado `BOTIC-SOFOF-23-30-13-0006-M3 - EVALUACIÓN FINAL DEL MÓDULO-52388.zip` en el mismo directorio.
3. Ejecute el script `main.py`.
4. Se generará un nuevo archivo Excel con las hojas copiadas, llamado `Final_Rubrica_Alumnos_v3.xlsx`.

## Estructura del Proyecto

- `main.py`: Código principal que realiza la lectura del archivo ZIP, extrae los nombres de los alumnos y copia la hoja del archivo Excel.
- `README.md`: Este archivo de documentación.
- `requirements.txt`: Lista de dependencias necesarias para ejecutar el proyecto.
