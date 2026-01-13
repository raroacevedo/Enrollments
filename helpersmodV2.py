#!/usr/bin/python
import json
import pandas as pd
from datetime import datetime
import os
import csv
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet") # Para evitar advertencias de openpyxl: Excel no contiene un "estilo por defecto" definido en sus metadatos

# Cargar configuración desde JSON - directorio es (opcional). Si no existe, se usan valores por defecto.
def load_config(path="config.json"):
    """
    Intenta cargar un archivo JSON de configuración ubicado en el directorio del script.
    Valores soportados (POR DEFECTO):
      {
        "banner_directory": "./",
        "bdusuarios_file": "./BDUsuarios/Listados Usuarios.xlsx",
        "salida_directory": "./salida/"
      }
    Devuelve un dict con la configuración (vacío si no existe o hay errores).
    """

    print("📂 Cargando configuración...")
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()

    cfg_path = os.path.join(base, path)
    if not os.path.isfile(cfg_path):
        # No hay archivo de configuración, devolvemos dict vacío
        return {}

    try:
        with open(cfg_path, encoding="utf8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️ Error al leer la configuración {cfg_path}: {e}")
        return {}

# Cargar la configuración global una sola vez, los directorios y archivos de origen de datos y salida
CONFIG = load_config()

#leer el archivo de NRC/LC: CSV
def leer_nrc():
    """
    Lee el archivo 'ListaCursos.csv' ubicado en el directorio actual.
    Retorna un DataFrame con la información de NRC y nombre de curso.
    
    El archivo debe contener al menos las columnas 'NRC' y 'Nombre_Curso'.
    """
    filename = 'shortnames.csv'
    filepath = os.path.join('.', filename)

    # Validación de existencia del archivo
    if not os.path.isfile(filepath):
        print(f"❌ Error: El archivo requerido '{filename}' no se encuentra en el directorio actual.")
        return None

    # Intentar lectura del archivo
    try:
        print(f"📥 Leyendo archivo: {filename}")
        df = pd.read_csv(
                filepath,
                encoding='utf-8',
                sep=',',
                header=0,
                names=['Nombre', 'NRC', 'Periodo'],
                dtype={'NRC': str}
        )

    except Exception as e:
        print(f"❌ Error al leer el archivo '{filename}': {e}")
        return None

    # Validación de columnas esperadas
    columnas_esperadas = {'Nombre','NRC','Periodo'}
    if not columnas_esperadas.issubset(df.columns):
        print(f"❌ Error: El archivo debe contener las columnas: {columnas_esperadas}. Columnas actuales: {df.columns.tolist()}")
        return None
  
    return df

#Generar el archivo de registro unico (resumen)
def merge_archivos():
    # Usar la ruta de salida desde la configuración si está definida
    base = os.path.dirname(os.path.abspath(__file__))
    directory = CONFIG.get('salida_directory', './salida/')
    if not os.path.isabs(directory):
        directory = os.path.join(base, directory)

    # iteramos sobre los .txt
    for filename in os.listdir(directory):
        if filename.endswith('.txt'):
            with open(directory + filename, encoding='utf8') as fp:
                data = fp.read()

            with open('registro_unicoMOD.txt', 'a', encoding='utf8') as fp:
                fp.write(data)            

    return

#Leer la BD de usuarios de BS: XLSX
def leer_BDUsuarios_BS(ruta_archivo=None):
    """
    Función para cargar un archivo Excel en un DataFrame.
    
    Parámetros:
    ruta_archivo (str): Ruta del archivo Excel a cargar.
    
    Retorna:
    pd.DataFrame: DataFrame con los datos del archivo o None si ocurre un error.
    """
    # Determinar ruta del archivo por la configuración si no se proporcionó una por defecto
    base = os.path.dirname(os.path.abspath(__file__))
    ruta_default = CONFIG.get('bdusuarios_file', "./BDUsuarios/Listados Usuarios.xlsx")
    if ruta_archivo is None:
        ruta_archivo = ruta_default
    # Resolver rutas relativas respecto a la carpeta del script
    if not os.path.isabs(ruta_archivo):
        ruta_archivo = os.path.join(base, ruta_archivo)

    try:
        # Leer el Excel
        df = pd.read_excel(
            ruta_archivo,
            sheet_name=0,        # Leer la primera hoja
        )

        #Se promueven las columnas necesarias
        # Asegurar que los valores sean cadenas de texto y completar con ceros a la izquierda, se asume la longitud de 9
        df=df[['UserName', 'FirstName', 'LastName', 'OrgRoleId']]
        df['UserName'] = df['UserName'].astype(str).str.zfill(9)
        
        print(f"[✓] Archivo '{ruta_archivo}' cargado exitosamente.")
        print(f"El archivo contiene {df.shape[0]} filas y {df.shape[1]} columnas.")
        
        return df

    except FileNotFoundError:
        print(f"❌ Error: El archivo '{ruta_archivo}' no fue encontrado.")
        return None
    except Exception as e:
        print(f"❌ Error al cargar el archivo: {e}")
        return None

#leer el archivo de moderadores a inscribir: EXCEL fuente QLIK
def leer_moderadores(date='nodate'):
    """
    Lee múltiples archivos .xlsx con información de estudiantes desde el directorio origen o actual.
    Optimizado para grandes volúmenes. Aplica limpieza y validación.
    
    Parámetros:
        date (str): Fecha mínima (YYYY-MM-DD) para filtrar la columna 'FECHA_ACTIVIDAD_EST'.
    
    Retorna:
        pd.DataFrame consolidado.
    """
    
    # Directorio con archivos .xlsx — puede definirse en config.json
    base = os.path.dirname(os.path.abspath(__file__))
    directory = CONFIG.get('banner_directory', './')
    if not os.path.isabs(directory):
        directory = os.path.join(base, directory)

    # Validación de existencia del directorio
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"❌ No se encontró el directorio de archivos .xlsx: {directory}")

    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

    # Validación de archivos excel en el directorio
    if not excel_files:
        raise FileNotFoundError("❌ No se encontró ningún archivo .xlsx en el directorio actual.")

    # Definición de solo las columnas necesarias para optimizar la carga del excel
    columnas_objetivo = {
        'PERIODO', 'NRC', 'LISTA_CRUZADA', 'ID_DOCENTE', 'TIPO_DOCUMENTO',
        'DOCUMENTO', 'CORREO_DOCENTE', 'NOMBRE_DOCENTE', 'APELLIDO_DOCENTE',"FECHA_ACTIVIDAD_DOC"
    }
    dataframes = []

    # Iterar sobre cada archivo Excel
    for file in excel_files:
        filepath = os.path.join(directory, file)
        print(f"📥 Leyendo archivo: {filepath}")

        try:
            # Cargar todo el archivo
            df = pd.read_excel(filepath, 
                               sheet_name=0,  # primera hoja (docentes)
                               dtype={'ID_DOCENTE': str},
                               engine='openpyxl'
                )      
        except Exception as e:
            print(f"❌ Error al leer el archivo {file}: {e}")
            continue

        # Validar columnas requeridas existan en el excel
        if not columnas_objetivo.issubset(set(df.columns)):
            faltantes = columnas_objetivo - set(df.columns)
            print(f"⚠️ Advertencia: El archivo {file} no tiene todas las columnas requeridas: {faltantes}")
            continue

        # Filtrar solo las columnas necesarias
        df = df[list(columnas_objetivo)]

        # Limpieza y transformación de datos :espacios, tipos de datos y tipo titulo
        df['ID_DOCENTE'] = df['ID_DOCENTE'].astype(str).str.zfill(9)
        df['LISTA_CRUZADA'] = df['LISTA_CRUZADA'].astype(str)
        df['NOMBRE_DOCENTE'] = df['NOMBRE_DOCENTE'].astype(str).str.strip().str.title()
        df['APELLIDO_DOCENTE'] = df['APELLIDO_DOCENTE'].astype(str).str.strip().str.title()

        #se verifica que las columnas de fecha y periodo existan
        try:
            df['FECHA_ACTIVIDAD_DOC'] = pd.to_datetime(df['FECHA_ACTIVIDAD_DOC'], errors='coerce')
            if date != 'nodate':
                df = df[df['FECHA_ACTIVIDAD_DOC'] >= pd.to_datetime(date)]
        except Exception as e:
            print(f"❌ Error al procesar fechas en {file}: {e}")
            continue

        dataframes.append(df)

    if not dataframes:
        raise ValueError("❌ Ningún archivo válido fue procesado correctamente.")

    # Concatenar todos los DataFrames en uno solo
    resultado = pd.concat(dataframes, ignore_index=True)
    
    #retornar el DataFrame consolidado
    return resultado

#se crea el archivo de registro para cada curso
def crearArchivos(data, course_name, course_nrc, BDUsuBS, log_file_path='log_creacion_moderadores.txt'):
    """
    Genera comandos de inscripción y creación/actualización de moderadores en Brightspace.

    Parámetros:
        data (pd.DataFrame): Datos de los docentes por curso.
        course_name (str): Nombre del curso.
        course_nrc (str): NRC del curso.
        BDUsuBS (pd.DataFrame): Base de usuarios de Brightspace.
        log_file_path (str): Ruta al archivo de log.
    """
    Rol = "Moderador"
    line_count = 0

    # Crear carpeta de salida si no existe
    os.makedirs('./salida', exist_ok=True)
    archivo_comandos = f'./salida/registro_{course_name}.txt'

    # Abrir archivos de salida y log
    with open(archivo_comandos, 'a', encoding='utf8') as fptr, \
         open(log_file_path, 'a', encoding='utf8') as log, \
         open('moderadores.csv', 'a', encoding='utf8', newline='') as moderadores:

        # Escribir encabezados en el archivo de comandos
        writer = csv.writer(moderadores)
        log.write(f"\n=== PROCESAMIENTO CURSO: {course_name} - NRC: {course_nrc} ===\n")
        log.write(f"Fecha: {datetime.now()}\n")

        for _, row in data.iterrows():
            idBanner = str(row.get('ID_DOCENTE', '')).strip()
            
            # Validación de ID_Banner
            if idBanner in ["000000nan","nan", "", "0", "-", "000000000",None]:
                log.write(f"❌ ID inválido: '{idBanner}' para curso {course_name}\n")
                continue
            
            # extraer el rol del moderador y verificar si es nuevo o existente
            Unuevo = idBanner not in BDUsuBS['UserName'].values
            RolModerador = BDUsuBS.loc[BDUsuBS['UserName'] == idBanner, 'OrgRoleId'].values

            # Se formatea el tipo de documento y número de documento
            try:
                ndocu = "{:,}".format(int(row['DOCUMENTO'])).replace(',', '.')
            except:
                ndocu = str(row['DOCUMENTO'])

            try:
                docuusu = f"{row['TIPO_DOCUMENTO']}. {ndocu}"
            except:
                docuusu = ndocu

            #Se limpian los nombres y apellidos
            first_name = str(row.get('NOMBRE_DOCENTE', '')).strip()
            last_name = str(row.get('APELLIDO_DOCENTE', '')).strip()
            email = str(row.get('CORREO_DOCENTE', '')).strip()

            # Validación del usuario: Nuevo o existente
            if Unuevo:
                fptr.write(f'CREATE,{idBanner},{docuusu},{first_name},{last_name},,{Rol},1,{email}\n')
            else:
                fptr.write(f'UPDATE,{idBanner},{docuusu},{first_name},{last_name},,1,{email}\n')
                # Si el usuario ya existe, se cambia el rol a moderador
                if RolModerador.size > 0:
                    rol = str(RolModerador[0])
                    mapeo_roles = {
                        '150': 'CVTE', '143': 'CVLA', '138': 'CVPR',
                        '137': 'CVFC', '136': 'CVFA', '135': 'CVFA'
                    }
                    if rol in mapeo_roles:
                        fptr.write(f'UNENROLL,{idBanner},,{mapeo_roles[rol]}\n')

                fptr.write(f'ENROLL,{idBanner},,{Rol},UPBV\n')
            
            # Inscripción al curso
            fptr.write(f'ENROLL,{idBanner},,{Rol},{course_name}\n')
            line_count += 1

        # Registros de inscripción
        writer.writerow([course_name, course_nrc, line_count])

        print(f"[✓] Se han inscrito: {line_count} moderadores en el curso: {course_name} NRC: {course_nrc}")

        log.write(f"[✓] Total moderadores inscritos: {line_count}\n")
