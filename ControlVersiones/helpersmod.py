#!/usr/bin/python

import pandas as pd
from datetime import datetime
import os
import csv
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet") # Para evitar advertencias de openpyxl: Excel no contiene un "estilo por defecto" definido en sus metadatos

#leer el archivo de NRC/LC: CSV
def leer_nrc():
    '''
    Lee un archivo CSV que debe estar en el mismo directorio desde el que se ejecuta el script.
    Devuelve un DataFrame que contiene la información de los NRC y el nombre del curso en Brightspace.
    '''
    # Variables
    directory = './'
    global count

    # Buscamos el archivo .csv en el directorio
    for filename in os.listdir(directory):
        if filename.endswith(".csv") :
            print('Leyendo... ' + filename)
            data = pd.read_csv(filename)
        else:
            continue
    return data

#Generar el archivo de registro unico (resumen)
def merge_archivos():

    # Variables
    directory = './salida/'

    # iteramos sobre los .txt
    for filename in os.listdir(directory):
        if filename.endswith('.txt'):
            with open(directory + filename, encoding='utf8') as fp:
                data = fp.read()

            with open('registro_unicoMOD.txt', 'a', encoding='utf8') as fp:
                fp.write(data)            

    return

#Leer la BD de usuarios de BS: XLSX
def leer_BDUsuarios_BS(ruta_archivo="./BDUsuarios/Listados Usuarios.xlsx"):
    """
    Función para cargar un archivo Excel de usuarios en un DataFrame.
    
    Parámetros:
    ruta_archivo (str): Ruta del archivo Excel a cargar.
    
    Retorna:
    pd.DataFrame: DataFrame con los datos del archivo o None si ocurre un error.
    """
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
    '''
    Función que no recibe entradas.
    Devuelve el DataFrame con la información del archivo Excel, extensión .xlsx con la información descargada de Banner.
    '''
    # Variables
    directory = './'

    # Busqueda del .xlsx dentro del mismo directorio
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            #Lectura del archivo
            print('\nLeyendo BD moderadores y NRC :' + filename)
            data = pd.read_excel(filename, sheet_name = 0)

            # Reemplazar espacios con "_" de las columnas
            data.columns = [c.replace(' ', '_') for c in data.columns]

            if (date != 'nodate'):
                data = data.loc[(data.Fecha_Actividad >= date)]
            
            #el ID_Docente se convierte a string para evitar problemas de formato
          
            data['ID_Docente']  = data.ID_Docente.astype(str)

            #Lista cruzada se convierte a string para evitar problemas de formato
            data['Lista_Cruzada'] = data['Lista_Cruzada'].astype(str)

            #Asegurar que los valores sean cadenas de texto y completar con ceros a la izquierda, se asume la longitud de 9
            data['ID_Docente'] = data['ID_Docente'].astype(str).str.zfill(9)

            #El nombre y apellidos se convierten a formato title y impieza y formateo de nombres y apellidos
            data['Nombre_Docente'] = data['Nombre_Docente'].str.strip().str.title()
            data['Apellidos_Docente'] = data['Apellidos_Docente'].str.strip().str.title()

            #COLUMNAS PROMOVIDAS
            data = data[['Periodo_Académico', 'NRC', 'Lista_Cruzada', 'ID_Docente', 'Tipo_Documento', 'Documento', 'Correo_Principal','Nombre_Docente', 'Apellidos_Docente', 'Fecha_Actividad']]
        else:
            continue

    return data

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
            idBanner = str(row.get('ID_Docente', '')).strip()

            # Validación de ID_Banner
            if idBanner in ["", "0", "-", "000000000",None]:
                log.write(f"❌ ID inválido: '{idBanner}' para curso {course_name}\n")
                continue
            
            # extraer el rol del moderador y verificar si es nuevo o existente
            Unuevo = idBanner not in BDUsuBS['UserName'].values
            RolModerador = BDUsuBS.loc[BDUsuBS['UserName'] == idBanner, 'OrgRoleId'].values

            # Se formatea el tipo de documento y número de documento
            try:
                ndocu = "{:,}".format(int(row['Documento'])).replace(',', '.')
            except:
                ndocu = str(row['Documento'])

            try:
                docuusu = f"{row['Tipo_Documento']}. {ndocu}"
            except:
                docuusu = ndocu

            #Se limpian los nombres y apellidos
            first_name = str(row.get('Nombre_Docente', '')).strip()
            last_name = str(row.get('Apellidos_Docente', '')).strip()
            email = str(row.get('Correo_Principal', '')).strip()

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
