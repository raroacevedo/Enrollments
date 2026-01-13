#!/usr/bin/python

import pandas as pd
import os
import csv
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet") # Para evitar advertencias de openpyxl: Excel no contiene un "estilo por defecto" definido en sus metadatos

#
#leer el archivo shortnames.csv con NRC/LC: CSV
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
        df = pd.read_csv(filepath)
    except Exception as e:
        print(f"❌ Error al leer el archivo '{filename}': {e}")
        return None

    # Validación de columnas esperadas
    columnas_esperadas = {'Nombre','NRC','Periodo'}
    if not columnas_esperadas.issubset(df.columns):
        print(f"❌ Error: El archivo debe contener las columnas: {columnas_esperadas}. Columnas actuales: {df.columns.tolist()}")
        return None
  
    return df

#
#Generar el archivo de registro unico (resumen)
def merge_archivos():
    directory = './salida/'

    # iteramos sobre los .txt
    for filename in os.listdir(directory):
        if filename.endswith('.txt'):
            with open(directory + filename, encoding='utf8') as fp:
                data = fp.read()

            with open('registro_unicoEst.txt', 'a', encoding='utf8') as fp:
                fp.write(data)            

    return

#
#Leer la BD de estudiantes de BS: XLSX
def leer_BDUsuarios_BS(ruta_archivo="./BDUsuarios/Listados Usuarios.xlsx"):
    """
    Función para cargar un archivo Excel en un DataFrame.
    
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
        df=df[['UserName', 'FirstName', 'LastName']]
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

#
#leer el archivo de estudiantes a inscribir: EXCEL fuente QLIK
def leer_estudiantes(date='nodate'):
    """
    Lee múltiples archivos .xlsx con información de estudiantes desde el directorio actual.
    Optimizado para grandes volúmenes. Aplica limpieza y validación.
    
    Parámetros:
        date (str): Fecha mínima (YYYY-MM-DD) para filtrar la columna 'Fecha_Actividad'.
    
    Retorna:
        pd.DataFrame consolidado.
    """
    
    #directorio actual y lista de archivos excel
    directory = './'
    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

    # Validación de archivos excel en el directorio
    if not excel_files:
        raise FileNotFoundError("❌ No se encontró ningún archivo .xlsx en el directorio actual.")

    # Definición de solo las columnas necesarias para optimizar la carga del excel
    columnas_objetivo = {
        'Periodo', 'NRC', 'Lista_Cruzada', 'ID_Estudiante', 'Tipo_Documento',
        'Documento', 'Correo_Principal_Estudiante', 'Nombre_Estudiante',
        'Apellidos_Estudiante', 'Tipo_Cancelación_Curso', 'Fecha_Actividad'
    }
    dataframes = []

    # Iterar sobre cada archivo Excel
    for file in excel_files:
        filepath = os.path.join(directory, file)
        print(f"📥 Leyendo archivo: {filepath}")

        try:
            # Cargar todo el archivo y Se renombran todas las columnas de " " con _ antes de procesar
            df = pd.read_excel(filepath, sheet_name=0, engine='openpyxl')
            df.columns = [col.strip().replace(' ', '_') for col in df.columns]
        except Exception as e:
            print(f"❌ Error al leer el archivo {file}: {e}")
            continue

        # Validar columnas requeridas existan el el excel
        if not columnas_objetivo.issubset(set(df.columns)):
            faltantes = columnas_objetivo - set(df.columns)
            print(f"⚠️ Advertencia: El archivo {file} no tiene todas las columnas requeridas: {faltantes}")
            continue

        # Filtrar solo las columnas necesarias
        df = df[list(columnas_objetivo)]

        # Limpieza y transformación de datos :espacios, tipos de datos y tipo titulo
        df['ID_Estudiante'] = df['ID_Estudiante'].astype(str).str.zfill(9)
        df['Lista_Cruzada'] = df['Lista_Cruzada'].astype(str)
        df['Tipo_Cancelación_Curso'] = df['Tipo_Cancelación_Curso'].astype(str)
        df['Nombre_Estudiante'] = df['Nombre_Estudiante'].astype(str).str.strip().str.title()
        df['Apellidos_Estudiante'] = df['Apellidos_Estudiante'].astype(str).str.strip().str.title()

        #se verifica que las columnas de fecha y periodo existan
        try:
            df['Fecha_Actividad'] = pd.to_datetime(df['Fecha_Actividad'], errors='coerce')
            if date != 'nodate':
                df = df[df['Fecha_Actividad'] >= pd.to_datetime(date)]
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


#
#Se crea el archivo de registro para cada curso NRC/LC
def crearArchivos(data, course_name, course_nrc, course_periodo, BDEstuBS):
    '''
    Función que recibe como entrada un dataframe del archivo de Excel leído, y el nombre del curso.
    No devuelve ningún valor.
    Recorre las filas del dataframe, y genera los comandos para la creación y registro de usuarios en Brightspace.
    '''
    #se evalua el tipo de formacion a inscribir
    #FA, PR, EX y TE 
    tipformacion = str(course_periodo)[-2:]

    if tipformacion in ["41", "42"]:
          Rol = "Student_fa"
          OrgUnid = "CVFA"
    elif tipformacion in ["10", "11", "20", "21"]:
          Rol = "Student_pr"
          OrgUnid = "CVPR"
    elif (tipformacion == "50"):
          Rol = "Student_ex"
          OrgUnid = "CVFC"
    elif tipformacion in ["17", "27", "37"]:
          Rol = "Student_te"
          OrgUnid = "CVTE"
  
    # Creamos los archivos
    file    = './salida/registro' + '_' + course_name + '.txt'
    fptr    = open(file, 'a', encoding='utf8')
    line_count = 0

    # Ciclo para recorrer el dataframe - estudiantes Banner
    for index, row in data.iterrows():

        ###Guardar los datos que necesitamos en variables

        #ID_Estudiante
        idBanner       = row['ID_Estudiante']

        # Verificar si el idBanner existe en la base de estudiantes BS
        Enuevo = idBanner not in BDEstuBS['UserName'].values

        #tipo documento + numero documento
        try:
             ndocu = "{:,}".format(int(row['Documento'])).replace(',', '.')  #Formatear con separador de miles
        except:
             ndocu = row['Documento']                                        # Si hay un error, deja el valor original
        
        try:
             docuusu     = row['Tipo_Documento']+". "+ndocu                  #se concatena el tipo de documento con el numero
        except:
             docuusu     = ndocu                                             #Si hay un error, se deja solo el numero

        #nombre y apellidos del estudiante.
        first_name  = row['Nombre_Estudiante']
        last_name   = row['Apellidos_Estudiante']

        #correo principal del estudiante
        email       = row['Correo_Principal_Estudiante']
        
        #tipo de cancelacion
        cancelacion = row['Tipo_Cancelación_Curso']

        if(cancelacion == 'nan'): ##si no hay cancelacion se procede a la inscripcion
                
            if Enuevo: ##si el estudiante no existe en la base de datos de estudiantes BS SE crea el usuario
                fptr.write('CREATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + Rol + ',' + '1' + ',' + email + '\n')
                # Generamos la inscripción  en la Unidad para la pagina de inicio
                fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol + ',' + OrgUnid + '\n')
            else: ##si el estudiante ya existe en la base de datos de estudiantes BS SE actualiza el usuario
                # Generamos la actualización de los datos del usuario y SE ACTIVA EL USUARIO
                fptr.write('UPDATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + '1' + ',' + email + '\n')
                # Generamos la inscripción  en la Unidad UPBV - CAMBIO ROL ARQUETIPO
                fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol + ',' + "UPBV" + '\n')

            # Generamos las lineas al archivo para inscripción en el curso
            fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + 'Student' +',' + course_name + '\n')

            line_count = line_count + 1
        else:
            # Generamos los registros para desmatricular al estudiante
            fptr.write('UNENROLL' + ',' + idBanner + ',' +',' + course_name + '\n')

    numberStudents = [course_name, course_nrc, line_count]
    estudiantes = open('students.csv', 'a', encoding='utf8')
    writer = csv.writer(estudiantes)
    writer.writerow(numberStudents)

    print("\n[✓] Se han inscrito:" + str(line_count) + " estudiantes en el curso:" + course_name + " NRC:" + course_nrc)

    # Cerramos los archivos
    fptr.close()
    estudiantes.close()