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
    Devuelve un dic con la configuración (vacío si no existe o hay errores).
    """
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
        df = pd.read_csv(
                filepath,
                encoding='utf-8',
                sep=',',
                header=0,
                names=['Nombre', 'NRC', 'Periodo'],
                dtype={'NRC': str, 'Periodo': str}
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

#
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

            with open('registro_unicoEst.txt', 'a', encoding='utf8') as fp:
                fp.write(data)            

    return

#
#Leer la BD de estudiantes de BS: XLSX - Origen BS INSIGHT
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
#leer el archivo de estudiantes a inscribir: EXCEL fuente BANNER - SZREINS 
def leer_estudiantesBanner(date='nodate'):
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
        'PERIODO', 'NRC', 'LISTA_CRUZADA', 'ID_ESTUDIANTE', 'TIPO_DOCUMENTO',
        'DOCUMENTO', 'CORREO_ESTUDIANTE', 'NOMBRE_ESTUDIANTE','APELLIDO_ESTUDIANTE', 
        'COD_INSCRIPCIÓN', 'ESTADO_INSCRIPCIÓN', 'FECHA_ACTIVIDAD_EST', "PAGO", "SOCIO_INTEGRADOR"
    }
    dataframes = []

    # Iterar sobre cada archivo Excel
    for file in excel_files:
        filepath = os.path.join(directory, file)
        print(f"📥 Leyendo archivo: {filepath}")

        try:
            # Cargar todo el archivo
            df = pd.read_excel(filepath, 
                               sheet_name=1,  # Segunda hoja (estudiantes)
                               dtype={'ID_ESTUDIANTE': str},
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
        df['ID_ESTUDIANTE'] = df['ID_ESTUDIANTE'].astype(str).str.zfill(9)
        df['LISTA_CRUZADA'] = df['LISTA_CRUZADA'].astype(str)
        df['COD_INSCRIPCIÓN'] = df['COD_INSCRIPCIÓN'].astype(str)
        df['ESTADO_INSCRIPCIÓN'] = df['ESTADO_INSCRIPCIÓN'].astype(str)
        df['NOMBRE_ESTUDIANTE'] = df['NOMBRE_ESTUDIANTE'].astype(str).str.strip().str.title()
        df['APELLIDO_ESTUDIANTE'] = df['APELLIDO_ESTUDIANTE'].astype(str).str.strip().str.title()
        df['SOCIO_INTEGRADOR'] = df['SOCIO_INTEGRADOR'].astype(str).str.strip() #socio integrador
        df['PAGO'] = df['PAGO'].astype(str)                                     #pago Y/N

        #se verifica que las columnas de fecha y periodo existan
        try:
            df['FECHA_ACTIVIDAD_EST'] = pd.to_datetime(df['FECHA_ACTIVIDAD_EST'], errors='coerce')
            if date != 'nodate':
                df = df[df['FECHA_ACTIVIDAD_EST'] >= pd.to_datetime(date)]
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
      
    # se evalua que tipo de proceso x defecto es Matricular
    tipproceso = CONFIG.get('Tipo_proceso', 'Matricular')
    
    #se evalua el tipo de formacion a inscribir
    #FA, PR, EX y TE 
    tipformacion = str(course_periodo)[-2:]

    if tipformacion in ["41", "42"]:                #formacion avanzada
          Rol = "Student_fa"
          OrgUnid = "CVFA"
    elif tipformacion in ["10", "11", "20", "21"]: #formacion pregrado
          Rol = "Student_pr"
          OrgUnid = "CVPR"
    elif (tipformacion == "50"):                    #formacion continua
          Rol = "Student_ex"
          OrgUnid = "CVFC"
    elif tipformacion in ["17", "27", "37"]:        #formacion tecnologica
          Rol = "Student_te"
          OrgUnid = "CVTE"
  
    # Creamos los archivos distintos por curso
    #file    = './salida/registro' + '_' + course_name + '.txt'
    directory = CONFIG.get('salida_directory', './salida/')
    file    = directory + 'registro_' + course_name + '.txt'
    fptr    = open(file, 'a', encoding='utf8')
    line_count = 0
 
    # Ciclo para recorrer el dataframe - estudiantes Banner
    for index, row in data.iterrows():

        ###Guardar los datos que necesitamos en variables

        #ID_Estudiante
        idBanner       = row['ID_ESTUDIANTE']
        
        # Verificar si el idBanner existe en la base de estudiantes BS
        Enuevo = idBanner not in BDEstuBS['UserName'].values

        #tipo documento + numero documento
        try:
             ndocu = "{:,}".format(int(row['DOCUMENTO'])).replace(',', '.')  #Formatear con separador de miles
        except:
             ndocu = row['DOCUMENTO']                                        # Si hay un error, deja el valor original

        try:
             docuusu     = row['TIPO_DOCUMENTO']+". "+ndocu                  #se concatena el tipo de documento con el numero
        except:
             docuusu     = ndocu                                             #Si hay un error, se deja solo el numero

        #nombre y apellidos del estudiante.
        first_name  = row['NOMBRE_ESTUDIANTE']
        last_name   = row['APELLIDO_ESTUDIANTE']

        #correo principal del estudiante
        email       = row['CORREO_ESTUDIANTE']
        
        #ESTADO_INSCRIPCIÓN
        estado = row['ESTADO_INSCRIPCIÓN']

        #ESTUDIANTE PAGO
        pago = row['PAGO']
  
        #validar socio integrador : Si es APLATAM o BS (UPBVIRTUAL)
        socio_integrador = row['SOCIO_INTEGRADOR']
        if socio_integrador == "nan" : socio_integrador = "BS"  # Si no tiene socio integrador, se asume BS
    
        #se valida el caso de "APLATAM" en formacion avanzada
        if tipformacion in ["41", "42"]:
            if socio_integrador == "AP":
                Rol = "Student_ap"
                OrgUnid = "CVLA"
            else:
                Rol = "Student_fa"
                OrgUnid = "CVFA"
               
        #Proceso de inscripcion o cancelacion
        if(tipproceso == 'Matricular'): #Definido en el JSON de configuracion
            if(estado == 'Inscrito'):   #Se procede a la inscripcion - BANNER|SZREINS
                if (socio_integrador == "BS") or (socio_integrador == "AP" and pago == "Y") :  # Solo se inscribe si es APLATAM y ha pagado o si es BS
                    
                    if Enuevo: ##si el estudiante no existe en la base de datos de estudiantes BS SE crea el usuario
                        fptr.write('CREATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + Rol + ',' + '1' + ',' + email + '\n')
                        # Generamos la inscripción  en la Unidad (nivel de formacion) para la pagina de inicio
                        fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol + ',' + OrgUnid + '\n')
                    else: ##si el estudiante ya existe en la base de datos de estudiantes BS SE actualiza el usuario
                        # Generamos la actualización de los datos del usuario y SE ACTIVA EL USUARIO
                        fptr.write('UPDATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + '1' + ',' + email + '\n')
                        # Generamos la inscripción  en la Unidad UPBV - CAMBIO ROL ARQUETIPO
                        fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol + ',' + "UPBV" + '\n')

                    # Generamos las lineas al archivo para inscripción en el curso
                    fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + 'Student' +',' + course_name + '\n')

                    line_count = line_count + 1
        #Proceso de desmatriculacion o Limpieza
        else:
            #print("\n[→] Limpiando estudiante ID_Banner: " + idBanner + " del curso: " + course_name + " NRC: " + course_nrc)
            if ((tipproceso == 'Desmatricular') and (estado == 'Cancelado')): ##si hay cancelacion se procede a la desmatriculacion
                fptr.write('UNENROLL' + ',' + idBanner + ',' +',' + course_name + '\n')
                line_count = line_count + 1
            elif ((tipproceso == 'Limpieza') and (estado == 'Eliminado')): ##si hay eliminacion se procede a la desmatriculacion
                  # Generamos los registros para desmatricular al estudiante - LIMPIEZA DE LISTA (DL)            
                  fptr.write('UNENROLL' + ',' + idBanner + ',' +',' + course_name + '\n')
                  line_count = line_count + 1
    
    # Generamos el archivo resumen de inscritos por curso
    numberStudents = [course_name, course_nrc, line_count]
    estudiantes = open('students.csv', 'a', encoding='utf8')
    writer = csv.writer(estudiantes)
    writer.writerow(numberStudents)

    print("\n[✓] Se han inscrito:" + str(line_count) + " estudiantes en el curso:" + course_name + " NRC:" + course_nrc)

    # Cerramos los archivos
    fptr.close()
    estudiantes.close()    