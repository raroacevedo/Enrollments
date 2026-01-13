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
            data = data[['Periodo', 'NRC', 'Lista_Cruzada', 'ID_Docente', 'Tipo_Documento', 'Documento', 'Correo_Principal','Nombre_Docente', 'Apellidos_Docente', 'Fecha_Actividad']]
        else:
            continue

    return data

#se crea el archivo de registro para cada curso
def crearArchivos(data, course_name, course_nrc, BDUsuBS):
    '''
    Función que recibe como entrada un dataframe del archivo de Excel leído, y el nombre del curso.
    No devuelve ningún valor.
    Recorre las filas del dataframe, y genera los comandos para la creación y registro de usuarios en Brightspace.
    '''
    # Rol del moderador
    Rol = "Moderador"
     
    # Creamos los archivos
    file    = './salida/registro' + '_' + course_name + '.txt'
    fptr    = open(file, 'a', encoding='utf8')
    line_count = 0

    # Ciclo para recorrer el dataframe - estudiantes Banner
    for index, row in data.iterrows():

        ###Guardar los datos que necesitamos en variables

        #ID_Docente
        idBanner       = row['ID_Docente']

        # Verificar si el idBanner existe en la base de usuarios BS
        Enuevo = idBanner not in BDUsuBS['UserName'].values

        #El rol del moderador en la base de datos de usuarios de Brightspace
        RolModerador = BDUsuBS.loc[BDUsuBS['UserName'] == idBanner, 'OrgRoleId'].values

        #tipo documento + numero documento
        try:
             ndocu = "{:,}".format(int(row['Documento'])).replace(',', '.')  #Formatear con separador de miles
        except:
             ndocu = row['Documento']                                        # Si hay un error, deja el valor original

        try:
             docuusu     = row['Tipo_Documento']+". "+ndocu                  #se concatena el tipo de documento con el numero
        except:
             docuusu     = ndocu                                             #Si hay un error, se deja solo el numero

        #nombre y apellidos del Moderador.
        first_name  = row['Nombre_Documento']
        last_name   = row['Apellidos_Documento']

        #correo principal del moderador
        email       = row['Correo_Principal']
   
        #se genera el comando de creación o actualización del usuario                  
        if Enuevo: ##si el moderador no existe en la base de datos de usuarios en BS SE crea el usuario
            fptr.write('CREATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + Rol + ',' + '1' + ',' + email + '\n')

        else: ##si el moderador ya existe en la base de datos de estudiantes BS SE actualiza el usuario
            # Generamos la actualización de los datos del usuario y SE ACTIVA EL USUARIO
            fptr.write('UPDATE' + ',' + idBanner + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + '1' + ',' + email + '\n')

            # Si el rol del moderador ya existe en la base de datos de usuarios de Brightspace y es un estudiante se da de baja del una unidad organizativa
            if RolModerador == '150':
               fptr.write('UNENROLL' + ',' + idBanner + ',' + ',' + 'CVTE' + '\n')
            elif RolModerador == '143':
                fptr.write('UNENROLL' + ',' + idBanner + ',' + ',' + 'CVLA' + '\n')
            elif RolModerador == '138':
                fptr.write('UNENROLL' + ',' + idBanner + ',' + ',' + 'CVPR' + '\n')
            elif RolModerador == '137':
                fptr.write('UNENROLL' + ',' + idBanner + ','  + ',' + 'CVFC' + '\n')
            elif RolModerador == '136':
                fptr.write('UNENROLL' + ',' + idBanner + ','  + ',' + 'CVFA' + '\n')
            elif RolModerador == '135':
                fptr.write('UNENROLL' + ',' + idBanner + ','  + ',' + 'CVFA' + '\n')

            # Generamos la inscripción  en la Unidad UPBV - CAMBIO ROL ARQUETIPO A MODERADOR
            fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol + ',' + "UPBV" + '\n')

        # Generamos las lineas al archivo para inscripción en el curso
        fptr.write('ENROLL' + ',' + idBanner + ',' + '' + ',' + Rol +',' + course_name + '\n')

        line_count = line_count + 1
  
    # Guardamos el número de moderareos inscritos en el curso
    numberModeradores = [course_name, course_nrc, line_count]
    moderadores = open('moderadores.csv', 'a', encoding='utf8')
    writer = csv.writer(moderadores)
    writer.writerow(numberModeradores)

    print("\n[✓] Se han inscrito:" + str(line_count) + " moderadores en el curso:" + course_name + " NRC:" + course_nrc)

    # Cerramos los archivos
    fptr.close()
    moderadores.close()