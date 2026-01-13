#!/usr/bin/python

import pandas as pd
import os
import csv

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
        if filename.endswith(".csv") and filename != 'students.csv' and filename != 'test.csv':
            print('Leyendo ' + filename)
            data = pd.read_csv(filename)
        else:
            continue
    return data

def merge_archivos():

    # Variables
    directory = './salida/'

    # iteramos sobre los .txt
    for filename in os.listdir(directory):
        if filename.endswith('.txt'):
            with open(directory + filename, encoding='utf8') as fp:
                data = fp.read()

            with open('registro_unico.txt', 'a', encoding='utf8') as fp:
                fp.write(data)            

    return

#
def leer_estudiantes(date='nodate'):
    '''
    Función que no recibe entradas.
    Devuelve el DataFrame con la información del archivo Excel, extensión .xlsx con la 
    información descargada de Banner.
    '''

    # Variables
    directory = './'

    # Busqueda del .xlsx dentro del mismo directorio
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            #Lectura del archivo
            print('Leyendo ' + filename)
            data = pd.read_excel(filename, sheet_name = 0)

            # Reemplazar espacios con "_"
            data.columns = [c.replace(' ', '_') for c in data.columns]

            if (date != 'nodate'):
                data = data.loc[(data.Fecha_Actividad >= date)]
            
            #el ID_Estudiante se convierte a string para evitar problemas de formato
            data['ID_Estudiante']  = data.ID_Estudiante.astype(str)

            #Lista cruzada se convierte a string para evitar problemas de formato
            data['Lista_Cruzada'] = data['Lista_Cruzada'].astype(str)

            #Asegurar que los valores sean cadenas de texto y completar con ceros a la izquierda, se asume la longitud de 9
            data['ID_Estudiante'] = data['ID_Estudiante'].astype(str).str.zfill(9)

            #El nombre y apellidos se convierten a formato title 
            data['Nombre_Estudiante'] = data.Nombre_Estudiante.str.title()
            data['Apellidos_Estudiante'] = data.Apellidos_Estudiante.str.title()

            data['Tipo_Cancelación_Curso']  = data.Tipo_Cancelación_Curso.astype(str)

            #COLUMNAS PROMOVIDAS
            data = data[['Periodo', 'NRC', 'Lista_Cruzada', 'ID_Estudiante', 'Tipo_Documento', 'Documento', 'Correo_Principal_Estudiante','Nombre_Estudiante', 'Apellidos_Estudiante', 'Tipo_Cancelación_Curso', 'Fecha_Actividad']]
        else:
            continue

    return data

#
def crearArchivos(data, course_name, course_nrc, tform, tacc):
    '''
    Función que recibe como entrada un dataframe del archivo de Excel leído,
    y el nombre del curso.
    No devuelve ningún valor.
    Recorre las filas del dataframe, y genera los comandos para la creación y
    registro de usuarios en Brightspace.
    '''
    
    #se evalua el tipo de formacion a inscribir
    Rol = "Student"
    OrgUnid=""

    if (tform == "1"): 
        Rol = "Student_fa"
        OrgUnid = "CVFA"
    elif (tform == "2"):
        Rol = "Student_pr"
        OrgUnid = "CVPR"
    elif (tform == "3") :
        Rol = "Student_ex"
        OrgUnid = "CVFC"
    elif (tform == "4") :
        Rol = "Student_ap"
        OrgUnid = "CVLA"

    # Creamos los archivos
    file    = './salida/registro' + '_' + course_name + '.txt'
    fptr    = open(file, 'a', encoding='utf8')
    line_count = 0

    # Ciclo para recorrer el dataframe - estudiantes
    for index, row in data.iterrows():

        # Guardar los datos que necesitamos en variables
        id          = row['ID_Estudiante']

        #tipo documento + numero documento
        #docuusu     = row['Tipo_Documento']+". "+ndocu
        try:
             ndocu = "{:,}".format(int(row['Documento'])).replace(',', '.')  #Formatear con separador de miles
        except:
             ndocu = row['Documento']                                        # Si hay un error, deja el valor original

        try:
             docuusu     = row['Tipo_Documento']+". "+ndocu                  #se concatena el tipo de documento con el numero
        except:
             docuusu     = ndocu                                             #Si hay un error, se deja solo el numero

        #nombre y apellidos. Se elimina espacios en blanco al inicio y final
        first_name  = str.title(row['Nombre_Estudiante']).strip()
        last_name   = str.title(row['Apellidos_Estudiante']).strip()

        #correo
        email       = row['Correo_Principal_Estudiante']
        
        #tipo de cancelacion
        cancelacion = row['Tipo_Cancelación_Curso']

        if(cancelacion == 'nan'):

            if (tacc == "1"): 
                # Escribimos las lineas al archivo para registro
                fptr.write('CREATE' + ',' + id + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + Rol + ',' + '1' + ',' + email + '\n')
                # Escribimos las lineas al archivo de actualización
                fptr.write('UPDATE' + ',' + id + ',' + docuusu + ',' + first_name + ',' + last_name+ ',,' + '1' + ',' + email + '\n')
                 # Escribimos las lineas al archivo para inscripción  en la Unidad UPBV
                fptr.write('ENROLL' + ',' + id + ',' + '' + ',' + Rol + ',' + "UPBV" + '\n')
                # Escribimos las lineas al archivo para inscripción  en la Unidad para la pagina de inicio
                fptr.write('ENROLL' + ',' + id + ',' + '' + ',' + Rol + ',' + OrgUnid + '\n')
                # Escribimos las lineas al archivo para inscripción
                fptr.write('ENROLL' + ',' + id + ',' + '' + ',' + 'Student' +',' + course_name + '\n')

            elif (tacc == "2") :
                # Escribimos las lineas al archivo de actualización
                fptr.write('UPDATE' + ',' + id + docuusu + ',' + ',' + first_name + ',' + last_name+ ',,' + '1' + ',' + email + '\n')
                # Escribimos las lineas al archivo para inscripción  en la Unidad para la pagina de inicio
                fptr.write('ENROLL' + ',' + id + ',' + '' + ',' + Rol + ',' + OrgUnid + '\n')
                # Escribimos las lineas al archivo para inscripción
                fptr.write('ENROLL' + ',' + id + ',' + '' + ',' + 'Student' +',' + course_name + '\n')

            line_count = line_count + 1
        else:
            # Escribimos las lineas al archivo para desmatricular
            fptr.write('UNENROLL' + ',' + id + ',' + id + ',' + course_name + '\n')

    numberStudents = [course_nrc, line_count]
    estudiantes = open('students.csv', 'a', encoding='utf8')
    writer = csv.writer(estudiantes)
    writer.writerow(numberStudents)

    # Cerramos los archivos
    fptr.close()
    estudiantes.close()