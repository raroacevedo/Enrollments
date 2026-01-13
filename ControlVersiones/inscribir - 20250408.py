#!/usr/bin/python

import sys
import helpers
from datetime import datetime

def main():

    # Validación de entradas y cargar EXCEL con los datos de los NRC - QLIK
    if(len(sys.argv) > 2):
        print('Número de entradas inválido. Saliendo...')
        return

    if(len(sys.argv) == 2):
        date_time_str = sys.argv[1] + ' 00:00:00'
        date_time_obj = datetime.strptime(date_time_str, '%d/%m/%y %H:%M:%S')

        if(date_time_obj > datetime.now()):
            print('La fecha de entrada no puede ser mayor a la fecha actual.')
            return
        
        BDestudiantes = helpers.leer_estudiantes(date_time_obj)
    else:
        BDestudiantes = helpers.leer_estudiantes()

    print("\n-------------------------------")
    print("BD Estudiantes a inscribir")
    print("-------------------------------")
    print(BDestudiantes)

    # captura el tipo de inscripcion 1. FA, 2. PR , 3. EX y 4. AP
    print ("\n")
    tform = input ("Tipo de formacion a inscribir: \n 1. FA \n 2. PR \n 3. EX \n 4. AP \n\n")
    if tform == "": tform = "1"

    # captura tipo de acciones
    #print ("\n-------------------------------")
    #tacc = input ("Tipo acciones: \n 1. Todo(Crear usuarios e inscribir) \n 2. Solo inscribir\n\n")
    #if tacc == "": 
    tacc = "1"

    # Leer datos de los NRC o LC que se incribiran - shortname.csv
    nrc = helpers.leer_nrc()
    #convertir a string el NRC o LC para evitar problemas de formato
    nrc['NRC'] = nrc['NRC'].astype(str)
    
    print("\n-------------------------------")
    print("\nCursos a inscribir")
    print("-------------------------------")
    print(nrc)

    # Remover duplicados x NRC y ID_Estudiante
    data_sin_duplicados = BDestudiantes.drop_duplicates(subset=['NRC', 'ID_Estudiante'])

    # Crear los archivos: CSV para inscripcion, uno por NRC y se genera resumen de inscripcion (student.csv)
    for index, row in nrc.iterrows():
        course_name = row['Nombre']
        course_nrc = row['NRC']

        #Se filtra el dataframe "EstudiantesInscribir"por el NRC o LC del curso
        EstudiantesInscribir = data_sin_duplicados[data_sin_duplicados['Lista_Cruzada'] == course_nrc]

        helpers.crearArchivos(EstudiantesInscribir, course_name, course_nrc, tform, tacc)

    #se crea un solo archvivo con todos los cursos
    helpers.merge_archivos()

    print("\n-------------------------------")
    print("Finaliza proceso de inscripcion")
    print("-------------------------------")

    return

if __name__ == "__main__":
    main()