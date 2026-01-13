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
    print ("\n")

    print("\n-------------------------------")
    print("BD Estudiantes de BS")
    print("-------------------------------")

    print(BDestudiantes)
    print ("\n")
 
    # captura tipo de acciones
    #print ("\n-------------------------------")
    #tacc = input ("Tipo acciones: \n 1. Todo(Crear usuarios e inscribir) \n 2. Solo inscribir\n\n")
    #if tacc == "": 
    tacc = "1"

    # Leer datos de los NRC/LC que se incribiran - shortname.csv
    nrc = helpers.leer_nrc()
    #convertir a string el NRC/LC para evitar problemas de formato
    nrc['NRC'] = nrc['NRC'].astype(str)
     
    print("\n-------------------------------")
    print("Cursos a inscribir")
    print("-------------------------------")
    print(nrc)

    # Remover duplicados x Pediodo, NRC/LC y ID_Estudiante
    data_sin_duplicados = BDestudiantes.drop_duplicates(subset=['Periodo', 'NRC', 'ID_Estudiante'])

    # Crear los archivos: CSV para inscripcion, uno por NRC y se genera resumen de inscripcion (student.csv)
    for index, row in nrc.iterrows():
        course_name = row['Nombre']
        course_nrc = row['NRC']
        course_periodo = row['Periodo']

        #Se filtra el dataframe "EstudiantesInscribir"por el Periudo y NRC/LC del curso
        EstudiantesInscribir = data_sin_duplicados[ 
                (data_sin_duplicados['Periodo'] == course_periodo) & 
                (data_sin_duplicados['Lista_Cruzada'] == course_nrc)
        ]

        helpers.crearArchivos(EstudiantesInscribir, course_name, course_nrc, course_periodo, tacc)

    #se crea un solo archvivo con todos los cursos
    helpers.merge_archivos()

    print("\n-------------------------------")
    print("Finaliza proceso de inscripcion")
    print("-------------------------------")

    return

if __name__ == "__main__":
    main()