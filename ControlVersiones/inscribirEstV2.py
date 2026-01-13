#!/usr/bin/python
import sys
import gc
import helpersestV2 as helpers
from datetime import datetime

def main():

    # Validación de entradas y cargar EXCEL con los datos de los NRC Y ESTUDIANTES - QLIK
    if(len(sys.argv) > 2):
        print('Número de entradas inválido. Saliendo...')
        return

    if(len(sys.argv) == 2):
        date_time_str = sys.argv[1] + ' 00:00:00'
        date_time_obj = datetime.strptime(date_time_str, '%d/%m/%y %H:%M:%S')

        if(date_time_obj > datetime.now()):
            print('La fecha de entrada no puede ser mayor a la fecha actual.')
            return
        
        BDEstudiantesNRC = helpers.leer_estudiantes(date_time_obj)
    else:
        BDEstudiantesNRC = helpers.leer_estudiantes()

    print("\n---------------------------------")
    print("BD Estudiantes y NRC a inscribir")
    print("-----------------------------------")
    print(BDEstudiantesNRC)
    print ("\n")

    print("\n-------------------------------")
    print("BD Estudiantes de BS")
    print("-------------------------------")
    BDestudiantes = helpers.leer_BDUsuarios_BS()
    print(BDestudiantes)
    print ("\n")
 
    # Leer datos de los NRC/LC que se incribiran - shortname.csv
    nrc = helpers.leer_nrc()
    if nrc is None: 
        # Si no se pudo leer el archivo de NRC/LC, se sale del programa   
        return
    
    #convertir a string el NRC/LC para evitar problemas de formato
    nrc['NRC'] = nrc['NRC'].astype(str)
     
    print("\n-------------------------------")
    print("Cursos a inscribir")
    print("-------------------------------")
    print(nrc)

    # Remover duplicados x Pediodo, NRC/LC y ID_Estudiante
    data_sin_duplicados = BDEstudiantesNRC.drop_duplicates(subset=['PERIODO', 'NRC', 'ID_ESTUDIANTE'])

    # Crear los archivos: CSV para inscripcion, uno por NRC y se genera resumen de inscripcion (student.csv)
    for index, row in nrc.iterrows():
        course_name = row['Nombre']     #Codigo del curso
        course_nrc = row['NRC']         #NRC/LC del curso
        course_periodo = row['Periodo'] #Periodo del curso

        #Se filtra el dataframe "EstudiantesInscribir"por el Periudo y NRC/LC del curso
        EstudiantesInscribir = data_sin_duplicados[ 
                (data_sin_duplicados['PERIODO'] == course_periodo) & 
                (data_sin_duplicados['LISTA_CRUZADA'] == course_nrc)
        ] 
   
        helpers.crearArchivos(EstudiantesInscribir, course_name, course_nrc, course_periodo, BDestudiantes)

    #se crea un solo archvivo con todos los cursos
    helpers.merge_archivos()

    print("\n-------------------------------")
    print("Limpieza de memoria...")
    print("-------------------------------")

    #Limpiar los DataFrames para liberar memoria
    BDEstudiantesNRC.drop(BDEstudiantesNRC.index, inplace=True)             # Limpiar el DataFrame QLIK para liberar memoria
    BDestudiantes.drop(BDestudiantes.index, inplace=True)                   # Limpiar el DataFrame de estudiantes BS
    del BDEstudiantesNRC, BDestudiantes                                     # Eliminar las variables para liberar memoria
    gc.collect()                                                            # Liberar memoria

    print("\n-------------------------------")
    print("Finaliza proceso de inscripcion")
    print("-------------------------------")
                                                     
    return

if __name__ == "__main__":
    main()