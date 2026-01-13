#!/usr/bin/python
import sys
import helpersmod as helpers
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
        
        BDModeradoresNRC = helpers.leer_moderadores(date_time_obj)
    else:
        BDModeradoresNRC = helpers.leer_moderadores()

    print("\n---------------------------------")
    print("BD Moderadores y NRC a inscribir")
    print("-----------------------------------")
    print(BDModeradoresNRC)
    print ("\n")

    print("\n-------------------------------")
    print("BD Usuarios de BS")
    print("-------------------------------")
    BDUsuarios = helpers.leer_BDUsuarios_BS()
    print(BDUsuarios)
    print ("\n")
 
    # Leer datos de los NRC/LC que se incribiran - shortname.csv
    nrc = helpers.leer_nrc()
    #convertir a string el NRC/LC para evitar problemas de formato
    nrc['NRC'] = nrc['NRC'].astype(str)
     
    print("\n-------------------------------")
    print("Cursos a inscribir")
    print("-------------------------------")
    print(nrc)

    # Remover duplicados x Pediodo, NRC/LC y ID_Docente
    data_sin_duplicados = BDModeradoresNRC.drop_duplicates(subset=['Periodo_Académico', 'NRC', 'ID_Docente'])

    print("\n----------------------------------")
    print("Iniciando proceso de inscripcion")
    print("----------------------------------")

    # Crear los archivos: CSV para inscripcion, uno por NRC y se genera resumen de inscripcion (moderadores.csv)
    for index, row in nrc.iterrows():
        course_name = row['Nombre']     #Codigo del curso
        course_nrc = row['NRC']         #NRC/LC del curso
        course_periodo = row['Periodo'] #Periodo del curso

        #Se filtra el dataframe "ModeradoresInscribir" por el Periodo y NRC/LC del curso
        ModeradoresInscribir = data_sin_duplicados[ 
                (data_sin_duplicados['Periodo_Académico'] == course_periodo) & 
                (data_sin_duplicados['Lista_Cruzada'] == course_nrc)
        ] 

        helpers.crearArchivos(ModeradoresInscribir, course_name, course_nrc, BDUsuarios)

    #se crea un solo archvivo con todos los cursos
    helpers.merge_archivos()

    print("\n------------------------------------------")
    print("Finaliza proceso de inscripcion de moderadores")
    print("------------------------------------------")

    return

if __name__ == "__main__":
    main()