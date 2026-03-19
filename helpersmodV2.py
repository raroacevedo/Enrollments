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
        "coordinadores_file": "./BDUsuarios/Coordinadores.xlsx",
        "salida_directory": "./salida/"
      }
    Devuelve un dict con la configuración (vacío si no existe o hay errores).
    """

    print("[INFO] Cargando configuración...")
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
        print(f"[WARN] Error al leer la configuración {cfg_path}: {e}")
        return {}

# Cargar la configuración global una sola vez, los directorios y archivos de origen de datos y salida
CONFIG = load_config()

INVALID_IDS = {"000000nan", "nan", "", "0", "-", "000000000", "none"}

# Funciones auxiliares comunes para la lectura de archivos, limpieza de datos, resolución de coordinadores y creación de archivos de inscripción.
def _resolve_path(path_value, default_path):
    base = os.path.dirname(os.path.abspath(__file__))
    final_path = path_value if path_value else default_path
    if not os.path.isabs(final_path):
        final_path = os.path.join(base, final_path)
    return final_path

# Limpia un valor convirtiéndolo a string, eliminando espacios y manejando valores nulos.
def _to_clean_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

# Normaliza un ID de banner a formato string de 9 dígitos, sin decimales ni caracteres no numéricos.
def _normalizar_id_banner(value):
    valor = _to_clean_str(value).lower()
    if valor in {"", "nan", "none"}:
        return ""

    # Si llega como float por Excel (ej: 138144.0) se limpia el decimal.
    if valor.endswith(".0"):
        valor = valor[:-2]

    return valor.zfill(9)

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
        print(f"[ERROR] Error: El archivo requerido '{filename}' no se encuentra en el directorio actual.")
        return None

    # Intentar lectura del archivo
    try:
        print(f"[INFO] Leyendo archivo: {filename}")
        df = pd.read_csv(
                filepath,
                encoding='utf-8',
                sep=',',
                header=0,
                names=['Nombre', 'NRC', 'Periodo'],
                dtype={'Nombre': str,'NRC': str}
        )

    except Exception as e:
        print(f"[ERROR] Error al leer el archivo '{filename}': {e}")
        return None

    # Validación de columnas esperadas
    columnas_esperadas = {'Nombre','NRC','Periodo'}
    if not columnas_esperadas.issubset(df.columns):
        print(f"[ERROR] Error: El archivo debe contener las columnas: {columnas_esperadas}. Columnas actuales: {df.columns.tolist()}")
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

#Leer la BD de usuarios de BS: XLSX - Fuente DOMO
def leer_BDUsuarios_BS(ruta_archivo=None):
    """
    Función para cargar un archivo Excel en un DataFrame.
    
    Parámetros:
    ruta_archivo (str): Ruta del archivo Excel a cargar.
    
    Retorna:
    pd.DataFrame: DataFrame con los datos del archivo o None si ocurre un error.
    """
    # Determinar ruta del archivo por la configuración si no se proporcionó una por defecto
    ruta_default = CONFIG.get('bdusuarios_file', "./BDUsuarios/Listados Usuarios.xlsx")
    if ruta_archivo is None:
        ruta_archivo = ruta_default
    ruta_archivo = _resolve_path(ruta_archivo, "./BDUsuarios/Listados Usuarios.xlsx")

    try:
        # Leer el Excel
        df = pd.read_excel(
            ruta_archivo,
            sheet_name=0,        # Leer la primera hoja
        )

        # Se promueven las columnas necesarias (incluye datos para creación de coordinadores)
        columnas_necesarias = ['UserName', 'FirstName', 'LastName', 'OrgRoleId', 'OrgDefinedId', 'ExternalEmail']
        for col in columnas_necesarias:
            if col not in df.columns:
                df[col] = ''

        df = df[columnas_necesarias]
        df['UserName'] = df['UserName'].apply(_normalizar_id_banner)
        df['FirstName'] = df['FirstName'].astype(str).str.strip()
        df['LastName'] = df['LastName'].astype(str).str.strip()
        df['OrgRoleId'] = df['OrgRoleId'].astype(str).str.strip()
        df['OrgDefinedId'] = df['OrgDefinedId'].apply(_to_clean_str)
        df['ExternalEmail'] = df['ExternalEmail'].apply(_to_clean_str)
        
        print(f"[OK] Archivo '{ruta_archivo}' cargado exitosamente.")
        print(f"El archivo contiene {df.shape[0]} filas y {df.shape[1]} columnas.")
        
        return df

    except FileNotFoundError:
        print(f"[ERROR] Error: El archivo '{ruta_archivo}' no fue encontrado.")
        return None
    except Exception as e:
        print(f"[ERROR] Error al cargar el archivo: {e}")
        return None

#leer el archivo de moderadores a inscribir: EXCEL fuente BANNER
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
    directory = _resolve_path(CONFIG.get('banner_directory', './'), './')

    # Validación de existencia del directorio
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"[ERROR] No se encontró el directorio de archivos .xlsx: {directory}")

    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx") and not f.startswith("~$")]

    # Validación de archivos excel en el directorio
    if not excel_files:
        raise FileNotFoundError("[ERROR] No se encontró ningún archivo .xlsx en el directorio actual.")

    # Definición de solo las columnas necesarias para optimizar la carga del excel
    columnas_objetivo = {
        'PERIODO', 'NRC', 'LISTA_CRUZADA', 'ID_DOCENTE', 'TIPO_DOCUMENTO',
        'DOCUMENTO', 'CORREO_DOCENTE', 'NOMBRE_DOCENTE', 'APELLIDO_DOCENTE',"FECHA_ACTIVIDAD_DOC"
    }
    dataframes = []

    # Iterar sobre cada archivo Excel
    for file in excel_files:
        filepath = os.path.join(directory, file)
        print(f"[INFO] Leyendo archivo: {filepath}")

        try:
            # Cargar todo el archivo
            df = pd.read_excel(filepath, 
                               sheet_name=0,  # primera hoja (docentes)
                               dtype={'ID_DOCENTE': str, 'LISTA_CRUZADA': str},
                               engine='openpyxl'
                )    

            print(f"[OK] Archivo '{file}' cargado con {df.shape[0]} filas y {df.shape[1]} columnas.")  
        except Exception as e:
            print(f"[ERROR] Error al leer el archivo {file}: {e}")
            continue

        # Validar columnas requeridas existan en el excel
        if not columnas_objetivo.issubset(set(df.columns)):
            faltantes = columnas_objetivo - set(df.columns)
            print(f"[WARN] Advertencia: El archivo {file} no tiene todas las columnas requeridas: {faltantes}")
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
            print(f"[ERROR] Error al procesar fechas en {file}: {e}")
            continue

        dataframes.append(df)

    if not dataframes:
        raise ValueError("[ERROR] Ningún archivo válido fue procesado correctamente.")

    # Concatenar todos los DataFrames en uno solo
    resultado = pd.concat(dataframes, ignore_index=True)
    
    #retornar el DataFrame consolidado
    return resultado

#Cargar el archivo de centro de costos estudiante desde Excel, con columnas 'PERIODO', 'LISTA_CRUZADA', 'ESTADO_INSCRIPCIÓN' y 'COD_PROGRAMA_ESTUDIANTE'.
#  Se filtra solo ESTADO_INSCRIPCIÓN='Inscrito' y se eliminan duplicados por: PERIODO, LISTA_CRUZADA, ESTADO_INSCRIPCIÓN, COD_PROGRAMA_ESTUDIANTE.
def leer_centrocostos_estudiante():
    """
    Construye el DataFrame CENTROCOSTOSESTUDIANTE desde banner_directory/hoja Estudiantes.
    Se filtra solo ESTADO_INSCRIPCIÓN='Inscrito' y se eliminan duplicados por:
    PERIODO, LISTA_CRUZADA, ESTADO_INSCRIPCIÓN, COD_PROGRAMA_ESTUDIANTE.
    """
    directory = _resolve_path(CONFIG.get('banner_directory', './'), './')

    if not os.path.isdir(directory):
        raise FileNotFoundError(f"[ERROR] No se encontró el directorio de archivos .xlsx: {directory}")

    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx") and not f.startswith("~$")]
    if not excel_files:
        raise FileNotFoundError("[ERROR] No se encontró ningún archivo .xlsx en el directorio actual.")

    columnas_objetivo = ['PERIODO', 'LISTA_CRUZADA', 'ESTADO_INSCRIPCIÓN', 'COD_PROGRAMA_ESTUDIANTE']
    dataframes = []

    for file in excel_files:
        filepath = os.path.join(directory, file)
        print(f"[INFO] Leyendo hoja Estudiantes (CentroCostos): {filepath}")

        try:
            # Se lee por nombre de hoja para cumplir la regla de negocio.
            df = pd.read_excel(filepath, sheet_name='Estudiantes', engine='openpyxl')
        except Exception as e:
            print(f"[ERROR] Error al leer hoja Estudiantes de {file}: {e}")
            continue

        if not set(columnas_objetivo).issubset(set(df.columns)):
            faltantes = set(columnas_objetivo) - set(df.columns)
            print(f"[WARN] Advertencia: El archivo {file} no tiene columnas requeridas para centro de costos: {faltantes}")
            continue

        df = df[columnas_objetivo].copy()
        df['PERIODO'] = df['PERIODO'].apply(_to_clean_str)
        df['LISTA_CRUZADA'] = df['LISTA_CRUZADA'].apply(_to_clean_str)
        df['ESTADO_INSCRIPCIÓN'] = df['ESTADO_INSCRIPCIÓN'].apply(_to_clean_str)
        df['COD_PROGRAMA_ESTUDIANTE'] = df['COD_PROGRAMA_ESTUDIANTE'].apply(_to_clean_str)

        # Solo se conservan estudiantes inscritos.
        df = df[df['ESTADO_INSCRIPCIÓN'].str.lower() == 'inscrito']
        df = df[
            (df['PERIODO'] != '') &
            (df['LISTA_CRUZADA'] != '') &
            (df['COD_PROGRAMA_ESTUDIANTE'] != '')
        ]

        dataframes.append(df)

    if not dataframes:
        raise ValueError("[ERROR] No fue posible construir CENTROCOSTOSESTUDIANTE desde los archivos de banner.")

    centro_costos_estudiante = pd.concat(dataframes, ignore_index=True)
    centro_costos_estudiante = centro_costos_estudiante.drop_duplicates(
        subset=['PERIODO', 'LISTA_CRUZADA', 'ESTADO_INSCRIPCIÓN', 'COD_PROGRAMA_ESTUDIANTE']
    ).reset_index(drop=True)

    print(f"[OK] CENTROCOSTOSESTUDIANTE cargado con {centro_costos_estudiante.shape[0]} filas únicas.")
    return centro_costos_estudiante

# Leer el archivo de coordinadores desde Excel, con columnas 'Centro de Costos' e 'ID COORDINADOR'.
def leer_coordinadores(ruta_archivo=None):
    """
    Lee el archivo de coordinadores parametrizado en JSON.
    Columnas requeridas: 'Centro de Costos' e 'ID COORDINADOR'.
    """
    ruta_default = CONFIG.get('coordinadores_file')
    if not ruta_default:
        # Fallback: misma carpeta de bdusuarios_file
        ruta_bdusuarios = _resolve_path(CONFIG.get('bdusuarios_file', "./BDUsuarios/Listados Usuarios.xlsx"),
                                        "./BDUsuarios/Listados Usuarios.xlsx")
        ruta_default = os.path.join(os.path.dirname(ruta_bdusuarios), 'Coordinadores.xlsx')

    if ruta_archivo is None:
        ruta_archivo = ruta_default
    ruta_archivo = _resolve_path(ruta_archivo, ruta_default)

    try:
        df = pd.read_excel(ruta_archivo, sheet_name=0, engine='openpyxl')
        columnas_requeridas = ['Centro de Costos', 'ID COORDINADOR']
        columnas_opcionales = ['Coordinador(a)', 'Correo Electrónico']

        if not set(columnas_requeridas).issubset(set(df.columns)):
            faltantes = set(columnas_requeridas) - set(df.columns)
            raise ValueError(f"Faltan columnas requeridas en archivo coordinadores: {faltantes}")

        for col in columnas_opcionales:
            if col not in df.columns:
                df[col] = ''

        df = df[columnas_requeridas + columnas_opcionales].copy()
        df['Centro de Costos'] = df['Centro de Costos'].apply(_to_clean_str)
        df['ID COORDINADOR'] = df['ID COORDINADOR'].apply(_normalizar_id_banner)
        df['Coordinador(a)'] = df['Coordinador(a)'].apply(_to_clean_str)
        df['Correo Electrónico'] = df['Correo Electrónico'].apply(_to_clean_str)

        df = df[
            (df['Centro de Costos'] != '') &
            (~df['ID COORDINADOR'].str.lower().isin(INVALID_IDS))
        ]
        df = df.drop_duplicates(subset=['Centro de Costos', 'ID COORDINADOR']).reset_index(drop=True)

        print(f"[OK] Archivo de coordinadores '{ruta_archivo}' cargado con {df.shape[0]} registros.")
        return df

    except FileNotFoundError:
        print(f"[ERROR] Error: El archivo '{ruta_archivo}' no fue encontrado.")
        return None
    except Exception as e:
        print(f"[ERROR] Error al cargar coordinadores: {e}")
        return None

#Buscar el coordinador del curso a partir del NRC/LC y Periodo, usando el DataFrame CENTROCOSTOSESTUDIANTE 
# para obtener el COD_PROGRAMA_ESTUDIANTE y luego buscar el ID del coordinador en el archivo de coordinadores. 
# Se devuelve un dict con la información del coordinador o None si no se encuentra.
def resolver_coordinador_curso(course_nrc, course_periodo, centro_costos_estudiante, bd_coordinadores, log):
    """
    Obtiene el ID del coordinador para un curso a partir de:
      1) NRC + PERIODO -> COD_PROGRAMA_ESTUDIANTE (CENTROCOSTOSESTUDIANTE)
      2) COD_PROGRAMA_ESTUDIANTE -> ID COORDINADOR (archivo coordinadores)
    """
    if centro_costos_estudiante is None or bd_coordinadores is None:
        return None, None

    nrc = _to_clean_str(course_nrc)
    periodo = _to_clean_str(course_periodo)

    match_cc = centro_costos_estudiante[
        (centro_costos_estudiante['LISTA_CRUZADA'].astype(str) == nrc) &
        (centro_costos_estudiante['PERIODO'].astype(str) == periodo)
    ]

    if match_cc.empty:
        log.write(f"[WARN] Sin COD_PROGRAMA_ESTUDIANTE para NRC={nrc}, PERIODO={periodo}\n")
        return None, None

    centros = match_cc['COD_PROGRAMA_ESTUDIANTE'].dropna().astype(str).str.strip().unique().tolist()
    centro_costo = centros[0]
    if len(centros) > 1:
        log.write(f"[WARN] NRC={nrc} PERIODO={periodo} tiene múltiples centros {centros}. Se usa: {centro_costo}\n")

    match_coord = bd_coordinadores[bd_coordinadores['Centro de Costos'] == centro_costo]
    if match_coord.empty:
        log.write(f"[WARN] Sin coordinador para Centro de Costos={centro_costo}\n")
        return None, centro_costo

    row_coord = match_coord.iloc[0]
    id_coordinador = _normalizar_id_banner(row_coord.get('ID COORDINADOR', ''))
    if id_coordinador.lower() in INVALID_IDS:
        log.write(f"[WARN] ID COORDINADOR inválido para Centro de Costos={centro_costo}\n")
        return None, centro_costo

    return row_coord, centro_costo

#Se obtiene la información del coordinador desde BDUsuarios (hoja 0) para el ID de banner dado. 
# Si no se encuentra, se devuelve None.
def obtener_datos_coordinador(id_banner, bd_usuarios):
    """
    Obtiene información del coordinador desde BDUsuarios (hoja 0).
    """
    user = bd_usuarios.loc[bd_usuarios['UserName'] == id_banner]
    if user.empty:
        return None

    row = user.iloc[0]
    return {
        'docuusu': _to_clean_str(row.get('OrgDefinedId', '')) or id_banner,
        'first_name': _to_clean_str(row.get('FirstName', '')),
        'last_name': _to_clean_str(row.get('LastName', '')),
        'email': _to_clean_str(row.get('ExternalEmail', ''))
    }

#se crea el archivo de registro para cada curso
def crearArchivos(data, course_name, course_nrc, course_periodo, BDUsuBS, centro_costos_estudiante,
                  bd_coordinadores, log_file_path='log_creacion_moderadores.txt'):
    """
    Genera comandos de inscripción y creación/actualización para:
      1) Docente con rol Moderador (flujo original).
      2) Coordinador con rol Coordinador (nuevo flujo).

    Parámetros:
        data (pd.DataFrame): Datos de los docentes por curso.
        course_name (str): Nombre del curso.
        course_nrc (str): NRC del curso.
        course_periodo (str): Periodo del curso.
        BDUsuBS (pd.DataFrame): Base de usuarios de Brightspace.
        centro_costos_estudiante (pd.DataFrame): CENTROCOSTOSESTUDIANTE.
        bd_coordinadores (pd.DataFrame): Archivo de coordinadores.
        log_file_path (str): Ruta al archivo de log.
    """
    rol_moderador = "Moderador"
    rol_coordinador = "Coordinador"
    line_count = 0

    directory = _resolve_path(CONFIG.get('salida_directory', './salida/'), './salida/')
    os.makedirs(directory, exist_ok=True)
    archivo_comandos = os.path.join(directory, f"registro_{course_name}.txt")
    usuarios_bs = set(BDUsuBS['UserName'].astype(str).tolist())

    with open(archivo_comandos, 'a', encoding='utf8') as fptr, \
         open(log_file_path, 'a', encoding='utf8') as log, \
         open('moderadores.csv', 'a', encoding='utf8', newline='') as moderadores:

        writer = csv.writer(moderadores)
        log.write(f"\n=== PROCESAMIENTO CURSO: {course_name} - NRC: {course_nrc} ===\n")
        log.write(f"Fecha: {datetime.now()}\n")

        # 1) Inscripción de docentes moderadores (flujo existente).
        for _, row in data.iterrows():
            idBanner = _normalizar_id_banner(row.get('ID_DOCENTE', ''))
            if idBanner.lower() in INVALID_IDS:
                log.write(f"[ERROR] ID inválido: '{idBanner}' para curso {course_name}\n")
                continue

            Unuevo = idBanner not in usuarios_bs
            RolModerador = BDUsuBS.loc[BDUsuBS['UserName'] == idBanner, 'OrgRoleId'].values

            try:
                ndocu = "{:,}".format(int(row['DOCUMENTO'])).replace(',', '.')
            except Exception:
                ndocu = str(row['DOCUMENTO'])

            try:
                docuusu = f"{row['TIPO_DOCUMENTO']}. {ndocu}"
            except Exception:
                docuusu = ndocu

            first_name = str(row.get('NOMBRE_DOCENTE', '')).strip()
            last_name = str(row.get('APELLIDO_DOCENTE', '')).strip()
            email = str(row.get('CORREO_DOCENTE', '')).strip()

            if Unuevo:
                fptr.write(f'CREATE,{idBanner},{docuusu},{first_name},{last_name},,{rol_moderador},1,{email}\n')
            else:
                fptr.write(f'UPDATE,{idBanner},{docuusu},{first_name},{last_name},,1,{email}\n')
                if RolModerador.size > 0:
                    rol = str(RolModerador[0])
                    mapeo_roles = {
                        '150': 'CVTE', '143': 'CVLA', '138': 'CVPR',
                        '137': 'CVFC', '136': 'CVFA', '135': 'CVFA'
                    }
                    if rol in mapeo_roles:
                        fptr.write(f'UNENROLL,{idBanner},,{mapeo_roles[rol]}\n')

                fptr.write(f'ENROLL,{idBanner},,{rol_moderador},UPBV\n')

            fptr.write(f'ENROLL,{idBanner},,{rol_moderador},{course_name}\n')
            line_count += 1

        # 2) Inscripción de coordinador por curso (nuevo flujo).
        row_coord, centro_costo = resolver_coordinador_curso(
            course_nrc, course_periodo, centro_costos_estudiante, bd_coordinadores, log
        )
        if row_coord is not None:
            id_coord = _normalizar_id_banner(row_coord.get('ID COORDINADOR', ''))
            if id_coord.lower() not in INVALID_IDS:
                coord_nuevo = id_coord not in usuarios_bs
                datos_coord = obtener_datos_coordinador(id_coord, BDUsuBS)

                if datos_coord is None:
                    # Fallback mínimo cuando el coordinador no está en BDUsuarios.
                    datos_coord = {
                        'docuusu': id_coord,
                        'first_name': _to_clean_str(row_coord.get('Coordinador(a)', '')),
                        'last_name': '',
                        'email': _to_clean_str(row_coord.get('Correo Electrónico', ''))
                    }

                if coord_nuevo:
                    fptr.write(
                        f"CREATE,{id_coord},{datos_coord['docuusu']},{datos_coord['first_name']},"
                        f"{datos_coord['last_name']},,{rol_coordinador},1,{datos_coord['email']}\n"
                    )

                fptr.write(f'ENROLL,{id_coord},,{rol_coordinador},{course_name}\n')
                log.write(
                    f"[OK] Coordinador inscrito NRC={course_nrc}, PERIODO={course_periodo}, "
                    f"CentroCosto={centro_costo}, ID={id_coord}\n"
                )

        writer.writerow([course_name, course_nrc, line_count])

        print(f"[OK] Se han inscrito: {line_count} moderadores en el curso: {course_name} NRC: {course_nrc}")
        log.write(f"[OK] Total moderadores inscritos: {line_count}\n")

