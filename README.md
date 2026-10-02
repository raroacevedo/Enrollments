# README - Documentacion de `config.json`

## Proposito
El archivo `config.json` centraliza la configuracion operativa del bot de inscripciones.
Su objetivo es desacoplar rutas, fuentes de datos y modo de ejecucion del codigo Python, para que el proceso pueda moverse entre ambientes sin editar scripts.

En este proyecto, `config.json` controla principalmente:

- Donde leer archivos fuente de Banner y Brightspace.
- Donde escribir archivos de salida para carga masiva.
- Que tipo de operacion ejecutar en el flujo de estudiantes.
- Que archivo usar para resolver coordinadores y mentores en el flujo de docentes.

## Alcance
La configuracion aplica a estos procesos:

- Proceso de estudiantes: `inscribirEstV2.py` + `helpersestV2.py`.
- Proceso de docentes/moderadores/coordinadores/mentores: `inscribirModV2.py` + `helpersmodV2.py`.

No aplica automaticamente a scripts que no usen `load_config()` o `CONFIG.get(...)`.

## Ubicacion y carga
- Archivo: `./config.json` (raiz del proyecto).
- Formato: objeto JSON plano.
- Carga: cada helper usa `load_config(path="config.json")`.
- Resolucion de rutas:
  - Si la ruta en JSON es absoluta, se usa tal cual.
  - Si es relativa, se resuelve respecto a la carpeta del script helper.

## Estructura actual del JSON
```json
{
  "banner_directory": "C:/.../ListadosEstudiantesDocentesBanner/",
  "bdusuarios_file": "C:/.../Listados Usuarios.xlsx",
  "coordinadores_file": "C:/.../Coordinadores.xlsx",
  "salida_directory": "C:/.../salida/2026/",
  "Tipo_proceso": "Matricular"
}
```

## Llaves del JSON

### 1) `banner_directory`
- Tipo: `string` (ruta de directorio).
- Requerido: si.
- Uso:
  - Estudiantes: leer Excel Banner (hojas de `Estudiantes`).
  - Docentes: leer Excel Banner (hoja `Docentes`).
- Esperado en origen:
  - Archivos `.xlsx` de Banner con hojas `Docentes` y `Estudiantes`.
- Si falla:
  - Si el directorio no existe, se lanza error de archivo no encontrado.
  - Si no hay `.xlsx` validos, se detiene el proceso.

### 2) `bdusuarios_file`
- Tipo: `string` (ruta de archivo Excel).
- Requerido: si.
- Uso:
  - Base de usuarios de Brightspace para validar si un usuario ya existe.
  - Soporte para `CREATE`/`UPDATE` y validacion de roles.
- Hoja usada:
  - Hoja 0 (primera hoja), segun regla operativa actual.
- Columnas consumidas:
  - `UserName`, `FirstName`, `LastName`, `OrgRoleId`, `OrgDefinedId`, `ExternalEmail`.
- Si falla:
  - Sin este archivo no se puede determinar existencia de usuarios ni construir comandos confiables.

### 3) `coordinadores_file`
- Tipo: `string` (ruta de archivo Excel). IMPORTANTE! En local apunta el archivo en ONEDRIVE: https://upbeduco.sharepoint.com/sites/SharepointUPBVirtual/Documentos%20compartidos/_COORDINADORES%20PROGRAMAS%20VIRTUALES.xlsx?web=1
- Requerido:
  - Recomendado como obligatorio para el flujo docente V2. Se usa el proceso de inscripcion de moderadores para inscribir los coordinadores de cada NRC
  - Si no se define, el helper intenta fallback automatico en la misma carpeta de `bdusuarios_file` con nombre `Coordinadores.xlsx`.
- Uso:
  - Resolver coordinador y hasta tres mentores por centro de costos para cada curso.
  - Flujo: `NRC + Periodo` -> `COD_PROGRAMA_ESTUDIANTE` -> `Centro de Costos` -> responsables.
- Hoja usada:
  - Hoja 0.
- Columnas requeridas:
  - `Centro de Costos`, `Coordinador_ID`, `Mentor1_Id`, `Mentor2_Id`, `Mentor3_Id`.
- Columnas opcionales (fallback de datos para `CREATE`):
  - `Coordinador_Nombre`, `Coordinador_Correo`.
  - `Mentor1_Nombres`, `Mentor1_Correo`.
  - `Mentor2_Nombres`, `Mentor2_Correo`.
  - `Mentor3_Nombres`, `Mentor3_Correo`.
- Validación:
  - Las filas sin centro de costos o sin ningún ID válido se descartan.
  - Si falta una columna requerida, la fuente se rechaza y el proceso se detiene.

### 4) `salida_directory`
- Tipo: `string` (ruta de directorio).
- Requerido: si.
- Uso:
  - Carpeta destino para los archivos `registro_<curso>.txt`.
  - Base para generar consolidado `registro_unico*.txt`.
- Recomendacion:
  - Usar una carpeta por anio/periodo para trazabilidad.
  - Garantizar permisos de escritura antes de ejecutar.

### 5) `Tipo_proceso`
- Tipo: `string`.
- Requerido: si en flujo de estudiantes.
- Uso:
  - Controla el comportamiento del proceso de estudiantes.
- Valores esperados:
  - `Matricular` - Estado en Banner Inscrito
  - `Desmatricular` Estado en Banner Cancelado
  - `Limpiar`  - Estado en Banner Eliminado (DL)
- Nota:
  - En el flujo docente actual, esta llave no altera la logica principal de inscripcion de moderadores/coordinadores.

## Resumen rapido por proceso

| Proceso | Llaves usadas |
|---|---|
| Estudiantes | `banner_directory`, `bdusuarios_file`, `salida_directory`, `Tipo_proceso` |
| Docentes (Moderador + Coordinador + Mentor) | `banner_directory`, `bdusuarios_file`, `coordinadores_file`, `salida_directory` |

## Reglas operativas importantes
- El archivo de usuarios (`bdusuarios_file`) se lee desde la hoja 0.
- En docentes V2, `CENTROCOSTOSESTUDIANTE` se precarga antes del loop de inscripcion para mejorar eficiencia.
- En docentes V2:
  - Primero se procesa rol `Moderador`.
  - Luego se procesan `Coordinador` y `Mentor` por curso.
  - La precedencia de rol es `Moderador > Coordinador > Mentor`.
  - Si el `Coordinador_ID` ya fue inscrito como `Moderador`, no se genera el
    `ENROLL` como `Coordinador`.
  - Los IDs repetidos entre `Mentor1_Id`, `Mentor2_Id` y `Mentor3_Id` se inscriben una sola vez.
  - Si un mentor ya tiene un rol de mayor precedencia en el curso, se omite el rol `Mentor`.
  - Para coordinadores y mentores se genera `CREATE` solo si el usuario no existe en `BDUsuarios`.

## Buenas practicas de mantenimiento
- Mantener rutas absolutas estables en ambientes productivos.
- Evitar espacios finales y errores de escritura en nombres de archivo.
- Validar que los archivos fuente tengan la estructura esperada antes de ejecutar.
- Versionar cambios de `config.json` en control de versiones.
- No exponer rutas o archivos con datos sensibles fuera del equipo operativo.

## Checklist previo a ejecucion
1. `banner_directory` existe y contiene los Excel correctos.
2. `bdusuarios_file` existe y es la version vigente de usuarios.
3. `coordinadores_file` existe y tiene columnas requeridas.
4. `salida_directory` existe o puede ser creada por el proceso.
5. `Tipo_proceso` coincide con la operacion planeada (en estudiantes).

## Ejemplo recomendado para nuevos ambientes
```json
{
  "banner_directory": "D:/Datos/Banner/",
  "bdusuarios_file": "D:/Datos/Brightspace/Listados Usuarios.xlsx",
  "coordinadores_file": "D:/Datos/Brightspace/Coordinadores.xlsx",
  "salida_directory": "D:/Enrollments/salida/2026/",
  "Tipo_proceso": "Matricular"
}
```
---
Si se agrega una nueva llave al JSON, documentarla aqui con: objetivo, tipo, valor por defecto, donde se usa y que pasa si falta.

# README - Documentación proceso de inscripción de Moderadores y Estudiantes

## Propósito

Los bots de inscripción de docentes y estudiantes generan archivos para la creación de usuarios nuevos, asignación de roles y enrolamiento en cursos de Brightspace/D2L.

El proceso permite preparar los archivos necesarios para inscribir usuarios en los respectivos NRC o listas cruzadas, de acuerdo con la información proveniente de Banner, Brightspace, Domo y archivos operativos de programación académica.

Los roles gestionados por los bots son:

- `Student`: rol utilizado para la inscripción de estudiantes.
- `Moderador`: rol utilizado para docentes o moderadores del curso.
- `Coordinador`: rol utilizado para coordinadores de programas virtuales.
- `Mentor`: rol utilizado para los mentores configurados por programa.

La inscripción de moderadores incluye tres procesos internos de enrolamiento:

- Docentes moderadores.
- Coordinadores.
- Mentores.

Los procesos de inscripción de docentes y estudiantes dependen obligatoriamente del archivo `shortnames.csv`, el cual se genera previamente mediante el script `get_shortname.py`.

## Alcance

Este README documenta el flujo general de los siguientes procesos:

- Obtención de fuente primaria de cursos: `get_shortname.py`.
- Inscripción de docentes, moderadores y coordinadores: `inscribirModV2.py`.
- Inscripción de estudiantes: `inscribirEstV2.py`.

El documento está orientado a usuarios operativos o técnicos que ejecutan los bots y requieren conocer las fuentes de entrada, archivos generados, rutas de configuración y orden correcto de ejecución.

## Flujo general del proceso

```text
Programación de cursos virtuales
        |
        v
ListaCursos.csv
        |
        v
get_shortname.py
        |
        v
shortnames.csv
        |
        +----------------------+
        |                      |
        v                      v
inscribirModV2.py       inscribirEstV2.py
        |                      |
        v                      v
Archivos de              Archivos de
moderadores              estudiantes
```

## Configuración requerida

Los scripts utilizan el archivo `config.json` para ubicar fuentes de entrada y carpetas de salida.

### Llaves principales del JSON

| Llave                | Uso | Proceso asociado |
|----------------------|-----|------------------|
| `banner_directory`   | Ruta donde se encuentra el archivo consolidado con información de Banner. Debe contener las hojas `Docentes` y `Estudiantes`, según el proceso ejecutado. | `inscribirModV2.py` / `inscribirEstV2.py` |
| `bdusuarios_file`    | Ruta del archivo `Listados Usuarios.xlsx`, obtenido desde Domo.                      | `inscribirModV2.py` / `inscribirEstV2.py` |
| `salida_directory`   | Ruta donde se almacenan los archivos generados por curso y los consolidados finales. | `inscribirModV2.py` / `inscribirEstV2.py` |
| `coordinadores_file` | Ruta del archivo de coordinadores y mentores de programas virtuales. | `inscribirModV2.py` |

## Paso 1. Obtener fuente primaria de cursos

### Script

```bash
python get_shortname.py
```

## Objetivo

Generar el archivo `shortnames.csv`, que será usado como fuente obligatoria para los procesos de inscripción de docentes y estudiantes.

Este paso debe ejecutarse antes de `inscribirModV2.py` e `inscribirEstV2.py`.

## Fuente de entrada

El script utiliza como fuente de datos el archivo:

```text
ListaCursos.csv
```

Este archivo debe contener los enlaces de los cursos que serán procesados.

## Origen de la información

Los enlaces de los cursos normalmente se obtienen del archivo de programación de cursos virtuales:

```text
ProgramacionDeAperturaCursosVirtuales.xlsx
```

Este archivo se encuentra alojado en OneDrive:

```text
https://upbeduco-my.sharepoint.com/:x:/r/personal/teamupbvirtual_upb_edu_co/Documents/AreaTecnologica/Plataformas/D2L/Matriculas/DuplicadosCursos/ProgramacionDeAperturaCursosVirtuales.xlsx?d=w4581ea79f1924a7f9487e9e0597b6828&csf=1&web=1&e=oHRNlv
```

IMPORTANTE: solo pueden acceder los usuarios que tengan permisos sobre el archivo.

## Preparación del archivo `ListaCursos.csv`

1. Abrir el archivo `ProgramacionDeAperturaCursosVirtuales.xlsx`.
2. Filtrar los cursos por fecha de inicio usando la columna `O`.
3. Copiar los enlaces de los cursos desde la columna `M`, denominada `URL BS`.
4. Pegar los enlaces copiados en el archivo `ListaCursos.csv`.
5. Guardar el archivo antes de ejecutar `get_shortname.py`.

## Proceso realizado por `get_shortname.py`

El script realiza scripting sobre los enlaces incluidos en `ListaCursos.csv` y extrae el código de oferta de cada curso.

La estructura esperada del código de oferta es:

```text
Materia-Curso-Periodo-LC/NRC
```

Con esta información, el proceso identifica:

| Dato    | Uso          |
|---------|--------------|
| Curso   | Permite identificar el curso donde se realizará la inscripción. |
| Periodo | Permite filtrar la información correspondiente en la fuente de Banner. |
| LC/NRC  | Permite filtrar los usuarios que deben ser inscritos. |

## Archivo generado

El proceso genera el archivo:

```text
shortnames.csv
```

Con los siguientes campos:

| Campo | Descripción |
|------ |-------------|
| `Nombre` | Nombre o shortname del curso. |
| `NRC` | NRC o lista cruzada asociada al curso. |
| `Periodo` | Periodo académico identificado en el enlace del curso. |

## Paso 2. Inscripción de docentes, moderadores y coordinadores

### Script

```bash
python inscribirModV2.py
```

## Objetivo

Generar los archivos requeridos para inscribir docentes, moderadores, mentores y coordinadores en los cursos identificados previamente mediante `shortnames.csv`.

## Fuentes de información requeridas

| Fuente | Ubicación / configuración | Descripción |
|--------|---------------------------|-------------|
| Consolidado de docentes | Clave `banner_directory` en `config.json` | Archivo Excel con hoja `Docentes`. Consolida la información ubicada en SharePoint `SZREINS/Docentes`. |
| `Listados Usuarios.xlsx` | Clave `bdusuarios_file` en `config.json` | Base de usuarios obtenida desde Domo. Permite consultar datos de docentes para creación e inscripción. |
| Archivo de responsables | Clave `coordinadores_file` en `config.json` | Archivo Excel usado para la inscripción de coordinadores y mentores de programas virtuales. |
| `shortnames.csv` | Generado por `get_shortname.py` | Contiene los cursos, NRC y periodos que serán procesados. |

## Archivo de coordinadores y mentores

El proceso de inscripción de moderadores también crea la inscripción de los
coordinadores y mentores de programas virtuales.

La fuente de información es un archivo Excel ubicado en la ruta configurada en la llave:

```json
"coordinadores_file"
```

Este archivo local se sincroniza con el archivo de OneDrive:

```text
https://upbeduco.sharepoint.com/sites/SharepointUPBVirtual/Documentos%20compartidos/_COORDINADORES%20PROGRAMAS%20VIRTUALES.xlsx?web=1
```

### Contrato de columnas

| Columna | Obligatoria | Uso |
|---|---|---|
| `Centro de Costos` | Sí | Relaciona el programa encontrado en Banner. |
| `Coordinador_ID` | Sí | ID Banner del coordinador; la celda puede estar vacía. |
| `Mentor1_Id` | Sí | Primer mentor opcional. |
| `Mentor2_Id` | Sí | Segundo mentor opcional. |
| `Mentor3_Id` | Sí | Tercer mentor opcional. |
| Columnas de nombre/correo | No | Respaldo para crear usuarios que no estén en `bdusuarios_file`. |

### Reglas de asignación de roles

1. Se inscriben primero los docentes con rol `Moderador`.
2. Si `Coordinador_ID` coincide con un moderador del curso, no se inscribe como
   `Coordinador` y el evento queda registrado en `log_creacion_moderadores.txt`.
3. Los mentores no vacíos se normalizan a nueve dígitos y se deduplican.
4. Para evitar que Brightspace reemplace un rol por otro dentro del mismo curso,
   se aplica la precedencia `Moderador > Coordinador > Mentor`.
5. Los usuarios que no existen en la base Brightspace reciben `CREATE` con el
   rol correspondiente antes del `ENROLL` al curso.

## Archivos generados

| Archivo | Descripción |
|---------|-------------|
| `registro_unicoMOD.txt` | Consolidado de todos los NRC a inscribir con moderadores, mentores y coordinadores. Este es el archivo principal que se sube para la inscripción. |
| `registro-codigocurso.txt` | Archivo individual por curso con los moderadores correspondientes. Se almacena en la ruta configurada en `salida_directory`. |
| `moderadores.csv` | Resumen con el número de docentes o moderadores inscritos por curso. |
| `log_creacion_moderadores.txt` | Log de moderadores, coordinadores y mentores, incluidas omisiones por precedencia. |

## Consideraciones del proceso de moderadores

- El archivo `shortnames.csv` debe existir antes de ejecutar el proceso.
- La hoja `Docentes` debe estar disponible en el archivo consolidado de Banner.
- El archivo `Listados Usuarios.xlsx` debe estar actualizado desde Domo.
- El archivo de coordinadores y mentores debe estar disponible, sincronizado y usar el contrato de columnas vigente.
- La ruta `salida_directory` debe existir o permitir escritura de archivos.

## Paso 3. Inscripción de estudiantes

### Script

```bash
python inscribirEstV2.py
```

## Objetivo

Generar los archivos requeridos para inscribir estudiantes en los cursos identificados previamente mediante `shortnames.csv`.

## Fuentes de información requeridas

| Fuente | Ubicación / configuración | Descripción |
|--------|---------------------------|-------------|
| Consolidado de estudiantes | Clave `banner_directory` en `config.json` | Archivo Excel con hoja `Estudiantes`. Consolida la información ubicada en SharePoint `SZREINS/Estudiantes`. |
| `Listados Usuarios.xlsx` | Clave `bdusuarios_file` en `config.json` | Base de usuarios obtenida desde Domo. Permite consultar datos de estudiantes para creación e inscripción. |
| `shortnames.csv` | Generado por `get_shortname.py` | Contiene los cursos, NRC y periodos que serán procesados. |

## Archivos generados

| Archivo | Descripción |
|---------|-------------|
| `registro_unicoEst.txt` | Consolidado de todos los NRC a inscribir con estudiantes. Este es el archivo principal que se sube para la inscripción. |
| `registro-codigocurso.txt` | Archivo individual por curso con los estudiantes correspondientes. Se almacena en la ruta configurada en `salida_directory`. |
| `students.csv` | Resumen con el número de estudiantes inscritos por curso. |

## Consideraciones del proceso de estudiantes

- El archivo `shortnames.csv` debe existir antes de ejecutar el proceso.
- La hoja `Estudiantes` debe estar disponible en el archivo consolidado de Banner.
- El archivo `Listados Usuarios.xlsx` debe estar actualizado desde Domo.
- La ruta `salida_directory` debe existir o permitir escritura de archivos.

## Orden recomendado de ejecución

| Orden | Script | Descripción |
|---|---|---|
| 1 | `get_shortname.py` | Genera el archivo base `shortnames.csv`. |
| 2 | `inscribirModV2.py` | Genera archivos de inscripción de docentes, moderadores, mentores y coordinadores. |
| 3 | `inscribirEstV2.py` | Genera archivos de inscripción de estudiantes. |

## Resultado esperado por proceso

| Proceso | Archivo principal para cargar | Archivos de control |
|---------|-------------------------------|---------------------|
| Moderadores / Docentes / Coordinadores / Mentores | `registro_unicoMOD.txt` | `moderadores.csv`, `log_creacion_moderadores.txt`, archivos `registro-codigocurso.txt`. |
| Estudiantes | `registro_unicoEst.txt` | `students.csv`, archivos `registro-codigocurso.txt`. |

## Checklist previo a la ejecución

1. Validar que `ListaCursos.csv` contenga únicamente los enlaces de los cursos a procesar.
2. Confirmar que los enlaces fueron tomados desde la columna `M` - `URL BS`.
3. Confirmar que el filtro inicial se realizó por fecha de inicio usando la columna `O`.
4. Ejecutar `get_shortname.py`.
5. Validar que se haya generado correctamente `shortnames.csv`.
6. Confirmar que `config.json` tenga rutas correctas.
7. Validar disponibilidad de `Listados Usuarios.xlsx`.
8. Validar disponibilidad de la hoja `Docentes`, si se ejecutará `inscribirModV2.py`.
9. Validar disponibilidad de la hoja `Estudiantes`, si se ejecutará `inscribirEstV2.py`.
10. Validar disponibilidad y columnas del archivo de coordinadores y mentores, si se ejecutará el proceso de moderadores.
11. Confirmar permisos de lectura sobre las fuentes de entrada.
12. Confirmar permisos de escritura sobre la carpeta definida en `salida_directory`.

## Buenas prácticas de operación

- Ejecutar siempre primero `get_shortname.py`.
- Usar archivos fuente actualizados antes de cada ejecución.
- Validar los archivos individuales por curso antes de cargar los consolidados.
- Conservar los archivos generados por periodo académico para trazabilidad.
- No modificar manualmente los archivos generados sin dejar registro del cambio.
- No exponer públicamente archivos que contengan datos personales o información académica.
- Mantener actualizado el archivo `config.json` según el ambiente de ejecución.

## Errores frecuentes

| Situación | Posible causa | Acción recomendada |
|-----------|---------------|--------------------|
| No se genera `shortnames.csv` | `ListaCursos.csv` está vacío o contiene enlaces incorrectos. | Validar que los enlaces provengan de la columna `URL BS`. |
| No se encuentran docentes | La hoja `Docentes` no existe o el archivo de Banner no está actualizado. | Validar el archivo ubicado en `banner_directory`. |
| No se encuentran estudiantes | La hoja `Estudiantes` no existe o el archivo de Banner no está actualizado. | Validar el archivo ubicado en `banner_directory`. |
| No se crean usuarios correctamente | `Listados Usuarios.xlsx` no está actualizado o tiene estructura diferente. | Actualizar la base desde Domo y validar columnas. |
| No se generan archivos de salida | La ruta `salida_directory` no existe o no tiene permisos de escritura. | Crear la carpeta o validar permisos. |
| No se inscriben coordinadores | El archivo definido en `coordinadores_file` no existe o no está sincronizado. | Validar la sincronización local con OneDrive. |

## Resumen rápido

El proceso inicia con la preparación de `ListaCursos.csv`, continúa con la ejecución de `get_shortname.py` para generar `shortnames.csv` y finaliza con la ejecución de los scripts de inscripción:

- `inscribirModV2.py` para docentes, moderadores, mentores y coordinadores.
- `inscribirEstV2.py` para estudiantes.

Los archivos principales resultantes son:

- `registro_unicoMOD.txt`
- `registro_unicoEst.txt`

Estos archivos son los consolidados que se usan para realizar la carga de inscripción.

---

# Proceso de inscripción de estudiantes V3

## Objetivo

`inscribirEstV3.py` reemplaza la preparación manual de `ListaCursos.csv` y la
ejecución previa de `get_shortname.py` únicamente para el flujo V3. Las versiones
V2 se conservan sin cambios en su orden de ejecución.

V3 realiza estas etapas, en orden:

1. Descarga `ProgramacionDeAperturaCursosVirtuales.xlsx` desde el enlace público
   de OneDrive y usa la copia local sincronizada como respaldo.
2. Abre la hoja indicada en `programacion_cursos_hoja`.
3. Filtra la columna `Inicio` usando el rango inclusivo configurado en
   `programacion_cursos_fecha_inicial` y `programacion_cursos_fecha_final`.
4. Conserva únicamente las filas cuya columna `Estado` contiene `✔️`.
5. Construye y reemplaza `shortnames.csv`.
6. Lee la fuente Banner en modo streaming y conserva solo filas de los
   periodos/NRC/listas cruzadas seleccionados.
7. Genera los archivos individuales, el consolidado y el resumen de estudiantes.

## Configuración V3

```json
{
  "programacion_cursos_url": "https://.../ProgramacionDeAperturaCursosVirtuales.xlsx?...",
  "programacion_cursos_preferir_remoto": true,
  "programacion_cursos_file": "C:/.../ProgramacionDeAperturaCursosVirtuales.xlsx",
  "programacion_cursos_hoja": "Duplicados 2026",
  "programacion_cursos_fecha_inicial": "2026-07-13",
  "programacion_cursos_fecha_final": "2026-07-13",
  "shortnames_file": "shortnames.csv",
  "log_file": "logs/inscribirEstV3.log"
}
```

| Llave | Uso |
|---|---|
| `programacion_cursos_url` | URL canónica del Excel en OneDrive/SharePoint. |
| `programacion_cursos_preferir_remoto` | Si es `true`, consulta primero el enlace público para obtener la versión más reciente. |
| `programacion_cursos_file` | Copia local sincronizada utilizada como respaldo si falla la descarga remota. |
| `programacion_cursos_hoja` | Hoja anual que contiene la programación. |
| `programacion_cursos_fecha_inicial` | Primer día incluido en el filtro de la columna `Inicio`, en formato `YYYY-MM-DD`. |
| `programacion_cursos_fecha_final` | Último día incluido en el filtro de la columna `Inicio`, en formato `YYYY-MM-DD`. Debe ser igual o posterior a la fecha inicial. |
| `shortnames_file` | Archivo CSV que V3 reemplaza en su primera etapa. |
| `log_file` | Log rotativo que almacena información, advertencias y errores. |

V3 agrega automáticamente el parámetro de descarga al enlace público. Si la red
o SharePoint no están disponibles, usa `programacion_cursos_file` como respaldo
y registra la advertencia correspondiente en el log.

## Reglas de construcción de `shortnames.csv`

- Solo se procesan filas con fecha `Inicio` dentro del rango configurado,
  incluyendo ambos límites.
- Solo se procesan filas cuyo encabezado `Estado` tenga el valor `✔️`. La
  validación se realiza por nombre de columna y no por su posición física.
- El código se forma como `Materia Curso-Periodo-Lista Cruzada/NRC`.
- `Lista Cruzada` tiene precedencia sobre `NRC`.
- Una lista cruzada repetida genera un solo registro por periodo.
- Una fila con `NRC Ciclo Integración` genera el registro principal y otro con
  el NRC de integración, usando el periodo de su fila auxiliar siguiente.
- La fila auxiliar de integración no se genera nuevamente como curso independiente.
- Textos operativos como `Revisar Listado` o `?` no se aceptan como códigos y
  quedan registrados como advertencias en el log.
- Si el Excel contiene metadatos de estilo inválidos, V3 sanea una copia temporal
  en memoria; nunca modifica el archivo sincronizado.

El encabezado generado siempre es:

```text
Nombre,NRC,Periodo
```

## Ejecución

Proceso completo:

```bash
python inscribirEstV3.py
```

Con fecha mínima de actividad Banner, manteniendo compatibilidad con V2:

```bash
python inscribirEstV3.py 13/07/26
```

Generar y validar solamente `shortnames.csv`:

```bash
python inscribirEstV3.py --solo-shortnames
```

## Archivos generados por V3

| Archivo | Descripción |
|---|---|
| `shortnames.csv` | Fuente de cursos generada al iniciar V3. |
| `salida/registro_<codigo-curso>.txt` | Comandos del curso de la ejecución actual. |
| `registro_unicoEst.txt` | Consolidado creado directamente, sin mezclar archivos antiguos de la carpeta de salida. |
| `students.csv` | Total único por curso y NRC/listas fuente asociados. |
| `logs/inscribirEstV3.log` | Trazabilidad completa de la ejecución. |

V3 abre los archivos Banner con `openpyxl` en modo de solo lectura, proyecta las
columnas requeridas, filtra antes de conservar estudiantes y usa conjuntos para
consultas de usuarios y deduplicación. Esto evita cargar el Excel completo en un
`DataFrame` y elimina búsquedas lineales repetidas por cada curso.

## Pruebas

```bash
python -m unittest discover -s tests -v
```

Las pruebas cubren ciclo de integración, lista cruzada, códigos inválidos,
estilos XML defectuosos y procesamiento streaming de Banner.
