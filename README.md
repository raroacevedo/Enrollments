# README - Documentacion de `config.json`

## Proposito
El archivo `config.json` centraliza la configuracion operativa del bot de inscripciones.
Su objetivo es desacoplar rutas, fuentes de datos y modo de ejecucion del codigo Python, para que el proceso pueda moverse entre ambientes sin editar scripts.

En este proyecto, `config.json` controla principalmente:

- Donde leer archivos fuente de Banner y Brightspace.
- Donde escribir archivos de salida para carga masiva.
- Que tipo de operacion ejecutar en el flujo de estudiantes.
- Que archivo usar para resolver coordinadores en el flujo de docentes.

## Alcance
La configuracion aplica a estos procesos:

- Proceso de estudiantes: `inscribirEstV2.py` + `helpersestV2.py`.
- Proceso de docentes/moderadores/coordinadores: `inscribirModV2.py` + `helpersmodV2.py`.

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
  - Estudiantes: leer Excel Banner (hojas de estudiantes).
  - Docentes: leer Excel Banner (hoja `Docentes`).
  - Docentes refactor V2: leer hoja `Estudiantes` para construir `CENTROCOSTOSESTUDIANTE`.
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
- Tipo: `string` (ruta de archivo Excel). En local apunta el archivo en ONEDRIVE: https://upbeduco.sharepoint.com/sites/SharepointUPBVirtual/Documentos%20compartidos/_COORDINADORES%20PROGRAMAS%20VIRTUALES.xlsx?web=1
- Requerido:
  - Recomendado como obligatorio para el flujo docente V2.
  - Si no se define, el helper intenta fallback automatico en la misma carpeta de `bdusuarios_file` con nombre `Coordinadores.xlsx`.
- Uso:
  - Resolver coordinador por centro de costos para cada curso.
  - Flujo: `NRC + Periodo` -> `COD_PROGRAMA_ESTUDIANTE` -> `Centro de Costos` -> `ID COORDINADOR`.
- Hoja usada:
  - Hoja 0.
- Columnas requeridas:
  - `Centro de Costos`, `ID COORDINADOR`.
- Columnas opcionales (fallback de datos para `CREATE`):
  - `Coordinador(a)`, `Correo Electronico`.

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
  - `Matricular`
  - `Desmatricular`
  - `Limpieza`
- Nota:
  - En el flujo docente actual, esta llave no altera la logica principal de inscripcion de moderadores/coordinadores.

## Resumen rapido por proceso

| Proceso | Llaves usadas |
|---|---|
| Estudiantes | `banner_directory`, `bdusuarios_file`, `salida_directory`, `Tipo_proceso` |
| Docentes (Moderador + Coordinador) | `banner_directory`, `bdusuarios_file`, `coordinadores_file`, `salida_directory` |

## Reglas operativas importantes
- El archivo de usuarios (`bdusuarios_file`) se lee desde la hoja 0.
- En docentes V2, `CENTROCOSTOSESTUDIANTE` se precarga antes del loop de inscripcion para mejorar eficiencia.
- En docentes V2:
  - Primero se procesa rol `Moderador`.
  - Luego se agrega rol `Coordinador` por curso.
  - Para coordinador:
    - Siempre se genera `ENROLL` al curso.
    - Se genera `CREATE` solo si el usuario no existe en `BDUsuarios`.

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
