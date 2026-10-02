#!/usr/bin/env python
"""Servicios del proceso V3 de inscripcion de estudiantes.

La version V3 genera ``shortnames.csv`` desde la programacion de cursos y
procesa las fuentes de Banner en modo streaming para no cargar cientos de
miles de filas que no pertenecen a los cursos seleccionados.
"""

from __future__ import annotations

import csv
import io
import json     #libreria para trabajar con archivos y cadenas JSON, incluyendo la serializacion y deserializacion de datos.
import logging  #libreria para registrar eventos y mensajes de depuracion, informacion y errores en aplicaciones Python.
from logging.handlers import RotatingFileHandler  #libreria para manejar archivos de log con rotacion y limite de tamaño
import os       #libreria para interactuar con el sistema operativo, incluyendo la manipulación de archivos y directorios, variables de entorno y rutas de archivos.
from contextlib import ExitStack    #libreria para manejar multiples contextos de manera segura y eficiente
from dataclasses import dataclass   #libreria para crear clases de datos inmutables y con slots
from datetime import date, datetime
from pathlib import Path
import re
import tempfile
import unicodedata                  #libreria para normalizar cadenas Unicode y eliminar acentos y caracteres especiales
from typing import Any, Iterable, Iterator, Mapping, Sequence       #libreria para definir tipos de datos y anotaciones de tipo en Python
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit #libreria para analizar y construir URLs, incluyendo la codificacion y decodificacion de parametros de consulta
from zipfile import ZIP_DEFLATED, ZipFile                           #libreria para trabajar con archivos ZIP, incluyendo la compresion y descompresion de archivos.

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

#variable que contiene la ruta base del directorio donde se encuentra el archivo helpersestV3.py
BASE_DIR = Path(__file__).resolve().parent

#variable que contiene la URL predeterminada de OneDrive para descargar la programacion de cursos.
DEFAULT_ONEDRIVE_URL = (
    "https://upbeduco-my.sharepoint.com/:x:/g/personal/"
    "teamupbvirtual_upb_edu_co/IQB56oFFkvF_SpSH6eBZe2goAcJVYitbGxH2ihdrwUqE3XA"
    "?e=TqWw9H"
)

#variable que contiene los encabezados requeridos para procesar la programacion de cursos desde el archivo XLSX.
PROGRAMACION_HEADERS = {
    "ESTADO",
    "NRC",
    "MATERIA_CURSO",
    "NRC_CICLO_INTEGRACION",
    "LISTA_CRUZADA",
    "PERIODO",
    "INICIO",
}

#variable que contiene los encabezados requeridos para procesar los archivos de Banner desde el directorio configurado.
BANNER_HEADERS = {
    "PERIODO",
    "NRC",
    "LISTA_CRUZADA",
    "ID_ESTUDIANTE",
    "TIPO_DOCUMENTO",
    "DOCUMENTO",
    "CORREO_ESTUDIANTE",
    "NOMBRE_ESTUDIANTE",
    "APELLIDO_ESTUDIANTE",
    "COD_INSCRIPCION",
    "ESTADO_INSCRIPCION",
    "FECHA_ACTIVIDAD_EST",
    "PAGO",
    "SOCIO_INTEGRADOR",
}

#clase de datos inmutable y con slots para representar un registro de shortname, 
#que contiene el nombre del curso, el NRC y el periodo.
@dataclass(frozen=True, slots=True)
class ShortnameRecord:
    nombre: str
    nrc: str
    periodo: str

#clase de datos inmutable y con slots para representar un curso objetivo,
@dataclass(frozen=True, slots=True)
class CourseTarget:
    nombre: str
    nrc_fuente: str
    periodo_fuente: str

#funcion que carga la configuracion desde un archivo JSON y devuelve un diccionario con los valores correspondientes.
def cargar_config(path: str | os.PathLike[str] = "config.json") -> dict[str, Any]:
    config_path = Path(path)
    if not config_path.is_absolute():
        config_path = BASE_DIR / config_path
    try:
        with config_path.open(encoding="utf-8") as handle:
            config = json.load(handle)
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"No existe el archivo de configuracion: {config_path}") from exc
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"El archivo de configuracion no es JSON valido: {config_path}: {exc}"
        ) from exc
    if not isinstance(config, dict):
        raise ValueError("config.json debe contener un objeto JSON en su nivel principal")
    return config

#funcion que resuelve una ruta relativa o absoluta y devuelve un objeto Path correspondiente,
def resolver_ruta(value: str | os.PathLike[str], base: Path = BASE_DIR) -> Path:
    path = Path(value).expanduser()
    return path if path.is_absolute() else base / path

#funcion que configura un logger con rotacion de archivos y devuelve un objeto Logger correspondiente.
def configurar_logging(config: Mapping[str, Any]) -> logging.Logger:
    log_path = resolver_ruta(config.get("log_file", "logs/inscribirEstV3.log"))
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("inscribirEstV3")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=int(config.get("log_max_bytes", 5_000_000)),
        backupCount=int(config.get("log_backup_count", 5)),
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logging.captureWarnings(True)
    warnings_logger = logging.getLogger("py.warnings")
    warnings_logger.handlers.clear()
    warnings_logger.addHandler(file_handler)
    warnings_logger.propagate = False
    return logger

#funcion que analiza un valor y devuelve una cadena de texto canónica para usar como encabezado de columna, eliminando acentos, caracteres especiales y espacios.
def _canonical_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^A-Z0-9]+", "_", text.upper()).strip("_")

#funcion que verifica si un valor es nulo, vacio o contiene valores especiales como "nan", "none" o "nat".   
def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    return str(value).strip().casefold() in {"", "nan", "none", "nat"}

#funcion que normaliza un identificador de curso, eliminando espacios, caracteres especiales y truncando a 6 caracteres si es necesario.
def normalizar_identificador(value: Any) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).replace("\u00a0", " ").strip()
    return re.sub(r"\.0$", "", text)

#funcion que normaliza el periodo de un curso, eliminando espacios y truncando a 6 caracteres si es necesario.
def normalizar_periodo(value: Any, *, banner: bool = False) -> str:
    period = normalizar_identificador(value).replace(" ", "")
    return period[:6] if banner and len(period) >= 6 else period

#funcion que normaliza la materia de un curso, reemplazando espacios y caracteres especiales por guiones y eliminando guiones consecutivos.
def normalizar_materia(value: Any) -> str:
    if _is_blank(value):
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    text = re.sub(r"\s+", "-", text)
    return re.sub(r"-+", "-", text).strip("-")

#funcio que parsea una fecha desde un valor de entrada, que puede ser un objeto datetime, date, 
#numero de Excel o cadena de texto en varios formatos, y devuelve un objeto date correspondiente.
def parse_fecha(value: Any, field_name: str = "fecha") -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted
        except Exception as exc:
            raise ValueError(f"{field_name} contiene una fecha Excel invalida: {value}") from exc

    text = str(value or "").strip()
    formats = ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d")
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(
        f"{field_name} debe usar YYYY-MM-DD o DD/MM/YYYY; valor recibido: {value!r}"
    )

#FUNCION QUE ANALIZA EL ENCABEZADO DE UNA HOJA DE EXCEL Y DEVUELVE UN DICCIONARIO CON LOS INDICES DE COLUMNA CORRESPONDIENTES A LOS ENCABEZADOS REQUERIDOS.
def _source_bytes(source: Path | bytes | bytearray | io.BytesIO) -> bytes:
    if isinstance(source, Path):
        return source.read_bytes()
    if isinstance(source, io.BytesIO):
        return source.getvalue()
    return bytes(source)

#FUNCION QUE SANEA LOS ESTILOS DE UN ARCHIVO XLSX, CORRIGIENDO VALORES DE FONT.FAMILY FUERA DEL RANGO OOXML ACEPTADO (0..14) Y DEVUELVE EL CONTENIDO SANEADO Y EL NUMERO DE REEMPLAZOS REALIZADOS.
def _sanear_estilos_xlsx(content: bytes) -> tuple[bytes, int]:
    """Corrige valores ``font.family`` fuera del rango OOXML aceptado (0..14)."""

    source = io.BytesIO(content)
    destination = io.BytesIO()
    replacements = 0
    family_pattern = re.compile(rb'(<family\s+val=")([^\"]+)("\s*/>)')

    def replace_family(match: re.Match[bytes]) -> bytes:
        nonlocal replacements
        try:
            value = float(match.group(2).decode("ascii"))
        except (ValueError, UnicodeDecodeError):
            return match.group(0)
        if 0 <= value <= 14:
            return match.group(0)
        replacements += 1
        return match.group(1) + b"2" + match.group(3)

    with ZipFile(source) as input_zip, ZipFile(destination, "w", ZIP_DEFLATED) as output_zip:
        for item in input_zip.infolist():
            data = input_zip.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = family_pattern.sub(replace_family, data)
            output_zip.writestr(item, data)
    return destination.getvalue(), replacements

#FUNCION QUE ABRE UN ARCHIVO XLSX DE MANERA TOLERANTE, CORRIGIENDO ESTILOS INVALIDOS SI ES NECESARIO, Y DEVUELVE UN OBJETO Workbook DE openpyxl.
def abrir_workbook_tolerante(
    source: Path | bytes | bytearray | io.BytesIO,
    logger: logging.Logger,
    *,
    read_only: bool = True,
):
    first_source: Any = source
    if isinstance(source, (bytes, bytearray)):
        first_source = io.BytesIO(bytes(source))
    try:
        return load_workbook(first_source, read_only=read_only, data_only=True)
    except ValueError as exc:
        message = str(exc).casefold()
        if "stylesheet" not in message and "style" not in message:
            raise
        sanitized, replacements = _sanear_estilos_xlsx(_source_bytes(source))
        if not replacements:
            raise
        logger.warning(
            "El Excel contiene %s valor(es) de estilo invalido(s); se usa una copia "
            "temporal saneada sin modificar el archivo origen.",
            replacements,
        )
        return load_workbook(io.BytesIO(sanitized), read_only=read_only, data_only=True)

#FUNCIO QUE DESCARGA UN ARCHIVO DESDE UNA URL DE ONEDRIVE O SHAREPOINT, AGREGANDO EL PARAMETRO ``download=1`` A LA URL, 
# Y DEVUELVE EL CONTENIDO DEL ARCHIVO COMO BYTES.
def _url_descarga(url: str) -> str:
    parts = urlsplit(url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment))

#FUNCIO QUE DESCARGA LA PROGRAMACION DE CURSOS DESDE UNA URL DE ONEDRIVE O SHAREPOINT, VERIFICA EL CODIGO DE ESTADO HTTP 
# Y EL TIPO DE CONTENIDO, Y DEVUELVE EL CONTENIDO DEL ARCHIVO COMO BYTES.
def _descargar_programacion(url: str, timeout: int, logger: logging.Logger) -> bytes:
    try:
        import requests
    except ImportError as exc:
        raise RuntimeError(
            "Se requiere requests para descargar la programacion desde OneDrive"
        ) from exc

    logger.info("Descargando programacion de cursos desde OneDrive/SharePoint")
    response = requests.get(_url_descarga(url), timeout=timeout, allow_redirects=True)
    if response.status_code in {401, 403}:
        raise PermissionError(
            "OneDrive/SharePoint rechazo la descarga. Configure "
            "'programacion_cursos_file' con la ruta local sincronizada del archivo."
        )
    response.raise_for_status()
    content = response.content
    content_type = response.headers.get("content-type", "").casefold()
    if not content.startswith(b"PK"):
        raise ValueError(
            "La URL de programacion no devolvio un archivo XLSX "
            f"(Content-Type: {content_type or 'desconocido'})"
        )
    return content

#FUNCION QUE OBTIENE EL ORIGEN DE LA PROGRAMACION DE CURSOS, PREFERIENDO LA DESCARGA REMOTA SI ESTA CONFIGURADA Y DISPONIBLE,
def obtener_origen_programacion(
    config: Mapping[str, Any], logger: logging.Logger
) -> Path | bytes:
    local_value = config.get("programacion_cursos_file")
    local_path = resolver_ruta(str(local_value)) if local_value else None
    url = str(config.get("programacion_cursos_url", DEFAULT_ONEDRIVE_URL)).strip()
    prefer_remote = bool(config.get("programacion_cursos_preferir_remoto", True))

    if prefer_remote and url:
        try:
            return _descargar_programacion(
                url, int(config.get("programacion_cursos_timeout", 60)), logger
            )
        except Exception as exc:
            if local_path and local_path.is_file():
                logger.warning(
                    "Fallo la descarga de OneDrive (%s). Se usara la copia local sincronizada: %s",
                    exc,
                    local_path,
                )
                return local_path
            raise

    if local_path and local_path.is_file():
        logger.info("Usando copia local sincronizada de OneDrive: %s", local_path)
        return local_path
    if local_path:
        logger.warning(
            "No existe la copia local configurada de la programacion: %s; se intentara la URL",
            local_path,
        )
    if not url:
        raise ValueError(
            "Debe configurar 'programacion_cursos_file' o 'programacion_cursos_url'"
        )
    return _descargar_programacion(
        url, int(config.get("programacion_cursos_timeout", 60)), logger
    )

#FUNCION QUE LEE LA PROGRAMACION DE CURSOS DESDE UN ARCHIVO XLSX, VERIFICA QUE CONTENGA LOS ENCABEZADOS REQUERIDOS Y DEVUELVE UNA LISTA DE FILAS COMO DICCIONARIOS.
def leer_programacion(
    config: Mapping[str, Any], logger: logging.Logger
) -> list[dict[str, Any]]:
    sheet_name = str(config.get("programacion_cursos_hoja", "")).strip()
    if not sheet_name:
        raise ValueError("Falta 'programacion_cursos_hoja' en config.json")

    workbook = abrir_workbook_tolerante(obtener_origen_programacion(config, logger), logger)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"No existe la hoja {sheet_name!r}. Hojas disponibles: {workbook.sheetnames}"
            )
        worksheet = workbook[sheet_name]
        iterator = worksheet.iter_rows(values_only=True)
        try:
            header_row = next(iterator)
        except StopIteration as exc:
            raise ValueError(f"La hoja {sheet_name!r} esta vacia") from exc

        indexes = {
            _canonical_header(value): index for index, value in enumerate(header_row)
        }
        missing = PROGRAMACION_HEADERS - indexes.keys()
        if missing:
            raise ValueError(
                f"La hoja {sheet_name!r} no contiene las columnas requeridas: {sorted(missing)}"
            )

        rows: list[dict[str, Any]] = []
        for excel_row, values in enumerate(iterator, start=2):
            row = {name: values[index] if index < len(values) else None for name, index in indexes.items()}
            row["_EXCEL_ROW"] = excel_row
            rows.append(row)
        logger.info("Programacion cargada: %s filas en la hoja %s", len(rows), sheet_name)
        return rows
    finally:
        workbook.close()

#FUNCION QUE INTEGRA TOKENS DE INTEGRACION DE CURSOS, BUSCANDO FILAS AUXILIARES EN LA PROGRAMACION Y ASOCIANDO LOS NRC CORRESPONDIENTES.
def _integration_tokens(value: Any) -> list[str]:
    if _is_blank(value):
        return []
    text = normalizar_identificador(value)
    tokens = re.findall(r"\d{4,10}", text)
    return list(dict.fromkeys(tokens or [text]))

#FUNCION QUE VALIDA EL ESTADO DE UN CURSO, INDICANDO SI CONTIENE EXACTAMENTE EL SIMBOLO OPERATIVO ``✔️``.
def estado_curso_activo(value: Any) -> bool:
    """Indica si Estado contiene exactamente el simbolo operativo ``✔️``."""

    if _is_blank(value):
        return False
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\ufe0f", "").replace("\u200d", "").strip()
    return text == "✔"

#FUNCION QUE CONSTRUYE LOS REGISTROS DE SHORTNAMES A PARTIR DE LAS FILAS DE LA PROGRAMACION, FILTRANDO POR FECHAS Y ESTADO, 
# Y ASOCIANDO LOS CURSOS AUXILIARES DE INTEGRACION.
def construir_registros_shortnames(
    rows: Sequence[Mapping[str, Any]],
    fecha_inicial: date,
    fecha_final: date,
    logger: logging.Logger,
) -> list[ShortnameRecord]:
    if fecha_inicial > fecha_final:
        raise ValueError("La fecha inicial de programacion no puede ser posterior a la fecha final")

    companion_rows: set[int] = set()
    integrations: dict[int, list[tuple[str, str]]] = {}

    selected_rows: set[int] = set()
    inactive_rows = out_of_range_rows = invalid_date_rows = 0
    for index, row in enumerate(rows):
        if not estado_curso_activo(row.get("ESTADO")):
            inactive_rows += 1
            continue
        raw_start = row.get("INICIO")
        if _is_blank(raw_start):
            invalid_date_rows += 1
            continue
        try:
            row_date = parse_fecha(raw_start, f"Inicio fila {row.get('_EXCEL_ROW', index + 2)}")
        except ValueError as exc:
            logger.warning("%s", exc)
            invalid_date_rows += 1
            continue
        if fecha_inicial <= row_date <= fecha_final:
            selected_rows.add(index)
        else:
            out_of_range_rows += 1

    logger.info(
        "Filtro de programacion: rango inclusivo %s a %s, Estado=✔️; "
        "%s fila(s) seleccionada(s), %s fuera de rango, %s sin Estado activo y %s sin fecha valida.",
        fecha_inicial.isoformat(),
        fecha_final.isoformat(),
        len(selected_rows),
        out_of_range_rows,
        inactive_rows,
        invalid_date_rows,
    )

    # Identifica las filas auxiliares del ciclo de integracion antes de aplicar
    # el filtro de fecha, para que no se conviertan tambien en cursos separados.
    for index, row in enumerate(rows):
        if index not in selected_rows:
            continue
        tokens = _integration_tokens(row.get("NRC_CICLO_INTEGRACION"))
        if not tokens:
            continue
        if len(tokens) > 1:
            logger.warning(
                "Fila %s: NRC Ciclo Integracion contiene varios NRC (%s); se procesaran todos.",
                row.get("_EXCEL_ROW", index + 2),
                ", ".join(tokens),
            )
        next_search = index + 1
        for token in tokens:
            match_index: int | None = None
            # La regla indica la fila siguiente. El margen adicional soporta
            # celdas con mas de un NRC de integracion sin perder trazabilidad.
            for candidate_index in range(next_search, min(len(rows), index + 6)):
                candidate = rows[candidate_index]
                if normalizar_identificador(candidate.get("NRC")) == token:
                    match_index = candidate_index
                    break
            if match_index is None:
                logger.warning(
                    "Fila %s: no se encontro una fila siguiente para el NRC de integracion %s; "
                    "no se generara ese registro.",
                    row.get("_EXCEL_ROW", index + 2),
                    token,
                )
                continue
            period = normalizar_periodo(rows[match_index].get("PERIODO"))
            if not period:
                logger.warning(
                    "Fila %s: el NRC de integracion %s no tiene periodo en su fila auxiliar.",
                    row.get("_EXCEL_ROW", index + 2),
                    token,
                )
                continue
            companion_rows.add(match_index)
            integrations.setdefault(index, []).append((token, period))
            next_search = match_index + 1

    records: list[ShortnameRecord] = []
    record_keys: set[tuple[str, str, str]] = set()
    cross_list_names: dict[tuple[str, str], str] = {}

    def add_record(record: ShortnameRecord) -> None:
        key = (record.nombre, record.nrc, record.periodo)
        if key not in record_keys:
            records.append(record)
            record_keys.add(key)

    for index, row in enumerate(rows):
        if index not in selected_rows or index in companion_rows:
            continue

        materia = normalizar_materia(row.get("MATERIA_CURSO"))
        period = normalizar_periodo(row.get("PERIODO"))
        nrc = normalizar_identificador(row.get("NRC"))
        cross_list = normalizar_identificador(row.get("LISTA_CRUZADA"))
        primary_id = cross_list or nrc
        excel_row = row.get("_EXCEL_ROW", index + 2)
        if not materia or not period or not primary_id:
            logger.warning(
                "Fila %s omitida: Materia Curso, Periodo y NRC/Lista Cruzada son obligatorios.",
                excel_row,
            )
            continue
        if not re.fullmatch(r"[A-Za-z0-9]+", primary_id):
            logger.warning(
                "Fila %s omitida: el NRC/Lista Cruzada %r no tiene formato de codigo valido.",
                excel_row,
                primary_id,
            )
            continue

        proposed_name = f"{materia}-{period}-{primary_id}"
        if cross_list:
            cross_key = (period, cross_list)
            course_name = cross_list_names.setdefault(cross_key, proposed_name)
            if course_name != proposed_name:
                logger.warning(
                    "Fila %s: la Lista Cruzada %s ya usa el curso %s; se omite el nombre alterno %s.",
                    excel_row,
                    cross_list,
                    course_name,
                    proposed_name,
                )
        else:
            course_name = proposed_name

        add_record(ShortnameRecord(course_name, primary_id, period))
        for integration_nrc, integration_period in integrations.get(index, []):
            add_record(ShortnameRecord(course_name, integration_nrc, integration_period))

    if not records:
        logger.warning(
            "No se encontraron cursos activos entre %s y %s; "
            "shortnames.csv quedara solo con encabezado.",
            fecha_inicial.isoformat(),
            fecha_final.isoformat(),
        )
    return records

#FUNCION QUE ESCRIBE LOS REGISTROS DE SHORTNAMES EN UN ARCHIVO CSV, CREANDO DIRECTORIOS SI ES NECESARIO 
# Y USANDO UN ARCHIVO TEMPORAL PARA EVITAR CORRUPCION.
def escribir_shortnames(
    records: Sequence[ShortnameRecord], output_path: Path
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="",
            dir=output_path.parent,
            prefix=f".{output_path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_name = handle.name
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(("Nombre", "NRC", "Periodo"))
            writer.writerows((record.nombre, record.nrc, record.periodo) for record in records)
        os.replace(temporary_name, output_path)
    finally:
        if temporary_name and os.path.exists(temporary_name):
            os.unlink(temporary_name)

#FUNCION PRINCIPAL QUE GENERA EL ARCHIVO shortnames.csv A PARTIR DE LA CONFIGURACION Y EL LOGGER,
# OBTENIENDO LA PROGRAMACION DE CURSOS, FILTRANDO POR FECHAS Y ESTADO, Y ESCRIBIENDO LOS REGISTROS EN EL ARCHIVO DE
def generar_shortnames(
    config: Mapping[str, Any], logger: logging.Logger
) -> list[ShortnameRecord]:
    configured_start = config.get("programacion_cursos_fecha_inicial")
    configured_end = config.get("programacion_cursos_fecha_final")
    legacy_date = config.get("programacion_cursos_fecha_inicio")
    if configured_start is None and configured_end is None and legacy_date is not None:
        configured_start = configured_end = legacy_date
        logger.warning(
            "La clave 'programacion_cursos_fecha_inicio' esta obsoleta; use "
            "'programacion_cursos_fecha_inicial' y 'programacion_cursos_fecha_final'."
        )
    if configured_start is None:
        raise ValueError("Falta 'programacion_cursos_fecha_inicial' en config.json")
    if configured_end is None:
        raise ValueError("Falta 'programacion_cursos_fecha_final' en config.json")

    start_date = parse_fecha(configured_start, "programacion_cursos_fecha_inicial")
    end_date = parse_fecha(configured_end, "programacion_cursos_fecha_final")
    if start_date > end_date:
        raise ValueError(
            "programacion_cursos_fecha_inicial no puede ser posterior a "
            "programacion_cursos_fecha_final"
        )
    records = construir_registros_shortnames(
        leer_programacion(config, logger), start_date, end_date, logger
    )
    output_path = resolver_ruta(config.get("shortnames_file", "shortnames.csv"))
    escribir_shortnames(records, output_path)
    logger.info(
        "shortnames.csv generado: %s registro(s) para el rango %s a %s en %s",
        len(records),
        start_date.isoformat(),
        end_date.isoformat(),
        output_path,
    )
    return records


def _find_header(
    worksheet: Any,
    required: set[str],
    max_rows: int = 20,
) -> tuple[dict[str, int], Iterator[tuple[Any, ...]]]:
    iterator = worksheet.iter_rows(values_only=True)
    for _ in range(max_rows):
        try:
            values = next(iterator)
        except StopIteration as exc:
            raise ValueError(f"La hoja {worksheet.title!r} esta vacia") from exc
        indexes = {_canonical_header(value): index for index, value in enumerate(values) if value is not None}
        if required.issubset(indexes):
            return indexes, iterator
    raise ValueError(
        f"No se encontro un encabezado valido en la hoja {worksheet.title!r}; "
        f"columnas requeridas: {sorted(required)}"
    )

#FUNCION QUE CARGA LA BASE DE USUARIOS DE BRIGHTSPACE DESDE UN ARCHIVO XLSX, VERIFICA QUE CONTENGA LA COLUMNA USERNAME Y DEVUELVE UN CONJUNTO DE USERNAMES NORMALIZADOS.
def cargar_usuarios_bs(config: Mapping[str, Any], logger: logging.Logger) -> set[str]:
    value = config.get("bdusuarios_file")
    if not value:
        raise ValueError("Falta 'bdusuarios_file' en config.json")
    path = resolver_ruta(str(value))
    if not path.is_file():
        raise FileNotFoundError(f"No existe la base de usuarios de Brightspace: {path}")

    workbook = abrir_workbook_tolerante(path, logger)
    try:
        worksheet = workbook[workbook.sheetnames[0]]
        indexes, iterator = _find_header(worksheet, {"USERNAME"})
        index = indexes["USERNAME"]
        users = {
            normalizar_identificador(row[index]).zfill(9)
            for row in iterator
            if index < len(row) and not _is_blank(row[index])
        }
        logger.info("Base de Brightspace cargada: %s UserName unicos", len(users))
        return users
    finally:
        workbook.close()

#FUNCION QUE OBTIENE LA LISTA DE ARCHIVOS XLSX DE BANNER EN EL DIRECTORIO CONFIGURADO, VERIFICA QUE EXISTAN 
# Y DEVUELVE UNA LISTA DE OBJETOS Path ORDENADOS.
def _banner_files(config: Mapping[str, Any]) -> list[Path]:
    value = config.get("banner_directory")
    if not value:
        raise ValueError("Falta 'banner_directory' en config.json")
    directory = resolver_ruta(str(value))
    if not directory.is_dir():
        raise FileNotFoundError(f"No existe el directorio Banner: {directory}")
    files = sorted(
        path for path in directory.glob("*.xlsx") if not path.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"No hay archivos .xlsx en el directorio Banner: {directory}")
    return files

#funcion que obtiene el valor de una celda de una fila de Banner, usando un diccionario de indices de columna y el nombre del encabezado.
def _cell(row: Sequence[Any], indexes: Mapping[str, int], name: str) -> Any:
    index = indexes[name]
    return row[index] if index < len(row) else None

#funcion que itera sobre las filas de los archivos de Banner, leyendo la hoja configurada 
# y devolviendo un diccionario con los valores de cada fila para los encabezados requeridos.
def _iter_banner_rows(
    config: Mapping[str, Any], logger: logging.Logger
) -> Iterator[dict[str, Any]]:
    sheet_name = str(config.get("banner_sheet", "Estudiantes"))
    valid_files = 0
    for path in _banner_files(config):
        logger.info("Leyendo Banner en modo streaming: %s", path)
        try:
            workbook = abrir_workbook_tolerante(path, logger)
        except Exception:
            logger.exception("No fue posible abrir el archivo Banner %s", path)
            continue
        try:
            if sheet_name not in workbook.sheetnames:
                logger.warning(
                    "El archivo %s no contiene la hoja %s; se omite.", path.name, sheet_name
                )
                continue
            worksheet = workbook[sheet_name]
            try:
                indexes, iterator = _find_header(worksheet, BANNER_HEADERS)
            except ValueError as exc:
                logger.warning("El archivo %s se omite: %s", path.name, exc)
                continue
            valid_files += 1
            for row in iterator:
                yield {name: _cell(row, indexes, name) for name in BANNER_HEADERS}
        finally:
            workbook.close()
    if not valid_files:
        raise ValueError("Ningun archivo Banner valido pudo ser procesado")

#funcion que determina el rol y la fuente de un estudiante en base al periodo y al socio integrador, 
# devolviendo una tupla con el rol y la fuente, o None si no se puede determinar.
def _role_for(period: str, partner: str) -> tuple[str, str] | None:
    period6 = normalizar_periodo(period, banner=True)
    formation = period6[-2:] if len(period6) >= 2 else ""
    if formation in {"41", "42"}:
        return ("Student_ap", "CVLA") if partner == "AP" else ("Student_fa", "CVFA")
    if formation in {"10", "11", "20", "21"}:
        return "Student_pr", "CVPR"
    if formation == "50":
        return "Student_ex", "CVFC"
    if formation in {"17", "27", "37"}:
        return "Student_te", "CVTE"
    return None

#funcion que construye una descripcion del documento de un estudiante a partir de los campos DOCUMENTO y TIPO_DOCUMENTO, 
# devolviendo una cadena de texto con el tipo y el numero de documento formateado.
def _document_description(row: Mapping[str, Any]) -> str:
    document = normalizar_identificador(row.get("DOCUMENTO"))
    if document.isdigit():
        document = f"{int(document):,}".replace(",", ".")
    document_type = str(row.get("TIPO_DOCUMENTO") or "").strip()
    return f"{document_type}. {document}" if document_type else document

#FUNCION QUE GENERA UN NOMBRE DE ARCHIVO SEGURO A PARTIR DEL NOMBRE DEL CURSO, REEMPLAZANDO CARACTERES INVALIDOS POR GUIONES BAJOS 
# Y ELIMINANDO PUNTOS O GUIONES AL PRINCIPIO O FINAL.
def _safe_filename(course_name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", course_name).strip("._")
    return safe or "curso_sin_nombre"

#FUNCION QUE NORMALIZA UN VALOR DE TEXTO PARA COMPARACIONES, ELIMINANDO ACENTOS, CARACTERES COMBINADOS Y CONVIRTIENDO A MINUSCULAS.
def _casefold(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").strip())
    return "".join(char for char in text if not unicodedata.combining(char)).casefold()

#funcion que procesa las inscripciones de estudiantes en cursos, leyendo los archivos de Banner, filtrando por fecha y estado,
# y generando archivos de registro para cada curso y un archivo consolidado, devolviendo un
def procesar_inscripciones(
    config: Mapping[str, Any],
    shortnames: Sequence[ShortnameRecord],
    logger: logging.Logger,
    fecha_actividad_desde: date | None = None,
) -> dict[str, int]:
    if not shortnames:
        logger.warning("No hay cursos para procesar; no se leera la base Banner.")
        return {}

    process_type = _casefold(config.get("Tipo_proceso", "Matricular"))
    allowed_processes = {"matricular", "desmatricular", "limpiar"}
    if process_type not in allowed_processes:
        raise ValueError(
            "Tipo_proceso debe ser Matricular, Desmatricular o Limpiar; "
            f"valor recibido: {config.get('Tipo_proceso')!r}"
        )

    targets_by_key: dict[tuple[str, str], list[CourseTarget]] = {}
    sources_by_course: dict[str, set[str]] = {}
    for record in shortnames:
        period6 = normalizar_periodo(record.periodo, banner=True)
        nrc = normalizar_identificador(record.nrc)
        target = CourseTarget(record.nombre, nrc, record.periodo)
        targets_by_key.setdefault((period6, nrc), []).append(target)
        sources_by_course.setdefault(record.nombre, set()).add(nrc)

    existing_users = cargar_usuarios_bs(config, logger) if process_type == "matricular" else set()
    output_directory = resolver_ruta(config.get("salida_directory", "salida"))
    output_directory.mkdir(parents=True, exist_ok=True)
    consolidated_path = resolver_ruta(
        config.get("registro_unico_est_file", "registro_unicoEst.txt")
    )
    students_path = resolver_ruta(config.get("students_file", "students.csv"))
    consolidated_path.parent.mkdir(parents=True, exist_ok=True)
    students_path.parent.mkdir(parents=True, exist_ok=True)

    counts = {name: 0 for name in sources_by_course}
    total_rows = relevant_rows = 0
    skipped = {
        "fecha": 0,
        "duplicado_fuente": 0,
        "duplicado_curso": 0,
        "sin_correo": 0,
        "estado": 0,
        "pago": 0,
        "periodo": 0,
    }
    missing_email_samples: list[str] = []
    invalid_periods: set[str] = set()
    source_seen: set[tuple[str, str, str]] = set()
    matched_source_keys: set[tuple[str, str]] = set()
    course_enrollments: set[tuple[str, str]] = set()
    prepared_users: set[str] = set()
    home_enrollments: set[tuple[str, str, str]] = set()

    with ExitStack() as stack:
        consolidated_handle = stack.enter_context(
            consolidated_path.open("w", encoding="utf-8", newline="")
        )
        consolidated_writer = csv.writer(consolidated_handle, lineterminator="\n")
        course_writers: dict[str, csv.writer] = {}
        output_names: dict[str, str] = {}
        for course_name in sources_by_course:
            safe_name = _safe_filename(course_name)
            previous = output_names.setdefault(safe_name.casefold(), course_name)
            if previous != course_name:
                raise ValueError(
                    f"Los cursos {previous!r} y {course_name!r} producen el mismo nombre de archivo"
                )
            handle = stack.enter_context(
                (output_directory / f"registro_{safe_name}.txt").open(
                    "w", encoding="utf-8", newline=""
                )
            )
            course_writers[course_name] = csv.writer(handle, lineterminator="\n")

        def emit(course_name: str, command: Sequence[Any]) -> None:
            course_writers[course_name].writerow(command)
            consolidated_writer.writerow(command)

        for row in _iter_banner_rows(config, logger):
            total_rows += 1
            if fecha_actividad_desde is not None:
                raw_activity = row.get("FECHA_ACTIVIDAD_EST")
                try:
                    activity_date = parse_fecha(raw_activity, "FECHA_ACTIVIDAD_EST")
                except ValueError:
                    skipped["fecha"] += 1
                    continue
                if activity_date < fecha_actividad_desde:
                    skipped["fecha"] += 1
                    continue

            period = normalizar_periodo(row.get("PERIODO"), banner=True)
            cross_list = normalizar_identificador(row.get("LISTA_CRUZADA"))
            nrc = normalizar_identificador(row.get("NRC"))
            source_nrc = cross_list or nrc
            targets = targets_by_key.get((period, source_nrc))
            if not targets:
                continue
            relevant_rows += 1
            matched_source_keys.add((period, source_nrc))

            student_id = normalizar_identificador(row.get("ID_ESTUDIANTE")).zfill(9)
            if not student_id.strip("0"):
                logger.warning(
                    "Se encontro una fila Banner relevante sin ID_ESTUDIANTE; se omite."
                )
                continue
            source_key = (period, source_nrc, student_id)
            if source_key in source_seen:
                skipped["duplicado_fuente"] += 1
                continue
            source_seen.add(source_key)

            status = _casefold(row.get("ESTADO_INSCRIPCION"))
            enrollment_code = _casefold(row.get("COD_INSCRIPCION"))
            partner = normalizar_identificador(row.get("SOCIO_INTEGRADOR")).upper() or "BS"
            payment = normalizar_identificador(row.get("PAGO")).upper()

            for target in targets:
                enrollment_key = (target.nombre, student_id)
                if enrollment_key in course_enrollments:
                    skipped["duplicado_curso"] += 1
                    continue

                if process_type == "matricular":
                    if status != "inscrito":
                        skipped["estado"] += 1
                        continue
                    if partner == "AP" and payment != "Y":
                        skipped["pago"] += 1
                        continue
                    if partner not in {"BS", "AP"}:
                        skipped["pago"] += 1
                        continue
                    email = str(row.get("CORREO_ESTUDIANTE") or "").strip()
                    if not email:
                        skipped["sin_correo"] += 1
                        if len(missing_email_samples) < 10:
                            missing_email_samples.append(student_id)
                        continue
                    role_info = _role_for(target.periodo_fuente, partner)
                    if role_info is None:
                        skipped["periodo"] += 1
                        invalid_periods.add(target.periodo_fuente)
                        continue
                    role, org_unit = role_info
                    first_name = str(row.get("NOMBRE_ESTUDIANTE") or "").strip().title()
                    last_name = str(row.get("APELLIDO_ESTUDIANTE") or "").strip().title()

                    if student_id not in prepared_users:
                        if student_id in existing_users:
                            emit(
                                target.nombre,
                                (
                                    "UPDATE",
                                    student_id,
                                    _document_description(row),
                                    first_name,
                                    last_name,
                                    "",
                                    "1",
                                    email,
                                ),
                            )
                        else:
                            emit(
                                target.nombre,
                                (
                                    "CREATE",
                                    student_id,
                                    _document_description(row),
                                    first_name,
                                    last_name,
                                    "",
                                    role,
                                    "1",
                                    email,
                                ),
                            )
                            existing_users.add(student_id)
                        prepared_users.add(student_id)

                    home_key = (student_id, role, org_unit)
                    if home_key not in home_enrollments:
                        emit(target.nombre, ("ENROLL", student_id, "", role, org_unit))
                        home_enrollments.add(home_key)
                    emit(target.nombre, ("ENROLL", student_id, "", "Student", target.nombre))
                elif process_type == "desmatricular":
                    if status != "cancelado":
                        skipped["estado"] += 1
                        continue
                    emit(target.nombre, ("UNENROLL", student_id, "", target.nombre))
                else:
                    if enrollment_code != "dl":
                        skipped["estado"] += 1
                        continue
                    emit(target.nombre, ("UNENROLL", student_id, "", target.nombre))

                course_enrollments.add(enrollment_key)
                counts[target.nombre] += 1

    with students_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\n")
        writer.writerow(("Nombre", "NRC", "Estudiantes"))
        for course_name in sorted(counts):
            writer.writerow(
                (course_name, "|".join(sorted(sources_by_course[course_name])), counts[course_name])
            )

    if skipped["sin_correo"]:
        logger.warning(
            "%s inscripcion(es) se omitieron por correo vacio. Muestra de ID: %s",
            skipped["sin_correo"],
            ", ".join(missing_email_samples),
        )
    if invalid_periods:
        logger.warning(
            "Se omitieron inscripciones con periodos sin mapeo de rol: %s",
            ", ".join(sorted(invalid_periods)),
        )
    unmatched_sources = sorted(set(targets_by_key) - matched_source_keys)
    if unmatched_sources:
        sample = ", ".join(f"{period}/{nrc}" for period, nrc in unmatched_sources[:20])
        suffix = " ..." if len(unmatched_sources) > 20 else ""
        logger.warning(
            "%s fuente(s) Periodo/NRC de shortnames no tuvieron filas Banner en esta ejecucion: %s%s",
            len(unmatched_sources),
            sample,
            suffix,
        )
    empty_courses = sorted(name for name, count in counts.items() if count == 0)
    if empty_courses:
        sample = ", ".join(empty_courses[:20])
        suffix = " ..." if len(empty_courses) > 20 else ""
        logger.warning(
            "%s curso(s) quedaron sin estudiantes para el Tipo_proceso actual: %s%s",
            len(empty_courses),
            sample,
            suffix,
        )
    logger.info(
        "Banner procesado: %s filas leidas, %s relevantes, %s estudiantes inscritos/desinscritos.",
        total_rows,
        relevant_rows,
        sum(counts.values()),
    )
    logger.info("Omisiones y deduplicaciones: %s", skipped)
    logger.info("Consolidado generado: %s", consolidated_path)
    logger.info("Resumen generado: %s", students_path)
    return counts
