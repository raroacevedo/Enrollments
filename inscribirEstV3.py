#!/usr/bin/env python
"""Punto de entrada del proceso V3 de inscripcion de estudiantes."""

from __future__ import annotations

import argparse
from datetime import date
import logging
import sys

import helpersestV3 as helpers

#funcion que analiza los argumentos de la linea de comandos y devuelve un objeto Namespace con los valores correspondientes.
def _arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Genera shortnames.csv desde la programacion de OneDrive y crea los "
            "archivos de inscripcion de estudiantes."
        )
    )
    parser.add_argument(
        "fecha_banner",
        nargs="?",
        help="Fecha minima de actividad Banner (DD/MM/YY, DD/MM/YYYY o YYYY-MM-DD)",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Ruta del archivo JSON de configuracion (por defecto: config.json)",
    )
    parser.add_argument(
        "--solo-shortnames",
        action="store_true",
        help="Genera shortnames.csv y no procesa estudiantes Banner",
    )
    return parser.parse_args()

#fUncion que configura un logger de respaldo para capturar errores incluso si la configuración falla.
def _fallback_logger() -> logging.Logger:
    # Garantiza que incluso los errores al abrir/interpretar config.json queden
    # persistidos en el log predeterminado.
    return helpers.configurar_logging({})

# Funcion principal del script, que se encarga de ejecutar el proceso de inscripcion de estudiantes
def run(args: argparse.Namespace) -> int:
    logger = _fallback_logger() #CONFIGURA UN LOGGER DE RESPALDO PARA CAPTURAR ERRORES INCLUSO SI LA CONFIGURACION FALLA
    try:
        config = helpers.cargar_config(args.config) #CARGA LA CONFIGURACION DEL ARCHIVO JSON
        logger = helpers.configurar_logging(config) #CONFIGURA EL LOGGER CON LA CONFIGURACION CARGADA
        logger.info("Inicia proceso V3 de inscripcion de estudiantes")

        activity_date: date | None = None #FECHA DE ACTIVIDAD
        if args.fecha_banner:
            activity_date = helpers.parse_fecha(args.fecha_banner, "fecha_banner") #PARSEA LA FECHA BANNER
            if activity_date > date.today():
                raise ValueError("La fecha Banner no puede ser posterior a la fecha actual")

        # Primera etapa obligatoria de V3.
        shortnames = helpers.generar_shortnames(config, logger) #GENERA EL ARCHIVO shortnames.csv A PARTIR DE LA PROGRAMACION DE ONEDRIVE
        if not args.solo_shortnames:
            helpers.procesar_inscripciones(config, shortnames, logger, activity_date) #PROCESA LAS INSCRIPCIONES DE ESTUDIANTES

        logger.info("Finaliza correctamente el proceso V3")
        return 0
    except Exception:
        logger.exception("El proceso V3 finalizo con error")
        return 1

#Actúa como punto de entrada del script.
#Llama a la función «run» con los argumentos analizados y finaliza con el código de estado correspondiente.
def main() -> None:
    sys.exit(run(_arguments()))

#inicio del script
if __name__ == "__main__":
    main()
