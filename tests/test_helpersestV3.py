import csv
from datetime import date, datetime
import io
import json
import logging
from pathlib import Path
import tempfile
import unittest
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

import helpersestV3 as helpers


def logger_for_tests() -> logging.Logger:
    logger = logging.getLogger("helpersestV3.tests")
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.setLevel(logging.DEBUG)
    return logger


class ShortnamesRulesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logger_for_tests()

    def test_main_integration_and_cross_list_rules(self) -> None:
        start = datetime(2026, 7, 13)
        rows = [
            {
                "_EXCEL_ROW": 2,
                "ESTADO": "✔️",
                "NRC": 53080,
                "MATERIA_CURSO": "5557-2065",
                "NRC_CICLO_INTEGRACION": 18200,
                "LISTA_CRUZADA": None,
                "PERIODO": 202642,
                "INICIO": start,
            },
            {
                "_EXCEL_ROW": 3,
                "ESTADO": "✔️",
                "NRC": 18200,
                "MATERIA_CURSO": "5557-2065",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": None,
                "PERIODO": 202620,
                "INICIO": start,
            },
            {
                "_EXCEL_ROW": 4,
                "ESTADO": "✔️",
                "NRC": 52958,
                "MATERIA_CURSO": "5556-2006",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": None,
                "PERIODO": 202642,
                "INICIO": start,
            },
            {
                "_EXCEL_ROW": 5,
                "ESTADO": "✔️",
                "NRC": 52963,
                "MATERIA_CURSO": "5569-2031",
                "NRC_CICLO_INTEGRACION": 18225,
                "LISTA_CRUZADA": None,
                "PERIODO": 202642,
                "INICIO": start,
            },
            {
                "_EXCEL_ROW": 6,
                "ESTADO": "✔️",
                "NRC": 18225,
                "MATERIA_CURSO": "5569-2031",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": None,
                "PERIODO": 202620,
                "INICIO": start,
            },
        ]

        records = helpers.construir_registros_shortnames(
            rows, date(2026, 7, 13), date(2026, 7, 13), self.logger
        )

        self.assertEqual(
            [(item.nombre, item.nrc, item.periodo) for item in records],
            [
                ("5557-2065-202642-53080", "53080", "202642"),
                ("5557-2065-202642-53080", "18200", "202620"),
                ("5556-2006-202642-52958", "52958", "202642"),
                ("5569-2031-202642-52963", "52963", "202642"),
                ("5569-2031-202642-52963", "18225", "202620"),
            ],
        )

    def test_cross_list_has_precedence_and_is_unique(self) -> None:
        rows = [
            {
                "_EXCEL_ROW": 2,
                "ESTADO": "✔️",
                "NRC": 10001,
                "MATERIA_CURSO": "AAAA-0001",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": "LC10",
                "PERIODO": 202642,
                "INICIO": "2026-08-03",
            },
            {
                "_EXCEL_ROW": 3,
                "ESTADO": "✔️",
                "NRC": 10002,
                "MATERIA_CURSO": "BBBB-0002",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": "LC10",
                "PERIODO": 202642,
                "INICIO": "2026-08-03",
            },
        ]
        records = helpers.construir_registros_shortnames(
            rows, date(2026, 8, 3), date(2026, 8, 3), self.logger
        )
        self.assertEqual(
            records, [helpers.ShortnameRecord("AAAA-0001-202642-LC10", "LC10", "202642")]
        )

    def test_invalid_operational_text_is_not_a_course_code(self) -> None:
        rows = [
            {
                "_EXCEL_ROW": 2,
                "ESTADO": "✔️",
                "NRC": "Revisar Listado",
                "MATERIA_CURSO": "5555-0001",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": None,
                "PERIODO": 202642,
                "INICIO": "2026-08-03",
            }
        ]
        self.assertEqual(
            helpers.construir_registros_shortnames(
                rows, date(2026, 8, 3), date(2026, 8, 3), self.logger
            ),
            [],
        )

    def test_inclusive_date_range_and_active_status_filter(self) -> None:
        def row(excel_row: int, nrc: int, start: str, status: object) -> dict[str, object]:
            return {
                "_EXCEL_ROW": excel_row,
                "ESTADO": status,
                "NRC": nrc,
                "MATERIA_CURSO": f"TEST-{nrc}",
                "NRC_CICLO_INTEGRACION": None,
                "LISTA_CRUZADA": None,
                "PERIODO": 202642,
                "INICIO": start,
            }

        rows = [
            row(2, 10001, "2026-07-12", "✔️"),
            row(3, 10002, "2026-07-13", "✔️"),
            row(4, 10003, "2026-07-15", "✔"),
            row(5, 10004, "2026-07-17", "✔️ "),
            row(6, 10005, "2026-07-18", "✔️"),
            row(7, 10006, "2026-07-15", None),
            row(8, 10007, "2026-07-15", "Pendiente"),
            row(9, 10008, "2026-07-15", "✅"),
        ]

        records = helpers.construir_registros_shortnames(
            rows, date(2026, 7, 13), date(2026, 7, 17), self.logger
        )

        self.assertEqual([record.nrc for record in records], ["10002", "10003", "10004"])

    def test_invalid_date_range_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "fecha inicial"):
            helpers.construir_registros_shortnames(
                [], date(2026, 7, 18), date(2026, 7, 17), self.logger
            )

    def test_invalid_font_family_can_be_sanitized(self) -> None:
        source = io.BytesIO()
        with ZipFile(source, "w", ZIP_DEFLATED) as archive:
            archive.writestr(
                "xl/styles.xml",
                b'<styleSheet><fonts><font><family val="18"/></font></fonts></styleSheet>',
            )
        sanitized, replacements = helpers._sanear_estilos_xlsx(source.getvalue())
        self.assertEqual(replacements, 1)
        with ZipFile(io.BytesIO(sanitized)) as archive:
            self.assertIn(b'<family val="2"/>', archive.read("xl/styles.xml"))

    def test_public_sharepoint_url_is_converted_to_download(self) -> None:
        url = "https://example.sharepoint.com/:x:/g/file?e=publicToken"
        download_url = helpers._url_descarga(url)
        self.assertIn("e=publicToken", download_url)
        self.assertIn("download=1", download_url)


class StreamingEnrollmentTests(unittest.TestCase):
    BANNER_COLUMNS = [
        "PERIODO",
        "NRC",
        "LISTA_CRUZADA",
        "ID_ESTUDIANTE",
        "TIPO_DOCUMENTO",
        "DOCUMENTO",
        "CORREO_ESTUDIANTE",
        "NOMBRE_ESTUDIANTE",
        "APELLIDO_ESTUDIANTE",
        "COD_INSCRIPCIÓN",
        "ESTADO_INSCRIPCIÓN",
        "FECHA_ACTIVIDAD_EST",
        "PAGO",
        "SOCIO_INTEGRADOR",
    ]

    def _save_workbook(self, path: Path, sheet: str, rows: list[list[object]]) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = sheet
        for row in rows:
            worksheet.append(row)
        workbook.save(path)
        workbook.close()

    def test_only_relevant_banner_rows_are_emitted_and_deduplicated(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            banner_dir = root / "banner"
            output_dir = root / "salida"
            banner_dir.mkdir()

            self._save_workbook(
                root / "usuarios.xlsx",
                "Usuarios",
                [["UserName"], [123]],
            )
            banner_rows = [
                self.BANNER_COLUMNS,
                [
                    202642,
                    53080,
                    None,
                    123,
                    "CC",
                    1000,
                    "existing@example.edu",
                    "ANA",
                    "PEREZ",
                    "IN",
                    "Inscrito",
                    datetime(2026, 7, 14),
                    "Y",
                    None,
                ],
                # Duplicado exacto de fuente: no debe volver a emitir comandos.
                [
                    202642,
                    53080,
                    None,
                    123,
                    "CC",
                    1000,
                    "existing@example.edu",
                    "ANA",
                    "PEREZ",
                    "IN",
                    "Inscrito",
                    datetime(2026, 7, 14),
                    "Y",
                    None,
                ],
                [
                    202620,
                    18200,
                    None,
                    456,
                    "CC",
                    2000,
                    "new@example.edu",
                    "LUIS",
                    "GOMEZ",
                    "IN",
                    "Inscrito",
                    datetime(2026, 7, 14),
                    "Y",
                    "BS",
                ],
                # NRC no seleccionado: se descarta antes de construir datos de salida.
                [
                    202642,
                    99999,
                    None,
                    999,
                    "CC",
                    3000,
                    "other@example.edu",
                    "OTRO",
                    "USUARIO",
                    "IN",
                    "Inscrito",
                    datetime(2026, 7, 14),
                    "Y",
                    "BS",
                ],
            ]
            self._save_workbook(banner_dir / "estudiantes.xlsx", "Estudiantes", banner_rows)

            config = {
                "banner_directory": str(banner_dir),
                "bdusuarios_file": str(root / "usuarios.xlsx"),
                "salida_directory": str(output_dir),
                "registro_unico_est_file": str(root / "registro_unicoEst.txt"),
                "students_file": str(root / "students.csv"),
                "Tipo_proceso": "Matricular",
            }
            shortnames = [
                helpers.ShortnameRecord("5557-2065-202642-53080", "53080", "202642"),
                helpers.ShortnameRecord("5557-2065-202642-53080", "18200", "202620"),
            ]

            counts = helpers.procesar_inscripciones(
                config, shortnames, logger_for_tests(), date(2026, 7, 13)
            )
            self.assertEqual(counts, {"5557-2065-202642-53080": 2})

            with (root / "registro_unicoEst.txt").open(encoding="utf-8", newline="") as handle:
                commands = list(csv.reader(handle))
            self.assertEqual(sum(row[0] == "ENROLL" and row[-1] == "5557-2065-202642-53080" for row in commands), 2)
            self.assertEqual(sum(row[0] == "UPDATE" for row in commands), 1)
            self.assertEqual(sum(row[0] == "CREATE" for row in commands), 1)
            self.assertTrue(any(row[-2:] == ["Student_fa", "CVFA"] for row in commands))
            self.assertTrue(any(row[-2:] == ["Student_pr", "CVPR"] for row in commands))


if __name__ == "__main__":
    unittest.main()
