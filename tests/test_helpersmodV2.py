import csv
import os
from pathlib import Path
import tempfile
import unittest

import pandas as pd

import helpersmodV2 as helpers


class ResponsablesProgramaTests(unittest.TestCase):
    def test_leer_coordinadores_usa_nuevo_contrato_y_conserva_mentor_sin_coordinador(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "Coordinadores.xlsx"
            pd.DataFrame(
                [
                    {
                        "Centro de Costos": "VE01",
                        "Coordinador_ID": 123,
                        "Mentor1_Id": 456,
                        "Mentor2_Id": None,
                        "Mentor3_Id": None,
                    },
                    {
                        "Centro de Costos": "VE02",
                        "Coordinador_ID": None,
                        "Mentor1_Id": 789,
                        "Mentor2_Id": None,
                        "Mentor3_Id": None,
                    },
                    {
                        "Centro de Costos": "VE03",
                        "Coordinador_ID": None,
                        "Mentor1_Id": None,
                        "Mentor2_Id": None,
                        "Mentor3_Id": None,
                    },
                ]
            ).to_excel(path, index=False)

            result = helpers.leer_coordinadores(path)

            self.assertEqual(result["Centro de Costos"].tolist(), ["VE01", "VE02"])
            self.assertEqual(result.loc[0, "Coordinador_ID"], "000000123")
            self.assertEqual(result.loc[1, "Mentor1_Id"], "000000789")

    def test_leer_coordinadores_rechaza_archivo_con_esquema_anterior(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "Coordinadores.xlsx"
            pd.DataFrame(
                [{"Centro de Costos": "VE01", "ID COORDINADOR": 123}]
            ).to_excel(path, index=False)

            self.assertIsNone(helpers.leer_coordinadores(path))

    def test_extraer_mentores_elimina_ids_repetidos(self):
        row = pd.Series(
            {
                "Mentor1_Id": "123",
                "Mentor1_Nombres": "Mentor Uno",
                "Mentor1_Correo": "uno@example.edu",
                "Mentor2_Id": 123.0,
                "Mentor2_Nombres": "Duplicado",
                "Mentor3_Id": "456",
                "Mentor3_Nombres": "Mentor Tres",
            }
        )

        result = helpers.extraer_mentores(row)

        self.assertEqual([mentor["id"] for mentor in result], ["000000123", "000000456"])
        self.assertEqual([mentor["origen"] for mentor in result], ["Mentor1_Id", "Mentor3_Id"])


class CrearArchivosRolesTests(unittest.TestCase):
    def setUp(self):
        self.original_config = helpers.CONFIG.copy()
        self.original_cwd = Path.cwd()
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        os.chdir(self.root)
        helpers.CONFIG["salida_directory"] = str(self.root / "salida")

    def tearDown(self):
        helpers.CONFIG.clear()
        helpers.CONFIG.update(self.original_config)
        os.chdir(self.original_cwd)
        self.temporary_directory.cleanup()

    @staticmethod
    def _usuarios(*ids):
        return pd.DataFrame(
            [
                {
                    "UserName": helpers._normalizar_id_banner(user_id),
                    "FirstName": f"Nombre {user_id}",
                    "LastName": "Apellido",
                    "OrgRoleId": "",
                    "OrgDefinedId": f"CC. {user_id}",
                    "ExternalEmail": f"{user_id}@example.edu",
                }
                for user_id in ids
            ]
        )

    @staticmethod
    def _moderador(user_id):
        return pd.DataFrame(
            [
                {
                    "ID_DOCENTE": helpers._normalizar_id_banner(user_id),
                    "DOCUMENTO": user_id,
                    "TIPO_DOCUMENTO": "CC",
                    "NOMBRE_DOCENTE": "Docente",
                    "APELLIDO_DOCENTE": "Moderador",
                    "CORREO_DOCENTE": "moderador@example.edu",
                }
            ]
        )

    @staticmethod
    def _centro_costos():
        return pd.DataFrame(
            [
                {
                    "PERIODO": "202642",
                    "LISTA_CRUZADA": "LC01",
                    "ESTADO_INSCRIPCIÓN": "Inscrito",
                    "COD_PROGRAMA_ESTUDIANTE": "VE01",
                }
            ]
        )

    @staticmethod
    def _responsables(coordinador_id="123"):
        return pd.DataFrame(
            [
                {
                    "Centro de Costos": "VE01",
                    "Coordinador_ID": helpers._normalizar_id_banner(coordinador_id),
                    "Coordinador_Nombre": "Coordinador Prueba",
                    "Coordinador_Correo": "coordinador@example.edu",
                    "Mentor1_Id": helpers._normalizar_id_banner("456"),
                    "Mentor1_Nombres": "Mentor Existente",
                    "Mentor1_Correo": "mentor456@example.edu",
                    "Mentor2_Id": helpers._normalizar_id_banner("789"),
                    "Mentor2_Nombres": "Mentor Nuevo",
                    "Mentor2_Correo": "mentor789@example.edu",
                    # También es moderador: debe prevalecer el rol Moderador.
                    "Mentor3_Id": helpers._normalizar_id_banner("123"),
                    "Mentor3_Nombres": "Mentor Moderador",
                    "Mentor3_Correo": "moderador@example.edu",
                }
            ]
        )

    def _commands(self):
        path = self.root / "salida" / "registro_CURSO-PRUEBA.txt"
        with path.open(encoding="utf-8", newline="") as handle:
            return list(csv.reader(handle))

    def test_coordinador_que_ya_es_moderador_no_se_inscribe_como_coordinador(self):
        log_path = self.root / "proceso.log"
        helpers.crearArchivos(
            self._moderador("123"),
            "CURSO-PRUEBA",
            "LC01",
            "202642",
            self._usuarios("123", "456"),
            self._centro_costos(),
            self._responsables("123"),
            str(log_path),
        )

        commands = self._commands()
        enrollments = [row for row in commands if row[0] == "ENROLL" and row[-1] == "CURSO-PRUEBA"]
        self.assertIn(["ENROLL", "000000123", "", "Moderador", "CURSO-PRUEBA"], enrollments)
        self.assertNotIn(["ENROLL", "000000123", "", "Coordinador", "CURSO-PRUEBA"], enrollments)
        self.assertEqual(
            [row for row in enrollments if row[3] == "Mentor"],
            [
                ["ENROLL", "000000456", "", "Mentor", "CURSO-PRUEBA"],
                ["ENROLL", "000000789", "", "Mentor", "CURSO-PRUEBA"],
            ],
        )
        self.assertEqual(
            len([row for row in commands if row[0] == "CREATE" and row[1] == "000000789"]),
            1,
        )
        self.assertIn("ya fue inscrito como Moderador", log_path.read_text(encoding="utf-8"))

    def test_coordinador_distinto_del_moderador_se_inscribe(self):
        helpers.crearArchivos(
            self._moderador("999"),
            "CURSO-PRUEBA",
            "LC01",
            "202642",
            self._usuarios("123", "456", "999"),
            self._centro_costos(),
            self._responsables("123"),
            str(self.root / "proceso.log"),
        )

        self.assertIn(
            ["ENROLL", "000000123", "", "Coordinador", "CURSO-PRUEBA"],
            self._commands(),
        )


if __name__ == "__main__":
    unittest.main()
