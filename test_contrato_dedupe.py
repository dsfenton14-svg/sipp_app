import unittest

from datos_Sipp import expedientes_laborales


class ContratoDedupeTests(unittest.TestCase):
    def test_generar_clave_de_duplicado(self):
        clave = expedientes_laborales._generar_clave_contrato(
            12,
            "Definido",
            "2024-01-01",
            "2024-12-31",
            "Tiempo completo",
            "08:00",
            "17:00",
            "Activo",
        )

        self.assertEqual(
            clave,
            (
                12,
                "definido",
                "2024-01-01",
                "2024-12-31",
                "tiempo completo",
                "08:00",
                "17:00",
                "activo",
            ),
        )


if __name__ == "__main__":
    unittest.main()
