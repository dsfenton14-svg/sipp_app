import unittest
from datetime import date, time

from horarios_laborales import dia_cuenta_como_ausencia, dia_cuenta_como_pagado, dia_excluye_ausencia, HorarioLaboral, calcular_jornada


class HorariosLaboralesTests(unittest.TestCase):
    def test_jornada_diurna_con_tolerancia(self):
        horario = HorarioLaboral(
            "Administrativo", time(8), time(17), tolerancia_entrada_minutos=10
        )
        resultado = calcular_jornada(horario, date(2026, 8, 27), "08:15", "17:00")

        self.assertEqual(resultado.horas_trabajadas, 8.75)
        self.assertEqual(resultado.horas_ordinarias, 8.0)
        self.assertEqual(resultado.horas_extra, 0.75)
        self.assertEqual(resultado.minutos_tardanza, 5)

    def test_jornada_nocturna_cruza_medianoche(self):
        horario = HorarioLaboral(
            "Nocturno", time(22), time(6), tipo_jornada="NOCTURNA", cruza_medianoche=True
        )
        resultado = calcular_jornada(horario, "2026-08-27", "22:00", "06:00")

        self.assertTrue(resultado.cruza_medianoche)
        self.assertEqual(resultado.horas_trabajadas, 8.0)
        self.assertEqual(resultado.horas_ordinarias, 7.0)
        self.assertEqual(resultado.horas_extra, 1.0)

    def test_salida_anticipada_respetando_tolerancia(self):
        horario = HorarioLaboral(
            "Administrativo", time(8), time(17), tolerancia_salida_minutos=15
        )
        resultado = calcular_jornada(horario, "2026-08-27", "08:00", "16:40")

        self.assertEqual(resultado.minutos_salida_anticipada, 5)

    def test_documentos_medicos_cuentan_como_dia_pagado(self):
        self.assertTrue(dia_cuenta_como_pagado("CERTIFICADO MÉDICO"))
        self.assertTrue(dia_cuenta_como_pagado("INCAPACIDAD"))
        self.assertFalse(dia_cuenta_como_ausencia("INCAPACIDAD"))
        self.assertTrue(dia_excluye_ausencia("INCAPACIDAD"))

    def test_constancia_no_es_ausencia_ni_dia_pagado(self):
        self.assertFalse(dia_cuenta_como_pagado("CONSTANCIA DE ASISTENCIA"))
        self.assertFalse(dia_cuenta_como_ausencia("CONSTANCIA DE ASISTENCIA"))
        self.assertTrue(dia_excluye_ausencia("CONSTANCIA DE ASISTENCIA"))


if __name__ == "__main__":
    unittest.main()
