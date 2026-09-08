import unittest

from SiPP import Deducciones


class DeduccionesNominaTests(unittest.TestCase):
    def setUp(self):
        self.deducciones = Deducciones()
        self.deducciones.valores_por_defecto.update({
            "CSS": "9.75%",
            "Seguro Educativo": "1.25%",
        })

    def test_pago_bruto_con_isr_progresivo(self):
        resultado = self.deducciones.calcular_deducciones_nomina(1500, incluir_isr=True)

        self.assertEqual(resultado["css"], 146.25)
        self.assertEqual(resultado["seguro_educativo"], 18.75)
        self.assertEqual(resultado["isr"], 87.5)
        self.assertEqual(resultado["total_deducciones"], 252.5)
        self.assertEqual(resultado["salario_neto"], 1247.5)

    def test_isr_proporcional_para_pago_parcial(self):
        salario_mensual = 3000.0
        pago_vacaciones = 1500.0
        deducciones = self.deducciones.calcular_deducciones_nomina(
            pago_vacaciones,
            incluir_isr=False,
        )
        isr_mensual = self.deducciones.formula_isr_panama(salario_mensual)
        isr_proporcional = self.deducciones.redondear(
            isr_mensual * (pago_vacaciones / salario_mensual)
        )

        total_deducciones = self.deducciones.redondear(
            deducciones["css"] + deducciones["seguro_educativo"] + isr_proporcional
        )
        pago_neto = self.deducciones.redondear(pago_vacaciones - total_deducciones)

        self.assertEqual(isr_mensual, 312.5)
        self.assertEqual(isr_proporcional, 156.25)
        self.assertEqual(total_deducciones, 321.25)
        self.assertEqual(pago_neto, 1178.75)

    def test_isr_no_aplica_hasta_11000_anuales(self):
        self.assertEqual(self.deducciones.formula_isr_panama(900), 0.0)

    def test_salario_periodo_no_supera_quince_dias(self):
        self.assertEqual(self.deducciones.calcular_salario_periodo(1500, 15), 750.0)
        self.assertEqual(self.deducciones.calcular_salario_periodo(1500, 20), 750.0)
        self.assertEqual(self.deducciones.calcular_salario_periodo(1500, 10), 500.0)

    def test_isr_se_prorratea_en_pago_quincenal(self):
        resultado = self.deducciones.calcular_deducciones_nomina(
            750,
            salario_mensual=1500,
            dias_pagados=15,
        )
        self.assertEqual(resultado["isr"], 43.75)


if __name__ == "__main__":
    unittest.main()
