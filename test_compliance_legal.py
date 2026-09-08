"""
Tests para validar cumplimiento legal de liquidaciones.

Prueba los cambios a:
- dias_ganados_por_meses(): Nueva escala de vacaciones (0, 15, 30+)
- calcular_liquidacion(): Validación de indemnización por causal
"""

import unittest
import sys
sys.path.insert(0, r'c:\Users\darie\OneDrive\Escritorio\Sipp2.3')

from datos_Sipp import liquidaciones_db, vacaciones_pagos
from validaciones_legales import ValidacionesLegales


class TestVacacionesLegales(unittest.TestCase):
    """Prueba la nueva escala legal de vacaciones."""
    
    def test_sin_derecho_menos_1_ano(self):
        """Menos de 1 año: 0 días."""
        for meses in [1, 3, 6, 11]:
            resultado = vacaciones_pagos.dias_ganados_por_meses(meses)
            self.assertEqual(resultado, 0, 
                f"Mes {meses}: debería ser 0, obtuvo {resultado}")
    
    def test_15_dias_1_a_4_anos(self):
        """Entre 1-4 años: 15 días."""
        for meses in [12, 24, 36, 48, 59]:
            resultado = vacaciones_pagos.dias_ganados_por_meses(meses)
            self.assertEqual(resultado, 15,
                f"Mes {meses}: debería ser 15, obtuvo {resultado}")
    
    def test_30_dias_5_anos(self):
        """A partir de 5 años: 30 días."""
        resultado = vacaciones_pagos.dias_ganados_por_meses(60)
        self.assertEqual(resultado, 30)
    
    def test_60_dias_10_anos(self):
        """A los 10 años: 30 + 30 = 60 días."""
        resultado = vacaciones_pagos.dias_ganados_por_meses(120)
        self.assertEqual(resultado, 60)
    
    def test_90_dias_15_anos(self):
        """A los 15 años: 30 + 60 = 90 días."""
        resultado = vacaciones_pagos.dias_ganados_por_meses(180)
        self.assertEqual(resultado, 90)


class TestValidacionesLegales(unittest.TestCase):
    """Prueba las validaciones legales del módulo."""
    
    def test_validar_escala_vacaciones(self):
        """Valida función de escala de vacaciones."""
        # Sin derecho
        resultado = ValidacionesLegales.validar_escala_vacaciones(11)
        self.assertEqual(resultado['dias_permitidos'], 0)
        
        # 15 días
        resultado = ValidacionesLegales.validar_escala_vacaciones(24)
        self.assertEqual(resultado['dias_permitidos'], 15)
        
        # 30 días
        resultado = ValidacionesLegales.validar_escala_vacaciones(60)
        self.assertEqual(resultado['dias_permitidos'], 30)
    
    def test_indemnizacion_causas_justificadas(self):
        """Indemnización 0 para causas justificadas."""
        resultado = ValidacionesLegales.validar_indemnizacion(
            "CAUSAS_JUSTIFICADAS", 24
        )
        self.assertEqual(resultado['dias_minimo'], 0)
        self.assertEqual(resultado['dias_maximo'], 0)
    
    def test_indemnizacion_despido_injustificado(self):
        """Indemnización 1-30 días para despido injustificado."""
        resultado = ValidacionesLegales.validar_indemnizacion(
            "DESPIDO_INJUSTIFICADO", 24
        )
        self.assertEqual(resultado['dias_minimo'], 1)
        self.assertEqual(resultado['dias_maximo'], 30)
    
    def test_indemnizacion_cierre_empresa(self):
        """Indemnización 30-60 días para cierre empresa."""
        resultado = ValidacionesLegales.validar_indemnizacion(
            "CIERRE_EMPRESA", 120
        )
        self.assertGreaterEqual(resultado['dias_minimo'], 30)
    
    def test_preaviso_valido(self):
        """Preaviso de 30+ días es válido."""
        resultado = ValidacionesLegales.validar_preaviso(30)
        self.assertTrue(resultado['cumple_ley'])
        self.assertIsNone(resultado['advertencia'])
    
    def test_preaviso_invalido(self):
        """Preaviso menor a 30 días genera advertencia."""
        resultado = ValidacionesLegales.validar_preaviso(15)
        self.assertFalse(resultado['cumple_ley'])
        self.assertIsNotNone(resultado['advertencia'])
    
    def test_permiso_luto(self):
        """Validar permiso de luto."""
        # 3 días (mínimo)
        resultado = ValidacionesLegales.validar_permiso("LUTO", 3)
        self.assertTrue(resultado['aprobado'])
        
        # 5 días (máximo)
        resultado = ValidacionesLegales.validar_permiso("LUTO", 5)
        self.assertTrue(resultado['aprobado'])
        
        # 6 días (fuera de rango)
        resultado = ValidacionesLegales.validar_permiso("LUTO", 6)
        self.assertFalse(resultado['aprobado'])
    
    def test_permiso_maternidad(self):
        """Validar permiso de maternidad."""
        resultado = ValidacionesLegales.validar_permiso("MATERNIDAD", 120)
        self.assertTrue(resultado['aprobado'])
    
    def test_permiso_matrimonio(self):
        """Validar permiso de matrimonio."""
        resultado = ValidacionesLegales.validar_permiso("MATRIMONIO", 3)
        self.assertTrue(resultado['aprobado'])
    
    def test_css_seguro(self):
        """Validar cálculo de CSS y Seguro Educativo."""
        resultado = ValidacionesLegales.validar_css_seguro(1000)
        self.assertEqual(resultado['css'], 97.5)  # 9.75%
        self.assertEqual(resultado['seguro_educativo'], 12.5)  # 1.25%


class TestLiquidacionConNuevaEscala(unittest.TestCase):
    """Prueba liquidación con nueva escala de vacaciones."""
    
    def test_liquidacion_menos_1_ano(self):
        """Liquidación para empleado con menos de 1 año."""
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1000,
            dias_trabajados=90,
            dias_vacaciones=0,  # Sin derecho
            dias_decimo=90,
            preaviso=0,
            indemnizacion=0,
            causal_terminacion="CAUSAS_JUSTIFICADAS"
        )
        
        # Validaciones
        self.assertEqual(resultado['vacaciones_pendientes'], 0)
        self.assertIn('advertencias', resultado)
    
    def test_liquidacion_con_indemnizacion_valida(self):
        """Liquidación con indemnización válida para despido injustificado."""
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1000,
            dias_trabajados=20,
            dias_vacaciones=5,
            dias_decimo=120,
            preaviso=1000,  # 30 días
            indemnizacion=500,
            causal_terminacion="DESPIDO_INJUSTIFICADO"
        )
        
        # Indemnización debe mantenerse
        self.assertAlmostEqual(resultado['indemnizacion'], 500)
        self.assertIn('causal_terminacion', resultado)
    
    def test_liquidacion_indemnizacion_rechazada(self):
        """Liquidación rechaza indemnización para causas justificadas."""
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1000,
            dias_trabajados=20,
            dias_vacaciones=5,
            dias_decimo=120,
            preaviso=1000,
            indemnizacion=500,  # Será rechazada
            causal_terminacion="CAUSAS_JUSTIFICADAS"
        )
        
        # Indemnización debe ser 0 (rechazada por ley)
        self.assertEqual(resultado['indemnizacion'], 0)
        # Debe haber advertencia
        self.assertTrue(len(resultado.get('advertencias', [])) > 0)


class TestCompatibilidadRegresiva(unittest.TestCase):
    """Valida que cambios no rompan funcionalidad existente."""
    
    def test_calcular_liquidacion_sin_parametro_causal(self):
        """Función funciona sin parámetro causal (valor por defecto)."""
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1000,
            dias_trabajados=20,
            dias_vacaciones=5,
            dias_decimo=120,
            preaviso=0,
            indemnizacion=0
            # NO se pasa causal_terminacion
        )
        
        # Debe funcionar y usar valor por defecto
        self.assertIsNotNone(resultado)
        self.assertIn('total_neto', resultado)
        self.assertIn('causal_terminacion', resultado)


if __name__ == '__main__':
    # Ejecutar tests
    suite = unittest.TestLoader().loadTestsFromModule(sys.modules[__name__])
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Resumen
    print("\n" + "=" * 60)
    print("RESUMEN DE PRUEBAS")
    print("=" * 60)
    print(f"Tests ejecutados: {result.testsRun}")
    print(f"Exitosos: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Fallos: {len(result.failures)}")
    print(f"Errores: {len(result.errors)}")
    
    if result.failures:
        print("\nFALLOS:")
        for test, traceback in result.failures:
            print(f"  - {test}")
    
    if result.errors:
        print("\nERRORES:")
        for test, traceback in result.errors:
            print(f"  - {test}")
