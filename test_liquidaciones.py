import unittest
from datetime import date

from datos_Sipp import gestion_empleado, liquidaciones_db


class LiquidacionesCalculoTests(unittest.TestCase):
    def limpiar_datos_prueba(self, empleado_id):
        servicio = liquidaciones_db()
        with servicio.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM liquidaciones WHERE empleado_id = %s", (empleado_id,))
            cursor.execute("DELETE FROM empleados WHERE id = %s", (empleado_id,))
        servicio.conexion.commit()

    def test_calcula_ingresos_y_deducciones(self):
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1500,
            dias_trabajados=10,
            dias_vacaciones=5,
            dias_decimo=90,
            preaviso=250,
            indemnizacion=500,
            otros_ingresos=100,
            otros_descuentos=50,
        )

        self.assertEqual(resultado["salario_pendiente"], 500.0)
        self.assertEqual(resultado["vacaciones_pendientes"], 250.0)
        self.assertEqual(resultado["decimo_proporcional"], 375.0)
        self.assertEqual(resultado["total_bruto"], 1975.0)
        self.assertEqual(resultado["base_gravable"], 1100.0)
        self.assertEqual(resultado["css"], 107.25)
        self.assertEqual(resultado["seguro_educativo"], 13.75)
        self.assertEqual(resultado["isr"], 43.75)
        self.assertEqual(resultado["total_neto"], 1760.25)

    def test_no_permite_valores_negativos(self):
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=1000,
            dias_trabajados=-10,
            dias_vacaciones=-5,
            preaviso=-100,
            indemnizacion=-100,
            otros_descuentos=-50,
        )

        self.assertEqual(resultado["total_bruto"], 0.0)
        self.assertEqual(resultado["total_deducciones"], 0.0)
        self.assertEqual(resultado["total_neto"], 0.0)

    def test_isr_no_aplica_bajo_limite_anual(self):
        resultado = liquidaciones_db.calcular_liquidacion(
            salario_mensual=900,
            dias_trabajados=30,
        )

        self.assertEqual(resultado["isr"], 0.0)

    def test_liquidacion_marca_empleado_como_liquidado_sin_borrar_historial(self):
        servicio_empleados = gestion_empleado()
        servicio_liquidaciones = liquidaciones_db()
        dui = "LQ-TEST-999"
        servicio_empleados.insertar_empleados(
            "Empleado de prueba",
            dui,
            "prueba@example.com",
            "Dirección de prueba",
            "9999-9999",
            "Supervisor",
            "Operaciones",
            "1200",
            "2023-01-15",
            "1990-01-01",
        )
        empleado = servicio_empleados.buscar_empleado(dui)
        self.assertIsNotNone(empleado)
        self.addCleanup(self.limpiar_datos_prueba, empleado[0])

        liquidacion_id, _ = servicio_liquidaciones.guardar_liquidacion(
            empleado,
            date(2024, 6, 30),
            "Renuncia",
            dias_trabajados=120,
            dias_vacaciones=0,
            dias_decimo=0,
            preaviso=0,
            indemnizacion=0,
            otros_ingresos=0,
            otros_descuentos=0,
            observaciones="Liquidación de prueba",
            estado="CALCULADA",
        )

        self.assertIsNotNone(liquidacion_id)
        self.assertEqual(servicio_empleados.obtener_estado_empleado(empleado[0]), "LIQUIDADO")
        self.assertNotIn(empleado[0], [fila[0] for fila in servicio_empleados.obtener_empleados()])
        self.assertIn(empleado[0], [fila[0] for fila in servicio_empleados.obtener_empleados_todos()])


if __name__ == "__main__":
    unittest.main()