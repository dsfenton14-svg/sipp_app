#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Test de cálculo de liquidación"""

from datos_Sipp import liquidaciones_db

# Test 1: Empleado con salario de 1000, 20 días trabajados, 5 días vacaciones
resultado = liquidaciones_db.calcular_liquidacion(
    salario_mensual=1000,
    dias_trabajados=20,
    dias_vacaciones=5,
    dias_decimo=120,
    preaviso=100,
    indemnizacion=200,
    otros_ingresos=50,
    otros_descuentos=10,
    incluir_deducciones=True
)

print('=== PRUEBA DE CÁLCULO DE LIQUIDACIÓN ===')
print(f'Salario Mensual: B/. {resultado["salario_mensual"]:.2f}')
print(f'Salario Pendiente (20 días): B/. {resultado["salario_pendiente"]:.2f}')
print(f'Vacaciones (5 días): B/. {resultado["vacaciones_pendientes"]:.2f}')
print(f'Décimo Proporcional (120 días): B/. {resultado["decimo_proporcional"]:.2f}')
print(f'Preaviso: B/. {resultado["preaviso"]:.2f}')
print(f'Indemnización: B/. {resultado["indemnizacion"]:.2f}')
print(f'Otros Ingresos: B/. {resultado["otros_ingresos"]:.2f}')
print()
print(f'Base Gravable: B/. {resultado["base_gravable"]:.2f}')
print(f'CSS (9.75%): B/. {resultado["css"]:.2f}')
print(f'Seguro Educativo (1.25%): B/. {resultado["seguro_educativo"]:.2f}')
print(f'ISR (proporcional): B/. {resultado["isr"]:.2f}')
print(f'Otros Descuentos: B/. {resultado["otros_descuentos"]:.2f}')
print()
print(f'Total Bruto: B/. {resultado["total_bruto"]:.2f}')
print(f'Total Deducciones: B/. {resultado["total_deducciones"]:.2f}')
print(f'Total Neto: B/. {resultado["total_neto"]:.2f}')
print()

# Validaciones
assert resultado["salario_pendiente"] == 666.67, "Salario pendiente incorrecto"
assert resultado["vacaciones_pendientes"] == 166.67, "Vacaciones incorrectas"
assert resultado["base_gravable"] == 983.33, "Base gravable incorrecta (debe incluir preaviso: 666.67+166.67+100+50)"
assert resultado["total_bruto"] > 0, "Total bruto debe ser positivo"
assert resultado["total_neto"] > 0, "Total neto debe ser positivo"
assert resultado["css"] > 0, "CSS debe ser positivo (incluye preaviso)"
assert resultado["isr"] > 0, "ISR debe ser positivo y proporcional"

print('✓ Todas las validaciones pasaron')
print('✓ Cálculo completado exitosamente')
