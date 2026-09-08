"""
Módulo de validaciones legales según Código de Trabajo de Panamá.

Este módulo contiene validaciones para garantizar cumplimiento con:
- Código de Trabajo (Ley 40 de 1981 y sus reformas)
- Reglamento sobre régimen de prestaciones sociales (Acuerdo 33 de 2010)
- Resoluciones de direcciones de trabajo y seguro social
"""

import logging
from datetime import datetime, timedelta
from enum import Enum

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class CausalTerminacion(Enum):
    """Causales de terminación según Código de Trabajo Panamá."""
    DESPIDO_INJUSTIFICADO = "DESPIDO_INJUSTIFICADO"
    CIERRE_EMPRESA = "CIERRE_EMPRESA"
    CAUSAS_JUSTIFICADAS = "CAUSAS_JUSTIFICADAS"
    RENUNCIA = "RENUNCIA"
    JUBILACION = "JUBILACION"
    MUERTE = "MUERTE"


class TipoPermiso(Enum):
    """Tipos de permisos legales según Código de Trabajo."""
    LUTO = ("LUTO", 3, 5)  # (tipo, mín, máx)
    MATERNIDAD = ("MATERNIDAD", 120, 120)
    PATERNIDAD = ("PATERNIDAD", 5, 5)
    MATRIMONIO = ("MATRIMONIO", 3, 3)
    LICENCIA_MEDICA = ("LICENCIA_MEDICA", 1, 365)


class ValidacionesLegales:
    """Validaciones de cumplimiento laboral según ley panameña."""
    
    @staticmethod
    def validar_escala_vacaciones(meses_servicio):
        """
        Valida y calcula días de vacaciones según ley.
        
        Código de Trabajo Art. 202:
        - < 1 año: 0 días
        - 1-4 años: 15 días
        - 5+ años: 30 días + 30/5 años adicionales
        
        Args:
            meses_servicio: Número de meses de servicio
            
        Returns:
            dict con dias_permitidos, categoria, y advertencia si aplica
        """
        if meses_servicio < 12:
            return {
                "dias_permitidos": 0,
                "categoria": "SIN_DERECHO",
                "descripcion": "Menos de 1 año: sin derecho a vacaciones",
                "advertencia": None
            }
        
        if meses_servicio < 60:  # Menos de 5 años
            return {
                "dias_permitidos": 15,
                "categoria": "15_DIAS",
                "descripcion": "1 a 4 años de servicio: 15 días",
                "advertencia": None
            }
        
        # 5+ años
        años_adicionales = (meses_servicio - 60) // 60
        dias = 30 + (años_adicionales * 30)
        
        return {
            "dias_permitidos": min(dias, 120),  # Máximo 120 días
            "categoria": "30_PLUS",
            "descripcion": f"5+ años de servicio: 30 + {años_adicionales * 30} días",
            "advertencia": None
        }
    
    @staticmethod
    def validar_indemnizacion(causal, meses_servicio):
        """
        Valida indemnización según causal de terminación.
        
        Código de Trabajo Art. 203-204:
        - Causas justificadas: 0 días
        - Renuncia: 0 días
        - Despido injustificado: 1-30 días según antigüedad
        - Cierre empresa: 30+ días según antigüedad
        
        Args:
            causal: CausalTerminacion enum value
            meses_servicio: Meses de servicio
            
        Returns:
            dict con dias_minimo, dias_maximo, y criterios
        """
        causal = str(causal).upper().strip()
        
        escalas = {
            "CAUSAS_JUSTIFICADAS": {
                "dias_minimo": 0,
                "dias_maximo": 0,
                "descripcion": "Causas justificadas: sin derecho",
                "legal": True
            },
            "RENUNCIA": {
                "dias_minimo": 0,
                "dias_maximo": 0,
                "descripcion": "Renuncia voluntaria: sin derecho",
                "legal": True
            },
            "DESPIDO_INJUSTIFICADO": {
                "dias_minimo": 1,
                "dias_maximo": 30,
                "descripcion": "Despido injustificado: 1-30 días",
                "legal": True,
                "criterios": [
                    "< 3 meses: 1 día",
                    "3-6 meses: 1 día",
                    "6-12 meses: 5 días",
                    "1-5 años: 10 días",
                    "5+ años: 30 días"
                ]
            },
            "CIERRE_EMPRESA": {
                "dias_minimo": 30,
                "dias_maximo": 60,
                "descripcion": "Cierre de empresa: 30-60 días",
                "legal": True,
                "criterios": [
                    "< 1 año: 30 días",
                    "1-5 años: 45 días",
                    "5+ años: 60 días"
                ]
            }
        }
        
        if causal not in escalas:
            return {
                "dias_minimo": 0,
                "dias_maximo": 0,
                "descripcion": f"Causal '{causal}' no reconocida",
                "legal": False,
                "advertencia": "Causal desconocida - revisar con RR.HH."
            }
        
        return escalas[causal]
    
    @staticmethod
    def validar_preaviso(dias_preaviso, causal=None):
        """
        Valida cumplimiento de preaviso.
        
        Código de Trabajo Art. 234: Mínimo 30 días de preaviso
        
        Args:
            dias_preaviso: Días de preaviso efectivos
            causal: CausalTerminacion (opcional)
            
        Returns:
            dict con validacion y advertencias
        """
        es_valido = dias_preaviso >= 30
        
        return {
            "es_valido": es_valido,
            "dias_minimos": 30,
            "dias_efectivos": dias_preaviso,
            "cumple_ley": es_valido,
            "advertencia": None if es_valido else 
                          f"INCUMPLIMIENTO: Preaviso de {dias_preaviso} días (mínimo: 30 días)",
            "consecuencia": None if es_valido 
                          else "Empleador debe pagar diferencia como salario adicional"
        }
    
    @staticmethod
    def validar_permiso(tipo_permiso, dias_solicitados, fecha_evento=None):
        """
        Valida permiso legal según tipo.
        
        Código de Trabajo Art. 207-211:
        - Luto: 3-5 días (familiar directo)
        - Maternidad: 120 días (8 antes + 56 después)
        - Paternidad: 5 días
        - Matrimonio: 3 días
        
        Args:
            tipo_permiso: TipoPermiso enum
            dias_solicitados: Días que se piden
            fecha_evento: Fecha del evento (para luto/matrimonio)
            
        Returns:
            dict con aprobación, dias_permitidos, y restricciones
        """
        permisos = {
            "LUTO": {
                "dias_minimo": 3,
                "dias_maximo": 5,
                "descripcion": "Luto por familiar directo",
                "requisitos": "Certificado de defunción",
                "plazo": "Dentro de 30 días del fallecimiento"
            },
            "MATERNIDAD": {
                "dias_minimo": 120,
                "dias_maximo": 120,
                "descripcion": "Maternidad",
                "distribución": "8 días antes + 56 después del parto",
                "requisitos": "Certificado médico y carnet prenatal",
                "plazo": "Notificación mínimo 30 días antes"
            },
            "PATERNIDAD": {
                "dias_minimo": 5,
                "dias_maximo": 5,
                "descripcion": "Paternidad",
                "requisitos": "Certificado de nacimiento",
                "plazo": "Dentro de 30 días del nacimiento"
            },
            "MATRIMONIO": {
                "dias_minimo": 3,
                "dias_maximo": 3,
                "descripcion": "Matrimonio",
                "requisitos": "Cédula + invitación o comprobante",
                "plazo": "Notificación mínimo 15 días antes"
            },
            "LICENCIA_MEDICA": {
                "dias_minimo": 1,
                "dias_maximo": 365,
                "descripcion": "Licencia médica",
                "requisitos": "Certificado médico válido",
                "plazo": "Según lo ordenado por médico"
            }
        }
        
        tipo = str(tipo_permiso).upper().strip()
        
        if tipo not in permisos:
            return {
                "aprobado": False,
                "tipo": tipo,
                "advertencia": f"Tipo de permiso '{tipo}' no reconocido"
            }
        
        permiso = permisos[tipo]
        es_valido = permiso["dias_minimo"] <= dias_solicitados <= permiso["dias_maximo"]
        
        return {
            "aprobado": es_valido,
            "tipo": tipo,
            "dias_permitidos": permiso.get("dias_minimo"),
            "dias_solicitados": dias_solicitados,
            "descripcion": permiso["descripcion"],
            "requisitos": permiso.get("requisitos"),
            "advertencia": None if es_valido 
                          else f"Días fuera de rango: {permiso['dias_minimo']}-{permiso['dias_maximo']}"
        }
    
    @staticmethod
    def validar_decimo_proporcional(salario_anual, dias_trabajados_periodo):
        """
        Valida cálculo de décimo proporcional.
        
        Código de Trabajo Art. 211-212:
        Décimo = (Salario anual / 360) × días del período
        
        Se calcula cada cuatrimestre (enero-abril, mayo-agosto, septiembre-diciembre)
        
        Args:
            salario_anual: Salario anual base
            dias_trabajados_periodo: Días trabajados en período cuatrimestral
            
        Returns:
            dict con validación y cálculo
        """
        diario = salario_anual / 360
        decimo = diario * dias_trabajados_periodo
        
        return {
            "salario_anual": salario_anual,
            "salario_diario": diario,
            "dias_periodo": dias_trabajados_periodo,
            "decimo_proporcional": decimo,
            "formula": "(Salario anual ÷ 360) × días",
            "legal": True,
            "observacion": "Se paga el 24 de diciembre, equivalente acumulado"
        }
    
    @staticmethod
    def validar_css_seguro(base_gravable):
        """
        Valida cálculo de aportes CSS y Seguro Educativo.
        
        Regulación:
        - CSS: 9.75% sobre base gravable
        - Seguro Educativo: 1.25% sobre base gravable
        - Base gravable = Salario + Vacaciones + Preaviso + Otros (excluye décimo, indemnización)
        
        Args:
            base_gravable: Base sobre la cual calcular aportes
            
        Returns:
            dict con validación de cálculos
        """
        css = base_gravable * 0.0975
        seguro = base_gravable * 0.0125
        total_aportes = css + seguro
        
        return {
            "base_gravable": base_gravable,
            "css": css,
            "tasa_css": "9.75%",
            "seguro_educativo": seguro,
            "tasa_seguro": "1.25%",
            "total_aportes": total_aportes,
            "tasa_total": "11.00%",
            "legal": True
        }


def generar_reporte_compliance(datos_empleado, datos_liquidacion):
    """
    Genera reporte de cumplimiento legal para un empleado.
    
    Args:
        datos_empleado: dict con info del empleado (meses_servicio, causal, etc)
        datos_liquidacion: dict con cálculos de liquidación
        
    Returns:
        str con reporte formatted
    """
    reporte = []
    reporte.append("=" * 60)
    reporte.append("REPORTE DE CUMPLIMIENTO LEGAL")
    reporte.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    reporte.append("=" * 60)
    
    # Validar vacaciones
    reporte.append("\n1. VALIDACIÓN DE VACACIONES")
    vacaciones_val = ValidacionesLegales.validar_escala_vacaciones(
        datos_empleado.get("meses_servicio", 0)
    )
    reporte.append(f"   - {vacaciones_val['descripcion']}")
    reporte.append(f"   - Días permitidos: {vacaciones_val['dias_permitidos']}")
    
    # Validar indemnización
    reporte.append("\n2. VALIDACIÓN DE INDEMNIZACIÓN")
    indem_val = ValidacionesLegales.validar_indemnizacion(
        datos_empleado.get("causal", "DESPIDO_INJUSTIFICADO"),
        datos_empleado.get("meses_servicio", 0)
    )
    reporte.append(f"   - {indem_val['descripcion']}")
    reporte.append(f"   - Rango: {indem_val['dias_minimo']}-{indem_val['dias_maximo']} días")
    
    # Validar preaviso
    reporte.append("\n3. VALIDACIÓN DE PREAVISO")
    preaviso_val = ValidacionesLegales.validar_preaviso(
        datos_liquidacion.get("preaviso_dias", 0)
    )
    reporte.append(f"   - Cumple ley: {'SÍ' if preaviso_val['cumple_ley'] else 'NO'}")
    if preaviso_val['advertencia']:
        reporte.append(f"   - ⚠ {preaviso_val['advertencia']}")
    
    # Advertencias de liquidación
    if datos_liquidacion.get("advertencias"):
        reporte.append("\n4. ADVERTENCIAS")
        for adv in datos_liquidacion["advertencias"]:
            reporte.append(f"   - ⚠ {adv}")
    
    reporte.append("\n" + "=" * 60)
    return "\n".join(reporte)


if __name__ == "__main__":
    # Pruebas básicas
    print("Prueba 1: Validar vacaciones para 20 meses")
    print(ValidacionesLegales.validar_escala_vacaciones(20))
    
    print("\nPrueba 2: Validar indemnización por despido injustificado")
    print(ValidacionesLegales.validar_indemnizacion("DESPIDO_INJUSTIFICADO", 24))
    
    print("\nPrueba 3: Validar preaviso de 15 días")
    print(ValidacionesLegales.validar_preaviso(15))
