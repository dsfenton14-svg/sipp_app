"""Reglas de jornada laboral para SiPP.

El modulo calcula duracion y novedades de una jornada sin depender de la interfaz
ni de la base de datos. Los porcentajes de recargo se mantienen fuera del modelo
porque dependen del concepto legal o contractual que se este liquidando.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta


LIMITES_SEMANALES = {
    "DIURNA": 48.0,
    "NOCTURNA": 42.0,
    "MIXTA": 45.0,
}

LIMITES_DIARIOS = {
    "DIURNA": 8.0,
    "NOCTURNA": 7.0,
    "MIXTA": 7.5,
}

TIPOS_DIA_NO_ASISTENCIA = {"DÍA LIBRE", "DIA LIBRE", "AUSENCIA"}
TIPOS_DIA_PAGADOS = {"JORNADA LABORAL", "CERTIFICADO MÉDICO", "CERTIFICADO MEDICO", "INCAPACIDAD"}
TIPOS_DIA_NO_PAGADOS = {"CONSTANCIA DE ASISTENCIA"}


def normalizar_tipo_dia(valor: str) -> str:
    return str(valor or "JORNADA LABORAL").strip().upper()


def dia_cuenta_como_pagado(valor: str) -> bool:
    return normalizar_tipo_dia(valor) in TIPOS_DIA_PAGADOS


def dia_cuenta_como_ausencia(valor: str) -> bool:
    return normalizar_tipo_dia(valor) == "AUSENCIA"


def dia_excluye_ausencia(valor: str) -> bool:
    return normalizar_tipo_dia(valor) in TIPOS_DIA_NO_ASISTENCIA | TIPOS_DIA_PAGADOS | TIPOS_DIA_NO_PAGADOS


def feriados_panama(anio: int) -> dict[date, str]:
    """Devuelve feriados nacionales usuales; la empresa puede ampliar este mapa."""
    feriados = {
        date(anio, 1, 1): "Año Nuevo",
        date(anio, 1, 9): "Día de los Mártires",
        date(anio, 5, 1): "Día del Trabajo",
        date(anio, 11, 3): "Separación de Panamá de Colombia",
        date(anio, 11, 5): "Consolidación de la Separación de Colombia",
        date(anio, 11, 10): "Primer Grito de Independencia",
        date(anio, 11, 28): "Independencia de Panamá de España",
        date(anio, 12, 8): "Día de las Madres",
        date(anio, 12, 25): "Navidad",
    }
    domingo_pascua = _domingo_pascua(anio)
    feriados[domingo_pascua - timedelta(days=2)] = "Viernes Santo"
    return feriados


def _domingo_pascua(anio: int) -> date:
    """Calcula el domingo de Pascua con el algoritmo gregoriano."""
    a = anio % 19
    b = anio // 100
    c = anio % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = (h + l - 7 * m + 114) % 31 + 1
    return date(anio, mes, dia)


@dataclass(frozen=True)
class HorarioLaboral:
    nombre: str
    entrada: time
    salida: time
    tipo_jornada: str = "DIURNA"
    tolerancia_entrada_minutos: int = 0
    tolerancia_salida_minutos: int = 0
    cruza_medianoche: bool = False

    def __post_init__(self):
        tipo = self.tipo_jornada.strip().upper()
        if tipo not in LIMITES_SEMANALES:
            raise ValueError("El tipo de jornada debe ser DIURNA, NOCTURNA o MIXTA.")
        if self.tolerancia_entrada_minutos < 0 or self.tolerancia_salida_minutos < 0:
            raise ValueError("Las tolerancias no pueden ser negativas.")
        object.__setattr__(self, "tipo_jornada", tipo)

    @property
    def maximo_diario(self) -> float:
        return LIMITES_DIARIOS[self.tipo_jornada]

    @property
    def maximo_semanal(self) -> float:
        return LIMITES_SEMANALES[self.tipo_jornada]


@dataclass(frozen=True)
class ResultadoJornada:
    horas_trabajadas: float
    horas_ordinarias: float
    horas_extra: float
    minutos_tardanza: int
    minutos_salida_anticipada: int
    jornada_incompleta: bool
    cruza_medianoche: bool


def _como_hora(valor: time | str) -> time:
    if isinstance(valor, time):
        return valor.replace(second=0, microsecond=0)
    texto = str(valor or "").strip().upper()
    for formato in ("%I:%M %p", "%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(texto, formato).time().replace(second=0, microsecond=0)
        except ValueError:
            continue
    raise ValueError(f"Hora invalida: {valor}")


def _momento(fecha: date, hora: time, siguiente_dia: bool = False) -> datetime:
    resultado = datetime.combine(fecha, hora)
    return resultado + timedelta(days=1 if siguiente_dia else 0)


def calcular_jornada(
    horario: HorarioLaboral,
    fecha: date | str,
    hora_entrada_real: time | str,
    hora_salida_real: time | str,
) -> ResultadoJornada:
    """Compara una marcacion completa contra su horario asignado."""
    if isinstance(fecha, str):
        fecha = date.fromisoformat(fecha[:10])

    entrada_programada = _como_hora(horario.entrada)
    salida_programada = _como_hora(horario.salida)
    entrada_real = _como_hora(hora_entrada_real)
    salida_real = _como_hora(hora_salida_real)

    cruza_medianoche = horario.cruza_medianoche or salida_programada <= entrada_programada
    inicio_programado = _momento(fecha, entrada_programada)
    fin_programado = _momento(fecha, salida_programada, cruza_medianoche)
    inicio_real = _momento(fecha, entrada_real)
    salida_real_es_dia_siguiente = cruza_medianoche or salida_real <= entrada_real
    fin_real = _momento(fecha, salida_real, salida_real_es_dia_siguiente)

    duracion = max(0.0, (fin_real - inicio_real).total_seconds() / 3600)
    horas_ordinarias = min(duracion, horario.maximo_diario)
    horas_extra = max(0.0, duracion - horario.maximo_diario)

    tolerancia_entrada = timedelta(minutes=horario.tolerancia_entrada_minutos)
    tardanza = max(0, int((inicio_real - inicio_programado - tolerancia_entrada).total_seconds() // 60))
    tolerancia_salida = timedelta(minutes=horario.tolerancia_salida_minutos)
    salida_anticipada = max(0, int((fin_programado - fin_real - tolerancia_salida).total_seconds() // 60))

    return ResultadoJornada(
        horas_trabajadas=round(duracion, 2),
        horas_ordinarias=round(horas_ordinarias, 2),
        horas_extra=round(horas_extra, 2),
        minutos_tardanza=tardanza,
        minutos_salida_anticipada=salida_anticipada,
        jornada_incompleta=duracion < horario.maximo_diario,
        cruza_medianoche=cruza_medianoche,
    )
