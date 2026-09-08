"""Conectores genericos para importar marcaciones de lectores biometrico.

Los adaptadores entregan siempre registros normalizados. La identidad interna
que usa SiPP es el DUI; el identificador del lector solo sirve como referencia
cuando el dispositivo no envia el DUI directamente.
"""

from __future__ import annotations

import csv
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


@dataclass
class MarcacionLector:
    identificador: str
    fecha: str
    hora: str
    tipo: str = ""


class ErrorConector(Exception):
    """Error controlado producido por un conector."""


class ConectorLector:
    def probar_conexion(self) -> str:
        raise NotImplementedError

    def obtener_marcaciones(self) -> list[MarcacionLector]:
        raise NotImplementedError


class ConectorArchivo(ConectorLector):
    def __init__(self, ruta: str, formato: str, columna_id: str, columna_fecha: str, columna_hora: str, columna_tipo: str = ""):
        self.ruta = Path(ruta).expanduser()
        self.formato = formato
        self.columna_id = columna_id.strip()
        self.columna_fecha = columna_fecha.strip()
        self.columna_hora = columna_hora.strip()
        self.columna_tipo = columna_tipo.strip()

    def _leer_registros(self) -> list[dict[str, Any]]:
        if not self.ruta.is_file():
            raise ErrorConector(f"No existe el archivo: {self.ruta}")
        try:
            if self.formato == "CSV":
                with self.ruta.open("r", encoding="utf-8-sig", newline="") as archivo:
                    return list(csv.DictReader(archivo))
            with self.ruta.open("r", encoding="utf-8-sig") as archivo:
                datos = json.load(archivo)
        except (OSError, UnicodeError, json.JSONDecodeError, csv.Error) as exc:
            raise ErrorConector(f"No se pudo leer el archivo: {exc}") from exc

        if isinstance(datos, dict):
            datos = datos.get("data", datos.get("records", datos.get("marcaciones", [])))
        if not isinstance(datos, list) or not all(isinstance(item, dict) for item in datos):
            raise ErrorConector("El archivo debe contener una lista de registros.")
        return datos

    def probar_conexion(self) -> str:
        registros = self._leer_registros()
        return f"Archivo accesible. Registros detectados: {len(registros)}"

    def obtener_marcaciones(self) -> list[MarcacionLector]:
        registros = self._leer_registros()
        resultado = []
        for registro in registros:
            try:
                identificador = str(registro[self.columna_id]).strip()
                fecha = _normalizar_fecha(registro[self.columna_fecha])
                hora = _normalizar_hora(registro[self.columna_hora])
            except (KeyError, TypeError, ValueError) as exc:
                raise ErrorConector(f"Columnas invalidas o ausentes: {exc}") from exc
            if not identificador or not fecha or not hora:
                continue
            tipo = str(registro.get(self.columna_tipo, "") or "").strip() if self.columna_tipo else ""
            resultado.append(MarcacionLector(identificador, fecha, hora, tipo))
        return resultado


class ConectorRest(ConectorLector):
    def __init__(self, url: str, token: str = "", columna_id: str = "dui", columna_fecha: str = "fecha", columna_hora: str = "hora", columna_tipo: str = ""):
        self.url = url.strip()
        self.token = token.strip()
        self.columna_id = columna_id.strip()
        self.columna_fecha = columna_fecha.strip()
        self.columna_hora = columna_hora.strip()
        self.columna_tipo = columna_tipo.strip()

    def _solicitar(self) -> Any:
        if not self.url:
            raise ErrorConector("Debe indicar la URL del servicio REST.")
        headers = {"Accept": "application/json"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        try:
            request = Request(self.url, headers=headers)
            with urlopen(request, timeout=10) as respuesta:
                return json.loads(respuesta.read().decode("utf-8"))
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError, OSError) as exc:
            raise ErrorConector(f"No se pudo consultar el servicio REST: {exc}") from exc

    def _registros(self) -> list[dict[str, Any]]:
        datos = self._solicitar()
        if isinstance(datos, dict):
            datos = datos.get("data", datos.get("records", datos.get("marcaciones", [])))
        if not isinstance(datos, list) or not all(isinstance(item, dict) for item in datos):
            raise ErrorConector("La respuesta REST no contiene una lista de registros.")
        return datos

    def probar_conexion(self) -> str:
        return f"Servicio REST accesible. Registros recibidos: {len(self._registros())}"

    def obtener_marcaciones(self) -> list[MarcacionLector]:
        return _normalizar_registros(self._registros(), self.columna_id, self.columna_fecha, self.columna_hora, self.columna_tipo)


class ConectorSQLite(ConectorLector):
    def __init__(self, ruta: str, consulta: str, columna_id: str, columna_fecha: str, columna_hora: str, columna_tipo: str = ""):
        self.ruta = Path(ruta).expanduser()
        self.consulta = consulta.strip()
        self.columna_id = columna_id.strip()
        self.columna_fecha = columna_fecha.strip()
        self.columna_hora = columna_hora.strip()
        self.columna_tipo = columna_tipo.strip()

    def _registros(self) -> list[dict[str, Any]]:
        if not self.ruta.is_file():
            raise ErrorConector(f"No existe la base SQLite: {self.ruta}")
        if not self.consulta:
            raise ErrorConector("Debe indicar una consulta SQL de lectura.")
        try:
            with sqlite3.connect(self.ruta) as conexion:
                conexion.row_factory = sqlite3.Row
                return [dict(fila) for fila in conexion.execute(self.consulta).fetchall()]
        except (sqlite3.Error, OSError) as exc:
            raise ErrorConector(f"No se pudo consultar SQLite: {exc}") from exc

    def probar_conexion(self) -> str:
        return f"Base SQLite accesible. Registros recibidos: {len(self._registros())}"

    def obtener_marcaciones(self) -> list[MarcacionLector]:
        return _normalizar_registros(self._registros(), self.columna_id, self.columna_fecha, self.columna_hora, self.columna_tipo)


def _normalizar_registros(registros, columna_id, columna_fecha, columna_hora, columna_tipo):
    resultado = []
    for registro in registros:
        try:
            identificador = str(registro[columna_id]).strip()
            fecha = _normalizar_fecha(registro[columna_fecha])
            hora = _normalizar_hora(registro[columna_hora])
        except (KeyError, TypeError, ValueError) as exc:
            raise ErrorConector(f"Columnas invalidas o ausentes: {exc}") from exc
        if identificador and fecha and hora:
            tipo = str(registro.get(columna_tipo, "") or "").strip() if columna_tipo else ""
            resultado.append(MarcacionLector(identificador, fecha, hora, tipo))
    return resultado


def _normalizar_fecha(valor: Any) -> str:
    if hasattr(valor, "strftime"):
        return valor.strftime("%Y-%m-%d")
    texto = str(valor or "").strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto[:10], formato).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def _normalizar_hora(valor: Any) -> str:
    texto = str(valor or "").strip().upper()
    for formato in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"):
        try:
            return datetime.strptime(texto, formato).strftime("%I:%M:%S %p")
        except ValueError:
            continue
    return ""


def crear_conector(tipo: str, **configuracion) -> ConectorLector:
    if tipo in {"CSV", "JSON"}:
        return ConectorArchivo(formato=tipo, **configuracion)
    if tipo == "REST":
        return ConectorRest(**configuracion)
    if tipo == "SQLite":
        return ConectorSQLite(**configuracion)
    raise ErrorConector(f"Tipo de conexion no soportado: {tipo}")
