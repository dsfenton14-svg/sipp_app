"""Configuracion sensible obtenida desde variables de entorno."""

import os
import ctypes
import json
from ctypes import wintypes


DB_HOST = os.getenv("SIPP_DB_HOST", "localhost")
DB_PORT = os.getenv("SIPP_DB_PORT", "5432")
DB_USER = os.getenv("SIPP_DB_USER", "postgres")
DB_PASSWORD = os.getenv("SIPP_DB_PASSWORD")
DB_NAME = os.getenv("SIPP_DB_NAME", "SiPP-2")
DB_SSLMODE = os.getenv("SIPP_DB_SSLMODE", "require")
RUTA_CREDENCIALES = os.path.join(os.getenv("LOCALAPPDATA", os.path.expanduser("~")), "SiPP", "postgres_credentials.bin")


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.c_void_p)]


def _proteger_datos(datos):
    entrada = _DataBlob(len(datos), ctypes.cast(ctypes.create_string_buffer(datos), ctypes.c_void_p))
    salida = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(salida)):
        raise OSError("No se pudieron proteger las credenciales.")
    try:
        return ctypes.string_at(salida.pbData, salida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(ctypes.c_void_p(salida.pbData))


def _desproteger_datos(datos):
    entrada = _DataBlob(len(datos), ctypes.cast(ctypes.create_string_buffer(datos), ctypes.c_void_p))
    salida = _DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(salida)):
        raise OSError("No se pudieron recuperar las credenciales.")
    try:
        return ctypes.string_at(salida.pbData, salida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(ctypes.c_void_p(salida.pbData))


def cargar_credenciales():
    try:
        with open(RUTA_CREDENCIALES, "rb") as archivo:
            datos = json.loads(_desproteger_datos(archivo.read()).decode("utf-8"))
        if datos.get("usuario") and datos.get("contraseña"):
            datos.setdefault("host", DB_HOST)
            datos.setdefault("puerto", DB_PORT)
            datos.setdefault("base_datos", DB_NAME)
            return datos
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        pass
    return None


def guardar_credenciales(usuario, contraseña, host=None, puerto=None, base_datos=None):
    carpeta = os.path.dirname(RUTA_CREDENCIALES)
    os.makedirs(carpeta, exist_ok=True)
    contenido = json.dumps(
        {
            "usuario": usuario,
            "contraseña": contraseña,
            "host": host or DB_HOST,
            "puerto": str(puerto or DB_PORT),
            "base_datos": base_datos or DB_NAME,
        }
    ).encode("utf-8")
    temporal = RUTA_CREDENCIALES + ".tmp"
    with open(temporal, "wb") as archivo:
        archivo.write(_proteger_datos(contenido))
    os.replace(temporal, RUTA_CREDENCIALES)


def exigir_password_db():
    if not DB_PASSWORD:
        raise RuntimeError("Falta configurar la variable de entorno SIPP_DB_PASSWORD.")
    return DB_PASSWORD


def validar_usuario_db(usuario):
    return str(usuario).strip()
