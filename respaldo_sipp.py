"""Copias de seguridad y restauración de la base de datos SiPP mediante pg_dump/pg_restore."""

import glob
import os
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from config_seguridad import DB_HOST, DB_PORT, DB_USER, DB_NAME, exigir_password_db

HOST = DB_HOST
USUARIO = DB_USER
PUERTO = DB_PORT
BASE_DATOS = DB_NAME
TIEMPO_MAXIMO_SUBPROCESO = 120

CARPETA_RESPALDOS = os.path.join(os.path.abspath("."), "respaldos")
CARPETA_UPLOADS = os.path.join(os.path.abspath("."), "uploads")


def _entorno_con_clave():
    entorno = os.environ.copy()
    entorno["PGPASSWORD"] = exigir_password_db()
    return entorno


def _ubicar_binario(nombre):
    """Busca pg_dump/pg_restore en el PATH o en las instalaciones típicas de Windows."""
    encontrado = shutil.which(nombre)
    if encontrado:
        return encontrado
    patrones = glob.glob(rf"C:\Program Files\PostgreSQL\*\bin\{nombre}.exe")
    if patrones:
        return sorted(patrones, reverse=True)[0]
    raise FileNotFoundError(
        f"No se encontró {nombre}. Instala las herramientas cliente de PostgreSQL o agrégalas al PATH."
    )


def crear_copia_seguridad(carpeta_destino=CARPETA_RESPALDOS):
    """Genera un .zip con el dump de la base de datos y los archivos de uploads/."""
    os.makedirs(carpeta_destino, exist_ok=True)
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_zip = os.path.join(carpeta_destino, f"sipp_backup_{marca}.zip")
    pg_dump = _ubicar_binario("pg_dump")
    with tempfile.TemporaryDirectory() as temporal:
        ruta_dump = os.path.join(temporal, "base_datos.dump")
        comando = [pg_dump, "-h", HOST, "-p", PUERTO, "-U", USUARIO, "-Fc", "-f", ruta_dump, BASE_DATOS]
        resultado = subprocess.run(
            comando,
            env=_entorno_con_clave(),
            capture_output=True,
            text=True,
            timeout=TIEMPO_MAXIMO_SUBPROCESO,
        )
        if resultado.returncode != 0:
            raise RuntimeError(f"pg_dump falló: {resultado.stderr.strip()}")
        with zipfile.ZipFile(nombre_zip, "w", zipfile.ZIP_DEFLATED) as zip_archivo:
            zip_archivo.write(ruta_dump, arcname="base_datos.dump")
            if os.path.isdir(CARPETA_UPLOADS):
                for raiz, _, archivos in os.walk(CARPETA_UPLOADS):
                    for archivo in archivos:
                        ruta_completa = os.path.join(raiz, archivo)
                        ruta_relativa = os.path.join("uploads", os.path.relpath(ruta_completa, CARPETA_UPLOADS))
                        zip_archivo.write(ruta_completa, arcname=ruta_relativa)
    return nombre_zip


def listar_copias(carpeta=CARPETA_RESPALDOS):
    """Devuelve las rutas de las copias existentes, de la más reciente a la más antigua."""
    if not os.path.isdir(carpeta):
        return []
    copias = glob.glob(os.path.join(carpeta, "sipp_backup_*.zip"))
    return sorted(copias, key=os.path.getmtime, reverse=True)


def obtener_ultima_copia(carpeta=CARPETA_RESPALDOS):
    copias = listar_copias(carpeta)
    return copias[0] if copias else None


def restaurar_copia_seguridad(ruta_zip):
    """Restaura la base de datos y los uploads desde un .zip generado por crear_copia_seguridad."""
    if not os.path.isfile(ruta_zip):
        raise FileNotFoundError(f"No existe el archivo de respaldo: {ruta_zip}")
    pg_restore = _ubicar_binario("pg_restore")
    with tempfile.TemporaryDirectory() as temporal:
        with zipfile.ZipFile(ruta_zip, "r") as zip_archivo:
            raiz_segura = os.path.abspath(temporal)
            for miembro in zip_archivo.infolist():
                destino = os.path.abspath(os.path.join(temporal, miembro.filename))
                if os.path.commonpath((raiz_segura, destino)) != raiz_segura:
                    raise ValueError("El respaldo contiene una ruta no segura.")
            zip_archivo.extractall(temporal)
        ruta_dump = os.path.join(temporal, "base_datos.dump")
        if not os.path.isfile(ruta_dump):
            raise ValueError("El archivo de respaldo no contiene un dump válido.")
        comando = [
            pg_restore, "-h", HOST, "-p", PUERTO, "-U", USUARIO,
            "-d", BASE_DATOS, "--clean", "--if-exists", "--no-owner", ruta_dump,
        ]
        resultado = subprocess.run(
            comando,
            env=_entorno_con_clave(),
            capture_output=True,
            text=True,
            timeout=TIEMPO_MAXIMO_SUBPROCESO,
        )
        if resultado.returncode != 0:
            raise RuntimeError(f"pg_restore falló: {resultado.stderr.strip()}")
        carpeta_uploads_zip = os.path.join(temporal, "uploads")
        if os.path.isdir(carpeta_uploads_zip):
            os.makedirs(CARPETA_UPLOADS, exist_ok=True)
            for raiz, _, archivos in os.walk(carpeta_uploads_zip):
                for archivo in archivos:
                    origen = os.path.join(raiz, archivo)
                    destino = os.path.join(CARPETA_UPLOADS, os.path.relpath(origen, carpeta_uploads_zip))
                    os.makedirs(os.path.dirname(destino), exist_ok=True)
                    shutil.copy2(origen, destino)
    return True
