"""Cliente para autorizar una instalacion nueva mediante un codigo temporal."""

import getpass
import json
import os
import platform
import socket
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


RUTA_BASE = os.path.join(os.getenv("LOCALAPPDATA", os.path.expanduser("~")), "SiPP")
RUTA_IDENTIFICADOR = os.path.join(RUTA_BASE, "installation_id")
RUTA_AUTORIZACION = os.path.join(RUTA_BASE, "device_authorization.json")
RUTA_CONDICIONES = os.path.join(RUTA_BASE, "terms_acceptance.json")
URL_SERVIDOR = os.getenv("SIPP_AUTH_SERVER_URL", "https://autorizacion-1.onrender.com")
PROCESO_SERVIDOR_LOCAL = None


def obtener_api_key():
    api_key = os.getenv("SIPP_AUTH_API_KEY", "").strip()
    if api_key or os.name != "nt":
        return api_key
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as clave:
            api_key, _ = winreg.QueryValueEx(clave, "SIPP_AUTH_API_KEY")
        return str(api_key).strip()
    except (FileNotFoundError, OSError, TypeError):
        return ""


class ErrorAutorizacion(Exception):
    """Error controlado durante la autorizacion de la instalacion."""


def obtener_identificador_instalacion():
    try:
        with open(RUTA_IDENTIFICADOR, "r", encoding="ascii") as archivo:
            identificador = archivo.read().strip()
        uuid.UUID(identificador)
        return identificador
    except (OSError, ValueError):
        identificador = str(uuid.uuid4())
        os.makedirs(RUTA_BASE, exist_ok=True)
        temporal = RUTA_IDENTIFICADOR + ".tmp"
        with open(temporal, "w", encoding="ascii") as archivo:
            archivo.write(identificador)
        os.replace(temporal, RUTA_IDENTIFICADOR)
        return identificador


def obtener_datos_equipo():
    try:
        ip_local = socket.gethostbyname(socket.gethostname())
    except OSError:
        ip_local = "Desconocida"
    return {
        "installation_id": obtener_identificador_instalacion(),
        "device_name": socket.gethostname(),
        "windows_user": getpass.getuser(),
        "operating_system": platform.platform(),
        "ip_local": ip_local,
        "system_machine": platform.machine(),
        "system_release": platform.release(),
        "processor": platform.processor() or "Desconocido",
        "python_version": platform.python_version(),
        "app_version": "2.3.0",
    }


def cargar_autorizacion_local():
    try:
        with open(RUTA_AUTORIZACION, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
        return bool(
            datos.get("authorized")
            and datos.get("installation_id") == obtener_identificador_instalacion()
        )
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return False


def guardar_autorizacion_local():
    os.makedirs(RUTA_BASE, exist_ok=True)
    temporal = RUTA_AUTORIZACION + ".tmp"
    with open(temporal, "w", encoding="utf-8") as archivo:
        json.dump(
            {
                "authorized": True,
                "installation_id": obtener_identificador_instalacion(),
                "authorized_at": datetime.now(timezone.utc).isoformat(),
            },
            archivo,
        )
    os.replace(temporal, RUTA_AUTORIZACION)


def cargar_aceptacion_condiciones():
    try:
        with open(RUTA_CONDICIONES, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
        return datos.get("terms_version", "")
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return ""


def guardar_aceptacion_condiciones(version):
    os.makedirs(RUTA_BASE, exist_ok=True)
    temporal = RUTA_CONDICIONES + ".tmp"
    with open(temporal, "w", encoding="utf-8") as archivo:
        json.dump(
            {
                "terms_version": str(version),
                "installation_id": obtener_identificador_instalacion(),
                "accepted_at": datetime.now(timezone.utc).isoformat(),
            },
            archivo,
        )
    os.replace(temporal, RUTA_CONDICIONES)


class ClienteAutorizacion:
    def __init__(self, url_servidor=None, timeout=15):
        self.url_servidor = str(url_servidor or URL_SERVIDOR).strip().rstrip("/")
        self.timeout = timeout
        if not self.url_servidor.startswith("https://") and not self.url_servidor.startswith("http://localhost"):
            raise ErrorAutorizacion("Configure SIPP_AUTH_SERVER_URL con una URL HTTPS.")
        self._iniciar_servidor_local()

    def _iniciar_servidor_local(self):
        global PROCESO_SERVIDOR_LOCAL
        if not self.url_servidor.startswith("http://localhost"):
            return
        if PROCESO_SERVIDOR_LOCAL is None:
            self._detener_servidor_local_existente()
        elif PROCESO_SERVIDOR_LOCAL.poll() is not None:
            PROCESO_SERVIDOR_LOCAL = None
        try:
            with urlopen(f"{self.url_servidor}/health", timeout=1):
                if PROCESO_SERVIDOR_LOCAL is not None:
                    return
        except (HTTPError, URLError, TimeoutError, OSError):
            pass
        if PROCESO_SERVIDOR_LOCAL is None or PROCESO_SERVIDOR_LOCAL.poll() is not None:
            PROCESO_SERVIDOR_LOCAL = subprocess.Popen(
                [sys.executable, "-m", "uvicorn", "servidor_autorizacion.main:app", "--host", "127.0.0.1", "--port", "8000"],
                cwd=os.path.dirname(os.path.abspath(__file__)),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        for _ in range(40):
            try:
                with urlopen(f"{self.url_servidor}/health", timeout=1):
                    return
            except (HTTPError, URLError, TimeoutError, OSError):
                time.sleep(0.5)
        raise ErrorAutorizacion("No se pudo iniciar el servidor local de autorizacion.")

    @staticmethod
    def _detener_servidor_local_existente():
        if os.name == "nt":
            resultado = subprocess.run(
                ["netstat", "-ano", "-p", "tcp"],
                capture_output=True,
                text=True,
                check=False,
            )
            for linea in resultado.stdout.splitlines():
                partes = linea.split()
                if len(partes) >= 5 and partes[1].endswith(":8000") and partes[3] == "LISTENING":
                    subprocess.run(["taskkill", "/PID", partes[4], "/F"], capture_output=True, check=False)
        else:
            subprocess.run(["fuser", "-k", "8000/tcp"], capture_output=True, check=False)

    def _solicitud(self, metodo, ruta, datos=None):
        cuerpo = None
        cabeceras = {"Accept": "application/json"}
        api_key = obtener_api_key()
        if api_key:
            cabeceras["X-API-Key"] = api_key
        if datos is not None:
            cuerpo = json.dumps(datos).encode("utf-8")
            cabeceras["Content-Type"] = "application/json"
        solicitud = Request(
            f"{self.url_servidor}/{ruta.lstrip('/')}",
            data=cuerpo,
            headers=cabeceras,
            method=metodo,
        )
        try:
            with urlopen(solicitud, timeout=self.timeout) as respuesta:
                resultado = json.loads(respuesta.read().decode("utf-8"))
        except HTTPError as exc:
            try:
                detalle = json.loads(exc.read().decode("utf-8")).get("detail", "")
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                detalle = ""
            mensaje = f"El servidor devolvio HTTP {exc.code}"
            if detalle:
                mensaje += f": {detalle}"
            raise ErrorAutorizacion(mensaje) from exc
        except (URLError, TimeoutError, OSError, ValueError) as exc:
            raise ErrorAutorizacion(f"No se pudo contactar el servidor: {exc}") from exc
        if not isinstance(resultado, dict):
            raise ErrorAutorizacion("El servidor devolvio una respuesta invalida.")
        return resultado

    def solicitar_codigo(self):
        respuesta = self._solicitud(
            "POST",
            "v1/access-requests",
            {**obtener_datos_equipo(), "force_new": True},
        )
        request_id = str(respuesta.get("request_id", "")).strip()
        if not request_id:
            raise ErrorAutorizacion("El servidor no devolvio una solicitud valida.")
        if respuesta.get("code"):
            respuesta["message"] = f"Codigo local: {respuesta['code']}"
        return respuesta

    def validar_codigo(self, request_id, codigo):
        return self._solicitud(
            "POST",
            f"v1/access-requests/{request_id}/verify",
            {
                "installation_id": obtener_identificador_instalacion(),
                "code": str(codigo).strip(),
            },
        )

    def registrar_aceptacion_condiciones(self, version, texto):
        return self._solicitud(
            "POST",
            "v1/terms-acceptances",
            {
                **obtener_datos_equipo(),
                "terms_version": str(version),
                "terms_text": str(texto),
            },
        )