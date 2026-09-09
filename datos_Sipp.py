from tkinter import messagebox
import os

import psycopg as pc
import random
import threading
import hashlib
import hmac
import secrets
from datetime import date, datetime, timedelta, timezone
from config_seguridad import DB_HOST, DB_PORT, DB_USER, DB_NAME, DB_SSLMODE, cargar_credenciales, exigir_password_db, validar_usuario_db


_conexion_compartida = None
_bloqueo_conexion = threading.Lock()
_tablas_inicializadas = set()
_bloqueo_tablas = threading.Lock()

_OPCIONES_POSTGRES = "-c statement_timeout=30000 -c lock_timeout=10000"


def parametros_conexion():
    credenciales = cargar_credenciales()
    if credenciales and DB_HOST in ("localhost", "127.0.0.1"):
        return credenciales["usuario"], credenciales["contraseña"], "disable"
    if credenciales and not os.getenv("SIPP_DB_PASSWORD"):
        return credenciales["usuario"], credenciales["contraseña"], DB_SSLMODE
    return validar_usuario_db(DB_USER), exigir_password_db(), DB_SSLMODE


def generar_hash_contraseña(contraseña, sal=None):
    sal = sal or secrets.token_bytes(16)
    hash_contraseña = hashlib.scrypt(
        str(contraseña).encode("utf-8"),
        salt=sal,
        n=32768,
        r=8,
        p=1,
        maxmem=64 * 1024 * 1024,
    )
    return sal.hex(), hash_contraseña.hex()


def verificar_hash_contraseña(contraseña, sal_hex, hash_hex):
    try:
        sal = bytes.fromhex(sal_hex)
        esperado = bytes.fromhex(hash_hex)
        actual = hashlib.scrypt(
            str(contraseña).encode("utf-8"),
            salt=sal,
            n=32768,
            r=8,
            p=1,
            maxmem=64 * 1024 * 1024,
        )
        return hmac.compare_digest(actual, esperado)
    except (TypeError, ValueError):
        return False


def obtener_conexion():
    """Devuelve una sola conexión reutilizable para los servicios de SiPP."""
    global _conexion_compartida
    with _bloqueo_conexion:
        if _conexion_compartida is None or _conexion_compartida.closed:
            usuario, contraseña, sslmode = parametros_conexion()
            _conexion_compartida = pc.connect(
                host=DB_HOST,
                user=usuario,
                password=contraseña,
                port=DB_PORT,
                dbname=DB_NAME,
                sslmode=sslmode,
                connect_timeout=5,
                options=_OPCIONES_POSTGRES,
            )
        return _conexion_compartida


def inicializar_tabla(clave, crear_tabla):
    """Ejecuta migraciones de un servicio solo la primera vez que se usa."""
    with _bloqueo_tablas:
        if clave in _tablas_inicializadas:
            return
        crear_tabla()
        _tablas_inicializadas.add(clave)


def limpiar_datos_operativos():
    """Vacía los datos de trabajo y conserva acceso y configuración del sistema."""
    tablas_operativas = (
        "aprobaciones_usuario_corte",
        "contratos_trabajo",
        "cortes_planilla",
        "decimoxiii",
        "descuentos",
        "empleados",
        "empresa",
        "justificaciones_marcacion",
        "liquidaciones",
        "marcaciones",
        "periodos_vacaciones",
        "permisos_ausencias",
        "planilla",
        "remuneraciones",
        "solicitudes_vacaciones",
        "vacaciones",
    )
    conexion = obtener_conexion()
    with conexion.cursor() as cursor:
        cursor.execute(
            "TRUNCATE TABLE "
            + ", ".join(tablas_operativas)
            + " RESTART IDENTITY CASCADE"
        )
    conexion.commit()
    return len(tablas_operativas)


class basedatos:
    def __init__(self):
        self.crear_base_datos()

    def crear_base_datos(self):
        """Crea la base de datos 'SiPP-2' si no existe, de lo contrario la usa."""
        try:
            usuario, contraseña, sslmode = parametros_conexion()
            # Conectar a la base de datos por defecto (postgres)
            conexion_admin = pc.connect(
                host=DB_HOST,
                user=usuario,
                password=contraseña,
                port=DB_PORT,
                dbname="postgres",
                sslmode=sslmode,
                connect_timeout=5,
                options=_OPCIONES_POSTGRES,
            )
            conexion_admin.autocommit = True
            
            # Verificar si la base de datos ya existe
            with conexion_admin.cursor() as cursor:
                cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (DB_NAME,))
                existe = cursor.fetchone()
                
                if not existe:
                    # Crear la base de datos si no existe
                    cursor.execute("CREATE DATABASE " + pc.sql.Identifier(DB_NAME).as_string(cursor))
                    messagebox.showinfo("Bienvenido a SiPP", "Inicio con exito")
                else:
                    pass
            
            conexion_admin.close()
            return True
            
        except Exception as e:
            messagebox.showerror(
                "Sin conexion",
                "No se pudo conectar a la base de datos de SiPP.\n"
                "Revise su conexion a internet o la red del servidor e intentelo nuevamente.",
            )
            return False



def generar_codigo_8_digitos():
    return random.randint(10000000, 99999999)


def _usuario__adm():
    usuario = "Dariel Fenton"
    return usuario


def _clave__de__acceso():
    contrase = "dariel"
    return contrase



class ConexionDB_datos_usuarios:

    
    def __init__(self):
        self.conexion = obtener_conexion()
        self.crear_tabla_usuarios()
        

    def crear_tabla_usuarios(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS usuarios (
                    nombre text, genero text, contraseña text, codigo int)""")
            self.conexion.commit()

            cursor.execute("""
                DO $$
                BEGIN
                    IF EXISTS (
                        SELECT 1
                        FROM information_schema.columns
                        WHERE table_name = 'usuarios'
                          AND column_name = 'contrasena'
                    )
                    AND NOT EXISTS (
                        SELECT 1
                        FROM information_schema.columns
                        WHERE table_name = 'usuarios'
                          AND column_name = 'contraseña'
                    ) THEN
                        ALTER TABLE usuarios RENAME COLUMN contrasena TO contraseña;
                    END IF;
                END
                $$;
            """)
            self.conexion.commit()
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS sal TEXT")
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS hash_contraseña TEXT")
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS rol TEXT NOT NULL DEFAULT 'CONSULTA'")
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS intentos_fallidos INTEGER NOT NULL DEFAULT 0")
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS bloqueado_hasta TIMESTAMPTZ")
            cursor.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS ultimo_acceso TIMESTAMPTZ")
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS auditoria_seguridad(
                    id SERIAL PRIMARY KEY,
                    usuario TEXT NOT NULL,
                    accion TEXT NOT NULL,
                    detalle TEXT,
                    creado_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            """)
            self.conexion.commit()
            cursor.close()
    
    def insertar_usuario(self, nombre, genero, contraseña, codigo):
        sal, hash_contraseña = generar_hash_contraseña(contraseña)
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM usuarios WHERE COALESCE(rol, 'CONSULTA') = 'CONSULTA'")
            if cursor.fetchone()[0] >= 3:
                raise ValueError("Ya existe el máximo de 3 usuarios normales.")
            cursor.execute("""
                INSERT INTO usuarios (nombre, genero, contraseña, codigo, sal, hash_contraseña, rol)
                VALUES (%s, %s, NULL, %s, %s, %s, 'CONSULTA')""",
                (nombre, genero, codigo, sal, hash_contraseña))
            self.conexion.commit()
            cursor.close()

    def insertar_administrador(self, nombre, contraseña, codigo=0):
        sal, hash_contraseña = generar_hash_contraseña(contraseña)
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM usuarios WHERE rol = 'ADMIN'")
            if cursor.fetchone()[0] >= 3:
                raise ValueError("Ya existe el máximo de 3 administradores.")
            cursor.execute("""
                INSERT INTO usuarios (nombre, genero, contraseña, codigo, sal, hash_contraseña, rol)
                VALUES (%s, %s, NULL, %s, %s, %s, 'ADMIN')
                ON CONFLICT DO NOTHING
            """, (nombre, "", codigo, sal, hash_contraseña))
            if cursor.rowcount == 0:
                cursor.execute(
                    """
                    UPDATE usuarios
                    SET contraseña = NULL, sal = %s, hash_contraseña = %s, rol = 'ADMIN'
                    WHERE nombre = %s
                    """,
                    (sal, hash_contraseña, nombre),
                )
            self.conexion.commit()

    def autenticar(self, nombre, contraseña):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT nombre, genero, contraseña, sal, hash_contraseña, rol,
                       intentos_fallidos, bloqueado_hasta
                FROM usuarios WHERE nombre = %s
                """,
                (nombre,),
            )
            usuario = cursor.fetchone()
            if not usuario:
                return None
            nombre_db, genero, contraseña_antigua, sal, hash_contraseña, rol, intentos, bloqueado_hasta = usuario
            ahora = datetime.now(timezone.utc)
            if bloqueado_hasta and bloqueado_hasta > ahora:
                self.registrar_auditoria(nombre_db, "ACCESO_BLOQUEADO", "Cuenta temporalmente bloqueada")
                return None
            valido = verificar_hash_contraseña(contraseña, sal, hash_contraseña)
            if not valido and contraseña_antigua and hmac.compare_digest(str(contraseña_antigua), str(contraseña)):
                sal, hash_contraseña = generar_hash_contraseña(contraseña)
                cursor.execute(
                    "UPDATE usuarios SET contraseña = NULL, sal = %s, hash_contraseña = %s WHERE nombre = %s",
                    (sal, hash_contraseña, nombre_db),
                )
                self.conexion.commit()
                valido = True
            if not valido:
                nuevos_intentos = int(intentos or 0) + 1
                bloqueo = ahora + timedelta(minutes=10) if nuevos_intentos >= 5 else None
                cursor.execute(
                    "UPDATE usuarios SET intentos_fallidos = %s, bloqueado_hasta = %s WHERE nombre = %s",
                    (0 if bloqueo else nuevos_intentos, bloqueo, nombre_db),
                )
                self.conexion.commit()
                self.registrar_auditoria(nombre_db, "ACCESO_RECHAZADO", "Credenciales inválidas")
                return None
            cursor.execute(
                "UPDATE usuarios SET intentos_fallidos = 0, bloqueado_hasta = NULL, ultimo_acceso = %s WHERE nombre = %s",
                (ahora, nombre_db),
            )
            self.conexion.commit()
            self.registrar_auditoria(nombre_db, "INICIO_SESION", "Acceso autenticado")
            return {"nombre": nombre_db, "genero": genero, "rol": rol or "CONSULTA"}

    def registrar_auditoria(self, usuario, accion, detalle=""):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                "INSERT INTO auditoria_seguridad (usuario, accion, detalle) VALUES (%s, %s, %s)",
                (str(usuario or "desconocido"), str(accion), str(detalle or "")),
            )
        self.conexion.commit()
    
    def obtener_usuarios(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM usuarios")
            usuarios = cursor.fetchall()
            cursor.close()
            return usuarios

    def obtener_administradores(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT nombre FROM usuarios WHERE rol = 'ADMIN' ORDER BY nombre")
            administradores = [fila[0] for fila in cursor.fetchall()]
        return administradores

    def eliminar_administrador(self, nombre):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                DELETE FROM usuarios
                WHERE nombre = %s
                  AND LOWER(nombre) <> LOWER('@@22')
                  AND rol = 'ADMIN'
                """,
                (nombre,),
            )
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
        return eliminado
        
    def eliminar_usuario(self, nombre):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                DELETE FROM usuarios
                WHERE nombre = %s
                  AND NOT (
                      rol = 'ADMIN'
                      AND (SELECT COUNT(*) FROM usuarios WHERE rol = 'ADMIN') <= 1
                  )
                """,
                (nombre,),
            )
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return eliminado

    def limpiar_tabla_usuarios(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM usuarios")
            self.conexion.commit()
            cursor.close()

    def eliminar_tabla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute(f"DROP TABLE IF EXISTS usuarios")
            self.conexion.commit()
            cursor.close()


class estado_remuneracion:
    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("estados_remuneracion", self.crear_tabla_estados_remuneracion)

    def crear_tabla_estados_remuneracion(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS estados_remuneracion(
                    remuneracion text PRIMARY KEY,
                    estados text)""")
            self.conexion.commit()
            cursor.close()
            self.datos_remuneracion()

    def insertar_estados_remuneracion(self, datos):
        with self.conexion.cursor() as cursor:
            cursor.executemany("""
        INSERT INTO estados_remuneracion (remuneracion, estados)
        VALUES (%s, %s)
    """, datos)
            self.conexion.commit()
            cursor.close()

    def datos_remuneracion(self):
        obtener = self.obtener_estados_remuneracion()
        if not obtener:
            datos = [
        ('Horas Extras', 'desactivado'),
        ('Comisiones', 'desactivado'),
        ('Bonificaciones', 'desactivado'),
        ('Otros', 'desactivado')
    ]
            self.insertar_estados_remuneracion(datos)
        else:
            pass
            
    def datos_nombre(self, nombre):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
        SELECT remuneracion, estados
        FROM estados_remuneracion
        WHERE remuneracion = %s
    """, (nombre,))

            dato = cursor.fetchone()

            return dato
            

    def limpiar(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM estados_remuneracion")
            self.conexion.commit()
            cursor.close()
        


    def obtener_estados_remuneracion(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM estados_remuneracion")
            usuarios = cursor.fetchall()
            cursor.close()
            return usuarios
        

    def actualizar_estados_remuneracion_nombre(self, remuneracion, estados):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """UPDATE estados_remuneracion SET estados = %s  WHERE remuneracion = %s """,
                (estados, remuneracion)
            )
            self.conexion.commit()
            cursor.close()


class registro_remuneracion:
    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("remuneraciones", self.crear_tabla_remuneraciones)

    def crear_tabla_remuneraciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS remuneraciones(
                    id SERIAL PRIMARY KEY,
                    nombre TEXT,
                    tipo TEXT,
                    monto NUMERIC,
                    fecha TIMESTAMP,
                    motivo TEXT,
                    dui TEXT,
                    estado TEXT NOT NULL DEFAULT 'ACTIVO'
                )""")
            cursor.execute("ALTER TABLE remuneraciones ADD COLUMN IF NOT EXISTS dui TEXT")
            cursor.execute("ALTER TABLE remuneraciones ADD COLUMN IF NOT EXISTS estado TEXT NOT NULL DEFAULT 'ACTIVO'")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_remuneraciones_dui ON remuneraciones(dui)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_remuneraciones_fecha ON remuneraciones(fecha)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_remuneraciones_estado ON remuneraciones(estado)")
            self.conexion.commit()
            cursor.close()

    def insertar_remuneracion(self, nombre, tipo, monto, fecha, motivo, dui=None):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """INSERT INTO remuneraciones (nombre, tipo, monto, fecha, motivo, dui)
                   VALUES (%s, %s, %s, %s, %s, %s)""",
                (nombre, tipo, monto, fecha, motivo, dui)
            )
            self.conexion.commit()
            cursor.close()

    def existe_remuneracion(self, dui, tipo, monto, fecha, excluir_id=None):
        consulta = """SELECT id FROM remuneraciones
                     WHERE dui = %s AND tipo = %s AND monto = %s
                       AND fecha::date = %s AND estado = 'ACTIVO'"""
        parametros = [dui, tipo, monto, fecha]
        if excluir_id is not None:
            consulta += " AND id <> %s"
            parametros.append(excluir_id)
        consulta += " LIMIT 1"
        with self.conexion.cursor() as cursor:
            cursor.execute(consulta, parametros)
            return cursor.fetchone() is not None

    def obtener_remuneraciones_detalladas(self, dui=None, fecha_inicio=None, fecha_fin=None):
        consulta = """SELECT id, nombre, dui, tipo, monto, fecha, motivo, estado
                     FROM remuneraciones WHERE estado = 'ACTIVO'"""
        parametros = []
        if dui:
            consulta += " AND dui = %s"
            parametros.append(dui)
        if fecha_inicio is not None:
            consulta += " AND fecha::date >= %s"
            parametros.append(fecha_inicio)
        if fecha_fin is not None:
            consulta += " AND fecha::date <= %s"
            parametros.append(fecha_fin)
        consulta += " ORDER BY fecha DESC"
        with self.conexion.cursor() as cursor:
            cursor.execute(consulta, parametros)
            return cursor.fetchall()

    def actualizar_remuneracion(self, id_remuneracion, nombre, dui, tipo, monto, fecha, motivo):
        if self.existe_remuneracion(dui, tipo, monto, fecha, excluir_id=id_remuneracion):
            return False
        with self.conexion.cursor() as cursor:
            cursor.execute("""UPDATE remuneraciones
                            SET nombre = %s, dui = %s, tipo = %s, monto = %s,
                                fecha = %s, motivo = %s
                            WHERE id = %s AND estado = 'ACTIVO'
                            RETURNING id""", (nombre, dui, tipo, monto, fecha, motivo, id_remuneracion))
            actualizado = cursor.fetchone() is not None
        self.conexion.commit()
        return actualizado

    def anular_remuneracion(self, id_remuneracion):
        with self.conexion.cursor() as cursor:
            cursor.execute("""UPDATE remuneraciones SET estado = 'ANULADO'
                            WHERE id = %s AND estado = 'ACTIVO'
                            RETURNING id""", (id_remuneracion,))
            anulado = cursor.fetchone() is not None
        self.conexion.commit()
        return anulado

    def obtener_remuneraciones(self, limit=None):
        with self.conexion.cursor() as cursor:
            query = "SELECT id, nombre, tipo, monto, fecha, motivo FROM remuneraciones WHERE estado = 'ACTIVO' ORDER BY fecha DESC"
            if limit is not None:
                query += " LIMIT %s"
                cursor.execute(query, (limit,))
            else:
                cursor.execute(query)
            registros = cursor.fetchall()
            cursor.close()
            return registros

    def cantidad_remuneraciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM remuneraciones WHERE estado = 'ACTIVO'")
            cantidad = cursor.fetchone()[0]
            cursor.close()
            return cantidad

    def total_importe(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COALESCE(SUM(monto), 0) FROM remuneraciones WHERE estado = 'ACTIVO'")
            total = cursor.fetchone()[0]
            cursor.close()
            return total



class empresa:
    """Administra los datos de la empresa usados en documentos y reportes."""

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("empresa", self.crear_tabla_empresa)

    def crear_tabla_empresa(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS empresa(
                    id INTEGER PRIMARY KEY DEFAULT 1 CHECK (id = 1),
                    nombre TEXT NOT NULL DEFAULT '',
                    ruc TEXT NOT NULL DEFAULT '',
                    direccion TEXT NOT NULL DEFAULT '',
                    telefono TEXT NOT NULL DEFAULT '',
                    logo TEXT NOT NULL DEFAULT ''
                )
            """)
            self.conexion.commit()

    def obtener_empresa(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, nombre, ruc, direccion, telefono, logo
                FROM empresa
                WHERE id = 1
            """)
            return cursor.fetchone()

    def guardar_empresa(self, nombre, ruc, direccion, telefono, logo=""):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO empresa(id, nombre, ruc, direccion, telefono, logo)
                VALUES (1, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO UPDATE SET
                    nombre = EXCLUDED.nombre,
                    ruc = EXCLUDED.ruc,
                    direccion = EXCLUDED.direccion,
                    telefono = EXCLUDED.telefono,
                    logo = EXCLUDED.logo
            """, (nombre, ruc, direccion, telefono, logo))
            self.conexion.commit()
            return True

    def actualizar_empresa(self, nombre, ruc, direccion, telefono, logo=""):
        return self.guardar_empresa(nombre, ruc, direccion, telefono, logo)

    def eliminar_empresa(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM empresa WHERE id = 1")
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            return eliminado

    def cerrar_conexion(self):
        self.conexion.close()


class gestion_empleado:

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("empleados", self.crear_tabla_empleados)

    def _asegurar_columna(self, nombre_columna, definicion):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT 1
                FROM information_schema.columns
                WHERE table_name = 'empleados' AND column_name = %s
                """,
                (nombre_columna,),
            )
            if cursor.fetchone() is None:
                cursor.execute(f"ALTER TABLE empleados ADD COLUMN IF NOT EXISTS {nombre_columna} {definicion}")
        self.conexion.commit()

    def busqueda_escritura(self, busca):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT * FROM empleados
                WHERE Nombre LIKE %s
                OR Dui LIKE %s
            """, (f"{busca}%", f"{busca}%"))
            return cursor.fetchall()

    def crear_tabla_empleados(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS empleados(
                    id SERIAL PRIMARY KEY,
                    Nombre TEXT, Dui TEXT, Correo TEXT, Direccion TEXT, Telefono TEXT,
                    Puesto TEXT, Departamento TEXT, Salario TEXT, Ingreso TEXT, Nacimiento TEXT,
                    estado TEXT DEFAULT 'ACTIVO',
                    fecha_liquidacion DATE,
                    motivo_liquidacion TEXT)""")
            self.conexion.commit()
            cursor.close()

        self._asegurar_columna("estado", "TEXT DEFAULT 'ACTIVO'")
        self._asegurar_columna("fecha_liquidacion", "DATE")
        self._asegurar_columna("motivo_liquidacion", "TEXT")

    def insertar_empleados(self,nombre,dui,correo,direccion,
        telefono,puesto,departamento,salario,ingreso,nacimiento
    ):

        with self.conexion.cursor() as cursor:

            cursor.execute("""
                INSERT INTO empleados(Nombre ,Dui ,Correo ,Direccion ,Telefono ,Puesto ,
                    Departamento ,Salario ,Ingreso ,Nacimiento)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""", (
                nombre,dui,correo,direccion,telefono,puesto,
                departamento,salario,ingreso,nacimiento
            ))

            self.conexion.commit()
            cursor.close()

    def obtener_empleados(self, estado=None):
        with self.conexion.cursor() as cursor:
            if estado is None:
                cursor.execute("""
                    SELECT * FROM empleados
                    WHERE COALESCE(estado, 'ACTIVO') <> 'LIQUIDADO'
                    ORDER BY id
                """)
            else:
                cursor.execute("""
                    SELECT * FROM empleados
                    WHERE COALESCE(estado, 'ACTIVO') = %s
                    ORDER BY id
                """, (str(estado).upper(),))

            datos = cursor.fetchall()
            cursor.close()

            return datos

    def obtener_empleados_todos(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM empleados ORDER BY id")
            return cursor.fetchall()

    def obtener_estado_empleado(self, id_empleado):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COALESCE(estado, 'ACTIVO') FROM empleados WHERE id = %s", (id_empleado,))
            fila = cursor.fetchone()
            return fila[0] if fila else None

    def obtener_empleado_id(self, id):
        with self.conexion.cursor() as cursor:

            cursor.execute("""
                SELECT * FROM empleados
                WHERE id = %s
            """, (id,))

            empleado = cursor.fetchone()
            cursor.close()
            
            return empleado

    def actualizar_empleado(self, id_empleado, nombre, dui, correo, direccion, telefono,
                            puesto, departamento, salario, ingreso, nacimiento):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE empleados
                SET Nombre = %s,
                    Dui = %s,
                    Correo = %s,
                    Direccion = %s,
                    Telefono = %s,
                    Puesto = %s,
                    Departamento = %s,
                    Salario = %s,
                    Ingreso = %s,
                    Nacimiento = %s
                WHERE id = %s
            """, (nombre, dui, correo, direccion, telefono, puesto,
                   departamento, salario, ingreso, nacimiento, id_empleado))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return actualizado
        
    def limpiar_empleado(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM empleados")
            self.conexion.commit()
            cursor.close()

    def eliminar_tabla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DROP TABLE empleados")
            self.conexion.commit()
            cursor.close()


    def eliminar_empleado(self, id):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                UPDATE empleados
                SET estado = 'LIQUIDADO',
                    fecha_liquidacion = CURRENT_DATE,
                    motivo_liquidacion = 'Eliminado del sistema'
                WHERE id = %s
                """,
                (id,),
            )
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return eliminado

    def actualizar_estado_empleado(self, id_empleado, estado, fecha_liquidacion=None, motivo=None):
        estado = str(estado or "ACTIVO").upper()
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                UPDATE empleados
                SET estado = %s,
                    fecha_liquidacion = %s,
                    motivo_liquidacion = %s
                WHERE id = %s
                """,
                (estado, fecha_liquidacion, motivo, id_empleado),
            )
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return actualizado
        

    def buscar_empleado(self, busca):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT * FROM empleados
                WHERE (Nombre ILIKE %s OR Dui = %s)
                  AND COALESCE(estado, 'ACTIVO') <> 'LIQUIDADO'
            """, (f"%{busca}%", busca,))
            empleado = cursor.fetchone()
            cursor.close()
            return empleado


class gestion_planilla:
    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("planilla", self.crear_tabla_planilla)

    def _ensure_column(self, column_name, definition):
        with self.conexion.cursor() as cursor:
            cursor.execute(
                """
                SELECT 1
                FROM information_schema.columns
                WHERE table_name = 'planilla' AND column_name = %s
                """,
                (column_name,)
            )
            if cursor.fetchone() is None:
                cursor.execute(f"ALTER TABLE planilla ADD COLUMN IF NOT EXISTS {column_name} {definition}")
            self.conexion.commit()

    def crear_tabla_planilla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS planilla(
                    id SERIAL PRIMARY KEY,
                    nombre TEXT,
                    dui TEXT,
                    puesto TEXT,
                    salario TEXT,
                    horas_trabajadas TEXT DEFAULT '0',
                    horas_extras TEXT DEFAULT '0',
                    comisiones TEXT DEFAULT '0',
                    bonificaciones TEXT DEFAULT '0',
                    css TEXT DEFAULT '0',
                    seguro_educativo TEXT DEFAULT '0',
                    isr TEXT DEFAULT '0',
                    otros_descuentos TEXT DEFAULT '0',
                    otros TEXT DEFAULT '0',
                    deducciones TEXT DEFAULT '0',
                    total_pagar TEXT,
                    fecha DATE DEFAULT CURRENT_DATE,
                    corte_id INTEGER
                )
            """)
            self.conexion.commit()

        columnas_requeridas = {
            'nombre': 'TEXT',
            'dui': 'TEXT',
            'puesto': 'TEXT',
            'salario': 'TEXT',
            'horas_trabajadas': "TEXT DEFAULT '0'",
            'horas_extras': "TEXT DEFAULT '0'",
            'comisiones': "TEXT DEFAULT '0'",
            'bonificaciones': "TEXT DEFAULT '0'",
            'css': "TEXT DEFAULT '0'",
            'seguro_educativo': "TEXT DEFAULT '0'",
            'isr': "TEXT DEFAULT '0'",
            'otros_descuentos': "TEXT DEFAULT '0'",
            'otros': "TEXT DEFAULT '0'",
            'deducciones': "TEXT DEFAULT '0'",
            'total_pagar': 'TEXT',
            'fecha': 'DATE DEFAULT CURRENT_DATE',
            'corte_id': 'INTEGER',
        }
        for nombre_columna, definicion in columnas_requeridas.items():
            self._ensure_column(nombre_columna, definicion)

        # Compatibilidad con columnas heredadas creadas con nombres en mayúsculas o antiguos.
        columnas_compatibles = {
            'Nombre': 'TEXT',
            'Dui': 'TEXT',
            'Puesto': 'TEXT',
            'Salario': 'TEXT',
            'Horas_Extras': "TEXT DEFAULT '0'",
            'Comisiones': "TEXT DEFAULT '0'",
            'Bonificaciones': "TEXT DEFAULT '0'",
            'CSS': "TEXT DEFAULT '0'",
            'Seguro_Educativo': "TEXT DEFAULT '0'",
            'ISR': "TEXT DEFAULT '0'",
            'Otros_Descuentos': "TEXT DEFAULT '0'",
            'Otros': "TEXT DEFAULT '0'",
            'Deducciones': "TEXT DEFAULT '0'",
            'Total_Pagar': 'TEXT',
            'Corte_Id': 'INTEGER',
        }
        for nombre_columna, definicion in columnas_compatibles.items():
            self._ensure_column(nombre_columna, definicion)

    def insertar_planilla(self, nombre, dui, puesto, salario, horas_extras,
        comisiones, bonificaciones, css, seguro_educativo, isr, otros, deducciones, total_pagar, corte_id=None):

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO planilla(nombre, dui, puesto, salario, horas_trabajadas, horas_extras,
                    comisiones, bonificaciones, css, seguro_educativo, isr, otros_descuentos, otros,
                    deducciones, total_pagar, fecha, corte_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, CURRENT_DATE, %s)
            """, (
                nombre, dui, puesto, salario,
                0,
                horas_extras,
                comisiones, bonificaciones, css, seguro_educativo, isr, 0, otros, deducciones, total_pagar, corte_id
            ))

            self.conexion.commit()
            cursor.close()

    def obtener_planilla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, nombre, dui, puesto, salario, horas_trabajadas, horas_extras,
                    comisiones, bonificaciones, css, seguro_educativo, isr, otros_descuentos, otros,
                    deducciones, total_pagar, fecha, corte_id
                FROM planilla
                ORDER BY nombre ASC
            """)
            datos = cursor.fetchall()
            cursor.close()
            return datos

    def obtener_planilla_por_corte(self, corte_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, nombre, dui, puesto, salario, horas_trabajadas, horas_extras,
                    comisiones, bonificaciones, css, seguro_educativo, isr, otros_descuentos, otros,
                    deducciones, total_pagar, fecha, corte_id
                FROM planilla
                WHERE corte_id = %s
                ORDER BY nombre ASC
            """, (corte_id,))
            datos = cursor.fetchall()
            cursor.close()
            return datos

    def eliminar_planilla_por_corte(self, corte_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                DELETE FROM planilla
                WHERE corte_id = %s
            """, (corte_id,))
            eliminado = cursor.rowcount
            self.conexion.commit()
            cursor.close()
            return eliminado

    def limpiar_deducciones_por_corte(self, corte_id):
        """Pone a cero los campos de deducciones en la tabla planilla para un corte dado.

        Retorna la cantidad de filas afectadas.
        """
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE planilla
                SET css = '0',
                    seguro_educativo = '0',
                    isr = '0',
                    otros_descuentos = '0',
                    otros = '0',
                    deducciones = '0'
                WHERE corte_id = %s
            """, (corte_id,))
            afectadas = cursor.rowcount
            self.conexion.commit()
            cursor.close()
            return afectadas

    def actualizar_planilla(self, registro_id, css, seguro_educativo, isr, otros, deducciones, total_pagar, otros_descuentos=0):
        """Actualiza los campos de deducciones y total de una fila de planilla por su id."""
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE planilla
                SET css = %s,
                    seguro_educativo = %s,
                    isr = %s,
                    otros_descuentos = %s,
                    otros = %s,
                    deducciones = %s,
                    total_pagar = %s
                WHERE id = %s
            """, (str(css), str(seguro_educativo), str(isr), str(otros_descuentos), str(otros), str(deducciones), str(total_pagar), registro_id))
            afectadas = cursor.rowcount
            self.conexion.commit()
            cursor.close()
            return afectadas


class marcaciones:
    _conexion_compartida = None
    _tablas_inicializadas = False

    @classmethod
    def _obtener_conexion_compartida(cls):
        conexion = cls._conexion_compartida
        if conexion is not None:
            try:
                if not conexion.closed:
                    return conexion
            except Exception:
                pass

        cls._conexion_compartida = obtener_conexion()
        return cls._conexion_compartida

    def __init__(self):
        self.conexion = self.__class__._obtener_conexion_compartida()
        if not self.__class__._tablas_inicializadas:
            self._crear_tabla_marcaciones()
            self._crear_tabla_justificaciones_marcacion()
            self._crear_tabla_aprobaciones_usuario_corte()
            self.__class__._tablas_inicializadas = True

    def _crear_tabla_marcaciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS marcaciones(
                    id SERIAL PRIMARY KEY,
                    Nombre TEXT, Dui TEXT, Hora_Entrada TEXT, Hora_Salida TEXT, Fecha DATE,
                    Estado_Dia TEXT DEFAULT 'JORNADA LABORAL')""")
            cursor.execute("""
                ALTER TABLE marcaciones
                ADD COLUMN IF NOT EXISTS Estado_Dia TEXT DEFAULT 'JORNADA LABORAL'
            """)
            self.conexion.commit()
            cursor.close()

    def _crear_tabla_justificaciones_marcacion(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS justificaciones_marcacion(
                    id SERIAL PRIMARY KEY,
                    dui TEXT NOT NULL,
                    fecha DATE NOT NULL,
                    estado_horas_extras TEXT DEFAULT 'PENDIENTE',
                    motivo TEXT,
                    creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(dui, fecha)
                )
            """)
            cursor.execute("""
                ALTER TABLE justificaciones_marcacion
                ADD COLUMN IF NOT EXISTS estado_horas_extras TEXT DEFAULT 'PENDIENTE'
            """)
            self.conexion.commit()
            cursor.close()

    def _crear_tabla_aprobaciones_usuario_corte(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS aprobaciones_usuario_corte(
                    id SERIAL PRIMARY KEY,
                    corte_id INTEGER NOT NULL,
                    dui TEXT NOT NULL,
                    aprobado BOOLEAN DEFAULT TRUE,
                    fecha_aprobacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(corte_id, dui)
                )
            """)
            self.conexion.commit()
            cursor.close()

    @classmethod
    def crear_tabla_marcaciones(cls):
        instancia = cls()
        return instancia

    @classmethod
    def insertar_marcacion(cls, nombre, dui, hora_entrada, hora_salida, fecha, estado_dia="JORNADA LABORAL"):
        instancia = cls()
        with instancia.conexion.cursor() as cursor:
            cursor.execute("SELECT id FROM marcaciones WHERE Dui = %s AND Fecha = %s", (dui, fecha))
            existente = cursor.fetchone()
            if existente:
                cursor.close()
                return False
            cursor.execute("""
                INSERT INTO marcaciones(Nombre ,Dui ,Hora_Entrada ,Hora_Salida, Fecha, Estado_Dia)
                VALUES (%s, %s, %s, %s, %s, %s)""", (
                nombre, dui, hora_entrada, hora_salida, fecha, estado_dia
            ))
            instancia.conexion.commit()
            cursor.close()
        return True

    @classmethod
    def buscar_marcacion(cls, dui, fecha):
        instancia = cls()
        with instancia.conexion.cursor() as cursor:
            cursor.execute("SELECT id, Nombre, Hora_Entrada, Hora_Salida FROM marcaciones WHERE Dui = %s AND Fecha = %s", (dui, fecha))
            registro = cursor.fetchone()
            cursor.close()
        return registro

    def _normalizar_fecha(self, fecha):
        if fecha is None:
            return None
        if hasattr(fecha, "strftime"):
            return fecha

        texto = str(fecha).strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue

        try:
            return datetime.fromisoformat(texto).date()
        except ValueError:
            return texto

    def guardar_justificacion_marcacion(self, dui, fecha, motivo, estado_horas_extras="PENDIENTE"):
        fecha_normalizada = self._normalizar_fecha(fecha)
        estado_normalizado = str(estado_horas_extras or "PENDIENTE").strip().upper()
        if estado_normalizado not in ("PENDIENTE", "APROBADA", "NO APROBADA"):
            estado_normalizado = "PENDIENTE"

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO justificaciones_marcacion (dui, fecha, estado_horas_extras, motivo, actualizado_en)
                VALUES (%s, %s, %s, %s, CURRENT_TIMESTAMP)
                ON CONFLICT (dui, fecha)
                DO UPDATE SET
                    estado_horas_extras = EXCLUDED.estado_horas_extras,
                    motivo = EXCLUDED.motivo,
                    actualizado_en = CURRENT_TIMESTAMP
            """, (dui, fecha_normalizada, estado_normalizado, motivo))
            guardado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return guardado

    def obtener_justificacion_marcacion(self, dui, fecha):
        fecha_normalizada = self._normalizar_fecha(fecha)
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT motivo
                FROM justificaciones_marcacion
                WHERE dui = %s AND fecha = %s
                LIMIT 1
            """, (dui, fecha_normalizada))
            dato = cursor.fetchone()
            cursor.close()
            return dato[0] if dato else ""

    def obtener_revision_marcacion(self, dui, fecha):
        fecha_normalizada = self._normalizar_fecha(fecha)
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT estado_horas_extras, motivo
                FROM justificaciones_marcacion
                WHERE dui = %s AND fecha = %s
                LIMIT 1
            """, (dui, fecha_normalizada))
            dato = cursor.fetchone()
            cursor.close()
            if not dato:
                return "PENDIENTE", ""
            return str(dato[0] or "PENDIENTE").upper(), str(dato[1] or "")

    def obtener_revisiones_marcaciones_rango(self, fecha_inicio, fecha_fin):
        fecha_inicio = self._normalizar_fecha(fecha_inicio)
        fecha_fin = self._normalizar_fecha(fecha_fin)

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT dui, fecha, estado_horas_extras, motivo
                FROM justificaciones_marcacion
                WHERE fecha BETWEEN %s AND %s
            """, (fecha_inicio, fecha_fin))
            datos = cursor.fetchall()
            cursor.close()

        revisiones = {}
        for dui, fecha, estado, motivo in datos:
            revisiones[(str(dui or "").strip(), fecha)] = (
                str(estado or "PENDIENTE").upper(),
                str(motivo or "")
            )
        return revisiones

    def obtener_aprobaciones_por_corte(self, corte_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT dui, aprobado, fecha_aprobacion
                FROM aprobaciones_usuario_corte
                WHERE corte_id = %s
            """, (corte_id,))
            datos = cursor.fetchall()
            cursor.close()

        aprobaciones = {}
        for dui, aprobado, fecha_aprobacion in datos:
            aprobaciones[str(dui or "").strip()] = (bool(aprobado), fecha_aprobacion)
        return aprobaciones

    def aprobar_usuario_en_corte(self, corte_id, dui):
        dui_normalizado = str(dui or "").strip()
        if not dui_normalizado:
            return False

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO aprobaciones_usuario_corte (corte_id, dui, aprobado, fecha_aprobacion)
                VALUES (%s, %s, TRUE, CURRENT_TIMESTAMP)
                ON CONFLICT (corte_id, dui)
                DO UPDATE SET
                    aprobado = TRUE,
                    fecha_aprobacion = CURRENT_TIMESTAMP
            """, (corte_id, dui_normalizado))
            guardado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return guardado

    def obtener_marcaciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT * FROM marcaciones
            """)
            datos = cursor.fetchall()
            cursor.close()
            return datos

    def obtener_marcaciones_por_rango(self, fecha_inicio, fecha_fin):
        def normalizar_fecha(fecha):
            if fecha is None:
                return None
            if hasattr(fecha, "strftime"):
                return fecha

            texto = str(fecha).strip()
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(texto, formato).date()
                except ValueError:
                    continue

            try:
                return datetime.fromisoformat(texto).date()
            except ValueError:
                return texto

        fecha_inicio = normalizar_fecha(fecha_inicio)
        fecha_fin = normalizar_fecha(fecha_fin)

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT *
                FROM marcaciones
                WHERE Fecha BETWEEN %s AND %s
                ORDER BY Fecha DESC, Nombre ASC, Dui ASC
            """, (fecha_inicio, fecha_fin))
            datos = cursor.fetchall()
            cursor.close()
            return datos
        
    def actualizar_marcacion(self, dui, fecha, hora_entrada, hora_salida, nombre=None, estado_dia="JORNADA LABORAL"):
        with self.conexion.cursor() as cursor:
            if nombre is not None:
                cursor.execute("""
                    UPDATE marcaciones
                    SET Nombre = %s, Hora_Entrada = %s, Hora_Salida = %s, Estado_Dia = %s
                    WHERE Dui = %s AND Fecha = %s
                """, (nombre, hora_entrada, hora_salida, estado_dia, dui, fecha))
            else:
                cursor.execute("""
                    UPDATE marcaciones
                    SET Hora_Entrada = %s, Hora_Salida = %s, Estado_Dia = %s
                    WHERE Dui = %s AND Fecha = %s
                """, (hora_entrada, hora_salida, estado_dia, dui, fecha))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return actualizado

    def elimar_marcacion(self, dui, fecha):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                DELETE FROM marcaciones
                WHERE Dui = %s AND Fecha = %s
            """, (dui, fecha))
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return eliminado


class cortes_planilla:
    _conexion_compartida = None
    _tabla_inicializada = False

    @classmethod
    def _obtener_conexion_compartida(cls):
        conexion = cls._conexion_compartida
        if conexion is not None:
            try:
                if not conexion.closed:
                    return conexion
            except Exception:
                pass

        cls._conexion_compartida = obtener_conexion()
        return cls._conexion_compartida

    def __init__(self):
        self.conexion = self.__class__._obtener_conexion_compartida()
        if not self.__class__._tabla_inicializada:
            self._crear_tabla_cortes_planilla()
            self.__class__._tabla_inicializada = True

    def _crear_tabla_cortes_planilla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS cortes_planilla(
                    id SERIAL PRIMARY KEY,
                    Fecha_Inicio DATE,
                    Fecha_Fin DATE,
                    Estado TEXT DEFAULT 'ACTIVO',
                    Marcaciones_Aprobadas BOOLEAN DEFAULT FALSE,
                    Fecha_Aprobacion TIMESTAMP NULL)""")
            cursor.execute("""
                ALTER TABLE cortes_planilla
                ADD COLUMN IF NOT EXISTS Estado TEXT DEFAULT 'ACTIVO'
            """)
            cursor.execute("""
                ALTER TABLE cortes_planilla
                ADD COLUMN IF NOT EXISTS Marcaciones_Aprobadas BOOLEAN DEFAULT FALSE
            """)
            cursor.execute("""
                ALTER TABLE cortes_planilla
                ADD COLUMN IF NOT EXISTS Fecha_Aprobacion TIMESTAMP NULL
            """)
            self.conexion.commit()
            cursor.close()

    def _normalizar_fecha(self, fecha):
        if fecha is None:
            return None
        if hasattr(fecha, "strftime"):
            return fecha

        texto = str(fecha).strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue

        try:
            return datetime.fromisoformat(texto).date()
        except ValueError:
            return texto

        
    def insertar_corte_planilla(self, fecha_inicio, fecha_fin):
        if self.obtener_corte_activo():
            return False

        fecha_inicio = self._normalizar_fecha(fecha_inicio)
        fecha_fin = self._normalizar_fecha(fecha_fin)

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO cortes_planilla(Fecha_Inicio, Fecha_Fin, Estado)
                VALUES (%s, %s, 'ACTIVO')""", (
                fecha_inicio, fecha_fin
            ))
            self.conexion.commit()
            cursor.close()
        return True

    def obtener_corte_activo(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, Fecha_Inicio, Fecha_Fin, Estado
                FROM cortes_planilla
                WHERE Estado = 'ACTIVO'
                ORDER BY id DESC
                LIMIT 1
            """)
            corte = cursor.fetchone()
            cursor.close()
            return corte

    def obtener_aprobacion_marcaciones_corte(self, corte_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT Marcaciones_Aprobadas, Fecha_Aprobacion
                FROM cortes_planilla
                WHERE id = %s
                LIMIT 1
            """, (corte_id,))
            dato = cursor.fetchone()
            cursor.close()

        if not dato:
            return False, None
        return bool(dato[0]), dato[1]

    def aprobar_marcaciones_corte_activo(self):
        corte_activo = self.obtener_corte_activo()
        if not corte_activo:
            return False, "SIN_CORTE"

        corte_id = corte_activo[0]
        aprobado, _ = self.obtener_aprobacion_marcaciones_corte(corte_id)
        if aprobado:
            return False, "YA_APROBADO"

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE cortes_planilla
                SET Marcaciones_Aprobadas = TRUE,
                    Fecha_Aprobacion = CURRENT_TIMESTAMP
                WHERE id = %s
            """, (corte_id,))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()

        if not actualizado:
            return False, "NO_ACTUALIZADO"
        return True, "OK"

    def cerrar_corte_activo(self, fecha_fin):
        corte_activo = self.obtener_corte_activo()
        if not corte_activo:
            return False

        fecha_fin = self._normalizar_fecha(fecha_fin)

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE cortes_planilla
                SET Fecha_Fin = %s,
                    Estado = 'CERRADO'
                WHERE id = %s
            """, (fecha_fin, corte_activo[0]))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return actualizado


    def obtener_cortes_planilla(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT * FROM cortes_planilla
                ORDER BY id ASC
            """)
            datos = cursor.fetchall()
            cursor.close()
            return datos


class descuentos:
    _instancia = None

    def __new__(cls, *args, **kwargs):
        if cls._instancia is None:
            cls._instancia = super().__new__(cls)
        return cls._instancia

    def __init__(self):
        if getattr(self, "_inicializada", False):
            return
        self.conexion = obtener_conexion()
        inicializar_tabla("descuentos", self.crear_tabla_descuentos)
        self._inicializada = True

    def _rollback_si_hay_error(self):
        """Evita que una transacción abortada deje la conexión en estado inválido."""
        try:
            self.conexion.rollback()
        except Exception:
            pass

    def crear_tabla_descuentos(self):
        """Crea/actualiza la tabla de descuentos sin campo Fecha redundante"""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS descuentos(
                        id SERIAL PRIMARY KEY,
                        Nombre TEXT NOT NULL,
                        Dui TEXT NOT NULL,
                        Cargo TEXT,
                        Salario NUMERIC(12,2),
                        Tipo TEXT NOT NULL,
                        Acreedor TEXT,
                        Monto NUMERIC(12,2) NOT NULL,
                        Estado TEXT DEFAULT 'ACTIVO',
                        Observaciones TEXT,
                        Fecha_Creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        Fecha_Actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )""")

                cursor.execute("ALTER TABLE descuentos DROP COLUMN IF EXISTS Fecha")
                cursor.execute("ALTER TABLE descuentos ADD COLUMN IF NOT EXISTS Estado TEXT DEFAULT 'ACTIVO'")
                cursor.execute("ALTER TABLE descuentos ADD COLUMN IF NOT EXISTS Observaciones TEXT")
                cursor.execute("ALTER TABLE descuentos ADD COLUMN IF NOT EXISTS Fecha_Creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
                cursor.execute("ALTER TABLE descuentos ADD COLUMN IF NOT EXISTS Fecha_Actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_descuentos_dui ON descuentos(Dui)")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_descuentos_tipo ON descuentos(Tipo)")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_descuentos_estado ON descuentos(Estado)")
                self.conexion.commit()
        except Exception:
            self._rollback_si_hay_error()
            raise

    def validar_duplicado(self, dui, tipo, acreedor, monto, excluir_id=None):
        """Verifica si existe un descuento idéntico para evitar duplicados"""
        try:
            # Normalizar monto a float
            monto_norm = float(str(monto).strip().replace(',', '.'))
            
            with self.conexion.cursor() as cursor:
                consulta = """
                    SELECT id FROM descuentos
                    WHERE Dui = %s AND Tipo = %s AND Acreedor = %s 
                    AND Monto = %s::NUMERIC(12,2) AND Estado = 'ACTIVO'
                """
                parametros = [dui, tipo, acreedor, monto_norm]
                if excluir_id is not None:
                    consulta += " AND id <> %s"
                    parametros.append(excluir_id)
                consulta += " LIMIT 1"
                cursor.execute(consulta, parametros)
                resultado = cursor.fetchone()
            return resultado is not None
        except Exception as e:
            # Si hay error en validación de duplicado, registrarlo pero permitir continuar
            print(f"Error en validar_duplicado: {e}")
            self._rollback_si_hay_error()
            return False

    def insertar_descuento(self, nombre, dui, cargo, salario, tipo, acreedor, monto, observaciones=""):
        """Inserta un nuevo descuento con validación de duplicados"""
        if self.validar_duplicado(dui, tipo, acreedor, monto):
            return False, "DUPLICADO"

        try:
            # Convertir monto de diferentes formatos
            monto_str = str(monto).strip().replace(',', '.')
            monto_float = float(monto_str)
            
            # Convertir salario - si está vacío o es None, usar None
            try:
                if salario and str(salario).strip():
                    salario_float = float(str(salario).strip().replace(',', '.'))
                else:
                    salario_float = None
            except (ValueError, TypeError):
                salario_float = None
                
        except (ValueError, TypeError) as e:
            return False, "MONTO_INVALIDO"

        if monto_float <= 0:
            return False, "MONTO_NEGATIVO"

        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO descuentos
                    (Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado, Observaciones)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, 'ACTIVO', %s)
                    RETURNING id
                """, (nombre, dui, cargo, salario_float, tipo, acreedor, monto_float, observaciones))
                id_nuevo = cursor.fetchone()[0]
            self.conexion.commit()
            return True, id_nuevo
        except Exception as e:
            self._rollback_si_hay_error()
            return False, "ERROR_DB"

    def obtener_descuentos(self, estado="ACTIVO"):
        """Obtiene todos los descuentos, filtrados por estado"""
        try:
            with self.conexion.cursor() as cursor:
                if estado:
                    cursor.execute("""
                        SELECT id, Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado, 
                               Observaciones, Fecha_Creacion, Fecha_Actualizacion
                        FROM descuentos
                        WHERE Estado = %s
                        ORDER BY Fecha_Creacion DESC
                    """, (estado,))
                else:
                    cursor.execute("""
                        SELECT id, Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado,
                               Observaciones, Fecha_Creacion, Fecha_Actualizacion
                        FROM descuentos
                        ORDER BY Fecha_Creacion DESC
                    """)
                return cursor.fetchall()
        except Exception:
            self._rollback_si_hay_error()
            raise

    def obtener_descuentos_por_empleado(self, dui, estado="ACTIVO"):
        """Obtiene descuentos de un empleado específico"""
        try:
            with self.conexion.cursor() as cursor:
                if estado:
                    cursor.execute("""
                        SELECT id, Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado,
                               Observaciones, Fecha_Creacion, Fecha_Actualizacion
                        FROM descuentos
                        WHERE Dui = %s AND Estado = %s
                        ORDER BY Fecha_Creacion DESC
                    """, (dui, estado))
                else:
                    cursor.execute("""
                        SELECT id, Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado,
                               Observaciones, Fecha_Creacion, Fecha_Actualizacion
                        FROM descuentos
                        WHERE Dui = %s
                        ORDER BY Fecha_Creacion DESC
                    """, (dui,))
                return cursor.fetchall()
        except Exception:
            self._rollback_si_hay_error()
            raise

    def obtener_descuento_por_id(self, id):
        """Obtiene un descuento específico por ID"""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT id, Nombre, Dui, Cargo, Salario, Tipo, Acreedor, Monto, Estado,
                           Observaciones, Fecha_Creacion, Fecha_Actualizacion
                    FROM descuentos
                    WHERE id = %s
                """, (id,))
                return cursor.fetchone()
        except Exception:
            self._rollback_si_hay_error()
            raise

    def actualizar_descuento(self, id, nombre, dui, cargo, salario, tipo, acreedor, monto, observaciones=""):
        """Actualiza un descuento existente"""
        if self.validar_duplicado(dui, tipo, acreedor, monto, excluir_id=id):
            return False, "DUPLICADO"
        try:
            # Convertir monto de diferentes formatos
            monto_str = str(monto).strip().replace(',', '.')
            monto_float = float(monto_str)
            
            # Convertir salario - si está vacío o es None, usar None
            try:
                if salario and str(salario).strip():
                    salario_float = float(str(salario).strip().replace(',', '.'))
                else:
                    salario_float = None
            except (ValueError, TypeError):
                salario_float = None
                
        except (ValueError, TypeError):
            return False, "MONTO_INVALIDO"

        if monto_float <= 0:
            return False, "MONTO_NEGATIVO"

        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    UPDATE descuentos
                    SET Nombre = %s, Dui = %s, Cargo = %s, Salario = %s,
                        Tipo = %s, Acreedor = %s, Monto = %s,
                        Observaciones = %s, Fecha_Actualizacion = CURRENT_TIMESTAMP
                    WHERE id = %s
                    RETURNING id
                """, (nombre, dui, cargo, salario_float, tipo, acreedor,
                      monto_float, observaciones, id))
                actualizado = cursor.rowcount > 0
            self.conexion.commit()
            return actualizado, "OK"
        except Exception:
            self._rollback_si_hay_error()
            return False, "ERROR_DB"

    def cambiar_estado_descuento(self, id, nuevo_estado):
        """Cambia el estado de un descuento (ACTIVO, PAGADO, ANULADO)"""
        if nuevo_estado not in ["ACTIVO", "PAGADO", "ANULADO"]:
            return False, "ESTADO_INVALIDO"

        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    UPDATE descuentos
                    SET Estado = %s, Fecha_Actualizacion = CURRENT_TIMESTAMP
                    WHERE id = %s
                    RETURNING id
                """, (nuevo_estado, id))
                actualizado = cursor.rowcount > 0
            self.conexion.commit()
            return actualizado, "OK"
        except Exception:
            self._rollback_si_hay_error()
            return False, "ERROR_DB"

    def eliminar_descuento(self, id):
        """Soft delete - marca como ANULADO en lugar de eliminar"""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    UPDATE descuentos
                    SET Estado = 'ANULADO', Fecha_Actualizacion = CURRENT_TIMESTAMP
                    WHERE id = %s
                    RETURNING id
                """, (id,))
                eliminado = cursor.rowcount > 0
            self.conexion.commit()
            return eliminado
        except Exception:
            self._rollback_si_hay_error()
            return False

    def eliminar_descuentos_por_empleado(self, dui):
        """Elimina todos los descuentos de un empleado"""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    UPDATE descuentos
                    SET Estado = 'ANULADO', Fecha_Actualizacion = CURRENT_TIMESTAMP
                    WHERE Dui = %s AND Estado != 'ANULADO'
                """, (dui,))
                eliminados = cursor.rowcount
            self.conexion.commit()
            return eliminados > 0
        except Exception:
            self._rollback_si_hay_error()
            return False

    def obtener_total_descuentos_empleado(self, dui, estado="ACTIVO"):
        """Calcula el total de descuentos de un empleado"""
        try:
            with self.conexion.cursor() as cursor:
                if estado:
                    cursor.execute("""
                        SELECT COALESCE(SUM(CAST(Monto AS NUMERIC)), 0) FROM descuentos
                        WHERE Dui = %s AND Estado = %s
                    """, (dui, estado))
                else:
                    cursor.execute("""
                        SELECT COALESCE(SUM(CAST(Monto AS NUMERIC)), 0) FROM descuentos
                        WHERE Dui = %s
                    """, (dui,))
                total = cursor.fetchone()[0]
            return float(total)
        except Exception:
            self._rollback_si_hay_error()
            raise

    def obtener_total_descuentos_por_rango(self, dui, fecha_inicio, fecha_fin):
        """Obtiene descuentos no anulados de un trabajador dentro de un corte."""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT COALESCE(SUM(
                        CASE
                            WHEN TRIM(Monto::TEXT) ~ '^[+-]?[0-9]+([.,][0-9]+)?$'
                            THEN REPLACE(TRIM(Monto::TEXT), ',', '.')::NUMERIC
                            ELSE 0
                        END
                    ), 0)
                    FROM descuentos
                    WHERE Dui = %s
                      AND Estado IN ('ACTIVO', 'PAGADO')
                      AND Fecha_Creacion::date BETWEEN %s AND %s
                """, (dui, fecha_inicio, fecha_fin))
                total = cursor.fetchone()[0]
            return float(total or 0)
        except Exception:
            self._rollback_si_hay_error()
            raise

    def obtener_totales_por_rango(self, fecha_inicio, fecha_fin):
        """Obtiene en una sola consulta los descuentos agrupados por DUI."""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT Dui, COALESCE(SUM(
                        CASE
                            WHEN TRIM(Monto::TEXT) ~ '^[+-]?[0-9]+([.,][0-9]+)?$'
                            THEN REPLACE(TRIM(Monto::TEXT), ',', '.')::NUMERIC
                            ELSE 0
                        END
                    ), 0)
                    FROM descuentos
                    WHERE Estado IN ('ACTIVO', 'PAGADO')
                      AND Fecha_Creacion::date BETWEEN %s AND %s
                    GROUP BY Dui
                """, (fecha_inicio, fecha_fin))
                return {str(dui or '').strip(): float(total or 0) for dui, total in cursor.fetchall()}
        except Exception:
            self._rollback_si_hay_error()
            raise

    def obtener_resumen_descuentos(self):
        """Obtiene resumen estadístico de descuentos"""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT 
                        Estado,
                        COUNT(*) as cantidad,
                        COALESCE(SUM(CAST(Monto AS NUMERIC)), 0) as monto_total
                    FROM descuentos
                    WHERE Estado != 'ANULADO'
                    GROUP BY Estado
                """)
                return cursor.fetchall()
        except Exception:
            self._rollback_si_hay_error()
            raise

    def limpiar_descuentos(self):
        """Elimina permanentemente todos los descuentos (usar con cuidado)"""
        try:
            self._rollback_si_hay_error()
            with self.conexion.cursor() as cursor:
                cursor.execute("DELETE FROM descuentos")
                cursor.execute("ALTER SEQUENCE descuentos_id_seq RESTART WITH 1")
            self.conexion.commit()
            return True
        except Exception:
            self._rollback_si_hay_error()
            return False


class Permisos_y_ausencias:

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("permisos_ausencias", self.crear_tabla_permisos_ausencias)

    def crear_tabla_permisos_ausencias(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS permisos_ausencias(
                    id SERIAL PRIMARY KEY,
                    Nombre TEXT,
                    Dui TEXT,
                    Tipo TEXT,
                    Pagado TEXT,
                    Fecha_Inicio DATE,
                    Fecha_Fin DATE,
                    Motivo TEXT,
                    Imagen_Nombre TEXT,
                    Imagen_Bytes BYTEA,
                    Creado_At TIMESTAMP
                )""")
            cursor.execute("""
                ALTER TABLE permisos_ausencias
                    ADD COLUMN IF NOT EXISTS Tipo TEXT,
                    ADD COLUMN IF NOT EXISTS Pagado TEXT,
                    ADD COLUMN IF NOT EXISTS Fecha_Inicio DATE,
                    ADD COLUMN IF NOT EXISTS Fecha_Fin DATE,
                    ADD COLUMN IF NOT EXISTS Motivo TEXT,
                    ADD COLUMN IF NOT EXISTS Imagen_Nombre TEXT,
                    ADD COLUMN IF NOT EXISTS Imagen_Bytes BYTEA,
                    ADD COLUMN IF NOT EXISTS Creado_At TIMESTAMP
            """)
            self.conexion.commit()
            cursor.close()

    def insertar_permiso_ausencia(self, nombre, dui, fecha_inicio, fecha_fin, tipo, motivo, pagado, imagen_nombre=None, imagen_bytes=None, creado_at=None):
        if creado_at is None:
            creado_at = datetime.now()
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO permisos_ausencias(
                    Nombre, Dui, Tipo, Pagado, Fecha_Inicio, Fecha_Fin, Motivo,
                    Imagen_Nombre, Imagen_Bytes, Creado_At
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (
                nombre, dui, tipo, pagado, fecha_inicio, fecha_fin, motivo,
                imagen_nombre, imagen_bytes, creado_at
            ))
            self.conexion.commit()
            cursor.close()

    def obtener_permisos_ausencias(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, Nombre, Dui, Tipo, Pagado, Fecha_Inicio, Fecha_Fin,
                       Motivo, Imagen_Nombre, Imagen_Bytes, Creado_At
                FROM permisos_ausencias
                ORDER BY Creado_At DESC NULLS LAST, id DESC
            """)
            datos = cursor.fetchall()
            cursor.close()
            return datos

    def eliminar_permiso_ausencia(self, id):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM permisos_ausencias WHERE id = %s", (id,))
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            cursor.close()
            return eliminado

    def limpiar_permisos_ausencias(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM permisos_ausencias")
            self.conexion.commit()
            cursor.close()


class expedientes_laborales:

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("contratos_trabajo", self.crear_tabla_contratos)

    @staticmethod
    def _normalizar_contrato(valor):
        return str(valor or "").strip().lower()

    @classmethod
    def _generar_clave_contrato(cls, empleado_id, tipo_contrato, fecha_inicio, fecha_fin,
                               jornada, horario_entrada, horario_salida, estado):
        return (
            int(empleado_id),
            cls._normalizar_contrato(tipo_contrato),
            str(fecha_inicio) if fecha_inicio is not None else "",
            str(fecha_fin) if fecha_fin is not None else "",
            cls._normalizar_contrato(jornada),
            cls._normalizar_contrato(horario_entrada),
            cls._normalizar_contrato(horario_salida),
            cls._normalizar_contrato(estado),
        )

    def crear_tabla_contratos(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS contratos_trabajo(
                    id SERIAL PRIMARY KEY,
                    empleado_id INTEGER NOT NULL REFERENCES empleados(id) ON DELETE CASCADE,
                    tipo_contrato TEXT NOT NULL DEFAULT 'Indefinido',
                    fecha_inicio DATE,
                    fecha_fin DATE,
                    jornada TEXT DEFAULT 'Tiempo completo',
                    horario_entrada TEXT DEFAULT '08:00',
                    horario_salida TEXT DEFAULT '17:00',
                    estado TEXT NOT NULL DEFAULT 'Activo',
                    archivo_nombre TEXT,
                    archivo_bytes BYTEA,
                    creado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("DROP INDEX IF EXISTS ux_contratos_trabajo_unico")
            self.limpiar_duplicados_contratos()
            self.conexion.commit()

    def limpiar_duplicados_contratos(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                WITH ranked AS (
                    SELECT id,
                           ROW_NUMBER() OVER (
                               PARTITION BY empleado_id,
                                            lower(tipo_contrato),
                                            coalesce(fecha_inicio::text, ''),
                                            coalesce(fecha_fin::text, ''),
                                            lower(jornada),
                                            lower(horario_entrada),
                                            lower(horario_salida),
                                            lower(estado)
                               ORDER BY creado_at DESC, id DESC
                           ) AS rn
                    FROM contratos_trabajo
                )
                DELETE FROM contratos_trabajo
                WHERE id IN (SELECT id FROM ranked WHERE rn > 1)
            """)
            self.conexion.commit()

    def existe_contrato_duplicado(self, empleado_id, tipo_contrato, fecha_inicio, fecha_fin,
                                 jornada, horario_entrada, horario_salida, estado):
        clave = self._generar_clave_contrato(
            empleado_id, tipo_contrato, fecha_inicio, fecha_fin,
            jornada, horario_entrada, horario_salida, estado
        )
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, tipo_contrato, fecha_inicio, fecha_fin, jornada,
                       horario_entrada, horario_salida, estado
                FROM contratos_trabajo
                WHERE empleado_id = %s
                ORDER BY creado_at DESC, id DESC
            """, (empleado_id,))
            for registro in cursor.fetchall():
                actual = self._generar_clave_contrato(
                    empleado_id,
                    registro[1], registro[2], registro[3],
                    registro[4], registro[5], registro[6], registro[7]
                )
                if actual == clave:
                    return registro[0]
        return None

    def guardar_contrato(self, empleado_id, tipo_contrato, fecha_inicio, fecha_fin,
                         jornada, horario_entrada, horario_salida, estado,
                         archivo_nombre=None, archivo_bytes=None):
        contrato_id = self.existe_contrato_duplicado(
            empleado_id, tipo_contrato, fecha_inicio, fecha_fin,
            jornada, horario_entrada, horario_salida, estado
        )
        if contrato_id:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    UPDATE contratos_trabajo
                    SET tipo_contrato = %s,
                        fecha_inicio = %s,
                        fecha_fin = %s,
                        jornada = %s,
                        horario_entrada = %s,
                        horario_salida = %s,
                        estado = %s,
                        archivo_nombre = COALESCE(%s, archivo_nombre),
                        archivo_bytes = COALESCE(%s, archivo_bytes)
                    WHERE id = %s
                """, (tipo_contrato, fecha_inicio, fecha_fin, jornada,
                       horario_entrada, horario_salida, estado,
                       archivo_nombre, archivo_bytes, contrato_id))
            self.conexion.commit()
            return contrato_id

        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO contratos_trabajo(
                    empleado_id, tipo_contrato, fecha_inicio, fecha_fin, jornada,
                    horario_entrada, horario_salida, estado, archivo_nombre, archivo_bytes
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (empleado_id, tipo_contrato, fecha_inicio, fecha_fin, jornada,
                   horario_entrada, horario_salida, estado, archivo_nombre, archivo_bytes))
            contrato_id = cursor.fetchone()[0]
            self.conexion.commit()
            return contrato_id

    def obtener_contrato(self, empleado_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT id, tipo_contrato, fecha_inicio, fecha_fin, jornada,
                       horario_entrada, horario_salida, estado, archivo_nombre
                FROM contratos_trabajo
                WHERE empleado_id = %s
                ORDER BY creado_at DESC, id DESC
                LIMIT 1
            """, (empleado_id,))
            return cursor.fetchone()

    def guardar_archivo_contrato(self, empleado_id, archivo_nombre, archivo_bytes):
        """Guarda el documento contractual sin que la aplicación cree su contenido."""
        with self.conexion.cursor() as cursor:
            cursor.execute(
                "SELECT id FROM contratos_trabajo WHERE empleado_id = %s ORDER BY creado_at DESC, id DESC LIMIT 1",
                (empleado_id,),
            )
            contrato = cursor.fetchone()
            if contrato:
                cursor.execute(
                    """
                    UPDATE contratos_trabajo
                    SET archivo_nombre = %s, archivo_bytes = %s
                    WHERE id = %s
                    """,
                    (archivo_nombre, archivo_bytes, contrato[0]),
                )
                contrato_id = contrato[0]
            else:
                cursor.execute(
                    """
                    INSERT INTO contratos_trabajo(
                        empleado_id, tipo_contrato, estado, archivo_nombre, archivo_bytes
                    ) VALUES (%s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (empleado_id, "Documento contractual", "Activo", archivo_nombre, archivo_bytes),
                )
                contrato_id = cursor.fetchone()[0]
        self.conexion.commit()
        return contrato_id

    @staticmethod
    def _numero(valor):
        try:
            return float(str(valor or "0").replace("B/.", "").replace(",", "").strip())
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _hora(valor):
        if not valor:
            return None
        texto = str(valor).strip()
        for formato in ("%H:%M", "%H:%M:%S", "%I:%M %p"):
            try:
                return datetime.strptime(texto, formato).time()
            except ValueError:
                continue
        return None

    @staticmethod
    def _dias_periodo(fecha_inicio, fecha_fin):
        if not fecha_inicio or not fecha_fin:
            return 0
        return sum(
            1 for numero in range((fecha_fin - fecha_inicio).days + 1)
            if (fecha_inicio + timedelta(days=numero)).weekday() < 5
        )

    def obtener_expediente(self, empleado, fecha_inicio, fecha_fin):
        empleado_id, nombre, dui = empleado[0], str(empleado[1] or ""), str(empleado[2] or "")
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT fecha, hora_entrada, hora_salida, estado_dia
                FROM marcaciones
                WHERE dui = %s AND fecha BETWEEN %s AND %s
                ORDER BY fecha DESC
            """, (dui, fecha_inicio, fecha_fin))
            marcaciones = cursor.fetchall()
            cursor.execute("""
                SELECT tipo, pagado, fecha_inicio, fecha_fin, motivo
                FROM permisos_ausencias
                WHERE dui = %s AND fecha_fin >= %s AND fecha_inicio <= %s
                ORDER BY fecha_inicio DESC
            """, (dui, fecha_inicio, fecha_fin))
            permisos = cursor.fetchall()
            cursor.execute("""
                SELECT fecha_inicio, fecha_fin, dias, total_bruto, estado
                FROM solicitudes_vacaciones
                WHERE empleado_id = %s AND fecha_fin >= %s AND fecha_inicio <= %s
                ORDER BY fecha_inicio DESC
            """, (empleado_id, fecha_inicio, fecha_fin))
            vacaciones = cursor.fetchall()
            cursor.execute("""
                SELECT tipo, monto, fecha, motivo
                FROM remuneraciones
                WHERE lower(nombre) = lower(%s) AND fecha::date BETWEEN %s AND %s
                ORDER BY fecha DESC
            """, (nombre, fecha_inicio, fecha_fin))
            remuneraciones = cursor.fetchall()
            cursor.execute("""
                SELECT salario, css, seguro_educativo, isr, deducciones, total_pagar, fecha
                FROM planilla
                WHERE dui = %s AND fecha BETWEEN %s AND %s
                ORDER BY fecha DESC
            """, (dui, fecha_inicio, fecha_fin))
            planilla = cursor.fetchall()

        hora_entrada = self._hora("08:00")
        hora_salida = self._hora("17:00")
        tardanzas = 0
        sobretiempo_horas = 0.0
        for _, entrada, salida, _ in marcaciones:
            entrada_hora = self._hora(entrada)
            salida_hora = self._hora(salida)
            if entrada_hora and entrada_hora > hora_entrada:
                tardanzas += 1
            if salida_hora and salida_hora > hora_salida:
                sobretiempo_horas += (
                    datetime.combine(datetime.today(), salida_hora) -
                    datetime.combine(datetime.today(), hora_salida)
                ).seconds / 3600

        dias_permiso = sum(
            max(0, (fin - inicio).days + 1) for _, _, inicio, fin, _ in permisos
            if inicio and fin
        )
        dias_esperados = self._dias_periodo(fecha_inicio, fecha_fin)
        ausencias = max(0, dias_esperados - len(marcaciones) - dias_permiso)
        return {
            "marcaciones": marcaciones,
            "permisos": permisos,
            "vacaciones": vacaciones,
            "remuneraciones": remuneraciones,
            "planilla": planilla,
            "dias_esperados": dias_esperados,
            "dias_trabajados": len(marcaciones),
            "tardanzas": tardanzas,
            "ausencias": ausencias,
            "dias_permiso": dias_permiso,
            "sobretiempo_horas": round(sobretiempo_horas, 2),
            "total_remuneraciones": round(sum(self._numero(registro[1]) for registro in remuneraciones), 2),
            "total_vacaciones": round(sum(self._numero(registro[3]) for registro in vacaciones if registro[4] == "Aprobada"), 2),
            "total_planilla": round(sum(self._numero(registro[5]) for registro in planilla), 2),
            "total_deducciones": round(sum(self._numero(registro[4]) for registro in planilla), 2),
        }


class vacaciones_pagos:

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("vacaciones", self.crear_tablas_vacaciones)

    def crear_tablas_vacaciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS periodos_vacaciones(
                    id SERIAL PRIMARY KEY,
                    empleado_id INTEGER NOT NULL REFERENCES empleados(id) ON DELETE CASCADE,
                    fecha_inicio DATE NOT NULL,
                    fecha_fin DATE NOT NULL,
                    dias_ganados INTEGER NOT NULL DEFAULT 0,
                    dias_usados INTEGER NOT NULL DEFAULT 0,
                    creado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(empleado_id, fecha_inicio)
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS solicitudes_vacaciones(
                    id SERIAL PRIMARY KEY,
                    empleado_id INTEGER NOT NULL REFERENCES empleados(id) ON DELETE CASCADE,
                    fecha_inicio DATE NOT NULL,
                    fecha_fin DATE NOT NULL,
                    dias INTEGER NOT NULL CHECK (dias > 0),
                    salario_mensual NUMERIC(12, 2) NOT NULL DEFAULT 0,
                    valor_diario NUMERIC(12, 2) NOT NULL DEFAULT 0,
                    total_bruto NUMERIC(12, 2) NOT NULL DEFAULT 0,
                    estado TEXT NOT NULL DEFAULT 'Pendiente',
                    observaciones TEXT,
                    creado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    aprobado_at TIMESTAMP
                )
            """)
            self.conexion.commit()

    @staticmethod
    def dias_ganados_por_meses(meses):
        """Devuelve el saldo de vacaciones según Código de Trabajo de Panamá.
        
        Legislación:
        - Menos de 1 año: 0 días (sin derecho)
        - 1 a 4 años: 15 días
        - 5 años o más: 30 días + 30 días por cada 5 años adicionales
        
        Nota: Esta es la escala LEGAL. Consulta con RR.HH. si se usa escala diferente.
        """
        if meses < 12:
            return 0  # Sin derecho antes de 1 año completo
        if meses < 60:  # Menos de 5 años
            return 15  # 15 días entre 1-4 años
        # 5+ años: 30 días base + 30 días por cada 5 años adicionales
        return 30 + ((meses - 60) // 60) * 30

    @staticmethod
    def meses_completos(fecha_ingreso, fecha_corte=None):
        if not fecha_ingreso:
            return 0
        fecha_corte = fecha_corte or datetime.now().date()
        if hasattr(fecha_ingreso, "date"):
            fecha_ingreso = fecha_ingreso.date()
        if isinstance(fecha_ingreso, str):
            for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                try:
                    fecha_ingreso = datetime.strptime(fecha_ingreso[:10], formato).date()
                    break
                except ValueError:
                    continue
            else:
                return 0
        meses = (fecha_corte.year - fecha_ingreso.year) * 12 + fecha_corte.month - fecha_ingreso.month
        if fecha_corte.day < fecha_ingreso.day:
            meses -= 1
        return max(0, meses)

    def obtener_resumen_empleado(self, empleado):
        empleado_id = empleado[0]
        nombre = str(empleado[1] or "Sin nombre")
        salario_texto = str(empleado[8] or "0").replace("B/.", "").replace(",", "").strip()
        try:
            salario = float(salario_texto)
        except ValueError:
            salario = 0.0
        meses = self.meses_completos(empleado[9])
        dias_ganados = self.dias_ganados_por_meses(meses)
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT COALESCE(SUM(dias), 0)
                FROM solicitudes_vacaciones
                WHERE empleado_id = %s AND estado = 'Aprobada'
            """, (empleado_id,))
            dias_usados = int(cursor.fetchone()[0])
            cursor.execute("""
                SELECT COUNT(*), COALESCE(SUM(dias), 0)
                FROM solicitudes_vacaciones
                WHERE empleado_id = %s AND estado = 'Pendiente'
            """, (empleado_id,))
            pendientes, dias_pendientes = cursor.fetchone()
        disponibles = max(0, dias_ganados - dias_usados - int(dias_pendientes or 0))
        return {
            "empleado_id": empleado_id,
            "nombre": nombre,
            "salario": salario,
            "meses": meses,
            "dias_ganados": dias_ganados,
            "dias_usados": dias_usados,
            "dias_pendientes": int(dias_pendientes or 0),
            "solicitudes_pendientes": int(pendientes or 0),
            "dias_disponibles": disponibles,
            "valor_diario": salario / 30,
        }

    def insertar_solicitud(self, empleado_id, fecha_inicio, fecha_fin, dias,
                           salario_mensual, observaciones=""):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM empleados WHERE id = %s", (empleado_id,))
            empleado = cursor.fetchone()
        if not empleado:
            raise ValueError("El empleado seleccionado ya no existe.")
        
        # Validar que las fechas no estén en el pasado
        if fecha_inicio < datetime.now().date():
            raise ValueError("La fecha de inicio no puede ser anterior a hoy.")
        
        resumen = self.obtener_resumen_empleado(empleado)
        if dias > resumen["dias_disponibles"]:
            raise ValueError(f"El empleado solo tiene {resumen['dias_disponibles']} días disponibles.")
        
        # Validar que no haya solicitudes superpuestas
        if not self.validar_fechas_superpuestas(empleado_id, fecha_inicio, fecha_fin):
            raise ValueError("Existe una solicitud de vacaciones que se superpone con las fechas seleccionadas.")
        
        # Limpiar el salario
        salario_limpio = self.limpiar_salario(salario_mensual)
        if salario_limpio <= 0:
            raise ValueError("El salario del empleado no es válido.")
        
        valor_diario = round(salario_limpio / 30, 2)
        total_bruto = round(valor_diario * dias, 2)
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO solicitudes_vacaciones(
                    empleado_id, fecha_inicio, fecha_fin, dias,
                    salario_mensual, valor_diario, total_bruto, observaciones
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, (empleado_id, fecha_inicio, fecha_fin, dias, salario_limpio,
                   valor_diario, total_bruto, observaciones))
            solicitud_id = cursor.fetchone()[0]
            self.conexion.commit()
            return solicitud_id

    def obtener_solicitudes(self, estado=None):
        with self.conexion.cursor() as cursor:
            consulta = """
                SELECT s.id, e.Nombre, e.Dui, s.fecha_inicio, s.fecha_fin,
                       s.dias, s.salario_mensual, s.total_bruto, s.estado,
                       s.observaciones
                FROM solicitudes_vacaciones s
                JOIN empleados e ON e.id = s.empleado_id
            """
            parametros = ()
            if estado:
                consulta += " WHERE s.estado = %s"
                parametros = (estado,)
            consulta += " ORDER BY s.creado_at DESC, s.id DESC"
            cursor.execute(consulta, parametros)
            return cursor.fetchall()

    def actualizar_estado_solicitud(self, solicitud_id, estado):
        if estado not in ("Aprobada", "Rechazada", "Cancelada"):
            raise ValueError("Estado de solicitud no válido.")
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE solicitudes_vacaciones
                SET estado = %s, aprobado_at = CASE WHEN %s = 'Aprobada' THEN CURRENT_TIMESTAMP ELSE aprobado_at END
                WHERE id = %s AND estado = 'Pendiente'
            """, (estado, estado, solicitud_id))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            return actualizado

    def contar_resumen(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) FROM solicitudes_vacaciones WHERE estado = 'Pendiente'")
            pendientes = cursor.fetchone()[0]
            cursor.execute("SELECT COALESCE(SUM(dias), 0) FROM solicitudes_vacaciones WHERE estado = 'Aprobada'")
            dias_aprobados = cursor.fetchone()[0]
        return int(pendientes), int(dias_aprobados)

    @staticmethod
    def limpiar_salario(salario_texto):
        """Extrae solo números y decimales del salario.
        Ejemplos: 'B/. 1,500.00' -> 1500.00, '1500' -> 1500.0
        """
        if isinstance(salario_texto, (int, float)):
            return float(salario_texto)
        texto = str(salario_texto or "0").strip()
        # Remover B/., espacios y comas
        texto = texto.replace("B/.", "").replace(" ", "").replace(",", "")
        try:
            return float(texto) if texto else 0.0
        except ValueError:
            return 0.0

    def validar_fechas_superpuestas(self, empleado_id, fecha_inicio, fecha_fin):
        """Valida que no existan solicitudes aprobadas o pendientes que se superpongan."""
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                SELECT COUNT(*) FROM solicitudes_vacaciones
                WHERE empleado_id = %s
                AND estado IN ('Aprobada', 'Pendiente')
                AND fecha_inicio <= %s AND fecha_fin >= %s
            """, (empleado_id, fecha_fin, fecha_inicio))
            cantidad = cursor.fetchone()[0]
        return cantidad == 0

    def obtener_departamentos_unicos(self):
        """Obtiene lista de departamentos únicos de todos los empleados."""
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT DISTINCT Departamento FROM empleados
                    WHERE Departamento IS NOT NULL AND Departamento != ''
                    ORDER BY Departamento
                """)
                departamentos = [row[0] for row in cursor.fetchall()]
            return departamentos
        except Exception:
            return []

class liquidaciones_db:

    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("liquidaciones", self.crear_tabla_liquidaciones)

    def crear_tabla_liquidaciones(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS liquidaciones(
                    id SERIAL PRIMARY KEY,
                    empleado_id INTEGER NOT NULL REFERENCES empleados(id) ON DELETE CASCADE,
                    nombre TEXT NOT NULL,
                    dui TEXT NOT NULL,
                    fecha_salida DATE NOT NULL,
                    motivo TEXT NOT NULL,
                    salario_mensual NUMERIC(12,2) NOT NULL DEFAULT 0,
                    dias_trabajados INTEGER NOT NULL DEFAULT 0,
                    salario_pendiente NUMERIC(12,2) NOT NULL DEFAULT 0,
                    dias_vacaciones NUMERIC(10,2) NOT NULL DEFAULT 0,
                    vacaciones_pendientes NUMERIC(12,2) NOT NULL DEFAULT 0,
                    dias_decimo NUMERIC(10,2) NOT NULL DEFAULT 0,
                    decimo_proporcional NUMERIC(12,2) NOT NULL DEFAULT 0,
                    preaviso NUMERIC(12,2) NOT NULL DEFAULT 0,
                    indemnizacion NUMERIC(12,2) NOT NULL DEFAULT 0,
                    otros_ingresos NUMERIC(12,2) NOT NULL DEFAULT 0,
                    css NUMERIC(12,2) NOT NULL DEFAULT 0,
                    seguro_educativo NUMERIC(12,2) NOT NULL DEFAULT 0,
                    isr NUMERIC(12,2) NOT NULL DEFAULT 0,
                    otros_descuentos NUMERIC(12,2) NOT NULL DEFAULT 0,
                    total_bruto NUMERIC(12,2) NOT NULL DEFAULT 0,
                    total_deducciones NUMERIC(12,2) NOT NULL DEFAULT 0,
                    total_neto NUMERIC(12,2) NOT NULL DEFAULT 0,
                    observaciones TEXT,
                    estado TEXT NOT NULL DEFAULT 'BORRADOR',
                    creado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    actualizado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_liquidaciones_empleado ON liquidaciones(empleado_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_liquidaciones_estado ON liquidaciones(estado)")
            self.conexion.commit()

    @staticmethod
    def _numero(valor):
        try:
            return max(0.0, float(str(valor or 0).replace("B/.", "").replace(",", "").strip()))
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _redondear(valor):
        return round(max(0.0, float(valor or 0)), 2)

    @staticmethod
    def _fecha(valor):
        if hasattr(valor, "strftime"):
            return valor
        texto = str(valor or "").strip()
        for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
        raise ValueError("La fecha de salida no es válida.")

    @classmethod
    def calcular_liquidacion(cls, salario_mensual, dias_trabajados=0, dias_vacaciones=0,
                             dias_decimo=0, preaviso=0, indemnizacion=0,
                             otros_ingresos=0, otros_descuentos=0,
                             incluir_deducciones=True, causal_terminacion="DESPIDO_INJUSTIFICADO"):
        """
        Calcula liquidación según Código de Trabajo de Panamá.
        
        INGRESOS:
        - Salario pendiente: Diario × días trabajados
        - Vacaciones: Diario × días ganados
        - Preaviso: Mínimo 30 días según ley
        - Décimo proporcional: (Anual ÷ 360) × días
        - Indemnización: Según causal de terminación
        
        BASE GRAVABLE: Salario pendiente + vacaciones + preaviso + otros ingresos
        EXENTOS: Décimo, Indemnización
        
        DEDUCCIONES:
        - CSS: 9.75%
        - Seguro educativo: 1.25%
        - ISR: Proporcional a días pagados
        
        CAUSALES (Código de Trabajo):
        - DESPIDO_INJUSTIFICADO: Derecho a indemnización
        - CIERRE_EMPRESA: Mayor indemnización
        - CAUSAS_JUSTIFICADAS: Sin indemnización
        - RENUNCIA: Sin indemnización
        """
        salario = cls._numero(salario_mensual)
        diario = salario / 30
        salario_pendiente = diario * max(0, float(dias_trabajados or 0))
        vacaciones = diario * max(0, float(dias_vacaciones or 0))
        decimo = salario / 360 * max(0, float(dias_decimo or 0))
        preaviso_monto = cls._numero(preaviso)
        indemnizacion_monto = cls._numero(indemnizacion)
        otros_ingresos_monto = cls._numero(otros_ingresos)
        
        # Validación legal: Indemnización según causal (Código de Trabajo Panamá)
        causal = str(causal_terminacion or "DESPIDO_INJUSTIFICADO").upper().strip()
        advertencias = []
        
        # Validar que indemnización sea apropiada para la causal
        if causal in ["CAUSAS_JUSTIFICADAS", "RENUNCIA"]:
            if indemnizacion_monto > 0:
                advertencias.append(f"ADVERTENCIA LEGAL: Indemnización NO procede para '{causal}'")
                indemnizacion_monto = 0
        
        # Validar preaviso mínimo de 30 días si aplica
        if preaviso_monto > 0:
            dias_preaviso = preaviso_monto / diario if diario > 0 else 0
            if dias_preaviso < 30:
                advertencias.append(f"ADVERTENCIA: Preaviso de {dias_preaviso:.0f} días (mínimo legal: 30 días)")
        
        ingresos_adicionales = preaviso_monto + indemnizacion_monto + otros_ingresos_monto
        total_bruto = salario_pendiente + vacaciones + decimo + ingresos_adicionales

        # Base gravable incluye: salario pendiente + vacaciones + preaviso + otros_ingresos
        # Excluye: décimo (exento), indemnización (exenta)
        base_gravable = salario_pendiente + vacaciones + preaviso_monto + otros_ingresos_monto
        
        if incluir_deducciones:
            css = base_gravable * 0.0975
            seguro = base_gravable * 0.0125
            
            # ISR proporcional a días pagados, no a mes completo
            # Calcula base anual teórica si trabajara los 30 días del mes
            dias_pagados = max(0, float(dias_trabajados or 0)) + max(0, float(dias_vacaciones or 0))
            if dias_pagados > 0:
                salario_mensual_equivalente = (diario * 30)
                base_anual_teórica = salario_mensual_equivalente * 12
                
                # Aplica tasas ISR sobre base anual teórica
                if base_anual_teórica > 11000:
                    if base_anual_teórica <= 50000:
                        isr_anual = (base_anual_teórica - 11000) * 0.15
                    else:
                        isr_anual = 5850 + (base_anual_teórica - 50000) * 0.25
                    
                    # Proporciona ISR a los días realmente pagados
                    isr = (isr_anual / 360) * dias_pagados
                else:
                    isr = 0.0
            else:
                isr = 0.0
        else:
            css = 0.0
            seguro = 0.0
            isr = 0.0
        
        otros = cls._numero(otros_descuentos)
        total_deducciones = css + seguro + isr + otros
        
        return {
            "salario_mensual": cls._redondear(salario),
            "salario_pendiente": cls._redondear(salario_pendiente),
            "vacaciones_pendientes": cls._redondear(vacaciones),
            "decimo_proporcional": cls._redondear(decimo),
            "preaviso": cls._redondear(preaviso_monto),
            "indemnizacion": cls._redondear(indemnizacion_monto),
            "otros_ingresos": cls._redondear(otros_ingresos_monto),
            "base_gravable": cls._redondear(base_gravable),
            "css": cls._redondear(css),
            "seguro_educativo": cls._redondear(seguro),
            "isr": cls._redondear(isr),
            "otros_descuentos": cls._redondear(otros),
            "total_bruto": cls._redondear(total_bruto),
            "total_deducciones": cls._redondear(total_deducciones),
            "total_neto": cls._redondear(total_bruto - total_deducciones),
            "causal_terminacion": causal,
            "advertencias": advertencias,
        }

    def obtener_datos_automaticos(self, empleado, fecha_salida):
        """Obtiene los días que pueden determinarse con la información registrada."""
        try:
            if not empleado or len(empleado) < 11:
                raise ValueError("El empleado seleccionado no es válido.")
            fecha_salida = self._fecha(fecha_salida)
            fecha_ingreso = self._fecha(empleado[9])
            if fecha_salida < fecha_ingreso:
                raise ValueError("La fecha de salida no puede ser anterior al ingreso.")

            with self.conexion.cursor() as cursor:
                cursor.execute("""
                    SELECT COALESCE(MAX(fecha_fin), %s)
                    FROM cortes_planilla
                    WHERE estado = 'CERRADO' AND fecha_fin < %s
                """, (fecha_ingreso, fecha_salida))
                inicio_pendiente = cursor.fetchone()[0]
                if inicio_pendiente < fecha_salida:
                    inicio_pendiente += timedelta(days=1)

                cursor.execute("""
                    SELECT COUNT(*)
                    FROM marcaciones
                    WHERE dui = %s AND fecha BETWEEN %s AND %s
                      AND upper(coalesce(estado_dia, 'JORNADA LABORAL'))
                          NOT IN ('AUSENCIA', 'DÍA LIBRE', 'DIA LIBRE')
                """, (empleado[2], inicio_pendiente, fecha_salida))
                dias_trabajados = int(cursor.fetchone()[0] or 0)

                cursor.execute("""
                    SELECT COALESCE(SUM(dias), 0)
                    FROM solicitudes_vacaciones
                    WHERE empleado_id = %s AND estado = 'Aprobada'
                """, (empleado[0],))
                dias_vacaciones_usados = int(cursor.fetchone()[0] or 0)

                cursor.execute("""
                    SELECT COALESCE(SUM(CAST(monto AS NUMERIC)), 0)
                    FROM descuentos
                    WHERE dui = %s
                        AND estado IN ('ACTIVO', 'PAGADO')
                        AND fecha_creacion::date BETWEEN %s AND %s
                """, (empleado[2], inicio_pendiente, fecha_salida))
                otros_descuentos = self._numero(cursor.fetchone()[0])

                cursor.execute("""
                    SELECT COALESCE(SUM(monto), 0)
                    FROM remuneraciones
                    WHERE lower(nombre) = lower(%s)
                        AND fecha::date BETWEEN %s AND %s
                        AND lower(tipo) NOT IN ('salario', 'sueldo')
                """, (empleado[1], inicio_pendiente, fecha_salida))
                otros_ingresos = self._numero(cursor.fetchone()[0])

            fecha_inicio_decimo = max(fecha_ingreso, date(fecha_salida.year, ((fecha_salida.month - 1) // 4) * 4 + 1, 1))
            dias_decimo = min(120, max(0, (fecha_salida - fecha_inicio_decimo).days + 1))
            meses_antiguedad = (fecha_salida.year - fecha_ingreso.year) * 12 + fecha_salida.month - fecha_ingreso.month
            if fecha_salida.day < fecha_ingreso.day:
                meses_antiguedad -= 1
            if meses_antiguedad < 3:
                dias_vacaciones_ganados = 0
            elif meses_antiguedad < 6:
                dias_vacaciones_ganados = 7
            elif meses_antiguedad < 9:
                dias_vacaciones_ganados = 15
            elif meses_antiguedad < 12:
                dias_vacaciones_ganados = 22
            else:
                dias_vacaciones_ganados = 30 + ((meses_antiguedad - 12) // 12) * 30
            return {
                "dias_trabajados": dias_trabajados,
                "dias_vacaciones": max(0, dias_vacaciones_ganados - dias_vacaciones_usados),
                "dias_decimo": dias_decimo,
                "preaviso": 0.0,
                "indemnizacion": 0.0,
                "otros_ingresos": otros_ingresos,
                "otros_descuentos": otros_descuentos,
                "inicio_periodo_pendiente": inicio_pendiente,
            }
        except Exception:
            raise

    def guardar_liquidacion(self, empleado, fecha_salida, motivo, dias_trabajados=0,
                            dias_vacaciones=0, dias_decimo=0, preaviso=0,
                            indemnizacion=0, otros_ingresos=0, otros_descuentos=0,
                            observaciones="", estado="BORRADOR"):
        if not empleado or len(empleado) < 11:
            raise ValueError("El empleado seleccionado no es válido.")
        fecha_salida = self._fecha(fecha_salida)
        motivo = str(motivo or "").strip()
        if not motivo:
            raise ValueError("Debe indicar el motivo de salida.")
        calculo = self.calcular_liquidacion(
            empleado[8], dias_trabajados, dias_vacaciones, dias_decimo,
            preaviso, indemnizacion, otros_ingresos, otros_descuentos,
        )
        valores = (
            empleado[0], empleado[1], empleado[2], fecha_salida, motivo,
            calculo["salario_mensual"], dias_trabajados, calculo["salario_pendiente"],
            dias_vacaciones, calculo["vacaciones_pendientes"], dias_decimo,
            calculo["decimo_proporcional"], calculo["preaviso"], calculo["indemnizacion"],
            calculo["otros_ingresos"], calculo["css"], calculo["seguro_educativo"],
            calculo["isr"], calculo["otros_descuentos"], calculo["total_bruto"],
            calculo["total_deducciones"], calculo["total_neto"], observaciones, estado,
        )
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO liquidaciones(
                    empleado_id, nombre, dui, fecha_salida, motivo, salario_mensual,
                    dias_trabajados, salario_pendiente, dias_vacaciones, vacaciones_pendientes,
                    dias_decimo, decimo_proporcional, preaviso, indemnizacion, otros_ingresos,
                    css, seguro_educativo, isr, otros_descuentos, total_bruto,
                    total_deducciones, total_neto, observaciones, estado
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                          %s, %s, %s, %s, %s, %s, %s, %s, %s)
                RETURNING id
            """, valores)
            liquidacion_id = cursor.fetchone()[0]
            self.conexion.commit()

        servicio_empleados = gestion_empleado()
        servicio_empleados.actualizar_estado_empleado(
            empleado[0],
            "LIQUIDADO",
            fecha_liquidacion=fecha_salida,
            motivo=motivo,
        )
        return liquidacion_id, calculo

    def obtener_liquidaciones(self, estado=None):
        with self.conexion.cursor() as cursor:
            consulta = """
                SELECT id, empleado_id, nombre, dui, fecha_salida, motivo,
                       total_bruto, total_deducciones, total_neto, estado, creado_at
                FROM liquidaciones
            """
            parametros = ()
            if estado:
                consulta += " WHERE estado = %s"
                parametros = (estado,)
            consulta += " ORDER BY creado_at DESC, id DESC"
            cursor.execute(consulta, parametros)
            return cursor.fetchall()

    def obtener_liquidacion(self, liquidacion_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM liquidaciones WHERE id = %s", (liquidacion_id,))
            return cursor.fetchone()

    def actualizar_estado_liquidacion(self, liquidacion_id, estado):
        estados = {"BORRADOR", "CALCULADA", "APROBADA", "PAGADA", "ANULADA"}
        estado = str(estado or "").strip().upper()
        if estado not in estados:
            raise ValueError("Estado de liquidación no válido.")
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                UPDATE liquidaciones
                SET estado = %s, actualizado_at = CURRENT_TIMESTAMP
                WHERE id = %s
            """, (estado, liquidacion_id))
            actualizado = cursor.rowcount > 0
            self.conexion.commit()
            return actualizado

    def eliminar_liquidacion(self, liquidacion_id):
        """Elimina una liquidación registrada por su identificador."""
        with self.conexion.cursor() as cursor:
            cursor.execute(
                "SELECT empleado_id FROM liquidaciones WHERE id = %s",
                (liquidacion_id,),
            )
            fila = cursor.fetchone()
            if fila is None:
                return False

            empleado_id = fila[0]
            cursor.execute("DELETE FROM liquidaciones WHERE id = %s", (liquidacion_id,))
            eliminado = cursor.rowcount > 0
            if eliminado:
                cursor.execute(
                    """
                    UPDATE empleados
                    SET estado = 'ACTIVO',
                        fecha_liquidacion = NULL,
                        motivo_liquidacion = NULL
                    WHERE id = %s
                      AND NOT EXISTS (
                          SELECT 1
                          FROM liquidaciones
                          WHERE empleado_id = %s
                      )
                    """,
                    (empleado_id, empleado_id),
                )
            self.conexion.commit()
            return eliminado

class decimoXIII:
    def __init__(self):
        self.conexion = obtener_conexion()
        inicializar_tabla("decimoxiii", self.crear_tabla_decimoxiii)

    def crear_tabla_decimoxiii(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS decimoxiii(
                    id SERIAL PRIMARY KEY,
                    empleado_id INTEGER,
                    nombre TEXT NOT NULL,
                    dui TEXT NOT NULL,
                    departamento TEXT NOT NULL DEFAULT '',
                    periodo_inicio DATE NOT NULL,
                    periodo_fin DATE NOT NULL,
                    dias NUMERIC(10, 2) NOT NULL DEFAULT 0,
                    salario_computable NUMERIC(12, 2) NOT NULL DEFAULT 0,
                    decimo_bruto NUMERIC(12, 2) NOT NULL DEFAULT 0,
                    creado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    actualizado_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (dui, periodo_inicio, periodo_fin)
                )
            """)
            self.conexion.commit()

    def guardar_decimo(self, empleado_id, nombre, dui, departamento, periodo_inicio,
                       periodo_fin, dias, salario_computable, decimo_bruto):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                INSERT INTO decimoxiii(
                    empleado_id, nombre, dui, departamento, periodo_inicio, periodo_fin,
                    dias, salario_computable, decimo_bruto
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (dui, periodo_inicio, periodo_fin) DO UPDATE SET
                    empleado_id = EXCLUDED.empleado_id,
                    nombre = EXCLUDED.nombre,
                    departamento = EXCLUDED.departamento,
                    dias = EXCLUDED.dias,
                    salario_computable = EXCLUDED.salario_computable,
                    decimo_bruto = EXCLUDED.decimo_bruto,
                    actualizado_at = CURRENT_TIMESTAMP
                RETURNING id
            """, (
                empleado_id, nombre, dui, departamento, periodo_inicio, periodo_fin,
                dias, salario_computable, decimo_bruto
            ))
            registro_id = cursor.fetchone()[0]
            self.conexion.commit()
            return registro_id

    def obtener_decimos(self, periodo_inicio=None, periodo_fin=None, dui=None):
        consulta = """
            SELECT id, empleado_id, nombre, dui, departamento, periodo_inicio,
                periodo_fin, dias, salario_computable, decimo_bruto, creado_at,
                actualizado_at
            FROM decimoxiii
            WHERE 1 = 1
        """
        parametros = []
        if periodo_inicio is not None:
            consulta += " AND periodo_inicio = %s"
            parametros.append(periodo_inicio)
        if periodo_fin is not None:
            consulta += " AND periodo_fin = %s"
            parametros.append(periodo_fin)
        if dui is not None:
            consulta += " AND dui = %s"
            parametros.append(dui)
        consulta += " ORDER BY nombre ASC, periodo_inicio DESC"
        with self.conexion.cursor() as cursor:
            cursor.execute(consulta, parametros)
            return cursor.fetchall()

    def obtener_decimo_id(self, registro_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM decimoxiii WHERE id = %s", (registro_id,))
            return cursor.fetchone()

    def eliminar_decimo(self, registro_id):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM decimoxiii WHERE id = %s", (registro_id,))
            eliminado = cursor.rowcount > 0
            self.conexion.commit()
            return eliminado

    def eliminar_periodo(self, periodo_inicio, periodo_fin):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                DELETE FROM decimoxiii
                WHERE periodo_inicio = %s AND periodo_fin = %s
            """, (periodo_inicio, periodo_fin))
            eliminados = cursor.rowcount
            self.conexion.commit()
            return eliminados

    def cerrar_conexion(self):
        self.conexion.close()


