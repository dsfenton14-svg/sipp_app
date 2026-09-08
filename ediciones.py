import psycopg as pc
import sys


class Tema:
    def __init__(self):
        try:
            self.conexion = pc.connect(
                host="localhost",
                user="postgres",
                password="12sql@?",
                port="5432",
                dbname="SiPP-2"
            )
        except Exception as e:
            print("Error al conectar a la base de datos:\n", e)
            raise
        # Ensure required tables exist
        self.ensure_tables()
        
    def tabla_tema(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS tema_sipp(
                    tema text)""")
            self.conexion.commit()
            cursor.close()
            self.datos_tema()

    def datos_tema(self):
        obtener = self.obtener_tema()
        if not obtener:
            datos = ["extreme"]
            self.insertar_tema(datos)
        else:
            pass
            
    def insertar_tema(self, datos):
        # aceptar tanto una lista/tupla con un elemento como un string
        if isinstance(datos, (list, tuple)):
            params = (datos[0],) if len(datos) > 0 else (None,)
        else:
            params = (datos,)
        with self.conexion.cursor() as cursor:
            cursor.execute(
        "INSERT INTO tema_sipp (tema) VALUES (%s)", params)
            self.conexion.commit()
            cursor.close()
    
    def insertar(self, texto):
        with self.conexion.cursor() as cursor:
            cursor.execute(
    "INSERT INTO tema_sipp(tema) VALUES (%s)",
    (texto,)  # ← importante la coma
)

            self.conexion.commit()
            cursor.close()


    def limpiar(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM tema_sipp")
            self.conexion.commit()
            cursor.close()

    def obtener_tema(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("SELECT * FROM tema_sipp")
            usuarios = cursor.fetchall()
            cursor.close()
            return usuarios
        
    def tabla_vista(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS vista_sipp(
                    tema text)""")
            self.conexion.commit()
            cursor.close()
            self.datos_vista()

    def ensure_tables(self):
        """Crear las tablas necesarias si no existen."""
        try:
            self.tabla_tema()
            self.tabla_vista()
        except Exception as e:
            print("Error al asegurar tablas:", e)
            raise

    def datos_vista(self):
        obtener = self.obtener_vista()
        if not obtener:
            datos = ["system"]
            self.insertar_vista(datos)
        else:
            pass

    def insertar_vista(self, datos):
        # aceptar tanto lista/tupla con elemento como string
        if isinstance(datos, (list, tuple)):
            params = (datos[0],) if len(datos) > 0 else (None,)
        else:
            params = (datos,)
        with self.conexion.cursor() as cursor:
            cursor.execute("INSERT INTO vista_sipp (tema) VALUES (%s)", params)
            self.conexion.commit()
            cursor.close()

    def insert_vista(self, texto):
        with self.conexion.cursor() as cursor:
            cursor.execute(
    "INSERT INTO vista_sipp(tema) VALUES (%s)",
    (texto,)  # ← importante la coma
)

            self.conexion.commit()
            cursor.close()
    
    def limpiar_vista(self):
        with self.conexion.cursor() as cursor:
            cursor.execute("DELETE FROM vista_sipp")
            self.conexion.commit()
            cursor.close()

    def obtener_vista(self):
        try:
            with self.conexion.cursor() as cursor:
                cursor.execute("SELECT * FROM vista_sipp")
                usuarios = cursor.fetchall()
                cursor.close()
                return usuarios
        except Exception as e:
            # Detectar tabla inexistente y dar instrucciones claras
            if hasattr(pc, 'errors') and isinstance(e, pc.errors.UndefinedTable):
                print("La tabla 'vista_sipp' no existe. Ejecuta 'python ediciones.py' para crear las tablas.")
            raise


def crear_tablas():
    """Función utilitaria para crear las tablas ejecutando este archivo.

    Ejecutar desde la línea de comandos cuando la DB está accesible:
        python ediciones.py
    """
    try:
        Tema()
        print("Tablas creadas / aseguradas correctamente.")
    except Exception as e:
        print("No se pudieron crear las tablas:\n", e)
        sys.exit(1)

    
    