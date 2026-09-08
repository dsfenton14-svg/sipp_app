# Análisis de Congelamiento de SiPP

## Problemas Identificados

### 1. **Conexiones Múltiples a la Base de Datos (CRÍTICO)**

**Ubicación:** `SiPP.py` - función `cargar_datos()` línea 2621

**Problema:**
```python
def cargar_datos(self, texto_busqueda="", departamento="Todos"):
    self._actualizar_departamentos_combo()  # ← Conexión #1
    # ...
    datos = db.gestion_empleado().obtener_empleados()  # ← Conexión #2
```

Cada llamada a `db.gestion_empleado()` crea una **NUEVA instancia** que abre una **NUEVA conexión** a PostgreSQL en el `__init__`:

```python
class gestion_empleado:
    def __init__(self):
        self.conexion = pc.connect(...)  # ← Espera a conectar
        self.crear_tabla_empleados()     # ← Otra operación lenta
```

**Impacto:** 
- Cuando escribes en la caja de búsqueda, se dispara `buscar_por_escritura()` 
- Esto llama a `cargar_datos()` que hace **2 conexiones simultáneamente**
- Si PostgreSQL está lento o no está disponible, la interfaz se congela durante 5 segundos por conexión

### 2. **Caída en Cascada de Secciones**

**Ubicación:** Todas las secciones en `SiPP.py` (líneas 2200+)

Cada sección sigue el patrón:
```python
def gestion_personal(self, sipp):
    self._crear_area_contenido(sipp)
    gestionar = Gestion_Personal()    # ← Crea instancia
    gestionar.ventana_gestion_personal(self.frame2)  # ← Llama al método
```

En `ventana_gestion_personal()`, se llama a `_actualizar_departamentos_combo()` **inmediatamente**, lo que hace que:
1. El cambio de sección se congele mientras espera la conexión
2. El usuario ve una interfaz congelada sin respuesta

### 3. **Operaciones de Base de Datos en el Hilo Principal**

**Ubicación:** Todo el archivo `SiPP.py`

**Problema:** Todas las operaciones de base de datos se ejecutan en el hilo principal de Tkinter:
```python
# Esto congela la interfaz mientras se ejecuta:
datos = db.gestion_empleado().obtener_empleados()
```

Las operaciones de red/BD pueden tardar segundos, paralizando toda la interfaz.

---

## Soluciones Recomendadas

### Solución 1: Pool de Conexiones (Recomendado)

Implementar un pool de conexiones reutilizable en `datos_Sipp.py`:

```python
from psycopg_pool import ConnectionPool

class BaseDatos:
    _pool = None
    
    @classmethod
    def obtener_pool(cls):
        if cls._pool is None:
            cls._pool = ConnectionPool(
                "postgresql://postgres:12sql@?@localhost:5432/SiPP-2",
                min_size=2,
                max_size=10
            )
        return cls._pool

class gestion_empleado:
    def obtener_empleados(self):
        pool = BaseDatos.obtener_pool()
        with pool.connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute("SELECT * FROM empleados")
                return cursor.fetchall()
```

**Ventajas:**
- Reutiliza conexiones, no las crea nuevas cada vez
- Mejora drásticamente el rendimiento

### Solución 2: Caché de Datos

Cachear los datos de empleados en la aplicación:

```python
class GestionPersonal:
    _cache_empleados = None
    _cache_timestamp = None
    
    @staticmethod
    def obtener_empleados(force_refresh=False):
        ahora = time.time()
        if not force_refresh and GestionPersonal._cache_empleados and \
           (ahora - GestionPersonal._cache_timestamp) < 60:  # Cache 60 segundos
            return GestionPersonal._cache_empleados
        
        datos = db.gestion_empleado().obtener_empleados()
        GestionPersonal._cache_empleados = datos
        GestionPersonal._cache_timestamp = ahora
        return datos
```

### Solución 3: Operaciones Asincrónicas

Mover operaciones de BD a un hilo separado:

```python
import threading

def cargar_datos_async(self, callback):
    def cargar():
        datos = db.gestion_empleado().obtener_empleados()
        self.after(0, callback, datos)
    
    threading.Thread(target=cargar, daemon=True).start()

# Uso:
def en_cargar_datos_completado(datos):
    # Actualizar UI con datos
    self.tabla_personal.insert(...datos...)

self.cargar_datos_async(en_cargar_datos_completado)
```

---

## Pasos para Verificar

1. **Abre la aplicación y ve a "Gestión Personal"** → ¿Se congela?
2. **Escribe en la caja de búsqueda** → ¿Cada letra causa congelamiento?
3. **Cambia de sección** → ¿Se demora más de 1-2 segundos?
4. **Verifica el servidor PostgreSQL**:
   ```bash
   psql -U postgres -d postgres
   ```

Si PostgreSQL no está disponible o es lento, todo se congela.

---

## Recomendación Inmediata

**Implementar la Solución 1 (Pool de Conexiones)** es la más efectiva:
- Requiere cambios mínimos en el código actual
- Reduce drasticamente el tiempo de conexión
- Es la práctica estándar en Python

¿Quieres que implemente la solución del pool de conexiones?
