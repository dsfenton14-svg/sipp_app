#!/usr/bin/env python
"""
Script para limpiar la tabla de descuentos y resolver transacciones abortadas.
Ejecutar: python limpiar_descuentos.py
"""

import psycopg as pc
from tkinter import messagebox

def limpiar_tabla_descuentos():
    """Elimina todos los registros de la tabla descuentos y resetea la conexión"""
    try:
        # Conexión a la base de datos
        conexion = pc.connect(
            host="localhost",
            user="postgres",
            password="12sql@?",
            port="5432",
            dbname="SiPP-2"
        )
        
        print("✓ Conectado a la base de datos SiPP-2")
        
        with conexion.cursor() as cursor:
            # Primero, hacer rollback de cualquier transacción pendiente
            cursor.execute("ROLLBACK")
            print("✓ Rollback de transacciones pendientes")
            
            # Limpiar la tabla
            cursor.execute("DELETE FROM descuentos")
            print(f"✓ {cursor.rowcount} registros eliminados de la tabla descuentos")
            
            # Resetear la secuencia (ID vuelve a empezar desde 1)
            cursor.execute("ALTER SEQUENCE descuentos_id_seq RESTART WITH 1")
            print("✓ Secuencia de IDs reseteada")
            
            # Commit de los cambios
            conexion.commit()
            print("✓ Cambios guardados exitosamente")
        
        conexion.close()
        print("\n✅ La tabla de descuentos ha sido limpiada exitosamente")
        return True
        
    except pc.errors.InsufficientPrivilege as e:
        print(f"❌ Error de permisos: {e}")
        print("   Asegúrate de que el usuario 'postgres' tenga suficientes permisos")
        return False
    except pc.errors.UndefinedTable as e:
        print(f"❌ La tabla descuentos no existe: {e}")
        return False
    except Exception as e:
        print(f"❌ Error al limpiar la tabla: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("LIMPIEZA DE TABLA DE DESCUENTOS")
    print("=" * 60)
    print()
    
    resultado = limpiar_tabla_descuentos()
    
    if resultado:
        print("\n" + "=" * 60)
        print("Ahora puedes abrir SiPP.py y la tabla estará vacía.")
        print("=" * 60)
    else:
        print("\nIntenta resolver el problema y ejecuta nuevamente.")
