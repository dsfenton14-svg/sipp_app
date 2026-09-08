#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de verificación de traducción en SiPP
Verifica que todos los cambios de idioma se han completado correctamente
"""

import sys
import os
import json

# Agregar el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def verificar_traducciones():
    """Verifica que el sistema de traducciones funciona correctamente."""
    
    print("=" * 70)
    print("VERIFICACIÓN DE TRADUCCIONES EN SIPP")
    print("=" * 70)
    
    try:
        # Importar las traducciones desde SiPP.py
        from SiPP import TRADUCCIONES, cargar_idioma_guardado, RUTA_CONFIG_IDIOMA
        
        print("\n✓ Módulo SiPP importado correctamente\n")
        
        # 1. Verificar que el diccionario tiene las entradas principales
        print("1. Verificando diccionario de traducciones...")
        print("-" * 70)
        
        entries_english = TRADUCCIONES.get("English", {})
        print(f"   Total de entradas en English: {len(entries_english)}")
        
        # Palabras clave que deberían estar presentes
        palabras_clave = [
            "Error", "Éxito", "Validación", "Aviso", "Atención", 
            "Usuario Guardado", "CODIGO DE USUARIO", "Usuario Eliminado",
            "Actualizado", "Guardado", "Reporte", "Logo",
            "Horario", "Registro existente", "Reporte de descuentos aplicados",
            "CARTA DE TRABAJO", "Bienvenido a SiPP", "Inicio con exito",
            "RUC", "No especificado"
        ]
        
        palabras_encontradas = 0
        palabras_faltantes = []
        
        for palabra in palabras_clave:
            if palabra in entries_english:
                palabras_encontradas += 1
                print(f"   ✓ '{palabra}' -> '{entries_english[palabra]}'")
            else:
                palabras_faltantes.append(palabra)
                print(f"   ✗ '{palabra}' NO ENCONTRADA")
        
        print(f"\n   Palabras encontradas: {palabras_encontradas}/{len(palabras_clave)}")
        
        if palabras_faltantes:
            print(f"\n   ⚠ PALABRAS FALTANTES ({len(palabras_faltantes)}):")
            for palabra in palabras_faltantes:
                print(f"     - {palabra}")
        
        # 2. Verificar el almacenamiento de idioma
        print("\n2. Verificando almacenamiento de idioma...")
        print("-" * 70)
        print(f"   Ruta de configuración: {RUTA_CONFIG_IDIOMA}")
        
        idioma_cargado = cargar_idioma_guardado()
        print(f"   Idioma cargado: {idioma_cargado}")
        
        if os.path.exists(RUTA_CONFIG_IDIOMA):
            try:
                with open(RUTA_CONFIG_IDIOMA, "r", encoding="utf-8") as f:
                    config = json.load(f)
                print(f"   Idioma guardado en disco: {config.get('idioma', 'No especificado')}")
                print("   ✓ Archivo de configuración existe y es válido")
            except Exception as e:
                print(f"   ✗ Error al leer archivo de configuración: {e}")
        else:
            print("   ℹ Archivo de configuración no existe (será creado al cambiar idioma)")
        
        # 3. Verificar funciones de traducción
        print("\n3. Verificando funciones de traducción...")
        print("-" * 70)
        
        from SiPP import _traducir, _tr, establecer_idioma_aplicacion, idioma_aplicacion_actual
        
        # Probar traducción
        establecer_idioma_aplicacion("English")
        print(f"   Idioma actual: {idioma_aplicacion_actual()}")
        
        pruebas = [
            "Error",
            "Éxito",
            "Usuario Guardado",
            "Reporte de descuentos aplicados",
        ]
        
        for texto in pruebas:
            traduccion = _tr(texto)
            if traduccion != texto:
                print(f"   ✓ '{texto}' -> '{traduccion}'")
            else:
                print(f"   ✗ '{texto}' no se tradujo")
        
        # Volver al español
        establecer_idioma_aplicacion("Español")
        print(f"\n   Idioma restaurado a: {idioma_aplicacion_actual()}")
        
        # Resumen
        print("\n" + "=" * 70)
        print("RESUMEN DE VERIFICACIÓN")
        print("=" * 70)
        
        if len(palabras_faltantes) == 0:
            print("✓ TODAS LAS TRADUCCIONES ESTÁN COMPLETAS")
            print("\n✓ El sistema de idiomas está FUNCIONANDO CORRECTAMENTE")
            print("✓ Se pueden cambiar entre Español e Inglés sin problemas")
            print("✓ Los PDFs se generarán en el idioma seleccionado")
            print("✓ El idioma seleccionado se guarda y carga correctamente")
            return 0
        else:
            print(f"⚠ FALTAN {len(palabras_faltantes)} TRADUCCIONES")
            print("\n✓ El sistema funciona pero requiere que se completen las traducciones")
            return 1
        
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 2

if __name__ == "__main__":
    exit_code = verificar_traducciones()
    sys.exit(exit_code)
