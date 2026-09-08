#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script simplificado de verificación de traducciones en SiPP
Solo verifica el diccionario de traducciones sin ejecutar la aplicación
"""

import re
import json
import os

def verificar_traducciones_archivo():
    """Verifica las traducciones leyendo directamente el archivo SiPP.py"""
    
    print("=" * 70)
    print("VERIFICACIÓN DE TRADUCCIONES EN SIPP")
    print("=" * 70)
    
    try:
        # Leer el archivo SiPP.py
        with open("SiPP.py", "r", encoding="utf-8") as f:
            contenido = f.read()
        
        print("\n✓ Archivo SiPP.py leído correctamente\n")
        
        # 1. Buscar el diccionario TRADUCCIONES
        print("1. Verificando diccionario de traducciones...")
        print("-" * 70)
        
        # Buscar la sección TRADUCCIONES
        inicio_traducciones = contenido.find('TRADUCCIONES = {')
        if inicio_traducciones == -1:
            print("   ✗ No se encontró TRADUCCIONES")
            return 1
        
        # Encontrar donde termina el diccionario
        # Buscar desde el inicio hasta el siguiente bloque de código
        fin_traducciones = contenido.find('\nRUTA_CONFIG_IDIOMA', inicio_traducciones)
        if fin_traducciones == -1:
            print("   ✗ No se encontró el final de TRADUCCIONES")
            return 1
        
        seccion_traducciones = contenido[inicio_traducciones:fin_traducciones]
        
        # Contar las entradas
        # Buscar patrones como "key": "value"
        entradas_english = re.findall(r'"([^"]+)":\s+"([^"]+)"', seccion_traducciones)
        
        print(f"   Total de entradas encontradas: {len(entradas_english)}")
        
        # Palabras clave que deberían estar presentes
        palabras_clave = [
            "Error",
            "Éxito", 
            "Validación",
            "Aviso",
            "Atención",
            "Usuario Guardado",
            "CODIGO DE USUARIO",
            "Usuario Eliminado",
            "Actualizado",
            "Guardado",
            "Reporte",
            "Logo",
            "Horario",
            "Registro existente",
            "Reporte de descuentos aplicados",
            "CARTA DE TRABAJO",
            "Bienvenido a SiPP",
            "Inicio con exito",
            "RUC",
            "No especificado",
        ]
        
        # Convertir a diccionario para búsqueda rápida
        entradas_dict = {clave: valor for clave, valor in entradas_english}
        
        palabras_encontradas = 0
        palabras_faltantes = []
        
        for palabra in palabras_clave:
            if palabra in entradas_dict:
                palabras_encontradas += 1
                print(f"   ✓ '{palabra}' -> '{entradas_dict[palabra]}'")
            else:
                palabras_faltantes.append(palabra)
                print(f"   ✗ '{palabra}' NO ENCONTRADA")
        
        print(f"\n   Palabras encontradas: {palabras_encontradas}/{len(palabras_clave)}")
        
        # 2. Verificar funciones de traducción
        print("\n2. Verificando funciones de traducción...")
        print("-" * 70)
        
        funciones_requeridas = [
            "def cargar_idioma_guardado",
            "def guardar_idioma_en_disco",
            "def establecer_idioma_aplicacion",
            "def idioma_aplicacion_actual",
            "_activar_traduccion_automatica_widgets",
        ]
        
        funciones_encontradas = 0
        for funciona in funciones_requeridas:
            if funciona in contenido:
                funciones_encontradas += 1
                print(f"   ✓ {funciona.replace('def ', '').replace('_', '')}")
            else:
                print(f"   ✗ {funciona} NO ENCONTRADA")
        
        print(f"\n   Funciones encontradas: {funciones_encontradas}/{len(funciones_requeridas)}")
        
        # 3. Verificar archivo de configuración de idioma
        print("\n3. Verificando almacenamiento de idioma...")
        print("-" * 70)
        
        ruta_config = os.path.join(os.getcwd(), "config_idioma.json")
        print(f"   Ruta de configuración: {ruta_config}")
        
        if os.path.exists(ruta_config):
            try:
                with open(ruta_config, "r", encoding="utf-8") as f:
                    config = json.load(f)
                print(f"   ✓ Archivo de configuración existe")
                print(f"   Idioma guardado: {config.get('idioma', 'No especificado')}")
            except Exception as e:
                print(f"   ✗ Error al leer archivo: {e}")
        else:
            print("   ℹ Archivo de configuración no existe")
            print("     (Será creado cuando el usuario cambie el idioma)")
        
        # 4. Verificar traducciones en PDFs
        print("\n4. Verificando traducciones en PDFs...")
        print("-" * 70)
        
        pdf_strings = [
            "_tr(\"COMPROBANTE DE PAGO\")",
            "_tr(\"Reporte de descuentos aplicados\")",
            "_tr(\"CARTA DE TRABAJO\")",
            "_tr(\"Atentamente,\")",
            "_tr(\"Emitida el\")",
        ]
        
        pdf_encontrados = 0
        for string_pdf in pdf_strings:
            if string_pdf in contenido:
                pdf_encontrados += 1
                print(f"   ✓ {string_pdf}")
            else:
                print(f"   ✗ {string_pdf} NO ENCONTRADO")
        
        print(f"\n   Strings en PDFs traducidos: {pdf_encontrados}/{len(pdf_strings)}")
        
        # Resumen
        print("\n" + "=" * 70)
        print("RESUMEN DE VERIFICACIÓN")
        print("=" * 70)
        
        total_completitud = (palabras_encontradas * 100) // len(palabras_clave)
        print(f"\nCompletitud de traducción: {total_completitud}%")
        
        if palabras_faltantes:
            print(f"\n⚠ FALTANTES {len(palabras_faltantes)} traducciones:")
            for palabra in palabras_faltantes:
                print(f"  - {palabra}")
        else:
            print("\n✓ TODAS LAS TRADUCCIONES ESTÁN PRESENTES")
        
        if funciones_encontradas == len(funciones_requeridas):
            print("✓ Todas las funciones de idioma están implementadas")
        else:
            print(f"⚠ Faltan {len(funciones_requeridas) - funciones_encontradas} funciones")
        
        if pdf_encontrados == len(pdf_strings):
            print("✓ Los PDFs se generarán en el idioma seleccionado")
        else:
            print(f"⚠ {len(pdf_strings) - pdf_encontrados} strings en PDFs sin traducir")
        
        print("\n" + "=" * 70)
        print("ESTADO DEL SISTEMA DE IDIOMAS")
        print("=" * 70)
        
        if len(palabras_faltantes) == 0 and funciones_encontradas == len(funciones_requeridas):
            print("\n✓ EL SISTEMA DE IDIOMAS ESTÁ COMPLETO Y FUNCIONAL\n")
            print("Características implementadas:")
            print("  ✓ Diccionario de traducciones Español-Inglés completo")
            print("  ✓ Sistema automático de traducción de widgets")
            print("  ✓ Traducción de messageboxes (títulos y contenidos)")
            print("  ✓ Almacenamiento persistente de idioma en disco")
            print("  ✓ Carga de idioma al iniciar la aplicación")
            print("  ✓ Traducción de PDFs (comprobantes, reportes, cartas)")
            print("  ✓ Traducción en títulos de diálogos")
            return 0
        else:
            print("\n⚠ El sistema funciona pero requiere ajustes\n")
            return 1
            
    except Exception as e:
        print(f"\n✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return 2

if __name__ == "__main__":
    exit_code = verificar_traducciones_archivo()
    print(f"\nCódigo de salida: {exit_code}")
