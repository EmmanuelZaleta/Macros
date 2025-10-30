#!/usr/bin/env python3
"""
Script de Utilidad: Backup y Verificación de Macro_MPSTest.xlsm
Optimizado por Claude Code - 2025-10-30
"""

import os
import shutil
from datetime import datetime
import hashlib

def calcular_hash(archivo):
    """Calcula el hash MD5 de un archivo."""
    hash_md5 = hashlib.md5()
    try:
        with open(archivo, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        return f"Error: {e}"

def crear_backup():
    """Crea un backup del archivo original con timestamp."""
    archivo_original = "/home/user/Macros/Macro_MPSTest.xlsm"

    if not os.path.exists(archivo_original):
        print("❌ Error: No se encontró Macro_MPSTest.xlsm")
        return False

    # Crear nombre de backup con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archivo_backup = f"/home/user/Macros/Macro_MPSTest_BACKUP_{timestamp}.xlsm"

    try:
        print(f"📦 Creando backup...")
        shutil.copy2(archivo_original, archivo_backup)

        # Verificar que el backup se creó correctamente
        if os.path.exists(archivo_backup):
            size_orig = os.path.getsize(archivo_original)
            size_backup = os.path.getsize(archivo_backup)

            print(f"✅ Backup creado exitosamente:")
            print(f"   Original: {size_orig:,} bytes")
            print(f"   Backup:   {size_backup:,} bytes")
            print(f"   Ubicación: {archivo_backup}")

            # Calcular hashes para verificar integridad
            hash_orig = calcular_hash(archivo_original)
            hash_backup = calcular_hash(archivo_backup)

            if hash_orig == hash_backup:
                print(f"✅ Integridad verificada (MD5: {hash_orig[:8]}...)")
                return True
            else:
                print("⚠️  Advertencia: Los hashes no coinciden")
                return False
        else:
            print("❌ Error: El backup no se creó correctamente")
            return False

    except Exception as e:
        print(f"❌ Error al crear backup: {e}")
        return False

def verificar_archivos_optimizados():
    """Verifica que los archivos optimizados existan."""
    archivos_necesarios = [
        "/home/user/Macros/VBA_OPTIMIZADO/mdl_Query_OPTIMIZED.bas",
        "/home/user/Macros/VBA_OPTIMIZADO/ADODBProcess_OPTIMIZED.cls",
        "/home/user/Macros/VBA_OPTIMIZADO/mdl_Utilities_OPTIMIZED.bas"
    ]

    print("\n🔍 Verificando archivos optimizados...")
    todos_existen = True

    for archivo in archivos_necesarios:
        if os.path.exists(archivo):
            size = os.path.getsize(archivo)
            print(f"✅ {os.path.basename(archivo):<35} ({size:>6,} bytes)")
        else:
            print(f"❌ {os.path.basename(archivo):<35} (NO ENCONTRADO)")
            todos_existen = False

    return todos_existen

def mostrar_instrucciones():
    """Muestra instrucciones paso a paso."""
    print("\n" + "="*70)
    print("📋 INSTRUCCIONES DE IMPLEMENTACIÓN")
    print("="*70)
    print("""
Paso 1: Cerrar Excel si está abierto
   ⚠️  Asegúrate de que Macro_MPSTest.xlsm NO esté abierto

Paso 2: Abrir el archivo y VBA Editor
   1. Abrir Macro_MPSTest.xlsm
   2. Presionar Alt+F11

Paso 3: Importar módulos optimizados
   1. En VBA Editor: File → Import File...
   2. Importar estos archivos (en orden):
      ✅ mdl_Query_OPTIMIZED.bas
      ✅ ADODBProcess_OPTIMIZED.cls
      ✅ mdl_Utilities_OPTIMIZED.bas

Paso 4: Renombrar módulos
   1. Seleccionar "mdl_Query_OPTIMIZED" en Project Explorer
   2. En Propiedades (F4), cambiar Name a: mdl_Query
   3. Repetir para:
      - ADODBProcess_OPTIMIZED → ADODBProcess
      - mdl_Utilities_OPTIMIZED → mdl_Utilities

Paso 5: Guardar y probar
   1. Ctrl+S para guardar
   2. Cerrar VBA Editor
   3. ¡Probar las macros!

📄 Para más detalles, ver: GUIA_IMPLEMENTACION_RAPIDA.md
""")
    print("="*70)

def main():
    """Función principal."""
    print("\n" + "="*70)
    print("  SCRIPT DE BACKUP Y VERIFICACIÓN")
    print("  Macro_MPSTest.xlsm - Versión Optimizada")
    print("="*70 + "\n")

    # Paso 1: Crear backup
    backup_exitoso = crear_backup()

    if not backup_exitoso:
        print("\n⚠️  No se pudo crear el backup. ¿Deseas continuar? (s/n)")
        respuesta = input("> ").lower()
        if respuesta != 's':
            print("❌ Operación cancelada.")
            return

    # Paso 2: Verificar archivos optimizados
    archivos_ok = verificar_archivos_optimizados()

    if not archivos_ok:
        print("\n❌ Faltan archivos optimizados. Verifica la carpeta VBA_OPTIMIZADO.")
        return

    # Paso 3: Mostrar instrucciones
    mostrar_instrucciones()

    print("\n✅ Todo listo para implementar las mejoras!")
    print("\n💡 Tip: Guarda esta ventana abierta para consultar las instrucciones.")

if __name__ == "__main__":
    main()
