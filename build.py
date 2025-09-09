#!/usr/bin/env python3
"""
Script para construir el ejecutable del Extractor PDF a Excel usando PyInstaller
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# Ruta específica de PyInstaller
PYINSTALLER_PATH = r"C:\Users\jsantillana\AppData\Local\Programs\Python\Python312\Scripts\pyinstaller.exe"

def check_dependencies():
    """Verifica que las dependencias necesarias estén instaladas"""
    required_packages = [
        'tkinter',  # Viene con Python, pero verificamos que esté disponible
        'pandas',
        'pdfplumber',
        'xlsxwriter',
        'openpyxl'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            if package == 'tkinter':
                import tkinter
            else:
                __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    # Verificar que PyInstaller esté disponible en la ruta especificada
    if not os.path.exists(PYINSTALLER_PATH):
        print(f"❌ PyInstaller no encontrado en: {PYINSTALLER_PATH}")
        print("   💡 Verifica la ruta o instala PyInstaller con: pip install pyinstaller")
        return False
    else:
        print(f"✅ PyInstaller encontrado en: {PYINSTALLER_PATH}")
    
    if missing_packages:
        print(f"❌ Paquetes faltantes: {', '.join(missing_packages)}")
        print("Instálalos con: pip install " + " ".join([p for p in missing_packages if p != 'tkinter']))
        if 'tkinter' in missing_packages:
            print("⚠️  tkinter no está disponible. En Linux instala: sudo apt-get install python3-tk")
        return False
    
    print("✅ Todas las dependencias están instaladas")
    return True

def check_required_files():
    """Verifica que los archivos necesarios existan"""
    required_files = [
        'app.py',
        'test_pdf.py'
    ]
    
    missing_files = []
    
    # Verificar archivos
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"❌ Archivos faltantes: {', '.join(missing_files)}")
        return False
    
    print("✅ Todos los archivos necesarios están presentes")
    return True

def create_spec_file():
    """Crea el archivo .spec para PyInstaller"""
    spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('test_pdf.py', '.'),
    ],
    hiddenimports=[
        'pandas',
        'pdfplumber',
        'xlsxwriter',
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'threading',
        'logging',
        're'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExtractorPDF',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Sin ventana de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/icon.ico' if os.path.exists('assets/icon.ico') else None,
)
'''
    
    with open('ExtractorPDF.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content.strip())
    
    print("✅ Archivo .spec creado: ExtractorPDF.spec")

def prepare_build_environment():
    """Prepara el entorno para la construcción"""
    print("🔧 Preparando entorno de construcción...")
    
    # Crear directorios necesarios
    directories = ['assets', 'output', 'logs']
    
    for dir_name in directories:
        os.makedirs(dir_name, exist_ok=True)
        print(f"   📁 Directorio {dir_name} listo")
    
    # Crear un icono simple si no existe (opcional)
    icon_path = Path('assets/icon.ico')
    if not icon_path.exists():
        print(f"⚠️  Icono no encontrado en {icon_path}")
        print("   💡 Puedes agregar un archivo icon.ico en assets/ para personalizar el icono")
    
    # Crear archivo .spec
    create_spec_file()
    
    print("✅ Entorno preparado")

def clean_previous_builds():
    """Limpia construcciones anteriores"""
    print("🧹 Limpiando construcciones anteriores...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['ExtractorPDF.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   🗑️  Eliminado {dir_name}")
    
    # Limpiar archivos .pyc recursivamente
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.pyc'):
                os.remove(os.path.join(root, file))
        # También limpiar directorios __pycache__
        if '__pycache__' in dirs:
            shutil.rmtree(os.path.join(root, '__pycache__'))
    
    print("✅ Limpieza completada")

def build_executable():
    """Construye el ejecutable usando PyInstaller"""
    print("🚀 Iniciando construcción del ejecutable...")
    print(f"   🔧 Usando PyInstaller desde: {PYINSTALLER_PATH}")
    
    # Comando de PyInstaller usando la ruta específica
    cmd = [
        PYINSTALLER_PATH,
        'ExtractorPDF.spec',
        '--clean',
        '--noconfirm'
    ]
    
    try:
        print("   ⏳ Ejecutando PyInstaller...")
        print(f"   📋 Comando: {' '.join(cmd)}")
        
        # Ejecutar con output en tiempo real
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            universal_newlines=True
        )
        
        # Mostrar output en tiempo real
        for line in process.stdout:
            print(f"      {line.rstrip()}")
        
        process.wait()
        
        if process.returncode == 0:
            print("✅ Construcción completada exitosamente")
            return True
        else:
            print(f"❌ Error durante la construcción. Código de salida: {process.returncode}")
            return False
            
    except FileNotFoundError:
        print(f"❌ Error: No se pudo encontrar PyInstaller en {PYINSTALLER_PATH}")
        print("   💡 Verifica que la ruta sea correcta o instala PyInstaller")
        return False
    except Exception as e:
        print(f"❌ Error inesperado durante la construcción: {e}")
        return False

def verify_executable():
    """Verifica que el ejecutable se haya creado correctamente"""
    exe_path = Path('dist/ExtractorPDF.exe')
    
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"✅ Ejecutable creado: {exe_path}")
        print(f"   📊 Tamaño: {size_mb:.1f} MB")
        return True
    else:
        print(f"❌ Ejecutable no encontrado en {exe_path}")
        # Listar contenido de dist para debugging
        dist_dir = Path('dist')
        if dist_dir.exists():
            print("   📁 Contenido de dist/:")
            for item in dist_dir.iterdir():
                print(f"      • {item.name}")
        return False

def post_build_setup():
    """Configuración post-construcción"""
    print("🔧 Configuración post-construcción...")
    
    dist_dir = Path('dist')
    
    # Crear directorio de salida en dist
    output_dir = dist_dir / 'output'
    output_dir.mkdir(exist_ok=True)
    print(f"   📁 Creado directorio de salida: {output_dir}")
    
    # Crear archivo README.txt con instrucciones
    readme_content = """
🏦 EXTRACTOR PDF A EXCEL - Balance de Comprobación
================================================

INSTRUCCIONES DE USO:
1. Ejecuta ExtractorPDF.exe
2. Selecciona tu archivo PDF del balance de comprobación
3. Elige dónde guardar el archivo Excel
4. Haz clic en "Procesar PDF"
5. Los archivos se guardarán en la ubicación que elijas

REQUISITOS:
• Archivos PDF del Banco de la Nación (Balance de Comprobación)
• Sistema operativo Windows

NOTAS:
• Los archivos Excel generados tendrán formato .xlsx
• Se incluye una hoja de resumen con estadísticas
• El proceso puede tomar varios minutos dependiendo del tamaño del PDF
• Los montos con "CR" indican valores de crédito

CARACTERÍSTICAS:
• Procesa líneas con y sin nombre de cuenta
• Detecta automáticamente valores faltantes
• Calcula balances y validaciones
• Genera hoja de resumen con estadísticas

SOPORTE:
• Si encuentras algún problema, revisa el log de proceso en la aplicación
• Asegúrate de que el PDF no esté corrupto o protegido con contraseña

© 2025 - Extractor PDF Balance de Comprobación
    """.strip()
    
    readme_path = dist_dir / 'README.txt'
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"   📄 Creado archivo de ayuda: {readme_path}")
    
    # Crear script de prueba (opcional)
    test_script = '''@echo off
echo 🧪 Probando ExtractorPDF...
echo.
echo Ejecutando ExtractorPDF.exe...
ExtractorPDF.exe
echo.
echo ✅ Prueba completada
pause
'''
    
    test_path = dist_dir / 'test.bat'
    with open(test_path, 'w', encoding='utf-8') as f:
        f.write(test_script)
    print(f"   🧪 Creado script de prueba: {test_path}")
    
    print("✅ Configuración post-construcción completada")

def create_installer_script():
    """Crea un script opcional para crear un instalador"""
    installer_content = '''@echo off
echo 🎁 Creando paquete de distribución...

set PACKAGE_NAME=ExtractorPDF_v1.0
set PACKAGE_DIR=%PACKAGE_NAME%

if exist %PACKAGE_DIR% rmdir /s /q %PACKAGE_DIR%
mkdir %PACKAGE_DIR%

echo 📦 Copiando archivos...
copy /Y ExtractorPDF.exe %PACKAGE_DIR%\\
copy /Y README.txt %PACKAGE_DIR%\\
copy /Y test.bat %PACKAGE_DIR%\\

mkdir %PACKAGE_DIR%\\output

echo 🗜️ Creando archivo ZIP...
powershell Compress-Archive -Path %PACKAGE_DIR% -DestinationPath %PACKAGE_NAME%.zip -Force

echo ✅ Paquete creado: %PACKAGE_NAME%.zip
echo.
echo El paquete incluye:
echo   • ExtractorPDF.exe (ejecutable principal)
echo   • README.txt (instrucciones)
echo   • test.bat (script de prueba)
echo   • output\\ (carpeta para archivos de salida)
echo.
pause
'''
    
    installer_path = Path('dist/create_package.bat')
    with open(installer_path, 'w', encoding='utf-8') as f:
        f.write(installer_content)
    print(f"✅ Creado script de empaquetado: {installer_path}")

def main():
    """Función principal del script de construcción"""
    print("🏦 Extractor PDF a Excel - Constructor de Ejecutable")
    print("=" * 60)
    print(f"🔧 PyInstaller Path: {PYINSTALLER_PATH}")
    print("=" * 60)
    
    # Verificaciones previas
    if not check_dependencies():
        return False
    
    if not check_required_files():
        return False
    
    # Preparar entorno
    prepare_build_environment()
    
    # Limpiar construcciones anteriores
    clean_previous_builds()
    
    # Construir ejecutable
    if not build_executable():
        return False
    
    # Verificar resultado
    if not verify_executable():
        return False
    
    # Configuración final
    post_build_setup()
    
    # Crear script de empaquetado
    create_installer_script()
    
    print("\n🎉 ¡Construcción completada exitosamente!")
    print("=" * 60)
    print(f"📦 Ejecutable disponible en: dist/ExtractorPDF.exe")
    print(f"📄 Documentación: dist/README.txt")
    print(f"🧪 Script de prueba: dist/test.bat")
    print(f"🎁 Script de empaquetado: dist/create_package.bat")
    print()
    print("💡 Próximos pasos:")
    print("   1. Prueba el ejecutable: cd dist && ExtractorPDF.exe")
    print("   2. Para distribuir: ejecuta create_package.bat en dist/")
    print("   3. Comparte el archivo ZIP generado")
    print()
    print("📋 Características del ejecutable:")
    print("   • Interfaz gráfica completa")
    print("   • No requiere instalación de Python")
    print("   • Incluye todas las dependencias")
    print("   • Funciona sin conexión a internet")
    print("   • Maneja montos con formato CR")
    
    return True

if __name__ == "__main__":
    success = main()
    input("\nPresiona Enter para salir...")
    sys.exit(0 if success else 1)