# -*- coding: utf-8 -*-
"""
Script para generar el ejecutable con PyInstaller
"""

import PyInstaller.__main__
import os
from pathlib import Path

# Directorio actual
base_dir = Path(__file__).parent

# Configuración de PyInstaller
PyInstaller.__main__.run([
    str(base_dir / 'src' / 'main.py'),  # Script principal
    '--name=ExportadorDatosOpticos',  # Nombre del ejecutable
    '--onefile',  # Un solo archivo ejecutable
    '--windowed',  # Sin consola (usar --console si se prefiere con consola)
    '--add-data', f'{base_dir / "config"};config',  # Incluir carpeta config
    '--hidden-import=openpyxl',  # Importación oculta
    '--hidden-import=tkinter',  # Importación oculta
    # '--icon=icon.ico',  # Descomentar si tiene un icono
    '--clean',  # Limpiar archivos temporales
])

print("\n" + "=" * 60)
print("  ✅ Ejecutable generado en la carpeta 'dist/'")
print("=" * 60)
print("\nPara distribuir el programa:")
print("1. Copie el ejecutable de dist/ExportadorDatosOpticos.exe")
print("2. Copie la carpeta 'config/' junto al ejecutable")
print("3. Asegúrese de que el usuario edite config/config.json con sus paths")
print("\n" + "=" * 60)
