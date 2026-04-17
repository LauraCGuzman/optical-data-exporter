# -*- coding: utf-8 -*-
"""
Módulo de validaciones
Valida archivos, celdas y datos
"""

import os
from pathlib import Path


def validar_archivo_existe(path):
    """
    Valida que un archivo exista
    
    Args:
        path: Path del archivo a validar
    
    Returns:
        (existe: bool, mensaje: str)
    """
    if not os.path.exists(path):
        return False, f"El archivo no existe: {path}"
    
    if not os.path.isfile(path):
        return False, f"La ruta no corresponde a un archivo: {path}"
    
    return True, "OK"


def validar_archivo_accesible(path):
    """
    Valida que un archivo sea accesible (no esté bloqueado)
    
    Args:
        path: Path del archivo a validar
    
    Returns:
        (accesible: bool, mensaje: str)
    """
    if not os.path.exists(path):
        return False, f"El archivo no existe: {path}"
    
    # Intentar abrir el archivo para verificar que no esté bloqueado
    try:
        with open(path, 'r+b') as f:
            pass
        return True, "OK"
    except PermissionError:
        return False, (
            f"El archivo está en uso o bloqueado:\n{path}\n\n"
            "Por favor, cierre el archivo si está abierto en Excel e intente de nuevo."
        )
    except Exception as e:
        return False, f"Error al acceder al archivo: {e}"


def validar_extension_excel(path):
    """
    Valida que un archivo tenga extensión de Excel
    
    Args:
        path: Path del archivo
    
    Returns:
        (valida: bool, mensaje: str)
    """
    extensiones_validas = ['.xlsx', '.xlsm', '.xls']
    extension = Path(path).suffix.lower()
    
    if extension not in extensiones_validas:
        return False, (
            f"El archivo no tiene una extensión de Excel válida.\n"
            f"Extensión encontrada: {extension}\n"
            f"Extensiones válidas: {', '.join(extensiones_validas)}"
        )
    
    return True, "OK"


def validar_celda_formato(celda_str):
    """
    Valida que una string tenga formato de celda Excel válido (ej: 'C5', 'AB12')
    
    Args:
        celda_str: String con la referencia de celda
    
    Returns:
        (valida: bool, mensaje: str)
    """
    import re
    
    # Permitir rangos también (ej: 'C13:C23')
    if ':' in celda_str:
        partes = celda_str.split(':')
        if len(partes) != 2:
            return False, f"Formato de rango inválido: {celda_str}"
        
        for parte in partes:
            valida, msg = validar_celda_formato(parte)
            if not valida:
                return False, msg
        return True, "OK"
    
    # Formato de celda individual
    patron = r'^[A-Z]+\d+$'
    if not re.match(patron, celda_str):
        return False, f"Formato de celda inválido: {celda_str}"
    
    return True, "OK"
