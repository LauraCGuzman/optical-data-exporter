# -*- coding: utf-8 -*-
"""
Módulo de logging orientado a usuario
Genera logs de error en lenguaje simple y no técnico
"""

import os
import traceback
from datetime import datetime
from pathlib import Path


class Logger:
    """Gestiona la generación de logs de error para el usuario"""
    
    def __init__(self, logs_dir=None):
        """
        Inicializa el logger
        
        Args:
            logs_dir: Directorio donde guardar los logs. Si es None, usa ./logs/
        """
        if logs_dir is None:
            base_dir = Path(__file__).parent.parent.parent
            logs_dir = base_dir / "logs"
        
        self.logs_dir = Path(logs_dir)
        self.logs_dir.mkdir(exist_ok=True)
    
    def generar_log_error(self, error, contexto="", archivo_origen=None, tipo_medicion=None):
        """
        Genera un archivo de log con información del error
        
        Args:
            error: La excepción o mensaje de error
            contexto: Descripción de qué estaba haciendo el programa
            archivo_origen: Path del archivo Excel que se estaba procesando
            tipo_medicion: 'reflectancia' o 'transmitancia'
        
        Returns:
            Path del archivo de log generado
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f"error_{timestamp}.txt"
        log_path = self.logs_dir / log_filename
        
        # Preparar el contenido del log
        contenido = self._formatear_log(error, contexto, archivo_origen, tipo_medicion)
        
        # Escribir el log
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(contenido)
        
        return log_path
    
    def _formatear_log(self, error, contexto, archivo_origen, tipo_medicion):
        """Formatea el contenido del log de manera user-friendly"""
        lineas = []
        lineas.append("=" * 70)
        lineas.append("REPORTE DE ERROR - Programa de Exportación de Datos Ópticos")
        lineas.append("=" * 70)
        lineas.append("")
        
        # Fecha y hora
        lineas.append(f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        lineas.append("")
        
        # Contexto
        if contexto:
            lineas.append("¿QUÉ ESTABA HACIENDO EL PROGRAMA?")
            lineas.append("-" * 70)
            lineas.append(contexto)
            lineas.append("")
        
        # Información del proceso
        if tipo_medicion:
            lineas.append(f"Tipo de medición: {tipo_medicion.upper()}")
        if archivo_origen:
            lineas.append(f"Archivo origen: {archivo_origen}")
        lineas.append("")
        
        # Descripción del error
        lineas.append("¿QUÉ SALIÓ MAL?")
        lineas.append("-" * 70)
        lineas.append(self._explicar_error(error))
        lineas.append("")
        
        # Posibles soluciones
        lineas.append("POSIBLES SOLUCIONES")
        lineas.append("-" * 70)
        soluciones = self._sugerir_soluciones(error)
        for i, solucion in enumerate(soluciones, 1):
            lineas.append(f"{i}. {solucion}")
        lineas.append("")
        
        # Información técnica (colapsable)
        lineas.append("=" * 70)
        lineas.append("INFORMACIÓN TÉCNICA (para soporte)")
        lineas.append("=" * 70)
        lineas.append(f"Tipo de error: {type(error).__name__}")
        lineas.append(f"Mensaje de error: {str(error)}")
        lineas.append("")
        
        # Stack trace si está disponible
        if hasattr(error, '__traceback__'):
            lineas.append("Traza del error:")
            lineas.append("-" * 70)
            tb_lines = traceback.format_exception(type(error), error, error.__traceback__)
            lineas.extend(tb_lines)
        
        return "\n".join(lineas)
    
    def _explicar_error(self, error):
        """Explica el error en lenguaje simple"""
        error_str = str(error)
        error_type = type(error).__name__
        
        # Mapeo de tipos de error comunes a explicaciones
        explicaciones = {
            'FileNotFoundError': "No se pudo encontrar un archivo necesario.",
            'PermissionError': "No se tiene permiso para acceder a un archivo. "
                             "Puede que el archivo esté abierto en Excel.",
            'KeyError': "Falta un campo esperado en el archivo de configuración.",
            'ValueError': "Un valor tiene un formato incorrecto o no es válido.",
            'ConfigError': "Hay un problema con el archivo de configuración.",
        }
        
        explicacion = explicaciones.get(error_type, "Se produjo un error inesperado.")
        
        return f"{explicacion}\n\nDetalle: {error_str}"
    
    def _sugerir_soluciones(self, error):
        """Sugiere posibles soluciones según el tipo de error"""
        error_type = type(error).__name__
        error_str = str(error).lower()
        
        soluciones = []
        
        # Soluciones según el tipo de error
        if error_type == 'FileNotFoundError':
            soluciones.append(
                "Verifique que el archivo existe en la ubicación especificada"
            )
            soluciones.append(
                "Revise el path en el archivo config.json y asegúrese de que sea correcto"
            )
        
        elif error_type == 'PermissionError' or 'permission' in error_str:
            soluciones.append(
                "Cierre el archivo Excel si está abierto en Microsoft Excel"
            )
            soluciones.append(
                "Verifique que tiene permisos de escritura en la ubicación del archivo"
            )
            soluciones.append(
                "Si el archivo está en una red, asegúrese de tener conexión"
            )
        
        elif 'config' in error_type.lower() or 'config' in error_str:
            soluciones.append(
                "Revise el archivo config.json en la carpeta 'config'"
            )
            soluciones.append(
                "Asegúrese de que los paths de archivos destino son correctos"
            )
            soluciones.append(
                "Verifique que el formato JSON es válido (comas, llaves, etc.)"
            )
        
        elif error_type == 'KeyError':
            soluciones.append(
                "Revise que el archivo Excel fuente tenga todas las celdas esperadas"
            )
            soluciones.append(
                "Verifique que está usando la plantilla correcta"
            )
        
        # Soluciones generales si no hay específicas
        if not soluciones:
            soluciones.append(
                "Revise el archivo de configuración (config.json)"
            )
            soluciones.append(
                "Asegúrese de que todos los archivos necesarios existen y son accesibles"
            )
            soluciones.append(
                "Intente cerrar todos los archivos Excel abiertos y vuelva a intentar"
            )
        
        # Solución final siempre
        soluciones.append(
            "Si el problema persiste, contacte con el soporte técnico y "
            "proporcione este archivo de log"
        )
        
        return soluciones
