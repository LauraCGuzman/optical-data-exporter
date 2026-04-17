# -*- coding: utf-8 -*-
"""
Módulo de gestión de configuración
Lee y valida el archivo config.json
"""

import json
import os
from pathlib import Path
import sys

class ConfigError(Exception):
    """Excepción personalizada para errores de configuración"""
    pass


class Config:
    """Gestiona la configuración del programa"""

    def __init__(self, config_path=None):
        """
        Inicializa la configuración
        """
        if config_path is None:
            # --- LÓGICA PARA PYINSTALLER ---
            if getattr(sys, 'frozen', False):
                # Si es un .exe, buscamos en la carpeta donde está el .exe
                base_dir = Path(sys.executable).parent
            else:
                # Si es PyCharm, usamos tu ruta original
                base_dir = Path(__file__).parent.parent.parent

            config_path = base_dir / "config" / "config.json"

        self.config_path = Path(config_path)
        self.data = None
        self._cargar_configuracion()
    
    def _cargar_configuracion(self):
        """Carga el archivo de configuración JSON"""
        if not self.config_path.exists():
            raise ConfigError(
                f"No se encontró el archivo de configuración en:\n{self.config_path}\n\n"
                "Por favor, asegúrese de que el archivo config.json existe en la carpeta 'config'."
            )
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except json.JSONDecodeError as e:
            raise ConfigError(
                f"El archivo de configuración tiene un formato JSON inválido:\n{e}\n\n"
                "Por favor, revise la sintaxis del archivo config.json"
            )
        except Exception as e:
            raise ConfigError(
                f"Error al leer el archivo de configuración:\n{e}"
            )
        
        self._validar_configuracion()
    
    def _validar_configuracion(self):
        """Valida que la configuración tenga la estructura esperada"""
        if not isinstance(self.data, dict):
            raise ConfigError("El archivo de configuración debe ser un objeto JSON")
        
        # Validar que existan las secciones principales
        for seccion in ['reflectancia', 'transmitancia_csp', 'transmitancia_pv']:
            if seccion not in self.data:
                raise ConfigError(
                    f"Falta la sección '{seccion}' en el archivo de configuración"
                )
            
            # Validar campos obligatorios
            seccion_data = self.data[seccion]
            if 'archivo_destino' not in seccion_data:
                raise ConfigError(
                    f"Falta el campo 'archivo_destino' en la sección '{seccion}'"
                )
            
            if 'nombre_hoja' not in seccion_data:
                raise ConfigError(
                    f"Falta el campo 'nombre_hoja' en la sección '{seccion}'"
                )
            
            if 'celdas' not in seccion_data or not isinstance(seccion_data['celdas'], list):
                raise ConfigError(
                    f"Falta el campo 'celdas' o no es una lista en la sección '{seccion}'"
                )
    
    def get_config_reflectancia(self):
        """Obtiene la configuración de reflectancia"""
        return self.data['reflectancia']
    
    def get_config_transmitancia_csp(self):
        """Obtiene la configuración de transmitancia"""
        return self.data['transmitancia_csp']

    def get_config_transmitancia_pv(self):
        """Obtiene la configuración de transmitancia"""
        return self.data['transmitancia_pv']
    
    def get_archivo_destino(self, tipo_medicion):
        """
        Obtiene el path del archivo destino para el tipo de medición
        
        Args:
            tipo_medicion: 'reflectancia' o 'transmitancia'
        
        Returns:
            Path del archivo destino
        """
        if tipo_medicion not in ['reflectancia', 'transmitancia_csp', 'transmitancia_pv']:
            raise ValueError(f"Tipo de medición inválido: {tipo_medicion}")
        
        return self.data[tipo_medicion]['archivo_destino']
    
    def get_nombre_hoja(self, tipo_medicion):
        """Obtiene el nombre de la hoja destino"""
        if tipo_medicion not in ['reflectancia', 'transmitancia_csp', 'transmitancia_pv']:
            raise ValueError(f"Tipo de medición inválido: {tipo_medicion}")
        
        return self.data[tipo_medicion]['nombre_hoja']
    
    def get_celdas(self, tipo_medicion):
        """Obtiene la lista de celdas a copiar"""
        if tipo_medicion not in ['reflectancia', 'transmitancia_csp', 'transmitancia_pv']:
            raise ValueError(f"Tipo de medición inválido: {tipo_medicion}")
        
        return self.data[tipo_medicion]['celdas']
    
    def validar_archivo_destino_existe(self, tipo_medicion):
        """
        Valida que el archivo destino exista
        
        Returns:
            (existe: bool, mensaje: str)
        """
        archivo = self.get_archivo_destino(tipo_medicion)
        if not os.path.exists(archivo):
            return False, (
                f"El archivo destino no existe:\n{archivo}\n\n"
                f"Por favor, verifique el path en config.json y asegúrese de que "
                f"el archivo existe."
            )
        return True, "OK"
