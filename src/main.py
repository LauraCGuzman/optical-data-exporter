# -*- coding: utf-8 -*-
"""
Programa principal de exportación de datos ópticos
Reemplaza los macros VBA para exportar datos de reflectancia y transmitancia
"""

import sys
import os
import argparse

from pathlib import Path

# Configurar rutas para los imports
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    sys.path.insert(0, os.path.join(base_path, 'src'))
else:
    sys.path.insert(0, str(Path(__file__).parent))

from ui.menu import UI
from core.config import Config, ConfigError
from core.excel_reader import ExcelReader
from core.excel_writer import ExcelWriter
from utils.logger import Logger
from utils.validators import validar_archivo_existe, validar_archivo_accesible, validar_extension_excel


def main():
    """Función principal del programa"""

    parser = argparse.ArgumentParser(description="Exportador de Datos Ópticos")
    parser.add_argument(
        "--config",
        type=str,
        default=None,
        help="Ruta al archivo de configuración JSON (por defecto: config/config.json)"
    )
    args, _ = parser.parse_known_args()

    ui = UI()
    logger = Logger()

    try:
        print("\nCargando configuración...")
        try:
            config = Config(config_path=args.config)
            print("✅ Configuración cargada correctamente")
        except ConfigError as e:
            ui.mostrar_error(str(e))
            return 1
        
        tipo_medicion = ui.mostrar_menu_principal()
        
        if tipo_medicion is None:
            print("\nPrograma finalizado por el usuario.")
            return 0
        
        print(f"\n✅ Seleccionado: {tipo_medicion.upper()}")
        
        existe, mensaje = config.validar_archivo_destino_existe(tipo_medicion)
        if not existe:
            ui.mostrar_error(mensaje)
            log_path = logger.generar_log_error(
                Exception(mensaje),
                contexto=f"Validando archivo destino de {tipo_medicion}",
                tipo_medicion=tipo_medicion
            )
            print(f"\n📝 Log generado en: {log_path}")
            return 1
        
        print(f"\nSeleccione el archivo Excel con datos de {tipo_medicion}...")
        archivo_fuente = ui.seleccionar_archivo_excel(tipo_medicion)
        
        if archivo_fuente is None:
            print("\nOperación cancelada por el usuario.")
            return 0
        
        print(f"✅ Archivo seleccionado: {archivo_fuente}")
        
        existe, mensaje = validar_archivo_existe(archivo_fuente)
        if not existe:
            ui.mostrar_error(mensaje)
            return 1
        
        valida, mensaje = validar_extension_excel(archivo_fuente)
        if not valida:
            ui.mostrar_error(mensaje)
            return 1
        
        print("\nAbriendo archivo...")
        reader = ExcelReader()
        try:
            reader.abrir_workbook(archivo_fuente)
            sheet_names = reader.get_sheet_names()
            print(f"✅ Archivo abierto. Hojas encontradas: {len(sheet_names)}")
        except Exception as e:
            error_msg = f"Error al abrir el archivo Excel:\n{str(e)}"
            ui.mostrar_error(error_msg)
            log_path = logger.generar_log_error(
                e,
                contexto=f"Abriendo archivo Excel fuente: {archivo_fuente}",
                archivo_origen=archivo_fuente,
                tipo_medicion=tipo_medicion
            )
            print(f"\n📝 Log generado en: {log_path}")
            return 1
        
        print("\nSeleccione las hojas a procesar...")
        hojas_seleccionadas = ui.seleccionar_hojas(sheet_names)
        
        if hojas_seleccionadas is None or len(hojas_seleccionadas) == 0:
            print("\nOperación cancelada por el usuario.")
            reader.cerrar()
            return 0
        
        print(f"✅ Hojas seleccionadas: {len(hojas_seleccionadas)}")
        
        if not ui.confirmar_procesamiento(archivo_fuente, hojas_seleccionadas, tipo_medicion):
            print("\nOperación cancelada por el usuario.")
            reader.cerrar()
            return 0
        
        print("\n" + "=" * 60)
        print("  PROCESANDO DATOS...")
        print("=" * 60 + "\n")
        
        archivo_destino = config.get_archivo_destino(tipo_medicion)
        nombre_hoja_destino = config.get_nombre_hoja(tipo_medicion)
        config_celdas = config.get_celdas(tipo_medicion)
        
        accesible, mensaje = validar_archivo_accesible(archivo_destino)
        if not accesible:
            ui.mostrar_error(mensaje)
            reader.cerrar()
            return 1
        
        writer = ExcelWriter()
        try:
            writer.abrir_excel_destino(archivo_destino)
            writer.desactivar_filtros(nombre_hoja_destino)
            print(f"✅ Archivo destino abierto: {archivo_destino}")
        except Exception as e:
            error_msg = f"Error al abrir el archivo destino:\n{str(e)}"
            ui.mostrar_error(error_msg)
            log_path = logger.generar_log_error(
                e,
                contexto=f"Abriendo archivo Excel destino: {archivo_destino}",
                archivo_origen=archivo_fuente,
                tipo_medicion=tipo_medicion
            )
            print(f"\n📝 Log generado en: {log_path}")
            reader.cerrar()
            return 1
        
        # Acumular filas escritas para Addlosses
        filas_escritas = []
        hojas_procesadas = 0

        for i, nombre_hoja in enumerate(hojas_seleccionadas, 1):
            print(f"\nProcesando hoja {i}/{len(hojas_seleccionadas)}: {nombre_hoja}...")
            
            try:
                datos = reader.leer_datos_hoja(nombre_hoja, config_celdas)
                print(f"  ✅ Datos leídos de la hoja")
                
                fila_destino = writer.encontrar_ultima_fila(nombre_hoja_destino)
                writer.escribir_datos(nombre_hoja_destino, fila_destino, datos, config_celdas)
                print(f"  ✅ Datos escritos en fila {fila_destino} del archivo destino")
                
                filas_escritas.append(fila_destino)
                hojas_procesadas += 1
                
            except Exception as e:
                error_msg = f"Error al procesar la hoja '{nombre_hoja}':\n{str(e)}"
                print(f"  ❌ {error_msg}")
                
                log_path = logger.generar_log_error(
                    e,
                    contexto=f"Procesando hoja '{nombre_hoja}' del archivo {archivo_fuente}",
                    archivo_origen=archivo_fuente,
                    tipo_medicion=tipo_medicion
                )
                print(f"  📝 Log generado en: {log_path}")
                
                if i < len(hojas_seleccionadas):
                    continuar = input("\n  ¿Desea continuar con las siguientes hojas? (S/N): ").strip().lower()
                    if continuar not in ['s', 'si', 'sí', 'y', 'yes']:
                        print("\n  Procesamiento detenido por el usuario.")
                        break

        # Ejecutar Addlosses sobre las filas recién escritas (solo reflectancia)
        if filas_escritas and tipo_medicion == 'reflectancia':
            print("\nEjecutando cálculo de pérdidas (Addlosses)...")
            try:
                writer.ejecutar_addlosses(nombre_hoja_destino, filas_escritas)
            except Exception as e:
                print(f"  ⚠️  Error en Addlosses (los datos se guardarán igualmente): {e}")
                logger.generar_log_error(
                    e,
                    contexto="Ejecutando Addlosses",
                    archivo_origen=archivo_fuente,
                    tipo_medicion=tipo_medicion
                )

        print("\nGuardando cambios...")
        writer.guardar_y_cerrar()
        reader.cerrar()
        print("✅ Cambios guardados")
        
        if hojas_procesadas > 0:
            ui.mostrar_exito(hojas_procesadas, archivo_destino)
            return 0
        else:
            ui.mostrar_error("No se procesó ninguna hoja correctamente.")
            return 1
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Programa interrumpido por el usuario.")
        return 1
    
    except Exception as e:
        error_msg = f"Error inesperado:\n{str(e)}"
        ui.mostrar_error(error_msg)
        log_path = logger.generar_log_error(
            e,
            contexto="Error inesperado en el programa principal"
        )
        print(f"\n📝 Log generado en: {log_path}")
        return 1
    
    finally:
        ui.cerrar()


if __name__ == "__main__":
    exit_code = main()
    input("\nPresione Enter para salir...")
    sys.exit(exit_code)
