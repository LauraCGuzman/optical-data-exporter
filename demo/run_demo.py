"""
Demo del Exportador de Datos Ópticos
=====================================
Ejecuta el pipeline completo sin interfaz gráfica usando datos sintéticos.

Uso (desde la raíz del repositorio):
    python demo/run_demo.py

Resultado:
    - demo/master_demo.xlsx  →  tabla maestra con las mediciones exportadas
"""

import sys
import os
from pathlib import Path

# --- Rutas -------------------------------------------------------------------
REPO_ROOT = Path(__file__).parent.parent
DEMO_DIR  = Path(__file__).parent
sys.path.insert(0, str(REPO_ROOT / "src"))

# --- Imports del proyecto ----------------------------------------------------
from core.config import Config
from core.excel_reader import ExcelReader
from core.excel_writer import ExcelWriter

# --- Colores ANSI (compatibles con Windows 10+, macOS, Linux) ----------------
GREEN  = "\033[92m"
BLUE   = "\033[94m"
YELLOW = "\033[93m"
BOLD   = "\033[1m"
RESET  = "\033[0m"

def ok(msg):   print(f"  {GREEN}✓{RESET}  {msg}")
def info(msg): print(f"  {BLUE}→{RESET}  {msg}")
def header(msg):
    print(f"\n{BOLD}{YELLOW}{'─'*60}{RESET}")
    print(f"{BOLD}{YELLOW}  {msg}{RESET}")
    print(f"{BOLD}{YELLOW}{'─'*60}{RESET}")


# ---------------------------------------------------------------------------
# Recrear la tabla maestra limpia antes de cada demo
# ---------------------------------------------------------------------------

def reset_master():
    """Recrea master_demo.xlsx vacío con cabeceras para una demo limpia."""
    from demo.create_demo_files import create_master_table
    master_path = DEMO_DIR / "master_demo.xlsx"
    create_master_table(master_path)
    ok(f"Tabla maestra reiniciada: {master_path.name}")


# ---------------------------------------------------------------------------
# Pipeline de exportación (sin GUI)
# ---------------------------------------------------------------------------

def exportar(config: Config, tipo: str, archivo_fuente: Path) -> int:
    """
    Exporta una hoja de medición al archivo maestro.
    Devuelve el número de fila escrita.
    """
    config_celdas   = config.get_celdas(tipo)
    archivo_destino = config.get_archivo_destino(tipo)
    nombre_hoja     = config.get_nombre_hoja(tipo)

    # Leer
    reader = ExcelReader()
    reader.abrir_workbook(str(archivo_fuente))
    hojas  = reader.get_sheet_names()
    datos  = reader.leer_datos_hoja(hojas[0], config_celdas)
    reader.cerrar()

    # Escribir
    writer = ExcelWriter()
    writer.abrir_excel_destino(archivo_destino)
    writer.desactivar_filtros(nombre_hoja)
    fila = writer.encontrar_ultima_fila(nombre_hoja)
    writer.escribir_datos(nombre_hoja, fila, datos, config_celdas)

    # Addlosses solo para reflectancia
    if tipo == "reflectancia":
        writer.ejecutar_addlosses(nombre_hoja, [fila])

    writer.guardar_y_cerrar()
    return fila, datos


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print(f"\n{BOLD}{'='*60}")
    print("  EXPORTADOR DE DATOS ÓPTICOS — DEMO")
    print(f"{'='*60}{RESET}")
    print("  Pipeline: lectura Excel → validación → escritura en tabla maestra")
    print("  Datos:    sintéticos (no contienen mediciones reales)\n")

    # Cargar config demo
    config_path = REPO_ROOT / "config" / "config_demo.json"
    config = Config(config_path=str(config_path))
    ok(f"Configuración cargada: {config_path.name}")

    # Recrear tabla maestra limpia
    # (importamos aquí para evitar importar openpyxl antes del mensaje de inicio)
    sys.path.insert(0, str(DEMO_DIR.parent))
    from demo.create_demo_files import create_master_table
    master_path = DEMO_DIR / "master_demo.xlsx"
    create_master_table(master_path)
    ok(f"Tabla maestra inicializada: {master_path.name}")


    # ------------------------------------------------------------------
    # BLOQUE 1: Reflectancia — tres instantes temporales del mismo espejo
    # ------------------------------------------------------------------
    header("1 · Reflectancia  (espejo solar — ensayo de durabilidad)")

    print(f"\n  Escenario: misma muestra medida en t=0, t=1 año y t=2 años.")
    print(f"  El módulo Addlosses calculará la pérdida óptica acumulada.\n")

    archivos_refl = [
        (DEMO_DIR / "refl_demo.xlsx",    "t = 0  (medición inicial)"),
        (DEMO_DIR / "refl_demo_1y.xlsx", "t = 1 año"),
        (DEMO_DIR / "refl_demo_2y.xlsx", "t = 2 años"),
    ]

    for archivo, etiqueta in archivos_refl:
        info(f"Exportando: {archivo.name}  [{etiqueta}]")
        fila, datos = exportar(config, "reflectancia", archivo)
        duration = datos.get("Duration (hours)", 0) or 0
        astm     = datos.get("ASTM")
        ok(
            f"→ fila {fila}  |  duración: {duration:.0f} h  |  "
            f"ASTM: {astm:.3f}  |  Addlosses aplicado"
        )

    # ------------------------------------------------------------------
    # BLOQUE 2: Transmitancia PV
    # ------------------------------------------------------------------
    header("2 · Transmitancia PV  (vidrio de módulo fotovoltaico)")

    archivo_pv = DEMO_DIR / "trans_pv_demo.xlsx"
    info(f"Exportando: {archivo_pv.name}")
    fila, datos = exportar(config, "transmitancia_pv", archivo_pv)
    trans = datos.get("Solar-weighted Transmittance")
    ok(f"→ fila {fila}  |  Transmitancia solar ponderada: {trans:.3f}")


    # ------------------------------------------------------------------
    # BLOQUE 3: Transmitancia CSP
    # ------------------------------------------------------------------
    header("3 · Transmitancia CSP  (tubo de vidrio borosilicato)")

    archivo_csp = DEMO_DIR / "trans_csp_demo.xlsx"
    info(f"Exportando: {archivo_csp.name}")
    fila, datos = exportar(config, "transmitancia_csp", archivo_csp)
    trans = datos.get("Solar-weighted Transmittance")
    ok(f"→ fila {fila}  |  Transmitancia solar ponderada: {trans:.3f}")


    # ------------------------------------------------------------------
    # Resultado final
    # ------------------------------------------------------------------
    print(f"\n{BOLD}{GREEN}{'='*60}")
    print("  ✓  DEMO COMPLETADO")
    print(f"{'='*60}{RESET}")
    print(f"\n  Resultado guardado en:  {master_path}")
    print(f"  Abre el archivo para ver las filas exportadas:\n")
    print(f"    · Hoja 'ReflectorsALL'   → 3 filas de reflectancia + pérdidas")
    print(f"    · Hoja 'TransmittanceALL'→ 2 filas de transmitancia (PV y CSP)")
    print()


if __name__ == "__main__":
    main()
