# Optical Data Exporter/Exportador de Datos Ópticos

(Spanish version below)

A Python tool that automates the export of optical measurements (reflectance and transmittance) from individual Excel sheets into centralized master tables. Built to replace legacy VBA macros in an optical characterization laboratory for CSP solar materials.

Deployed in production as a standalone .exe distributed to laboratory colleagues — no Python installation required on end-user machines.


The problem it solved
The original workflow relied on VBA macros written back in 2008. Over time they accumulated serious limitations:

Fragile to Excel changes: an Office update could break the macro without warning.
Hard to distribute: every user needed a working VBA environment and macro permissions enabled.
Opaque logic: unmodularized code, no structured error handling, no traceability.
Manual process: the technician had to open the master file, run the macro, and trust that it worked.

This project re-implements that entire logic in Python, packages the result as a standalone executable, and adds a graphical interface, validation, error logs, and a modular, testable design.

Technical solution
Source Excel file             config.json              Destination master file
(measurement sheet)     →   (cell mapping)       →   (cumulative table)
     │                                                        │
     └──── ExcelReader ──── processing ──── ExcelWriter ─────┘
                                   │
                          UI (tkinter) + Logger
The program reads specific cells from each measurement sheet (sample metadata, dates, optical parameters, full 250–2500 nm spectrum), validates them, and writes them into the matching row of the master table, replicating the visual style of the existing rows.
For reflectance, it includes a re-implementation of the Addlosses calculation: it computes the relative optical loss and the combined uncertainty with respect to each sample's initial value.

Technical skills demonstrated

Python — full business logic, 0 dependency on Excel/VBA
openpyxl — read_only reading with cell caching, style-preserving writes, .xlsm support with embedded VBA
tkinter — native cross-platform GUI: main menu, sheet selector with multi-selection and an "all" checkbox, confirmation dialogs
PyInstaller — packaging as a standalone .exe with embedded data (--add-data)
Modular architecture — clean Reader / Writer / Config / UI / Validators / Logger separation
Config-driven design — all cell→column mapping logic externalized to JSON; changing the schema requires no code changes
Spectral data — handling of 451-point ranges (250–2500 nm) per measurement, with differential rounding (3 dec. for scalars / 4 dec. for the spectrum)
VBA→Python migration — includes a documented re-implementation of the original logic


Project structure
optical-data-exporter/
├── src/
│   ├── main.py               # Entry point and orchestration
│   ├── ui/
│   │   └── menu.py           # Graphical interface (tkinter)
│   ├── core/
│   │   ├── config.py         # Loading and validation of config.json
│   │   ├── excel_reader.py   # Reading with caching (read_only compatible)
│   │   └── excel_writer.py   # Writing, styles, and Addlosses calculation
│   └── utils/
│       ├── logger.py         # Error logging with context
│       └── validators.py     # File and extension validation
├── config/
│   ├── config.json           # Path configuration (edit before use)
│   └── config_demo.json      # Configuration to run the demo
├── demo/
│   ├── refl_demo.xlsx        # Synthetic reflectance measurement sheet
│   ├── trans_pv_demo.xlsx    # Synthetic PV transmittance sheet
│   ├── trans_csp_demo.xlsx   # Synthetic CSP transmittance sheet
│   └── master_demo.xlsx      # Destination master table (demo)
├── requirements.txt
├── build_exe.py              # Script to build the executable
└── README.md

Quick demo
bash# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the demo (no GUI, no real data)
python demo/run_demo.py
The script automatically processes the synthetic files included in demo/ and shows the full pipeline in the terminal:
============================================================
  OPTICAL DATA EXPORTER — DEMO
============================================================
  Pipeline: Excel reading → validation → write to master table
  Data:     synthetic (contains no real measurements)

  ✓  Configuration loaded: config_demo.json
  ✓  Master table initialized: master_demo.xlsx

────────────────────────────────────────────────────────────
  1 · Reflectance  (solar mirror — durability test)
────────────────────────────────────────────────────────────

  Scenario: same sample measured at t=0, t=1 year and t=2 years.
  The Addlosses module computes the accumulated optical loss.

  →  Exporting: refl_demo.xlsx     [t = 0  (initial measurement)]
  ✓  → row 3  |  duration: 0 h     |  ASTM: 0.937  |  Addlosses applied
  →  Exporting: refl_demo_1y.xlsx  [t = 1 year]
  ✓  → row 4  |  duration: 8760 h  |  ASTM: 0.924  |  Addlosses applied
  →  Exporting: refl_demo_2y.xlsx  [t = 2 years]
  ✓  → row 5  |  duration: 17520 h |  ASTM: 0.929  |  Addlosses applied
  ...
  ✓  DEMO COMPLETE
Open demo/master_demo.xlsx to see the exported rows with all their metadata, optical values, and full spectrum.

Development setup
bashgit clone https://github.com/LauraCGuzman/optical-data-exporter
cd optical-data-exporter
pip install -r requirements.txt
python src/main.py
Build the executable
bashpip install pyinstaller
python build_exe.py
# The .exe is generated in dist/ as ExportadorDatosOpticos.exe
For end-user instructions on running the distributed executable, see EXECUTABLE_GUIDE.md.

Configuration
Edit config/config.json and set the paths to your files:
json{
  "reflectance": {
    "destination_file": "C:/path/to/your/master_table_reflectance.xlsm",
    "sheet_name": "ReflectorsALL",
    "cells": [...]
  },
  "transmittance_pv": {
    "destination_file": "C:/path/to/your/master_table_transmittance.xlsx",
    ...
  }
}
The cell→column mapping is fully externalized: if the schema of your measurement sheet changes, you only need to update the JSON.

Supported measurement types
TypeExported parametersSpectrumReflectanceASTM, ISO, H660, Direct660 + uncertaintiesAJ79:AJ529 (251–2500 nm)PV transmittanceSolar-weighted transmittanceAJ79:AJ529CSP transmittanceSolar-weighted transmittanceAM80:AM530

Project context
This program emerged from work with optical characterization data of materials for concentrated solar power (CSP) systems: mirrors, glass covers, and antireflective coatings subjected to field durability tests.
The original workflow — manually exporting with VBA macros every time a new measurement arrived — did not scale as the sample volume grew. This tool removes the manual step, standardizes the output format, and can be distributed to any technician without setting up a development environment.

Dependencies
PackageUseopenpyxl >= 3.1Reading and writing Excel filespyinstaller >= 6.0Packaging as a standalone executable
Standard Python for the rest (tkinter, pathlib, logging, json).

License
MIT — free to use, modify, and distribute.

---

**Herramienta Python que automatiza la exportación de mediciones ópticas** (reflectancia y transmitancia) desde hojas Excel individuales a tablas maestras centralizadas. Desarrollada para reemplazar macros VBA heredados en un laboratorio de caracterización óptica de materiales solares CSP.

> Deployed in production as a standalone `.exe` distributed to laboratory colleagues — no Python installation required on end-user machines.

---

## El problema que resolvía

El flujo de trabajo original dependía de macros VBA escritos desde 2008. Con el tiempo acumularon limitaciones serias:

- **Frágiles ante cambios de Excel**: una actualización de Office podía romper el macro sin aviso.
- **No distribuibles fácilmente**: cada usuario necesitaba un entorno VBA funcional y permisos de macros habilitados.
- **Lógica opaca**: código sin modularizar, sin manejo de errores estructurado, sin trazabilidad.
- **Proceso manual**: el técnico tenía que abrir el archivo maestro, lanzar el macro, y confiar en que funcionaba.

Este proyecto reimplementa toda esa lógica en Python, empaqueta el resultado como ejecutable autónomo y añade interfaz gráfica, validaciones, logs de error y un diseño modular testeable.

---

## Solución técnica

```
Archivo Excel fuente          config.json              Archivo maestro destino
(hoja de medición)      →   (mapeo de celdas)    →   (tabla acumulativa)
     │                                                        │
     └──── ExcelReader ──── procesamiento ──── ExcelWriter ──┘
                                   │
                          UI (tkinter) + Logger
```

El programa lee celdas concretas de cada hoja de medición (metadatos de la muestra, fechas, parámetros ópticos, espectro completo 250–2500 nm), las valida y las escribe en la fila correspondiente de la tabla maestra, replicando el estilo visual de las filas existentes.

Para reflectancia, incluye la reimplementación del cálculo `Addlosses`: calcula la pérdida óptica relativa y la incertidumbre combinada respecto al valor inicial de cada muestra.

---

## Habilidades técnicas demostradas

- **Python** — lógica de negocio completa, 0 dependencias de Excel/VBA
- **openpyxl** — lectura en modo `read_only` con caché de celdas, escritura preservando estilos, soporte `.xlsm` con VBA embebido
- **tkinter** — GUI nativa multiplataforma: menú principal, selector de hojas con multi-selección y checkbox "todas", diálogos de confirmación
- **PyInstaller** — empaquetado como `.exe` autónomo con datos embebidos (`--add-data`)
- **Arquitectura modular** — separación limpia Reader / Writer / Config / UI / Validators / Logger
- **Config-driven design** — toda la lógica de mapeo celda→columna externalizada en JSON; cambiar el esquema no requiere modificar código
- **Datos espectrales** — manejo de rangos de 451 puntos (250–2500 nm) por medición, con redondeo diferencial (3 dec. escalares / 4 dec. espectro)
- **Migración VBA→Python** — incluye reimplementación documentada de la lógica original

---

## Estructura del proyecto

```
optical-data-exporter/
├── src/
│   ├── main.py               # Punto de entrada y orquestación
│   ├── ui/
│   │   └── menu.py           # Interfaz gráfica (tkinter)
│   ├── core/
│   │   ├── config.py         # Carga y validación de config.json
│   │   ├── excel_reader.py   # Lectura con caché (compatible read_only)
│   │   └── excel_writer.py   # Escritura, estilos y cálculo Addlosses
│   └── utils/
│       ├── logger.py         # Generación de logs de error con contexto
│       └── validators.py     # Validaciones de archivo y extensión
├── config/
│   ├── config.json           # Configuración de rutas (editar antes de usar)
│   └── config_demo.json      # Configuración para ejecutar el demo
├── demo/
│   ├── refl_demo.xlsx        # Hoja de medición de reflectancia sintética
│   ├── trans_pv_demo.xlsx    # Hoja de transmitancia PV sintética
│   ├── trans_csp_demo.xlsx   # Hoja de transmitancia CSP sintética
│   └── master_demo.xlsx      # Tabla maestra de destino (demo)
├── requirements.txt
├── build_exe.py              # Script para generar el ejecutable
└── README.md
```

---

## Demo rápido

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar el demo (sin GUI, sin datos reales)
python demo/run_demo.py
```

El script procesa automáticamente los archivos sintéticos incluidos en `demo/` y muestra el pipeline completo en terminal:

```
============================================================
  EXPORTADOR DE DATOS ÓPTICOS — DEMO
============================================================
  Pipeline: lectura Excel → validación → escritura en tabla maestra
  Datos:    sintéticos (no contienen mediciones reales)

  ✓  Configuración cargada: config_demo.json
  ✓  Tabla maestra inicializada: master_demo.xlsx

────────────────────────────────────────────────────────────
  1 · Reflectancia  (espejo solar — ensayo de durabilidad)
────────────────────────────────────────────────────────────

  Escenario: misma muestra medida en t=0, t=1 año y t=2 años.
  El módulo Addlosses calculará la pérdida óptica acumulada.

  →  Exportando: refl_demo.xlsx  [t = 0  (medición inicial)]
  ✓  → fila 3  |  duración: 0 h    |  ASTM: 0.937  |  Addlosses aplicado
  →  Exportando: refl_demo_1y.xlsx  [t = 1 año]
  ✓  → fila 4  |  duración: 8760 h |  ASTM: 0.924  |  Addlosses aplicado
  →  Exportando: refl_demo_2y.xlsx  [t = 2 años]
  ✓  → fila 5  |  duración: 17520 h|  ASTM: 0.929  |  Addlosses aplicado
  ...
  ✓  DEMO COMPLETADO
```

Abre `demo/master_demo.xlsx` para ver las filas exportadas con todos sus metadatos, valores ópticos y espectro completo.

---

## Instalación para desarrollo

```bash
git clone https://github.com/LauraCGuzman/optical-data-exporter
cd exportador-datos-opticos
pip install -r requirements.txt
python src/main.py
```

### Generar ejecutable

```bash
pip install pyinstaller
python build_exe.py
# El .exe se genera en dist/
```

---

## Configuración

Edita `config/config.json` y ajusta las rutas a tus archivos:

```json
{
  "reflectancia": {
    "archivo_destino": "C:/ruta/a/tu/tabla_maestra_reflectancia.xlsm",
    "nombre_hoja": "ReflectorsALL",
    "celdas": [...]
  },
  "transmitancia_pv": {
    "archivo_destino": "C:/ruta/a/tu/tabla_maestra_transmitancia.xlsx",
    ...
  }
}
```

El mapeo celda→columna está completamente externalizado: si el esquema de tu hoja de medición cambia, solo hay que actualizar el JSON.

---

## Tipos de medición soportados

| Tipo | Parámetros exportados | Espectro |
|------|----------------------|----------|
| Reflectancia | ASTM, ISO, H660, Direct660 + incertidumbres | AJ79:AJ529 (251–2500 nm) |
| Transmitancia PV | Solar-weighted transmittance | AJ79:AJ529 |
| Transmitancia CSP | Solar-weighted transmittance | AM80:AM530 |

---

## Contexto del proyecto

Este programa surgió en el marco del trabajo con datos de caracterización óptica de materiales para sistemas de concentración solar (CSP): espejos, cubiertas de vidrio y recubrimientos antirreflectantes sometidos a ensayos de durabilidad en campo.

El flujo original —exportar manualmente con macros VBA cada vez que llegaba una nueva medición— no escalaba cuando el volumen de muestras crecía. Esta herramienta elimina el proceso manual, estandariza el formato de salida y permite distribuirse a cualquier técnico sin necesidad de configurar entornos de desarrollo.

---

## Dependencias

| Paquete | Uso |
|---------|-----|
| `openpyxl >= 3.1` | Lectura y escritura de archivos Excel |
| `pyinstaller >= 6.0` | Empaquetado como ejecutable autónomo |

Python estándar para el resto (tkinter, pathlib, logging, json).

---

## Licencia

MIT — libre para uso, modificación y distribución.
