# Executable — User Guide

This guide is for **end users** who received the program as a standalone executable.
You do **not** need Python or any programming knowledge to use it.

> If you are a developer and want to build the executable from source, see the
> [README](README.md) instead.

---

## What you received

After unzipping the distributed folder, you should have:

```
ExportadorDatosOpticos/
├── ExportadorDatosOpticos.exe   ← the program
└── config/
    └── config.json              ← settings you must edit once
```

The `config/` folder **must stay next to the `.exe`**. Do not move the `.exe` out
on its own.

---

## One-time setup

Before the first run, open `config/config.json` in any text editor (Notepad is fine)
and set the paths to **your own** master tables:

```json
{
  "reflectance": {
    "destination_file": "C:/path/to/your/master_table_reflectance.xlsm",
    "sheet_name": "ReflectorsALL"
  },
  "transmittance_pv": {
    "destination_file": "C:/path/to/your/master_table_transmittance.xlsx"
  }
}
```

Notes:
- Use forward slashes `/` in the paths (e.g. `C:/Users/...`), or double backslashes `\\`.
- The destination files must already exist and contain the expected master sheet.
- You only need to do this once, unless your file locations change.

---

## How to use it

1. **Double-click `ExportadorDatosOpticos.exe`.**
   The main window opens. (No black console window appears — that is expected.)

2. **Choose the measurement type** to export (reflectance / PV transmittance / CSP
   transmittance).

3. **Select the source Excel file** with the new measurement when prompted.

4. **Select the sheets to export.**
   The sheet selector lets you pick several sheets at once, or tick the **"all
   sheets"** checkbox to export every sheet in the file.

5. **Confirm.**
   A confirmation dialog summarizes what will be written. The program reads the
   selected cells, validates them, and appends them as a new row in your master
   table — preserving the visual style of the existing rows.

6. **Done.**
   Open your master table to check the newly added row(s), including metadata,
   optical values, and the full spectrum.

For reflectance measurements, the program automatically computes the **Addlosses**
values (relative optical loss and combined uncertainty) against the sample's
initial measurement.

---

## If something goes wrong

- The program writes an **error log** describing what happened and where. Keep that
  file — it makes diagnosing the problem much faster.
- Most common cause: a wrong path in `config/config.json`, or the destination file
  being **open in Excel** at the same time. Close the master file in Excel and try
  again.
- If the master sheet name does not match the one in `config.json`, nothing is
  written. Check the `sheet_name` value.

---

## Frequently asked

**Do I need to install Python?**
No. The executable bundles everything it needs.

**Can I run it on macOS or Linux?**
The distributed `.exe` is built for Windows. On other systems, run it from source
(see the [README](README.md)).

**Will it overwrite existing rows?**
No. Each export is appended as a new row after the last filled row.
