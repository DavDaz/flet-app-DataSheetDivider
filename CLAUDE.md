# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Flet desktop GUI application that splits an Excel file's data by a selected column's unique values — generating one sheet per unique value in a combined output file, plus individual `.xlsx` files per value in a subfolder.

## Commands

### Setup
```bash
python -m venv venv
source venv/bin/activate       # macOS/Linux
pip install -r requirements.txt
```

### Run
```bash
flet run data_sheet_divider.py
```

### Build executable
```bash
flet pack data_sheet_divider.py --icon img.png
```

## Architecture

The entire application lives in a single file: `data_sheet_divider.py`.

**Core processing logic:**
- `exportar_ventanas_xlsx(file_path, carpeta_principal)` — reads a multi-sheet Excel file and exports each sheet as its own `.xlsx` file inside a `Separados-<filename>` subfolder.
- `sanitize_sheet_name(name)` — strips characters Excel disallows in sheet names and enforces the 31-character limit.
- `get_default_folder_name()` — returns a timestamped folder name (`Resultados_YYYY-MM-DD_HH-MM-SS`) used when the user leaves the folder field blank.

**UI flow (inside `main(page)`):**
1. User picks an `.xlsx`/`.xls` file via `ft.FilePicker` → `cargar_hojas_excel()` populates `sheets_dropdown`.
2. User selects a sheet → `on_sheet_change()` populates `columns_dropdown` with that sheet's column headers.
3. User optionally names the output folder, then clicks "Ejecutar Proceso!".
4. `btn_click()` reads the chosen sheet, groups rows by the unique values of the selected column, writes a combined Excel file (one sheet per unique value) into the output folder, then calls `exportar_ventanas_xlsx()` to also produce individual files.
5. Results (or errors) are shown in `resultados_container` with a "Procesar otro archivo" button that calls `limpiar_campos()` to reset the UI.

**Output structure** (created next to the source Excel file):
```
<output_folder>/
  <original_filename>.xlsx        # combined file, one sheet per unique value
  Separados-<original_filename>/
    <original_filename>-<sheet_name>.xlsx   # one file per sheet
```

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| flet | 0.15.0 | GUI framework |
| pandas | 2.2.2 | Excel read/write |
| openpyxl | 3.1.2 | pandas Excel engine |
| python-dotenv | 1.0.1 | `.env` loading (loaded at startup, no env vars currently required) |
| pyinstaller | 5.13.2 | Executable packaging (used by `flet pack`) |
| pillow | 11.1.0 | Icon support for packaging |
