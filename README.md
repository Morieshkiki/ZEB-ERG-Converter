# CSV Field Mapper (ERG)

This folder contains a Tkinter app to map columns from a CSV file to the ERG target fields and export the mapped result.

## Features

- Load a CSV (tries common encodings and attempts to sniff delimiter)
- Map **multiple target fields to the same CSV column**
- Auto-map by position
- Smart auto-map by name
- Export:
  - Excel (`.xlsx`) via `pandas` + `openpyxl`
  - Access (`.accdb`) via `pyodbc` (requires Microsoft Access/ACE driver)

## Install

Create/activate your Python environment, then install:

```powershell
pip install -r requirements.txt
```

## Run

```powershell
python .\6_csv_field_mapper.py
```

## Notes on Access export

- The `.accdb` export requires an installed ODBC driver:
  - **Microsoft Access Database Engine (ACE)** / "Microsoft Access Driver (*.mdb, *.accdb)"
- If the driver is missing, the app will offer an Excel fallback (which you can import into Access).

## Sample data

A sample CSV is included: `erg_202601060952.csv`.
