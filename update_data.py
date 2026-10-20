"""
Builds data/erp-data.js (which climate-risk-tarp.html loads) from the three CSV files
in the "data" folder:

    data/Database.csv               <- the "Database" sheet
    data/TARP_ERP.csv               <- the "TARP & ERP" sheet
    data/Project_Questionaire.csv   <- the "Project Questionaire" sheet

Save each sheet from Excel with File > Save As > "CSV UTF-8 (Comma delimited)",
then double-click "Update data.bat" (or run:  python update_data.py)
and refresh climate-risk-tarp.html.
"""
import os, sys

FILES = [("Database", "Database.csv"),
         ("TARP & ERP", "TARP_ERP.csv"),
         ("Project Questionaire", "Project_Questionaire.csv")]

data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

def read_csv_text(path):
    # Excel's "CSV UTF-8" adds a BOM; plain "CSV (Comma delimited)" uses the Windows code page
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    sys.exit("Could not read " + path)

def js_escape(s):
    return s.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

parts = []
for sheet, fname in FILES:
    path = os.path.join(data_dir, fname)
    if not os.path.exists(path):
        sys.exit(f"Missing data\\{fname}. Save the \"{sheet}\" sheet as CSV with that exact name.")
    text = read_csv_text(path)
    print(f"  {fname}: {text.count(chr(10))} lines")
    parts.append(f'"{sheet}": `{js_escape(text)}`')

with open(os.path.join(data_dir, "erp-data.js"), "w", encoding="utf-8", newline="\n") as f:
    f.write("/* Generated from the CSV files in this folder by update_data.py. Do not edit by hand. */\n")
    f.write("window.ERP_CSV = {\n" + ",\n".join(parts) + "\n};\n")
print("Done. data/erp-data.js updated. Refresh climate-risk-tarp.html in your browser.")
