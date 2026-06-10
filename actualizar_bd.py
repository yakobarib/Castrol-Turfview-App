"""
actualizar_bd.py
================
Actualiza los datos embebidos en la app HTML cuando cambias las bases de datos.

USO:
  1. Reemplaza los archivos Excel en "3-Bases de Datos\"
  2. Ejecuta este script: python actualizar_bd.py
  3. Hace commit y push a GitHub automáticamente

REQUISITOS:
  pip install openpyxl
"""

import json
import re
import subprocess
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

BASE = Path(__file__).parent
BD_DIR = BASE / "3-Bases de Datos"
HTML_PATH = BASE / "7- APP TURFVIEW" / "turfview_castrol_app_v2_3.html"

# ── Leer BD_SKU_Codes ────────────────────────────────────────────────────────
def leer_sku(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    sku_dict = {}
    for row in rows:
        if not row[0]:
            continue
        sku = str(row[0]).strip()
        if sku.upper() == "SKU" or sku.upper() == "SKU CODE":
            continue  # cabecera
        desc = str(row[1]).strip() if row[1] else ""
        sku_dict[sku] = desc
    wb.close()
    return sku_dict

# ── Leer BD_Precios_Tarifa ───────────────────────────────────────────────────
def leer_tarifa(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Detectar fila cabecera
    data_start = 0
    for i, row in enumerate(rows[:5]):
        r0 = str(row[0] or "").upper()
        r1 = str(row[1] or "").upper()
        if "SKU" in r0 or "PRODUCTO" in r1:
            data_start = i + 1
            break

    tarifa_dict = {}
    for row in rows[data_start:]:
        sku  = str(row[0] or "").strip()
        desc = str(row[1] or "").strip()
        try:
            L     = float(row[2]) if row[2] not in (None, "") else None
            coste = float(row[3]) if row[3] not in (None, "") else None
            venta = float(row[4]) if row[4] not in (None, "") else None
        except (TypeError, ValueError):
            continue
        if not sku or L is None or coste is None or venta is None:
            continue
        tarifa_dict[sku] = {"d": desc, "L": L, "p": venta, "v": coste}

    wb.close()
    return tarifa_dict

# ── Merge: añadir descripción de SKU_Codes a Tarifa ─────────────────────────
def merge(sku_dict, tarifa_dict):
    for sku, data in tarifa_dict.items():
        if not data["d"] and sku in sku_dict:
            data["d"] = sku_dict[sku]
    return tarifa_dict

# ── Inyectar en HTML ─────────────────────────────────────────────────────────
MARKER_START = "// <<<BD_EMBED_START>>>"
MARKER_END   = "// <<<BD_EMBED_END>>>"

def inyectar_en_html(catalog_dict, html_path):
    html = html_path.read_text(encoding="utf-8")

    catalog_json = json.dumps(catalog_dict, ensure_ascii=False, separators=(",", ":"))
    count = len(catalog_dict)

    nuevo_bloque = (
        f"{MARKER_START}\n"
        f"  const BD_EMBED_CATALOG = {catalog_json};\n"
        f"  const BD_EMBED_COUNT   = {count};\n"
        f"  {MARKER_END}"
    )

    if MARKER_START in html:
        # Reemplazar bloque existente
        patron = re.compile(
            re.escape(MARKER_START) + r".*?" + re.escape(MARKER_END),
            re.DOTALL
        )
        html_nuevo = patron.sub(nuevo_bloque, html)
    else:
        # Insertar antes del cierre </script> del primer bloque de scripts
        html_nuevo = html.replace(
            "// ─── BD DRAG & DROP",
            nuevo_bloque + "\n\n// ─── BD DRAG & DROP",
            1
        )

    html_path.write_text(html_nuevo, encoding="utf-8")
    return count

# ── Git commit + push ────────────────────────────────────────────────────────
def git_push(base_dir, count):
    try:
        subprocess.run(
            ["git", "add",
             "7- APP TURFVIEW/turfview_castrol_app_v2_3.html",
             "3-Bases de Datos/BD_SKU_Codes.xlsx",
             "3-Bases de Datos/BD_Precios_Tarifa.xlsx"],
            cwd=base_dir, check=True
        )
        subprocess.run(
            ["git", "commit", "-m",
             f"feat: actualizar BD embebida en app ({count} referencias)"],
            cwd=base_dir, check=True
        )
        subprocess.run(
            ["git", "push"],
            cwd=base_dir, check=True
        )
        print("✓ Cambios subidos a GitHub.")
    except subprocess.CalledProcessError as e:
        print(f"⚠ Git falló: {e}. Sube los cambios manualmente.")

# ── Main ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    sku_path    = BD_DIR / "BD_SKU_Codes.xlsx"
    tarifa_path = BD_DIR / "BD_Precios_Tarifa.xlsx"

    print("Leyendo BD_SKU_Codes.xlsx...")
    sku_dict = leer_sku(sku_path)
    print(f"  → {len(sku_dict)} SKUs encontrados")

    print("Leyendo BD_Precios_Tarifa.xlsx...")
    tarifa_dict = leer_tarifa(tarifa_path)
    print(f"  → {len(tarifa_dict)} referencias en tarifa")

    print("Combinando datos...")
    catalog = merge(sku_dict, tarifa_dict)

    print("Inyectando datos en el HTML...")
    total = inyectar_en_html(catalog, HTML_PATH)
    print(f"  → {total} referencias embebidas en la app")

    print("Subiendo a GitHub...")
    git_push(BASE, total)

    print("\n✓ Listo. La app en GitHub Pages ya tiene los datos actualizados.")
