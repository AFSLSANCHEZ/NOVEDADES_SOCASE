"""
script_procesamiento.py
─────────────────────────────────────────────────────────────────────────────
Convierte cambios.xlsx → data/data.json para el Portal de Novedades.

Estructura esperada del Excel:
  - Cada hoja = periodo  (ej: 2026-01, 2026-02)
  - Columnas fijas:  ID | TITULO | DESCRIPCION
  - Columnas dinámicas: TEXTO_1 | IMAGEN_1 | TEXTO_2 | IMAGEN_2 | ...

Uso:
  python script_procesamiento.py
  python script_procesamiento.py --excel data/mi_archivo.xlsx
  python script_procesamiento.py --excel data/cambios.xlsx --output data/data.json
"""

import argparse
import json
import os
import re
import sys
import unicodedata

# ─── DEPENDENCIA ─────────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    print("❌  Dependencia faltante. Instala con:\n    pip install openpyxl")
    sys.exit(1)


# ─── CONFIGURACIÓN POR DEFECTO ───────────────────────────────────────────────
DEFAULT_EXCEL  = os.path.join("data", "cambios.xlsx")
DEFAULT_OUTPUT = os.path.join("data", "data.json")
IMG_BASE_DIR   = "img"


# ─── PARSEO DE ARGUMENTOS ────────────────────────────────────────────────────
def parse_args():
    parser = argparse.ArgumentParser(
        description="Genera data.json a partir de cambios.xlsx"
    )
    parser.add_argument(
        "--excel", default=DEFAULT_EXCEL,
        help=f"Ruta al archivo Excel (default: {DEFAULT_EXCEL})"
    )
    parser.add_argument(
        "--output", default=DEFAULT_OUTPUT,
        help=f"Ruta del JSON de salida (default: {DEFAULT_OUTPUT})"
    )
    return parser.parse_args()


# ─── HELPERS ─────────────────────────────────────────────────────────────────
def normalize(text):
    """
    Normaliza un string: elimina tildes, convierte a mayúsculas y quita espacios.
    Ej: 'Título' → 'TITULO', 'descripción' → 'DESCRIPCION'
    """
    nfkd = unicodedata.normalize("NFKD", text)
    ascii_str = nfkd.encode("ascii", "ignore").decode("ascii")
    return ascii_str.upper().strip()


def cell_value(cell):
    """Devuelve el valor de la celda como string limpio, o None si está vacía."""
    val = cell.value
    if val is None:
        return None
    val = str(val).strip()
    return val if val else None


def parse_columns(header_row):
    """
    Analiza la fila de encabezados y devuelve:
      - fixed_cols: dict { nombre_normalizado: índice_columna }
      - dynamic_cols: lista ordenada de (índice, tipo, número)
                      tipo ∈ {'texto', 'imagen'}
                      ordenada para respetar TEXTO_1→IMAGEN_1→TEXTO_2→IMAGEN_2
    """
    fixed_cols   = {}
    dynamic_cols = []

    print("    🔎  Encabezados detectados:")
    for idx, cell in enumerate(header_row):
        raw = cell_value(cell)
        if raw is None:
            continue
        name_up = normalize(raw)  # tolera tildes y minúsculas
        print(f"        col {idx}: '{raw}' → normalizado: '{name_up}'")

        if name_up in ("ID", "TITULO", "DESCRIPCION"):
            fixed_cols[name_up] = idx
        else:
            # Acepta: TEXTO_1, TEXTO1, TEXTO_01, texto_1, Texto 1, etc.
            m = re.match(r'^(TEXTO|IMAGEN)[_\s]*(\d+)$', name_up)
            if m:
                tipo   = "texto" if m.group(1) == "TEXTO" else "imagen"
                numero = int(m.group(2))
                dynamic_cols.append((idx, tipo, numero))

    # ✅ FIX: ordenar por (número_de_bloque, orden_tipo)
    # 'imagen' viene DESPUÉS de 'texto' en el mismo número → invertir orden alfa
    # usando 0 para texto y 1 para imagen dentro del mismo grupo
    TYPE_ORDER = {"texto": 0, "imagen": 1}
    dynamic_cols.sort(key=lambda x: (x[2], TYPE_ORDER[x[1]]))
    return fixed_cols, dynamic_cols


def process_sheet(sheet):
    """
    Procesa una hoja completa y devuelve una lista de cambios.
    La primera fila es siempre el encabezado.
    """
    rows = list(sheet.iter_rows())
    if not rows:
        return []

    header_row = rows[0]
    fixed_cols, dynamic_cols = parse_columns(header_row)

    if not fixed_cols:
        print(f"  ⚠️  Hoja '{sheet.title}': no se encontraron columnas fijas (ID/TITULO/DESCRIPCION).")
        print(f"      Verifica que los encabezados estén en la FILA 1 y usen esos nombres exactos.")
        return []

    print(f"    ✅  Columnas fijas encontradas: {list(fixed_cols.keys())}")
    print(f"    ✅  Columnas dinámicas: {[(t, n) for (_, t, n) in dynamic_cols]}")

    cambios = []

    for row_idx, row in enumerate(rows[1:], start=2):
        row_cells = list(row)

        # Leer columnas fijas
        def get_fixed(name):
            col_idx = fixed_cols.get(name)
            if col_idx is None:
                return None
            return cell_value(row_cells[col_idx]) if col_idx < len(row_cells) else None

        cambio_id    = get_fixed("ID")
        titulo       = get_fixed("TITULO")
        descripcion  = get_fixed("DESCRIPCION")

        # Saltar filas completamente vacías
        if cambio_id is None and titulo is None and descripcion is None:
            continue

        # Normalizar ID (fallback si está vacío)
        if cambio_id is None:
            cambio_id = f"cambio-{row_idx:03d}"

        # Construir bloques dinámicos
        bloques = []
        for (col_idx, tipo, _numero) in dynamic_cols:
            if col_idx >= len(row_cells):
                continue
            valor = cell_value(row_cells[col_idx])
            if valor is None:
                continue  # Ignorar celdas vacías

            if tipo == "texto":
                bloques.append({"tipo": "texto", "contenido": valor})
            else:  # imagen
                # Construir ruta relativa: img/[ID_CAMBIO]/nombre_archivo
                img_path = os.path.join(IMG_BASE_DIR, cambio_id, valor).replace("\\", "/")
                bloques.append({"tipo": "imagen", "src": img_path})

        cambio = {
            "ID":          cambio_id,
            "TITULO":      titulo or "",
            "DESCRIPCION": descripcion or "",
            "BLOQUES":     bloques,
        }
        cambios.append(cambio)

    return cambios


# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    args = parse_args()

    # Validar existencia del Excel
    if not os.path.isfile(args.excel):
        print(f"❌  No se encontró el archivo: {args.excel}")
        print("    Crea el Excel o usa --excel para indicar la ruta correcta.")
        sys.exit(1)

    print(f"📖  Leyendo: {args.excel}")
    wb = openpyxl.load_workbook(args.excel, data_only=True)

    result = {}
    total_cambios = 0

    for sheet in wb.worksheets:
        periodo = sheet.title.strip()
        print(f"  📅  Procesando hoja: {periodo}")
        cambios = process_sheet(sheet)
        result[periodo] = cambios
        total_cambios += len(cambios)
        print(f"      ✅  {len(cambios)} cambio(s) procesado(s)")

    # Crear directorio de salida si no existe
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        print(f"📁  Directorio creado: {output_dir}")

    # Escribir JSON
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n✅  JSON generado: {args.output}")
    print(f"   Periodos: {len(result)} | Cambios totales: {total_cambios}")
    print("\n🌐  Abre index.html en tu navegador para ver el portal.")


if __name__ == "__main__":
    main()