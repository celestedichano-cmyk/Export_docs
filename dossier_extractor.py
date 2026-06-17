"""
dossier_extractor.py
Extrae datos de un Dossier de producto (.xlsx) de The Not Company
para autocompletar el Reporte de Fórmula.

Uso:
    with open("dossier.xlsx", "rb") as f:
        data = extract_dossier(f.read())
"""

from openpyxl import load_workbook
import io


def _fmt_pct(value):
    """
    Formatea una fracción decimal (ej 0.0303) como porcentaje legible.
    Usa suficientes decimales para no perder valores muy chicos
    (ej. vitaminas en trazas), recortando ceros finales innecesarios.
    """
    if value is None:
        return ""
    pct = value * 100
    if pct == 0:
        return "0"
    # Para valores muy chicos, usar notación con más decimales
    decimales = 6
    while round(pct, decimales) == 0 and decimales < 12:
        decimales += 2
    s = f"{pct:.{decimales}f}".rstrip("0").rstrip(".")
    return s if s else "0"


def extract_dossier(xlsx_bytes):
    """
    Extrae producto y fórmula ordenada (decreciente) de un Dossier .xlsx.

    Returns:
        dict con 'producto' y 'formula_rows' (formato Ingrediente|% por línea,
        listo para pegar en el campo del Reporte Fórmula).
    """
    result = {}
    wb = load_workbook(io.BytesIO(xlsx_bytes), data_only=True)

    # --- Producto ---
    if "Datos de producto" in wb.sheetnames:
        ws_datos = wb["Datos de producto"]
        for row in ws_datos.iter_rows(min_row=1, max_row=30, max_col=2):
            label_cell, value_cell = row[0], row[1]
            if label_cell.value and str(label_cell.value).strip() == "Producto":
                result["producto"] = str(value_cell.value).strip() if value_cell.value else ""
                break

    # --- Fórmula (orden decreciente, columnas X/Y) ---
    if "Fórmula" in wb.sheetnames:
        ws_formula = wb["Fórmula"]
        rows = []
        # Datos empiezan en fila 5; columna X=24, Y=25
        r = 5
        while True:
            ing = ws_formula.cell(row=r, column=24).value
            pct = ws_formula.cell(row=r, column=25).value
            if ing is None:
                break
            ing_clean = str(ing).strip()
            pct_str = _fmt_pct(pct) if isinstance(pct, (int, float)) else ""
            rows.append(f"{ing_clean} | {pct_str}")
            r += 1
        result["formula_rows"] = "\n".join(rows)

    return result


# ── Test rápido ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else \
        "/mnt/user-data/uploads/Dossier_NotMilk_Original_10001202.xlsx"

    with open(path, "rb") as f:
        data = extract_dossier(f.read())

    print("Producto:", data.get("producto"))
    print("\nFormula rows:")
    print(data.get("formula_rows"))
