"""
dossier_extractor.py
Extrae datos de un Dossier de producto (.xlsx) de The Not Company
para autocompletar el Reporte de Fórmula y el Informe de Aditivos.

Uso:
    with open("dossier.xlsx", "rb") as f:
        data = extract_dossier(f.read())
"""

from openpyxl import load_workbook
import io
import unicodedata


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


def _normalize(text):
    """Normaliza texto para comparación: sin tildes, minúsculas, sin espacios extra."""
    if not text:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


def extract_dossier(xlsx_bytes):
    """
    Extrae producto, fórmula ordenada y datos de aditivos de un Dossier .xlsx.

    Returns:
        dict con:
        - 'producto'
        - 'formula_rows' (Ingrediente|% por línea, para Reporte Fórmula)
        - 'ingredientes' (uno por línea, orden decreciente, para Informe Aditivos)
        - 'aditivos' (NOMBRE | Función, uno por línea)
        - 'aditivos_filas' (números de fila 1-based que son aditivos, ej "8, 13")
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

    if "Fórmula" in wb.sheetnames:
        ws_formula = wb["Fórmula"]

        # --- Fórmula (orden decreciente, columnas X/Y) → para Reporte Fórmula ---
        formula_rows = []
        ingredientes_orden = []
        r = 5
        while True:
            ing = ws_formula.cell(row=r, column=24).value
            pct = ws_formula.cell(row=r, column=25).value
            if ing is None:
                break
            ing_clean = str(ing).strip()
            pct_str = _fmt_pct(pct) if isinstance(pct, (int, float)) else ""
            formula_rows.append(f"{ing_clean} | {pct_str}")
            ingredientes_orden.append(ing_clean)
            r += 1
        result["formula_rows"] = "\n".join(formula_rows)

        # --- Mapa nombre normalizado → función (desde tabla A/E) ---
        funcion_por_ingrediente = {}
        r = 4
        while True:
            ing = ws_formula.cell(row=r, column=1).value
            if ing is None:
                break
            funcion = ws_formula.cell(row=r, column=5).value
            if funcion:
                funcion_por_ingrediente[_normalize(ing)] = str(funcion).strip()
            r += 1
        # "Saborizantes naturales" ya viene agrupado en columna X; si no está
        # mapeado individualmente en A/E (porque ahí está desglosado en aromas),
        # asignamos su función fija conocida.
        funcion_por_ingrediente.setdefault(_normalize("Saborizantes naturales"), "Saborización")

        # --- Cruce: ingredientes (orden decreciente) + aditivos + filas ---
        aditivos_list = []
        aditivos_filas = []
        for idx, nombre in enumerate(ingredientes_orden, start=1):
            funcion = funcion_por_ingrediente.get(_normalize(nombre), "")
            if funcion:
                aditivos_list.append(f"{nombre.upper()} | {funcion}")
                aditivos_filas.append(str(idx))

        result["ingredientes"] = "\n".join(ingredientes_orden)
        result["aditivos"] = "\n".join(aditivos_list)
        result["aditivos_filas"] = ", ".join(aditivos_filas)

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
    print("\nIngredientes (orden decreciente):")
    print(data.get("ingredientes"))
    print("\nAditivos:")
    print(data.get("aditivos"))
    print("\nAditivos filas:")
    print(data.get("aditivos_filas"))
