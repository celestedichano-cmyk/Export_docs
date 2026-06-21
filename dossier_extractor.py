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
import re
import unicodedata


def _fmt_pct(value):
    """
    Formatea una fracción decimal (ej 0.0303) como porcentaje legible.
    Usa suficientes decimales para conservar al menos 2 cifras significativas,
    evitando perder valores muy chicos (ej. vitaminas en trazas).
    """
    if value is None:
        return ""
    pct = value * 100
    if pct == 0:
        return "0"
    # Calcular decimales necesarios para mostrar al menos 2 cifras significativas
    import math
    magnitud = math.floor(math.log10(abs(pct)))
    decimales = max(6, -magnitud + 1)
    decimales = min(decimales, 15)
    s = f"{pct:.{decimales}f}".rstrip("0").rstrip(".")
    return s if s else "0"


def _fmt_ins(value):
    """Formatea el número INS: quita '.0' final si es un float entero, deja texto tal cual."""
    if value is None or str(value).strip() == "":
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value).strip()


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

        # También cargamos el workbook con fórmulas (sin resolver) para
        # poder seguir la cadena Y → V → D por número de fila, evitando
        # depender de valores cacheados que pueden estar desactualizados.
        wb_formulas = load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
        ws_formula_raw = wb_formulas["Fórmula"]

        def _resolve_pct_from_formula(y_formula):
            """
            Dada la fórmula de una celda Y (ej '=V11' o '=V15+V16+V17'),
            extrae los números de fila V referenciados y suma los valores
            reales de la columna D en esas mismas filas (D{n} = V{n} siempre).
            Devuelve None si la celda no es una fórmula de columna V.
            """
            if not isinstance(y_formula, str) or not y_formula.startswith("="):
                return None
            filas_v = re.findall(r"V(\d+)", y_formula)
            if not filas_v:
                return None
            total = 0.0
            for fila_str in filas_v:
                fila_n = int(fila_str)
                d_val = ws_formula.cell(row=fila_n, column=4).value
                if isinstance(d_val, (int, float)):
                    total += d_val
            return total

        # --- Mapa nombre normalizado → % (desde columna D, fuente real sin caché) ---
        # Se usa como respaldo cuando el nombre de X coincide con A.
        pct_por_ingrediente = {}
        r = 4
        while True:
            ing_a = ws_formula.cell(row=r, column=1).value
            if ing_a is None:
                break
            pct_d = ws_formula.cell(row=r, column=4).value
            if isinstance(pct_d, (int, float)):
                pct_por_ingrediente[_normalize(ing_a)] = pct_d
            r += 1

        # --- Orden decreciente (columna X) con % resuelto vía cadena de fórmulas ---
        formula_rows = []
        ingredientes_orden = []
        r = 5
        while True:
            ing = ws_formula.cell(row=r, column=24).value
            if ing is None:
                break
            ing_clean = str(ing).strip()
            y_formula = ws_formula_raw.cell(row=r, column=25).value
            pct_val = _resolve_pct_from_formula(y_formula)
            if pct_val is None:
                # Fallback: cruce por nombre contra columna A/D
                pct_val = pct_por_ingrediente.get(_normalize(ing_clean))
            pct_str = _fmt_pct(pct_val) if isinstance(pct_val, (int, float)) else ""
            formula_rows.append(f"{ing_clean} | {pct_str}")
            ingredientes_orden.append(ing_clean)
            r += 1
        result["formula_rows"] = "\n".join(formula_rows)

        # --- Mapa nombre normalizado → función y → INS (desde tabla A/E/F) ---
        funcion_por_ingrediente = {}
        ins_por_ingrediente = {}
        r = 4
        while True:
            ing = ws_formula.cell(row=r, column=1).value
            if ing is None:
                break
            funcion = ws_formula.cell(row=r, column=5).value
            ins = ws_formula.cell(row=r, column=6).value
            if funcion:
                funcion_por_ingrediente[_normalize(ing)] = str(funcion).strip()
            ins_fmt = _fmt_ins(ins)
            if ins_fmt:
                ins_por_ingrediente[_normalize(ing)] = ins_fmt
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
                ins = ins_por_ingrediente.get(_normalize(nombre), "")
                nombre_fmt = f"{nombre.upper()} (INS {ins})" if ins else nombre.upper()
                aditivos_list.append(f"{nombre_fmt} | {funcion}")
                aditivos_filas.append(str(idx))

        result["ingredientes"] = "\n".join(ingredientes_orden)
        result["aditivos"] = "\n".join(aditivos_list)
        result["aditivos_filas"] = ", ".join(aditivos_filas)

    # --- Información nutricional (hoja "Proyecto de rótulo...", columna 100 ml) ---
    rotulo_sheet_name = next(
        (s for s in wb.sheetnames if _normalize(s).startswith("proyecto de rotulo")),
        None
    )
    if rotulo_sheet_name:
        ws_rotulo = wb[rotulo_sheet_name]
        # Mapa etiqueta normalizada → campo del template
        nut_map = {
            "energia (kcal)": "energia",
            "proteinas (g)": "proteina",
            "grasas totales (g)": "grasa_total",
            "grasas saturadas (g)": "grasa_sat",
            "grasas monoinsat. (g)": "grasa_mono",
            "grasas poliinsat. (g)": "grasa_poli",
            "grasas trans (g)": "grasa_trans",
            "colesterol (mg)": "colesterol",
            "carbohidratos totales (g)": "carb_totales",
            "carbohidratos disp. (g)": "carb_disp",
            "azucares totales (g)": "azucares",
            "fibra dietetica total (g)": "fibra",
            "sodio (mg)": "sodio",
        }
        # Buscar la columna "100 ml" o "100 g" — suele estar en una fila
        # encabezado cercana a "INFORMACIÓN NUTRICIONAL"
        col_100 = None
        for row in ws_rotulo.iter_rows(min_row=1, max_row=ws_rotulo.max_row):
            for cell in row:
                if cell.value and _normalize(cell.value) in ("100 ml", "100 g", "100g", "100ml"):
                    col_100 = cell.column
                    break
            if col_100:
                break
        if col_100:
            # Delimitar el bloque de la tabla nutricional: desde la fila
            # "INFORMACIÓN NUTRICIONAL" más cercana hasta la nota al pie
            # "(*) En relación..." — así evitamos capturar basura de otras
            # secciones de la hoja como micronutrientes "extra" por error.
            fila_inicio = None
            fila_fin = None
            for row in ws_rotulo.iter_rows(min_row=1, max_row=ws_rotulo.max_row, max_col=col_100):
                label_cell = row[1] if len(row) > 1 else None
                if label_cell is None or not label_cell.value:
                    continue
                texto = _normalize(label_cell.value)
                if "informacion nutricional" in texto and fila_inicio is None:
                    fila_inicio = label_cell.row
                if texto.startswith("(*) en relacion") and fila_inicio is not None and fila_fin is None:
                    fila_fin = label_cell.row
                    break
            if fila_inicio is None:
                fila_inicio = 1
            if fila_fin is None:
                fila_fin = ws_rotulo.max_row

            micronutrientes_extra = []
            for row in ws_rotulo.iter_rows(min_row=fila_inicio, max_row=fila_fin, max_col=col_100):
                label_cell = row[1] if len(row) > 1 else None
                if label_cell is None or not label_cell.value:
                    continue
                label_norm = _normalize(label_cell.value)
                campo = nut_map.get(label_norm)
                val_cell = ws_rotulo.cell(row=label_cell.row, column=col_100)
                if campo:
                    if campo not in result and isinstance(val_cell.value, (int, float)):
                        result[campo] = str(val_cell.value)
                elif isinstance(val_cell.value, (int, float)):
                    nombre_limpio = str(label_cell.value).strip()
                    micronutrientes_extra.append(f"{nombre_limpio} | {val_cell.value}")
            if micronutrientes_extra:
                result["micronutrientes_extra"] = "\n".join(micronutrientes_extra)

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
