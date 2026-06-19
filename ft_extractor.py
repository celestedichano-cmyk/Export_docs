"""
ft_extractor.py
Extrae datos de una Ficha Técnica PDF de The Not Company
y los devuelve como diccionario compatible con los templates de export docs.

Uso:
    with open("ficha.pdf", "rb") as f:
        data = extract_ft(f.read())
"""

import pdfplumber
import io
import re


def _clean(text):
    """Limpia texto: elimina saltos de línea y espacios extra."""
    if text is None:
        return ""
    return " ".join(str(text).split())


def _get_table(pdf, page_num):
    """
    Devuelve la primera tabla de la página lógica indicada (1-based).
    Busca por contenido "Página X de 3" para tolerar páginas en blanco
    intercaladas que desplacen la numeración física.
    """
    target = f"Página {page_num} de 3"
    for page in pdf.pages:
        tables = page.extract_tables()
        if not tables:
            continue
        # Check header row for page marker
        header = str(tables[0][0]) if tables[0] else ""
        if target in header:
            return tables[0]
    # Fallback: positional (original behavior)
    return pdf.pages[page_num - 1].extract_tables()[0]


def _find_value(table, key):
    """
    Busca una fila cuya primera celda contenga 'key' (case-insensitive)
    y devuelve el valor de la segunda celda.
    """
    key_lower = key.lower()
    for row in table:
        if row and row[0] and key_lower in str(row[0]).lower():
            return _clean(row[1]) if len(row) > 1 else ""
    return ""


def _remove_parens(s):
    """Remove all parenthetical content from a string."""
    result_s = ""
    depth = 0
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0:
            result_s += ch
    return result_s


def _parse_ingredientes(raw):
    """Parse ingredient string: remove parens, split by comma and 'y'."""
    clean = _remove_parens(raw)
    items = []
    for part in clean.split(","):
        part = part.strip().rstrip(".")
        if " y " in part.lower():
            subparts = re.split(r' y ', part, flags=re.IGNORECASE)
            for sp in subparts:
                sp = sp.strip().rstrip(".")
                if sp:
                    items.append(sp)
        else:
            if part:
                items.append(part)
    return "\n".join(items)


def _first_number(text):
    """Extract first numeric value from a string."""
    if not text:
        return ""
    m = re.search(r'[\d]+(?:[.,][\d]+)?', str(text))
    return m.group(0).replace(',', '.') if m else ""


def _extract_pagina1(table):
    """Extrae producto, ingredientes y datos nutricionales de página 1."""
    result = {}

    # --- Nombre del producto ---
    result["producto"] = _find_value(table, "Nombre Comercial")

    # --- Ingredientes ---
    # "Ingredientes" row may contain legal text (CHOCO) or actual list (PEANUT)
    # "Descripción del producto" may contain legal text too
    # Heuristic: the real ingredient list contains commas and no "SUPLEMENTO"
    raw_ing = ""
    for key in ["Ingredientes", "Descripción del producto"]:
        candidate = _find_value(table, key)
        if candidate and "," in candidate and "SUPLEMENTO" not in candidate:
            raw_ing = candidate
            break
    if raw_ing:
        result["ingredientes"] = _parse_ingredientes(raw_ing)

    # --- Tabla nutricional: estrategia robusta ---
    # Busca cada nutriente por nombre de fila y extrae el valor de 100g.
    # Funciona aunque la estructura de columnas varíe entre FTs.
    nut_search = {
        "energia":     ["Energía"],
        "proteina":    ["Proteínas"],
        "grasa_total": ["Grasa total"],
        "grasa_sat":   ["Grasas saturadas"],
        "grasa_mono":  ["Grasas monoinsat"],
        "grasa_poli":  ["Grasas poliinsat"],
        "grasa_trans": ["Grasas trans"],
        "colesterol":  ["Colesterol"],
        "carb_disp":   ["Carbohidratos disp"],
        "azucares":    ["Azúcares totales"],
        "fibra":       ["Fibra dietética total"],
        "sodio":       ["Sodio"],
    }

    # Build a flat map: keyword → campo
    kw_map = {}
    for campo, keywords in nut_search.items():
        for kw in keywords:
            kw_map[kw] = campo

    # Collect all text across all cells per row
    # Strategy: for each nutrient row, find the first standalone number
    # that looks like a 100g value (not in a multi-value stacked cell for the
    # first occurrence which is always 100g)
    
    # First pass: try the "stacked values in Energía row" pattern (original FT)
    for row in table:
        if row and row[0] and "Energía" in str(row[0]):
            # Find the cell with stacked values (contains \n with multiple numbers)
            for cell in row[1:]:
                if cell and "\n" in str(cell):
                    vals = str(cell).split("\n")
                    nutrientes_orden = [
                        "energia", "proteina", "grasa_total", "grasa_sat",
                        "grasa_mono", "grasa_poli", "grasa_trans", "colesterol",
                        "carb_disp", "azucares", "_az_g", "_az_an",
                        "fibra", "_fib_s", "_fib_i", "sodio"
                    ]
                    if len(vals) >= 10:  # Enough values to be the stacked column
                        for i, campo in enumerate(nutrientes_orden):
                            if i < len(vals) and not campo.startswith("_"):
                                result[campo] = vals[i].strip()
                        result["fibra_total"] = result.get("fibra", "")
                        return result

    # Second pass: row-by-row search (new FT layout)
    # col[1] = 100g values, col[2] = 1 porción values
    # Only use col[1] — col[2] are porción values, not 100g
    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0])

        matched_campo = None
        for kw, campo in kw_map.items():
            if kw in cell0:
                matched_campo = campo
                break

        if matched_campo and matched_campo not in result:
            # Only use col[1] for 100g values
            cell1 = row[1] if len(row) > 1 else None
            if cell1 and str(cell1).strip() and str(cell1) != "None":
                cell_str = str(cell1).strip()
                if "\n" in cell_str:
                    val = _first_number(cell_str.split("\n")[0])
                else:
                    val = _first_number(cell_str)
                if val:
                    result[matched_campo] = val

        # Fibra row: col[1] stacked = fibra, fibra_sol, fibra_insol, sodio (100g)
        if "Fibra dietética total" in cell0:
            cell1 = row[1] if len(row) > 1 else None
            if cell1 and "\n" in str(cell1):
                lines = str(cell1).strip().split("\n")
                if len(lines) >= 1 and "fibra" not in result:
                    result["fibra"] = _first_number(lines[0])
                if len(lines) >= 4:
                    result["sodio"] = _first_number(lines[3])

    result["fibra_total"] = result.get("fibra", "")
    return result


def _extract_pagina2(table):
    """Extrae sensoriales, fisicoquímicos y microbiológicos de página 2."""
    result = {}

    # --- 7.1 Sensoriales ---
    sensoriales_map = {
        "color": "color",
        "sabor": "sabor",
        "aroma": "aroma",
        "textura": "textura",
    }
    for row in table:
        if not row or not row[0]:
            continue
        param = str(row[0]).strip().lower()
        for key, campo in sensoriales_map.items():
            if param == key:
                result[campo] = _clean(row[1]) if len(row) > 1 else ""

    # --- 7.2 Fisicoquímicos ---
    fq_rows = []
    in_fq = False
    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()
        if "7.2 PARÁMETROS FISICOQUÍMICOS" in cell0:
            in_fq = True
            continue
        if in_fq and "7.3" in cell0:
            break
        if in_fq and cell0 and cell0 != "PARÁMETRO":
            parametro = _clean(row[0])
            especificacion = _clean(row[1]) if len(row) > 1 and row[1] else ""
            # columna de método: no suele figurar para fisicoquímicos en esta FT
            metodologia = ""
            fq_rows.append(f"{parametro}|{metodologia}|{especificacion}")

    result["fq_rows"] = "\n".join(fq_rows)

    # --- 7.3 Microbiológicos ---
    mb_rows = []
    in_mb = False
    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()
        if "7.3 PARÁMETROS MICROBIOLÓGICOS" in cell0:
            in_mb = True
            continue
        if in_mb and "7.4" in cell0:
            break
        if in_mb:
            # Saltar filas de encabezado/meta
            if cell0 in ("RSA, ARTÍCULO 173", "PARÁMETRO", ""):
                continue
            # Saltar fila secundaria de encabezado (CLASES, n, c, m, M)
            if row[1] is None and "CLASES" in str(row).upper():
                continue
            # Fila de dato real: [parametro, categoria, clases, n, c, m, M]
            parametro = _clean(row[0])
            if not parametro:
                continue
            valor_M = _clean(row[6]) if len(row) > 6 else ""
            if valor_M == "-" or valor_M == "":
                resultado = "0"
            else:
                resultado = f"<{valor_M}"
            mb_rows.append(f"{parametro}||{resultado}")

    result["mb_rows"] = "\n".join(mb_rows)

    return result


def _extract_pagina3(table):
    """Extrae datos de envase primario y secundario de página 3."""
    result = {}

    # Tipo de envase: viene del encabezado de sección
    # "ENVASE PRIMARIO (FLOWPACK)" → tipo = FLOWPACK
    # "ENVASE SECUNDARIO (ESTUCHE)" → tipo = ESTUCHE
    in_primario = False
    in_secundario = False
    in_caja = False

    primario = {}
    secundario = {}
    caja = {}

    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()

        # Detectar sección
        if "ENVASE PRIMARIO" in cell0:
            in_primario = True
            in_secundario = False
            in_caja = False
            # Extraer tipo entre paréntesis
            m = re.search(r'\(([^)]+)\)', cell0)
            result["tipo_envase_primario"] = m.group(1).title() if m else ""
            continue

        if "ENVASE SECUNDARIO" in cell0:
            in_primario = False
            in_secundario = True
            in_caja = False
            m = re.search(r'\(([^)]+)\)', cell0)
            result["tipo_envase_secundario"] = m.group(1).title() if m else ""
            continue

        if "CAJA MASTER" in cell0:
            in_primario = False
            in_secundario = False
            in_caja = True
            continue

        if "PALLETIZADO" in cell0:
            in_primario = False
            in_secundario = False
            in_caja = False
            continue

        val = _clean(row[1]) if len(row) > 1 and row[1] else ""

        if in_primario:
            if "Ancho" in cell0:
                primario["ancho"] = val
            elif "Largo" in cell0:
                primario["largo"] = val
            elif "Alto" in cell0:
                primario["alto"] = val
            elif "Peso neto" in cell0:
                primario["peso_neto"] = val
            elif "Peso bruto" in cell0:
                primario["peso_bruto"] = val
            elif "Material" in cell0:
                primario["material"] = val

        elif in_secundario:
            if "Ancho" in cell0:
                secundario["ancho"] = val
            elif "Largo" in cell0:
                secundario["largo"] = val
            elif "Alto" in cell0:
                secundario["alto"] = val
            elif "Peso neto" in cell0:
                secundario["peso_neto"] = val
            elif "Peso bruto" in cell0:
                secundario["peso_bruto"] = val
            elif "Material" in cell0:
                secundario["material"] = val

        elif in_caja:
            if "Ancho" in cell0:
                caja["ancho"] = val
            elif "Largo" in cell0:
                caja["largo"] = val
            elif "Alto" in cell0:
                caja["alto"] = val
            elif "Peso neto" in cell0:
                caja["peso_neto"] = val
            elif "Peso bruto" in cell0:
                caja["peso_bruto"] = val
            elif "Material" in cell0:
                # En la sección Caja Master a veces pdfplumber corta la celda
                # de etiqueta justo en "envase C" + "ARTÓN..." en vez de
                # "envase" + "CARTÓN...". Reconstruimos uniendo el resto de
                # la etiqueta (después de "Material de envase") con el valor.
                label_extra = cell0.split("envase", 1)[-1].strip()
                material_val = (label_extra + val).strip() if label_extra else val
                # El material de Caja Master casi siempre es Cartón Corrugado,
                # pero pdfplumber suele truncar el número de código entre
                # paréntesis de forma inconsistente según la FT. Normalizamos
                # al código estándar conocido (20) cuando se detecta el patrón.
                if "ARTÓN CORRUGADO" in material_val.upper() or "CARTÓN CORRUGADO" in material_val.upper():
                    material_val = "CARTÓN CORRUGADO (20)"
                elif material_val.count("(") > material_val.count(")"):
                    material_val += ")"
                caja["material"] = material_val
            elif "Cantidad" in cell0:
                caja["cantidad_unidades"] = val

    # Mapear a nombres de campos del template
    result["ancho_p"] = primario.get("ancho", "")
    result["largo_p"] = primario.get("largo", "")
    result["alto_p"] = primario.get("alto", "")
    result["peso_neto_p"] = primario.get("peso_neto", "")
    result["peso_bruto_p"] = primario.get("peso_bruto", "")
    result["material_p"] = primario.get("material", "")

    result["ancho_s"] = secundario.get("ancho", "")
    result["largo_s"] = secundario.get("largo", "")
    result["alto_s"] = secundario.get("alto", "")
    result["peso_neto_s"] = secundario.get("peso_neto", "")
    result["peso_bruto_s"] = secundario.get("peso_bruto", "")
    result["material_s"] = secundario.get("material", "")

    result["ancho_cm"] = caja.get("ancho", "")
    result["largo_cm"] = caja.get("largo", "")
    result["alto_cm"] = caja.get("alto", "")
    result["peso_neto_cm"] = caja.get("peso_neto", "")
    result["peso_bruto_cm"] = caja.get("peso_bruto", "")
    result["material_cm"] = caja.get("material", "")
    result["cantidad_unidades"] = caja.get("cantidad_unidades", "")

    return result


def extract_ft(pdf_bytes):
    """
    Extrae todos los campos relevantes de una Ficha Técnica PDF de NotCo.

    Args:
        pdf_bytes: contenido del PDF como bytes

    Returns:
        dict con todos los campos para los templates de export docs
    """
    result = {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        table_p1 = _get_table(pdf, 1)
        table_p2 = _get_table(pdf, 2)
        table_p3 = _get_table(pdf, 3)

    result.update(_extract_pagina1(table_p1))
    result.update(_extract_pagina2(table_p2))
    result.update(_extract_pagina3(table_p3))

    return result


# ── Test rápido ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import json
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else \
        "/mnt/user-data/uploads/POE_10_D_010001377_-__FT__NOTPROTEIN_CRUNCHY_CHOCOINTENSE_14X4X30G_CL.pdf"

    with open(path, "rb") as f:
        data = extract_ft(f.read())

    print(json.dumps(data, ensure_ascii=False, indent=2))
