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
    """Devuelve la primera tabla de la página indicada (1-based)."""
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


def _extract_pagina1(table):
    """Extrae producto y datos nutricionales de página 1."""
    result = {}

    # --- Nombre del producto ---
    result["producto"] = _find_value(table, "Nombre Comercial")

    # --- Tabla nutricional ---
    # La fila de Energía tiene los valores 100g apilados en una sola celda
    # Ejemplo: '274\n34\n10\n3.0\n5.3\n0.6\n0.0\n0.0\n12\n4.3\n0.1\n0.0\n17\n16\n0.4\n293'
    nutrientes_orden = [
        "energia", "proteina", "grasa_total", "grasa_sat",
        "grasa_mono", "grasa_poli", "grasa_trans", "colesterol",
        "carb_disp", "azucares",
        # azúcar y azúcar añadido se ignoran (no son campos del template)
        # pero están en la secuencia, así que los leemos y descartamos
        "_azucar_g", "_azucar_anadido",
        "fibra", "_fibra_soluble", "_fibra_insoluble", "sodio"
    ]

    for row in table:
        if row and row[0] and "Energía" in str(row[0]) and row[1]:
            valores_raw = str(row[1]).split("\n")
            for i, campo in enumerate(nutrientes_orden):
                if i < len(valores_raw) and not campo.startswith("_"):
                    result[campo] = valores_raw[i].strip()
            break

    # fibra_total: mismo valor que "fibra" extraído arriba
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

    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()

        # Detectar sección
        if "ENVASE PRIMARIO" in cell0:
            in_primario = True
            in_secundario = False
            # Extraer tipo entre paréntesis
            m = re.search(r'\(([^)]+)\)', cell0)
            result["tipo_envase_primario"] = m.group(1) if m else ""
            continue

        if "ENVASE SECUNDARIO" in cell0:
            in_primario = False
            in_secundario = True
            m = re.search(r'\(([^)]+)\)', cell0)
            result["tipo_envase_secundario"] = m.group(1) if m else ""
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
            if "Cantidad" in cell0:
                secundario["cantidad_unidades"] = val

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
    result["cantidad_unidades"] = secundario.get("cantidad_unidades", "")

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
