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
    raw = re.sub(r'^\s*Ingredientes\s*:\s*', '', raw, flags=re.IGNORECASE)
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

    result["producto"] = _find_value(table, "Nombre Comercial")

    raw_ing = ""
    for key in ["Ingredientes", "Descripción del producto"]:
        candidate = _find_value(table, key)
        if candidate and "," in candidate and "SUPLEMENTO" not in candidate:
            raw_ing = candidate
            break
    if raw_ing:
        result["ingredientes"] = _parse_ingredientes(raw_ing)

    etiqueta_a_campo = {
        "energia": "energia",
        "proteinas": "proteina",
        "proteina": "proteina",
        "grasa total": "grasa_total",
        "grasas totales": "grasa_total",
        "grasa saturada": "grasa_sat",
        "grasas saturadas": "grasa_sat",
        "grasa monoinsat": "grasa_mono",
        "grasas monoinsat": "grasa_mono",
        "grasa poliinsat": "grasa_poli",
        "grasas poliinsat": "grasa_poli",
        "grasa trans": "grasa_trans",
        "grasas trans": "grasa_trans",
        "acidos grasos trans": "grasa_trans",
        "colesterol": "colesterol",
        "carbohidratos totales": "carb_totales",
        "hidratos de carbono totales": "carb_totales",
        "carbohidratos disp": "carb_disp",
        "hidratos de carbono disp": "carb_disp",
        "azucares totales": "azucares",
        "azucares": "azucares",
        "azucar": None,
        "lactosa": None,
        "fibra dietetica total": "fibra",
        "fibra soluble": None,
        "fibra insoluble": None,
        "sodio": "sodio",
    }

    def _norm_simple(s):
        s = str(s).strip().lower()
        s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
               .replace("ó", "o").replace("ú", "u"))
        return s

    def _campo_de_etiqueta(etiqueta):
        et_norm = _norm_simple(etiqueta)
        mejor = None
        for clave, c in etiqueta_a_campo.items():
            if _norm_simple(clave) in et_norm:
                if mejor is None or len(clave) > len(mejor[0]):
                    mejor = (clave, c)
        return mejor[1] if mejor else None

    def _split_valor_doble(linea):
        partes = linea.strip().split()
        if len(partes) >= 2 and _first_number(partes[0]) and _first_number(partes[1]):
            return partes[0]
        return linea.strip()

    idx_inicio = None
    for i, row in enumerate(table):
        if row and row[0] and "INFORMACIÓN NUTRICIONAL" in str(row[0]).upper():
            idx_inicio = i
            break
    if idx_inicio is None:
        idx_inicio = 0

    for row in table[idx_inicio:]:
        if row and row[0] and "Energía" in str(row[0]):
            cell = row[1] if len(row) > 1 else None
            if cell and "\n" in str(cell):
                vals16 = str(cell).split("\n")
                if len(vals16) == 16:
                    orden16 = [
                        "energia", "proteina", "grasa_total", "grasa_sat",
                        "grasa_mono", "grasa_poli", "grasa_trans", "colesterol",
                        "carb_disp", "azucares", "_az", "_aza",
                        "fibra", "_fibs", "_fibi", "sodio"
                    ]
                    for idx_v, campo in enumerate(orden16):
                        if not campo.startswith("_"):
                            result[campo] = vals16[idx_v].strip()
                    result["fibra_total"] = result.get("fibra", "")
                    idx_inicio = None
            break

    secuencia = []
    pendientes_sin_etiqueta = []

    for i in range(idx_inicio if idx_inicio is not None else len(table), len(table)):
        row = table[i]
        if not row or not row[0]:
            continue
        cell0 = str(row[0])
        if "SELLOS" in cell0.upper() or "DESCRIPTORES" in cell0.upper():
            break
        etiquetas_aqui = [l.strip() for l in cell0.split("\n") if l.strip()]
        if not etiquetas_aqui:
            continue

        valor_cell = None
        if len(row) > 1 and row[1] and str(row[1]).strip():
            valor_cell = str(row[1]).strip()

        if valor_cell and "\n" in valor_cell:
            valores_aqui = [_split_valor_doble(l) for l in valor_cell.split("\n") if l.strip()]
        elif valor_cell:
            valores_aqui = [_split_valor_doble(valor_cell)]
        else:
            valores_aqui = []

        for idx_et, et in enumerate(etiquetas_aqui):
            if pendientes_sin_etiqueta:
                secuencia.append((et, pendientes_sin_etiqueta.pop(0)))
            elif idx_et < len(valores_aqui):
                secuencia.append((et, valores_aqui[idx_et]))
            else:
                secuencia.append((et, None))
        if len(valores_aqui) > len(etiquetas_aqui) and "\n" in (valor_cell or ""):
            ya_consumidos = max(0, len(etiquetas_aqui))
            pendientes_sin_etiqueta.extend(valores_aqui[ya_consumidos:])

    resultado_secuencia = []
    cola_valores = list(pendientes_sin_etiqueta)
    for et, val in secuencia:
        if val is None and cola_valores:
            val = cola_valores.pop(0)
        resultado_secuencia.append((et, val))

    for et, val in resultado_secuencia:
        campo = _campo_de_etiqueta(et)
        if campo and campo not in result and val:
            num = _first_number(val)
            if num:
                result[campo] = num.replace(".", ",") if "," in val and "." not in num else num
                result[campo] = num

    result["fibra_total"] = result.get("fibra", "")

    micronutrientes_extra = []
    for row in table:
        if row and row[0] and "Energía" in str(row[0]) and len(row) > 5:
            etiqueta_lateral = row[4]
            valor_lateral = row[5]
            if etiqueta_lateral and "Vitamina" in str(etiqueta_lateral) and valor_lateral:
                nombre = str(etiqueta_lateral).strip()
                valor = _first_number(str(valor_lateral).split("\n")[0])
                if valor:
                    micronutrientes_extra.append(f"{nombre} | {valor}")
            break

    if not micronutrientes_extra:
        for row in table:
            if not row or not row[0]:
                continue
            texto_bloque = str(row[0])
            if "(*)" in texto_bloque or any(
                m in texto_bloque for m in ["Calcio (mg)", "Fósforo (mg)", "Zinc (mg)", "Vitamina D", "Vitamina B12"]
            ):
                for linea in texto_bloque.split("\n"):
                    linea = linea.strip()
                    if not linea or linea.startswith("(*)"):
                        continue
                    m = re.match(r"^(.+?\([^)]+\))\s+([\d.,]+)", linea)
                    if m:
                        nombre = m.group(1).strip()
                        valor = m.group(2).strip()
                        micronutrientes_extra.append(f"{nombre} | {valor}")
                if micronutrientes_extra:
                    break

    if micronutrientes_extra:
        result["micronutrientes_extra"] = "\n".join(micronutrientes_extra)

    if "energia" not in result:
        etiqueta_a_campo_fallback = {
            "energia (kcal)": "energia",
            "proteinas (g)": "proteina",
            "proteina (g)": "proteina",
            "grasas totales (g)": "grasa_total",
            "grasa total (g)": "grasa_total",
            "grasas saturadas (g)": "grasa_sat",
            "grasa saturada (g)": "grasa_sat",
            "grasas monoinsat": "grasa_mono",
            "grasas poliinsat": "grasa_poli",
            "grasas trans (g)": "grasa_trans",
            "acidos grasos trans": "grasa_trans",
            "colesterol (mg)": "colesterol",
            "carbohidratos totales (g)": "carb_totales",
            "hidratos de carbono totales (g)": "carb_totales",
            "carbohidratos disp": "carb_disp",
            "hidratos de carbono disp": "carb_disp",
            "azucares totales (g)": "azucares",
            "fibra dietetica total (g)": "fibra",
            "sodio (mg)": "sodio",
        }
        for row in table:
            if not row or not row[0]:
                continue
            texto_bloque = str(row[0])
            if "INFORMACI" not in texto_bloque.upper() or "ENERG" not in texto_bloque.upper():
                continue
            for linea in texto_bloque.split("\n"):
                linea_norm = _norm_simple(linea)
                campo = None
                mejor_clave = ""
                for clave, c in etiqueta_a_campo_fallback.items():
                    if clave in linea_norm and len(clave) > len(mejor_clave):
                        mejor_clave, campo = clave, c
                if campo and campo not in result:
                    pos_etiqueta = linea_norm.find(mejor_clave)
                    texto_tras_etiqueta = linea[pos_etiqueta + len(mejor_clave):] if pos_etiqueta >= 0 else linea
                    numeros = re.findall(r'[\d]+(?:[.,][\d]+)?', texto_tras_etiqueta)
                    if numeros:
                        result[campo] = numeros[0]
            break

    result["fibra_total"] = result.get("fibra", result.get("fibra_total", ""))

    return result


def _extract_pagina2(table):
    """Extrae sensoriales, fisicoquímicos y microbiológicos de página 2."""
    result = {}

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
            metodologia = ""
            fq_rows.append(f"{parametro}|{metodologia}|{especificacion}")

    result["fq_rows"] = "\n".join(fq_rows)

    mb_rows = []
    in_mb = False
    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()
        if "7.3 PARÁMETROS MICROBIOLÓGICOS" in cell0:
            in_mb = True
            continue
        if in_mb and re.match(r'^7\.\d', cell0):
            break
        if in_mb and any(kw in cell0.upper() for kw in ("PESTICIDAS", "ALÉRGENOS", "ALERGENOS", "GLUTEN FREE", "METALES PESADOS", "MICOTOXINAS")):
            break
        if in_mb and re.match(r'^\*+\s*\S', cell0) and (len(row) <= 1 or not row[1] or not str(row[1]).strip()):
            # Notas al pie tipo "**: RSA Artículo..." o "* RSA, Art. 173,
            # punto 1.4." (con o sin dos puntos). Se distinguen de una fila
            # de dato real porque las notas no tienen valor en la columna
            # de categoría/clases/plan de muestreo.
            continue
        if in_mb:
            if cell0 in ("RSA, ARTÍCULO 173", "PARÁMETRO", ""):
                continue
            if cell0.startswith("RSA, ARTÍCULO"):
                continue
            if row[1] is None and "CLASES" in str(row).upper():
                continue
            parametro = _clean(row[0])
            if not parametro:
                continue
            valor_M = _clean(row[6]) if len(row) > 6 else ""
            if valor_M in ("-", "--", "") or valor_M.strip("-") == "":
                resultado = "0"
            else:
                resultado = f"<{valor_M}"
            mb_rows.append(f"{parametro}||{resultado}")

    result["mb_rows"] = "\n".join(mb_rows)

    return result


def _extract_pagina3(table):
    """Extrae datos de envase primario y secundario de página 3."""
    result = {
        "tipo_envase_primario": "",
        "tipo_envase_secundario": "",
    }

    in_primario = False
    in_secundario = False
    in_caja = False

    primario = {}
    secundario = {}
    caja = {}
    tipo_primario_de_parens = False
    tipo_secundario_de_parens = False

    for row in table:
        if not row or not row[0]:
            continue
        cell0 = str(row[0]).strip()

        if "ENVASE PRIMARIO" in cell0:
            in_primario = True
            in_secundario = False
            in_caja = False
            m = re.search(r'\(([^)]+)\)', cell0)
            if m:
                result["tipo_envase_primario"] = m.group(1).title()
                tipo_primario_de_parens = True
            continue

        if "ENVASE SECUNDARIO" in cell0:
            in_primario = False
            in_secundario = True
            in_caja = False
            m = re.search(r'\(([^)]+)\)', cell0)
            if m:
                result["tipo_envase_secundario"] = m.group(1).title()
                tipo_secundario_de_parens = True
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
                label_extra = cell0.split("envase", 1)[-1].strip()
                material_val = (label_extra + val).strip() if label_extra else val
                if "ARTÓN CORRUGADO" in material_val.upper() or "CARTÓN CORRUGADO" in material_val.upper():
                    material_val = "CARTÓN CORRUGADO (20)"
                elif material_val.count("(") > material_val.count(")"):
                    material_val += ")"
                caja["material"] = material_val
            elif "Cantidad" in cell0:
                caja["cantidad_unidades"] = val

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

    if not tipo_primario_de_parens:
        result["tipo_envase_primario"] = result.get("material_p", "") or "Bolsa plástica"
    if not tipo_secundario_de_parens and result.get("material_s"):
        result["tipo_envase_secundario"] = result.get("material_s", "")

    return result


def extract_ft(pdf_bytes):
    result = {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        table_p1 = _get_table(pdf, 1)
        table_p2 = _get_table(pdf, 2)
        table_p3 = _get_table(pdf, 3)

    result.update(_extract_pagina1(table_p1))
    result.update(_extract_pagina2(table_p2))
    result.update(_extract_pagina3(table_p3))

    return result
