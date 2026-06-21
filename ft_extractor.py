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

    # --- Tabla nutricional: estrategia unificada ---
    # Las FT varían mucho en cómo agrupan valores: algunas ponen cada
    # nutriente en su propia fila (1 valor en row[1]), otras apilan varios
    # valores con \n en una sola celda y dejan las filas de etiqueta
    # siguientes vacías (la cantidad de valores apilados varía: 3, 12, 16...),
    # y alguna pone "100g 1porción" pegados en la misma línea del stack.
    # Esta función recorre la tabla linealmente, consumiendo etiquetas y
    # valores en el orden en que aparecen, sin asumir un tamaño de stack fijo.

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
        # Buscar la clave más larga que matchee, para evitar que "azucar"
        # capture indebidamente parte de "azucares totales" o viceversa.
        mejor = None
        for clave, c in etiqueta_a_campo.items():
            if _norm_simple(clave) in et_norm:
                if mejor is None or len(clave) > len(mejor[0]):
                    mejor = (clave, c)
        return mejor[1] if mejor else None

    def _split_valor_doble(linea):
        """
        Si la línea tiene dos números separados por espacio (formato
        '318 38' = valor 100g + valor porción pegados), devuelve solo el
        primero (100g). Si es un solo número, lo devuelve tal cual.
        """
        partes = linea.strip().split()
        if len(partes) >= 2 and _first_number(partes[0]) and _first_number(partes[1]):
            return partes[0]
        return linea.strip()

    # Encontrar la fila de inicio de la tabla nutricional para no procesar
    # texto de otras secciones por error.
    idx_inicio = None
    for i, row in enumerate(table):
        if row and row[0] and "INFORMACIÓN NUTRICIONAL" in str(row[0]).upper():
            idx_inicio = i
            break
    if idx_inicio is None:
        idx_inicio = 0

    # Caso especial conocido: stack de exactamente 16 valores en la celda de
    # Energía (formato típico de FT de barras). Acá el orden interno incluye
    # posiciones que NO tienen fila de etiqueta propia visible (Azúcar,
    # Azúcar añadido, Fibra soluble, Fibra insoluble), por lo que el conteo
    # genérico de filas se desalinea. Usamos el orden fijo ya verificado.
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

    # Construir la secuencia de (etiqueta, valor_str_o_None) recorriendo
    # filas desde el inicio de la tabla nutricional. Cuando una celda trae
    # varios valores apilados con \n, se consumen como si fueran filas
    # adicionales con la etiqueta de las filas siguientes (que vienen vacías).
    # Si idx_inicio es None, ya se resolvió con el caso especial de 16 arriba.
    secuencia = []  # lista de (etiqueta, valor)
    pendientes_sin_etiqueta = []  # valores ya extraídos esperando etiqueta

    for i in range(idx_inicio if idx_inicio is not None else len(table), len(table)):
        row = table[i]
        if not row or not row[0]:
            continue
        cell0 = str(row[0])
        if "SELLOS" in cell0.upper() or "DESCRIPTORES" in cell0.upper():
            break
        # Una celda de etiqueta puede traer varias etiquetas apiladas
        # (ej. "Carbohidratos disp...\nAzúcares...\nSodio...")
        etiquetas_aqui = [l.strip() for l in cell0.split("\n") if l.strip()]
        if not etiquetas_aqui:
            continue

        # Buscar el valor disponible SOLO en la columna 1 (100g). Si esa
        # columna está vacía, NO se debe caer a la columna 2 (que es el
        # valor de "1 porción", no de 100g) — es preferible dejar el campo
        # vacío a inventar un dato incorrecto de la columna equivocada.
        valor_cell = None
        if len(row) > 1 and row[1] and str(row[1]).strip():
            valor_cell = str(row[1]).strip()

        if valor_cell and "\n" in valor_cell:
            valores_aqui = [_split_valor_doble(l) for l in valor_cell.split("\n") if l.strip()]
        elif valor_cell:
            valores_aqui = [_split_valor_doble(valor_cell)]
        else:
            valores_aqui = []

        # Emparejar: si hay valores pendientes de un stack anterior, esos
        # tienen PRIORIDAD sobre cualquier valor individual que la fila
        # pueda tener en su propia columna (que en algunos formatos es el
        # valor de "1 porción" colado en la posición equivocada, no 100g).
        for idx_et, et in enumerate(etiquetas_aqui):
            if pendientes_sin_etiqueta:
                secuencia.append((et, pendientes_sin_etiqueta.pop(0)))
            elif idx_et < len(valores_aqui):
                secuencia.append((et, valores_aqui[idx_et]))
            else:
                secuencia.append((et, None))
        # Si esta fila trajo un stack nuevo más grande que sus propias
        # etiquetas, el resto queda pendiente para las próximas filas.
        if len(valores_aqui) > len(etiquetas_aqui) and "\n" in (valor_cell or ""):
            ya_consumidos = max(0, len(etiquetas_aqui))
            pendientes_sin_etiqueta.extend(valores_aqui[ya_consumidos:])

    # Resolver pendientes contra futuras etiquetas sin valor (segunda pasada,
    # por si el emparejamiento quedó desalineado en el primer recorrido)
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
                # Mantener formato original (coma o punto) tal como viene
                result[campo] = num

    result["fibra_total"] = result.get("fibra", "")

    # --- Micronutrientes en columna lateral (ej. "Vitamina B₆ (mg)") ---
    # Aparecen en la misma fila que "Energía (kcal)", en una columna a la
    # derecha de la tabla principal (col índice 4=etiqueta, 5=valor 100g).
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

    # Micronutrientes empaquetados en una celda de texto libre debajo de la
    # tabla principal (ej. "Calcio (mg) 129 258 (32%*)\nFósforo (mg) 109...")
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

    # --- Fallback: tabla nutricional completa en una sola celda de texto ---
    # Algunas FT (ej. NotMayo Frasco) ponen toda la tabla "INFORMACIÓN
    # NUTRICIONAL ... Energía (kcal) 324 39 ... Sodio (mg) 730 88" en un
    # único bloque de texto con saltos de línea, sin columnas separadas.
    # Si los macronutrientes principales siguen vacíos, se intenta este
    # parser dedicado antes de devolver el resultado.
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
                    # El primer número que sigue al texto de la etiqueta es
                    # el valor de 100g (el segundo es 1 porción, se ignora).
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
