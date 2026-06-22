from docx import Document
import io
import unicodedata


def _normalize(text):
    if not text:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


# Palabras de relleno que aparecen en la FT pero no en el informe modelo
# (o viceversa) y que no aportan a la identidad del parámetro analítico.
_PALABRAS_RELLENO = {"recuento", "de", "del", "la", "el", "los", "las"}

# Sufijos de contexto de muestreo que algunas FT agregan al nombre del
# parámetro (p.ej. "Salmonella en 25g") y que el informe modelo no incluye.
import re as _re
_SUFIJO_MUESTREO_RE = _re.compile(r'\s+en\s+\d+\s*g\.?\s*$')


def _normalize_fuzzy(text):
    """
    Normalización tolerante para matchear parámetros entre la FT y el
    informe modelo, cuando la redacción exacta difiere (prefijos como
    "Recuento", sufijos como "en 25g", puntuación variable como "E.coli"
    vs "e. coli", asteriscos de nota al pie, u orden de palabras distinto
    como "aerobios mesófilos" vs "mesófilos aerobios").

    Devuelve las palabras significativas, sin puntuación ni acentos,
    ordenadas alfabéticamente, para que la comparación sea insensible al
    orden y a estas variaciones menores de redacción.
    """
    t = _normalize(text)
    t = _SUFIJO_MUESTREO_RE.sub("", t)
    # Quitar puntuación y asteriscos de nota al pie; mantener letras,
    # números y espacios.
    t = _re.sub(r'[^\w\s]', ' ', t)
    palabras = [p for p in t.split() if p and p not in _PALABRAS_RELLENO]
    return " ".join(sorted(palabras))


def _build_fuzzy_index(mapa):
    """
    Construye un índice {clave_fuzzy: metodologia} a partir del mapa ya
    normalizado (claves en formato _normalize). Si dos parámetros distintos
    del modelo colisionan en la misma clave fuzzy, se descarta esa clave
    (ambigua) en vez de aplicar una metodología potencialmente incorrecta.
    """
    indice = {}
    ambiguas = set()
    for clave_normal, metodologia in mapa.items():
        clave_fuzzy = _normalize_fuzzy(clave_normal)
        if not clave_fuzzy:
            continue
        if clave_fuzzy in indice and indice[clave_fuzzy] != metodologia:
            ambiguas.add(clave_fuzzy)
        else:
            indice[clave_fuzzy] = metodologia
    for clave in ambiguas:
        indice.pop(clave, None)
    return indice


def extract_metodologias(docx_bytes):
    doc = Document(io.BytesIO(docx_bytes))
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tbls = doc.element.body.findall(f'.//{{{WNS}}}tbl')

    mapa = {}
    tablas_usadas = 0

    for tbl in tbls:
        if tablas_usadas >= 2:
            break

        rows = tbl.findall(f'{{{WNS}}}tr')
        if len(rows) < 2:
            continue

        header_cells = rows[0].findall(f'{{{WNS}}}tc')
        header_texts = [
            ''.join(t.text or '' for t in c.findall(f'.//{{{WNS}}}t')).strip()
            for c in header_cells
        ]
        if len(header_texts) < 3:
            continue
        if "metodologia" not in _normalize(header_texts[1]):
            continue
        es_sensorial = False
        for row in rows[1:]:
            cells = row.findall(f'{{{WNS}}}tc')
            if len(cells) < 2:
                continue
            metodo_texto = ''.join(t.text or '' for t in cells[1].findall(f'.//{{{WNS}}}t')).strip()
            if "metodologia interna de analisis sensorial" in _normalize(metodo_texto):
                es_sensorial = True
                break
        if es_sensorial:
            continue

        for row in rows[1:]:
            cells = row.findall(f'{{{WNS}}}tc')
            if len(cells) < 2:
                continue
            parametro = ''.join(t.text or '' for t in cells[0].findall(f'.//{{{WNS}}}t')).strip()
            metodologia = ''.join(t.text or '' for t in cells[1].findall(f'.//{{{WNS}}}t')).strip()
            if parametro and metodologia:
                mapa[_normalize(parametro)] = metodologia

        tablas_usadas += 1

    return mapa


def aplicar_metodologias(rows_text, mapa, separador="|"):
    if not rows_text or not rows_text.strip():
        return rows_text

    indice_fuzzy = _build_fuzzy_index(mapa)

    lineas_nuevas = []
    for linea in rows_text.split("\n"):
        if not linea.strip():
            lineas_nuevas.append(linea)
            continue
        partes = [p.strip() for p in linea.split(separador)]
        if len(partes) < 2:
            lineas_nuevas.append(linea)
            continue
        parametro = partes[0]
        metodologia = partes[1] if len(partes) > 1 else ""
        resultado = partes[2] if len(partes) > 2 else ""

        if not metodologia:
            # 1) Match exacto (redacción idéntica salvo tildes/mayúsculas).
            metodologia = mapa.get(_normalize(parametro), "")
        if not metodologia:
            # 2) Match tolerante: ignora prefijos como "Recuento", sufijos
            # como "en 25g", puntuación variable y orden de palabras.
            metodologia = indice_fuzzy.get(_normalize_fuzzy(parametro), "")

        nueva_linea = f"{parametro} | {metodologia} | {resultado}" if resultado or len(partes) > 2 else f"{parametro} | {metodologia}"
        lineas_nuevas.append(nueva_linea)

    return "\n".join(lineas_nuevas)
