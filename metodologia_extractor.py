"""
metodologia_extractor.py
Extrae el mapa Parámetro -> Metodología desde un Informe de Análisis ya
generado (.docx), para reutilizar esas metodologías en nuevos análisis sin
tener que volver a tipearlas cada vez.

Solo se extraen las tablas de Fisicoquímicos y Microbiológicos (la de
Sensorial casi siempre usa la misma frase fija y no aporta valor real).

Uso:
    with open("informe_modelo.docx", "rb") as f:
        mapa = extract_metodologias(f.read())
    # mapa: {"ph": "Potenciometría", "solidos totales": "Gravimetría", ...}
"""

from docx import Document
import io
import unicodedata


def _normalize(text):
    """Normaliza texto para comparación: sin tildes, minúsculas, sin espacios extra."""
    if not text:
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


def extract_metodologias(docx_bytes):
    """
    Extrae {parámetro_normalizado: metodología} de las tablas de
    Fisicoquímicos y Microbiológicos de un Informe de Análisis ya generado.

    Identifica las tablas relevantes por su encabezado ('Metodología de
    análisis' en la segunda columna), sin asumir que sean siempre las
    primeras dos tablas del documento — así funciona aunque el orden o la
    cantidad de tablas varíe levemente entre versiones del template.
    """
    doc = Document(io.BytesIO(docx_bytes))
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tbls = doc.element.body.findall(f'.//{{{WNS}}}tbl')

    mapa = {}
    tablas_usadas = 0

    for tbl in tbls:
        if tablas_usadas >= 2:
            # Solo Fisicoquímicos y Microbiológicos — se ignora Sensorial
            # y cualquier otra tabla adicional que pueda aparecer después.
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
        # Es una tabla de Fisicoquímicos o Microbiológicos (formato
        # Parámetro | Metodología | Resultado). La de Sensorial tiene el
        # mismo encabezado pero su metodología es siempre fija, así que
        # también calza este filtro — se distingue más abajo por contenido.
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
    """
    Dado un texto multilínea "Parámetro | Metodología | Resultado" (el
    formato que ya usan fq_rows/mb_rows en el formulario), completa la
    columna de Metodología con el mapa de referencia cuando está vacía y
    el parámetro coincide (comparación normalizada, sin tildes/mayúsculas).

    No pisa una metodología que el usuario ya haya escrito.
    """
    if not rows_text or not rows_text.strip():
        return rows_text

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
            metodologia = mapa.get(_normalize(parametro), "")

        nueva_linea = f"{parametro} | {metodologia} | {resultado}" if resultado or len(partes) > 2 else f"{parametro} | {metodologia}"
        lineas_nuevas.append(nueva_linea)

    return "\n".join(lineas_nuevas)
