#!/usr/bin/env python3
"""
Export Documentation Tool - Flask Backend
"""

import os
import re
import copy
import json
import subprocess
from datetime import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file, send_from_directory
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__, static_folder='.')
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates')

MONTHS_ES = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
}

def today_es():
    d = datetime.today()
    return f"{d.day} de {MONTHS_ES[d.month]} del {d.year}"

# Campos de firmante compartidos por todos los templates
FIRMANTE_FIELDS = [
    {"id": "firmante_nombre", "label": "Nombre del firmante", "type": "text", "placeholder": "Mariana Brizzio"},
    {"id": "firmante_cargo", "label": "Cargo del firmante", "type": "text", "placeholder": "Head R&D South Cone"},
]

TEMPLATES = {
    "certificado_codificacion": {
        "label": "Certificado de Codificación de Fecha y Lote",
        "file": "Template_CERTIFICADO_DE_CODIFICACIO_N_DE_FECHA_Y_LOTE.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "copacker_abrev", "label": "Abreviatura del copacker", "type": "text", "placeholder": "PB"},
            {"id": "copacker_nombre", "label": "Nombre completo del copacker", "type": "text", "placeholder": "Pacific Blu"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "certificado_empaque": {
        "label": "Certificado de Empaque",
        "file": "Template_CERTIFICADO_DE_EMPAQUE.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "tipo_envase_primario", "label": "Tipo de envase primario", "type": "text", "placeholder": "Bolsa plástica"},
            {"id": "ancho_p", "label": "Primario – Ancho (cm)", "type": "number", "placeholder": ""},
            {"id": "largo_p", "label": "Primario – Largo (cm)", "type": "number", "placeholder": ""},
            {"id": "alto_p", "label": "Primario – Alto (cm)", "type": "number", "placeholder": ""},
            {"id": "peso_neto_p", "label": "Primario – Peso Neto (g)", "type": "number", "placeholder": ""},
            {"id": "peso_bruto_p", "label": "Primario – Peso Bruto (g)", "type": "number", "placeholder": ""},
            {"id": "material_p", "label": "Primario – Material", "type": "text", "placeholder": "PVDC/PE"},
            {"id": "tipo_envase_secundario", "label": "Tipo de envase secundario", "type": "text", "placeholder": "Caja de cartón"},
            {"id": "ancho_s", "label": "Secundario – Ancho (cm)", "type": "number", "placeholder": ""},
            {"id": "largo_s", "label": "Secundario – Largo (cm)", "type": "number", "placeholder": ""},
            {"id": "alto_s", "label": "Secundario – Alto (cm)", "type": "number", "placeholder": ""},
            {"id": "peso_neto_s", "label": "Secundario – Peso Neto (kg)", "type": "number", "placeholder": ""},
            {"id": "peso_bruto_s", "label": "Secundario – Peso Bruto (kg)", "type": "number", "placeholder": ""},
            {"id": "material_s", "label": "Secundario – Material", "type": "text", "placeholder": "Cartón corrugado"},
            {"id": "cantidad_unidades", "label": "Secundario – Cantidad de unidades", "type": "number", "placeholder": ""},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "certificado_proceso": {
        "label": "Certificado de Proceso de Producción",
        "file": "Template_CERTIFICADO_PROCESO_DE_PRODUCCIO_N.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto (código)", "type": "text", "placeholder": "NOT20012"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "informe_analisis": {
        "label": "Informe de Análisis",
        "file": "Template_INFORME_ANA_LISIS.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "fq_rows", "label": "Análisis fisicoquímicos (Parámetro | Metodología | Resultado, una fila por línea)", "type": "textarea", "placeholder": "pH | AOAC 943.02 | 6.8\nHumedad | AOAC 925.10 | 62%"},
            {"id": "mb_rows", "label": "Análisis microbiológicos (Parámetro | Metodología | Resultado)", "type": "textarea", "placeholder": "Recuento aeróbico | ISO 4833 | <10 UFC/g"},
            {"id": "apariencia", "label": "Sensorial – Apariencia", "type": "text", "placeholder": "Homogénea, sin defectos"},
            {"id": "color", "label": "Sensorial – Color", "type": "text", "placeholder": "Marrón característico"},
            {"id": "aroma", "label": "Sensorial – Aroma", "type": "text", "placeholder": "Característico a carne"},
            {"id": "sabor", "label": "Sensorial – Sabor", "type": "text", "placeholder": "Umami, levemente salado"},
            {"id": "textura", "label": "Sensorial – Textura", "type": "text", "placeholder": "Firme, jugosa"},
            {"id": "pb", "label": "Contaminantes – Plomo Pb (mg/kg)", "type": "text", "placeholder": "<0.1"},
            {"id": "cu", "label": "Contaminantes – Cobre Cu (mg/kg)", "type": "text", "placeholder": "<0.5"},
            {"id": "as_", "label": "Contaminantes – Arsénico As (mg/kg)", "type": "text", "placeholder": "<0.1"},
            {"id": "sn", "label": "Contaminantes – Estaño Sn (mg/kg)", "type": "text", "placeholder": "<1.0"},
            {"id": "fe", "label": "Contaminantes – Hierro Fe (mg/kg)", "type": "text", "placeholder": "<5.0"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "informe_nutricional": {
        "label": "Informe de Análisis Nutricional",
        "file": "Template_INFORME_ANA_LISIS_NUTRICIONAL.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "energia", "label": "Energía (kcal)", "type": "number", "placeholder": ""},
            {"id": "proteina", "label": "Proteína (g)", "type": "number", "placeholder": ""},
            {"id": "grasa_total", "label": "Grasa total (g)", "type": "number", "placeholder": ""},
            {"id": "grasa_sat", "label": "Grasa saturada (g)", "type": "number", "placeholder": ""},
            {"id": "grasa_mono", "label": "Grasa monoinsaturada (g)", "type": "number", "placeholder": ""},
            {"id": "grasa_poli", "label": "Grasa poliinsaturada (g)", "type": "number", "placeholder": ""},
            {"id": "grasa_trans", "label": "Grasa trans (g)", "type": "number", "placeholder": ""},
            {"id": "colesterol", "label": "Colesterol (mg)", "type": "number", "placeholder": ""},
            {"id": "carb_totales", "label": "Carbohidratos totales (g)", "type": "number", "placeholder": ""},
            {"id": "carb_disp", "label": "Carbohidratos disponibles (g)", "type": "number", "placeholder": ""},
            {"id": "azucares", "label": "Azúcares (g)", "type": "number", "placeholder": ""},
            {"id": "fibra", "label": "Fibra (g)", "type": "number", "placeholder": ""},
            {"id": "sodio", "label": "Sodio (mg)", "type": "number", "placeholder": ""},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "informe_aditivos": {
        "label": "Informe de Funcionalidad de Aditivos",
        "file": "Template_INFORME_FUNCIONALIDAD_ADITIVOS.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "ingredientes", "label": "Lista de ingredientes (uno por línea, aditivos en MAYÚSCULAS)", "type": "textarea", "placeholder": "Agua\nProteína de soya\nACEITE DE GIRASOL\nSal\nAROMATIZANTE NATURAL"},
            {"id": "aditivos", "label": "Aditivos y su función (Nombre | Función, uno por línea)", "type": "textarea", "placeholder": "ACEITE DE GIRASOL | Agente de relleno\nAROMATIZANTE NATURAL | Aromatizante"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "reporte_fibra": {
        "label": "Reporte de Fibra Dietética Total",
        "file": "Template_REPORTE_FIBRA_DIETE_TICA_TOTAL.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "fibra_total", "label": "Fibra Dietética Total (g/100g)", "type": "number", "placeholder": ""},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "reporte_formula": {
        "label": "Reporte de Fórmula de Producto",
        "file": "Template_REPORTE_FO_RMULA.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "formula_rows", "label": "Ingredientes y cantidades (Ingrediente | % en 100g, uno por línea)", "type": "textarea", "placeholder": "Agua | 55.0\nProteína de soya texturizada | 20.0\nAceite de girasol | 10.0"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },

    "reporte_saborizantes": {
        "label": "Reporte de Saborizantes",
        "file": "Template__REPORTE_SABORIZANTES.docx",
        "fields": [
            {"id": "producto", "label": "Nombre del producto", "type": "text", "placeholder": "NOT Burger 113g"},
            {"id": "sab_natural", "label": "Saborizante Natural (%)", "type": "text", "placeholder": "0.5"},
            {"id": "sab_identico", "label": "Saborizante Idéntico Natural (%)", "type": "text", "placeholder": "0.3"},
            {"id": "sab_total", "label": "Total saborizantes (%)", "type": "text", "placeholder": "0.8"},
            {"id": "fecha", "label": "Fecha del documento", "type": "text", "placeholder": today_es()},
        ] + FIRMANTE_FIELDS,
    },
}


# ─── Document generation helpers ─────────────────────────────────────────────

def replace_in_paragraph(para, replacements):
    full = ''.join(r.text for r in para.runs)
    changed = False
    for old, new in replacements.items():
        if old in full:
            full = full.replace(old, new)
            changed = True
    if changed and para.runs:
        para.runs[0].text = full
        for r in para.runs[1:]:
            r.text = ''

def replace_all(doc, replacements):
    for para in doc.paragraphs:
        replace_in_paragraph(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para, replacements)
    for section in doc.sections:
        for hdr in [section.header, section.footer]:
            if hdr:
                for para in hdr.paragraphs:
                    replace_in_paragraph(para, replacements)

def apply_firmante(doc, data):
    """Replace firmante name and cargo in all paragraphs."""
    nombre = data.get('firmante_nombre', '').strip()
    cargo = data.get('firmante_cargo', '').strip()
    if nombre:
        replace_all(doc, {'Mariana Brizzio': nombre, 'Javiera Mujica': nombre})
    if cargo:
        replace_all(doc, {
            'Head R&D South Cone': cargo,
            'Head R****&****D South Cone': cargo,
            'R&D Manager South Cone': cargo,
            'R****&****D Manager South Cone': cargo,
        })

def fill_table_rows(table, data_rows, start_row=1):
    while len(table.rows) > start_row:
        tr = table.rows[-1]._tr
        tr.getparent().remove(tr)
    for row_data in data_rows:
        row = copy.deepcopy(table.rows[start_row - 1])
        cells = row.cells
        for i, val in enumerate(row_data):
            if i < len(cells):
                for para in cells[i].paragraphs:
                    for run in para.runs:
                        run.text = ''
                    if para.runs:
                        para.runs[0].text = str(val)
                    else:
                        para.add_run(str(val))
        table._tbl.append(row._tr)

def parse_rows(text, separator='|'):
    rows = []
    for line in text.strip().split('\n'):
        if line.strip():
            parts = [p.strip() for p in line.split(separator)]
            rows.append(parts)
    return rows


# ─── Per-template generation ─────────────────────────────────────────────────

def generate_doc(template_id, data):
    tmpl = TEMPLATES[template_id]
    path = os.path.join(TEMPLATES_DIR, tmpl['file'])
    doc = Document(path)

    producto = data.get('producto', 'XXX')
    fecha = data.get('fecha', today_es())

    if template_id == 'certificado_codificacion':
        abrev = data.get('copacker_abrev', 'xx')
        nombre = data.get('copacker_nombre', 'xx')
        # First fill table row 3 (copacker fields use 'xx' lowercase)
        if doc.tables:
            t = doc.tables[0]
            row = t.rows[3]
            # Col 1 (Estructura): 'PB XXX (HH:MM)' -> replace PB with abrev
            cell1 = row.cells[1]
            for p in cell1.paragraphs:
                replace_in_paragraph(p, {'PB': abrev})
            # Col 2 (Significado): replace PB (abrev) and Pacificblu (nombre completo)
            cell2 = row.cells[2]
            for p in cell2.paragraphs:
                replace_in_paragraph(p, {'PB': abrev, 'Pacificblu': nombre})
        # Then replace product name in body text and date (not in table)
        for para in doc.paragraphs:
            replace_in_paragraph(para, {'XXX': producto, 'XX de XX del 2025': fecha})

    elif template_id == 'certificado_proceso':
        replace_all(doc, {'NOTXXX': producto, 'XX de mayo de 2025': fecha})

    elif template_id == 'certificado_empaque':
        replace_all(doc, {'XXX': producto, 'XX de XX del 2025': fecha})
        tables = doc.tables
        if len(tables) >= 1:
            t1 = tables[0]
            mapping_p = [
                ('Ancho (cm)', data.get('ancho_p','')),
                ('Largo (cm)', data.get('largo_p','')),
                ('Alto (cm)', data.get('alto_p','')),
                ('Peso Neto (g)', data.get('peso_neto_p','')),
                ('Peso Bruto (g)', data.get('peso_bruto_p','')),
                ('Material', data.get('material_p','')),
            ]
            for row in t1.rows[1:]:
                label = row.cells[0].text.strip()
                for lbl, val in mapping_p:
                    if lbl in label:
                        cell = row.cells[1]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(val)
                            else: p.add_run(str(val))
        if len(tables) >= 2:
            t2 = tables[1]
            mapping_s = [
                ('Ancho (cm)', data.get('ancho_s','')),
                ('Largo (cm)', data.get('largo_s','')),
                ('Alto (cm)', data.get('alto_s','')),
                ('Peso Neto (kg)', data.get('peso_neto_s','')),
                ('Peso Bruto (kg)', data.get('peso_bruto_s','')),
                ('Material', data.get('material_s','')),
                ('Cantidad de unidades', data.get('cantidad_unidades','')),
            ]
            for row in t2.rows[1:]:
                label = row.cells[0].text.strip()
                for lbl, val in mapping_s:
                    if lbl in label:
                        cell = row.cells[1]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(val)
                            else: p.add_run(str(val))
        found = 0
        for para in doc.paragraphs:
            if 'Tipo de envase:' in para.text:
                found += 1
                if found == 1:
                    replace_in_paragraph(para, {'Tipo de envase: XXX': f'Tipo de envase: {data.get("tipo_envase_primario","")}'})
                elif found == 2:
                    replace_in_paragraph(para, {'Tipo de envase: XXX': f'Tipo de envase: {data.get("tipo_envase_secundario","")}'})

    elif template_id == 'informe_analisis':
        replace_all(doc, {'XXX': producto, 'XX de XX del 2025': fecha})
        tables = doc.tables
        if len(tables) > 0 and data.get('fq_rows'):
            fill_table_rows(tables[0], parse_rows(data['fq_rows']))
        if len(tables) > 1 and data.get('mb_rows'):
            fill_table_rows(tables[1], parse_rows(data['mb_rows']))
        if len(tables) > 2:
            sensorial = {
                'Apariencia': data.get('apariencia',''),
                'Color': data.get('color',''),
                'Aroma': data.get('aroma',''),
                'Sabor': data.get('sabor',''),
                'Textura': data.get('textura',''),
            }
            for row in tables[2].rows[1:]:
                key = row.cells[0].text.strip()
                if key in sensorial:
                    cell = row.cells[2]
                    for p in cell.paragraphs:
                        for r in p.runs: r.text=''
                        if p.runs: p.runs[0].text = sensorial[key]
                        else: p.add_run(sensorial[key])
        if len(tables) > 3:
            contam = {
                'Plomo (Pb)': data.get('pb',''),
                'Cobre (Cu)': data.get('cu',''),
                'Arsénico (As)': data.get('as_',''),
                'Estaño (Sn)': data.get('sn',''),
                'Hierro (Fe)': data.get('fe',''),
            }
            for row in tables[3].rows[1:]:
                key = row.cells[0].text.strip()
                for k, v in contam.items():
                    if k in key:
                        cell = row.cells[2]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(v)
                            else: p.add_run(str(v))

    elif template_id == 'informe_nutricional':
        replace_all(doc, {'XXX': producto, 'XX de XX del 2025': fecha})
        nut_map = {
            'Energía (kcal)': 'energia',
            'Proteína (g)': 'proteina',
            'Grasa total (g)': 'grasa_total',
            'Grasa saturada (g)': 'grasa_sat',
            'Grasa monoinsaturada (g)': 'grasa_mono',
            'Grasa poliinsaturada (g)': 'grasa_poli',
            'Grasa trans (g)': 'grasa_trans',
            'Colesterol (mg)': 'colesterol',
            'Carbohidratos totales (g)': 'carb_totales',
            'Carbohidratos disponibles (g)': 'carb_disp',
            'Azúcares (g)': 'azucares',
            'Fibra (g)': 'fibra',
            'Sodio (mg)': 'sodio',
        }
        if doc.tables:
            for row in doc.tables[0].rows[1:]:
                key = row.cells[0].text.strip()
                for lbl, fid in nut_map.items():
                    if lbl in key:
                        cell = row.cells[1]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(data.get(fid,''))
                            else: p.add_run(str(data.get(fid,'')))

    elif template_id == 'informe_aditivos':
        replace_all(doc, {'XX': producto, 'XX de XX del 2025': fecha})
        if doc.tables and data.get('ingredientes'):
            lines = [l.strip() for l in data['ingredientes'].split('\n') if l.strip()]
            t = doc.tables[0]
            while len(t.rows) > 0:
                t._tbl.remove(t.rows[-1]._tr)
            for ing in lines:
                row = OxmlElement('w:tr')
                tc = OxmlElement('w:tc')
                p = OxmlElement('w:p')
                r = OxmlElement('w:r')
                t_el = OxmlElement('w:t')
                t_el.text = ing
                r.append(t_el); p.append(r); tc.append(p); row.append(tc)
                t._tbl.append(row)
        if data.get('aditivos'):
            pairs = parse_rows(data['aditivos'])
            aditivo_paras = [p for p in doc.paragraphs if 'ADITIVO' in p.text]
            for i, para in enumerate(aditivo_paras):
                if i < len(pairs):
                    nombre = pairs[i][0] if len(pairs[i]) > 0 else ''
                    funcion = pairs[i][1] if len(pairs[i]) > 1 else ''
                    if len(para.runs) >= 2:
                        para.runs[0].text = nombre
                        para.runs[1].text = funcion
                    else:
                        replace_in_paragraph(para, {'ADITIVO': nombre, 'Función tecnológica': funcion})

    elif template_id == 'reporte_fibra':
        replace_all(doc, {
            'xxxxxx': producto,
            'xx de febrero del 202x': fecha,
            'xx de febrero del 202X': fecha,
        })
        if doc.tables and data.get('fibra_total'):
            for row in doc.tables[0].rows[1:]:
                cell = row.cells[1]
                for p in cell.paragraphs:
                    for r in p.runs: r.text=''
                    if p.runs: p.runs[0].text = str(data['fibra_total'])
                    else: p.add_run(str(data['fibra_total']))

    elif template_id == 'reporte_formula':
        replace_all(doc, {'XXX': producto, 'XX de XX del 2025': fecha})
        if doc.tables and data.get('formula_rows'):
            rows = parse_rows(data['formula_rows'])
            t = doc.tables[0]
            while len(t.rows) > 2:
                t._tbl.remove(t.rows[-2]._tr)
            total_tr = t.rows[-1]._tr
            for row_data in rows:
                new_row = copy.deepcopy(t.rows[0])
                for i, val in enumerate(row_data):
                    if i < len(new_row.cells):
                        cell = new_row.cells[i]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(val)
                            else: p.add_run(str(val))
                t._tbl.insertBefore(new_row._tr, total_tr)

    elif template_id == 'reporte_saborizantes':
        replace_all(doc, {
            'XXXXXXX': producto,
            'XX de febrero del 202X': fecha,
            'XX de febrero del 202x': fecha,
        })
        sab_map = {
            'SABORIZANTE NATURAL': 'sab_natural',
            'SABORIZANTE IDENTICO NATURAL': 'sab_identico',
        }
        if doc.tables:
            for row in doc.tables[0].rows[1:]:
                key = row.cells[0].text.strip()
                for lbl, fid in sab_map.items():
                    if lbl in key:
                        cell = row.cells[1]
                        for p in cell.paragraphs:
                            for r in p.runs: r.text=''
                            if p.runs: p.runs[0].text = str(data.get(fid,'')) + ' %'
                            else: p.add_run(str(data.get(fid,'')) + ' %')
            for row in doc.tables[0].rows:
                if 'TOTAL' in row.cells[0].text:
                    cell = row.cells[1]
                    for p in cell.paragraphs:
                        for r in p.runs: r.text=''
                        val = str(data.get('sab_total','')) + ' %'
                        if p.runs: p.runs[0].text = val
                        else: p.add_run(val)

    # Apply firmante to ALL templates at the end
    apply_firmante(doc, data)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/templates')
def get_templates():
    return jsonify({
        tid: {'label': t['label'], 'fields': t['fields']}
        for tid, t in TEMPLATES.items()
    })

@app.route('/api/generate', methods=['POST'])
def generate():
    body = request.get_json()
    tid = body.get('template_id')
    data = body.get('data', {})
    if tid not in TEMPLATES:
        return jsonify({'error': 'Template no encontrado'}), 400
    try:
        buf = generate_doc(tid, data)
        label = TEMPLATES[tid]['label'].replace(' ', '_')
        producto = data.get('producto', 'producto').replace(' ', '_')
        filename = f"{label}_{producto}.docx"
        return send_file(buf, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/api/convert-pdf', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    f = request.files['file']
    tmp_docx = f'/tmp/upload_{f.filename}'
    tmp_pdf = tmp_docx.replace('.docx', '.pdf')
    f.save(tmp_docx)
    try:
        # Find libreoffice binary (path varies by environment)
        import shutil
        lo_bin = shutil.which('libreoffice') or shutil.which('soffice') or '/usr/bin/libreoffice'
        subprocess.run([lo_bin, '--headless', '--convert-to', 'pdf', '--outdir', '/tmp', tmp_docx],
                       capture_output=True)
        if os.path.exists(tmp_pdf):
            return send_file(tmp_pdf, as_attachment=True,
                             download_name=f.filename.replace('.docx', '.pdf'),
                             mimetype='application/pdf')
        else:
            return jsonify({'error': 'Conversión fallida. LibreOffice no está disponible en el servidor.'}), 500
    finally:
        if os.path.exists(tmp_docx): os.remove(tmp_docx)

@app.route('/api/import-docx', methods=['POST'])
def import_docx():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    f = request.files['file']
    try:
        doc = Document(f)
        campos = {}
        # Strategy 1: read a two-column table (campo | valor)
        for table in doc.tables:
            for row in table.rows:
                if len(row.cells) >= 2:
                    campo = row.cells[0].text.strip()
                    valor = row.cells[1].text.strip()
                    if campo and valor and campo != 'campo' and campo != 'Campo':
                        campos[campo] = valor
        # Strategy 2: read paragraphs with "campo: valor" or "campo | valor" format
        if not campos:
            for para in doc.paragraphs:
                txt = para.text.strip()
                for sep in ['|', ':']:
                    if sep in txt:
                        parts = txt.split(sep, 1)
                        if len(parts) == 2:
                            campo = parts[0].strip()
                            valor = parts[1].strip()
                            if campo and valor:
                                campos[campo] = valor
                        break
        return jsonify(campos)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
