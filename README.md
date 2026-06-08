# Export Docs Tool — The Not Company

Herramienta local para generar documentación técnica de exportación a partir de templates .docx.

## Requisitos

- Python 3.9+
- LibreOffice (para conversión a PDF)

### Instalar Python y dependencias

```bash
pip install flask python-docx
```

### Instalar LibreOffice (para conversión a PDF)

- **Windows:** https://www.libreoffice.org/download/download/
- **Mac:** `brew install libreoffice`
- **Linux:** `sudo apt install libreoffice`

---

## Estructura de carpetas

```
export-tool/
├── server.py          ← Backend Flask
├── index.html         ← Interfaz web
├── requirements.txt
└── templates/         ← Los 9 templates .docx (no mover ni renombrar)
    ├── Template_CERTIFICADO_DE_CODIFICACIO_N_DE_FECHA_Y_LOTE.docx
    ├── Template_CERTIFICADO_DE_EMPAQUE.docx
    ├── Template_CERTIFICADO_PROCESO_DE_PRODUCCIO_N.docx
    ├── Template_INFORME_ANA_LISIS.docx
    ├── Template_INFORME_ANA_LISIS_NUTRICIONAL.docx
    ├── Template_INFORME_FUNCIONALIDAD_ADITIVOS.docx
    ├── Template_REPORTE_FIBRA_DIETE_TICA_TOTAL.docx
    ├── Template_REPORTE_FO_RMULA.docx
    └── Template__REPORTE_SABORIZANTES.docx
```

---

## Cómo usar

### 1. Iniciar el servidor

```bash
cd export-tool
python server.py
```

Abre **http://localhost:5000** en el navegador.

### 2. Generar un documento

1. Elegí el template desde el panel izquierdo
2. Completá los campos del formulario
3. (Opcional) Importá los datos desde un Excel con columnas: `campo` | `valor`
4. Clic en **"Generar y descargar .docx"**

### 3. Importar datos desde Excel

El Excel debe tener dos columnas (sin encabezado requerido):

| campo | valor |
|-------|-------|
| producto | NOT Burger 113g |
| energia | 210 |
| proteina | 18 |
| ...    | ...  |

El nombre de la columna "campo" debe coincidir con los IDs del formulario
(producto, energia, proteina, grasa_total, etc.)

### 4. Convertir documento firmado a PDF

Una vez que el manager devuelva el .docx firmado:
1. Scroll hacia abajo en la misma página
2. Arrastrá el .docx o hacé clic para seleccionarlo
3. Se descarga automáticamente el PDF

---

## IDs de campos por template

### Análisis Nutricional
producto, energia, proteina, grasa_total, grasa_sat, grasa_mono, grasa_poli,
grasa_trans, colesterol, carb_totales, carb_disp, azucares, fibra, sodio, fecha

### Reporte de Fórmula
producto, formula_rows (Ingrediente | % por línea), fecha

### Informe de Análisis
producto, fq_rows, mb_rows, apariencia, color, aroma, sabor, textura,
pb, cu, as_, sn, fe, fecha

### Certificado de Empaque
producto, tipo_envase_primario, ancho_p, largo_p, alto_p, peso_neto_p,
peso_bruto_p, material_p, tipo_envase_secundario, ancho_s, largo_s, alto_s,
peso_neto_s, peso_bruto_s, material_s, cantidad_unidades, fecha

### Certificado de Codificación
producto, fecha

### Certificado de Proceso de Producción
producto (código NOTXXX), fecha

### Informe de Aditivos
producto, ingredientes (uno por línea), aditivos (Nombre | Función por línea), fecha

### Reporte de Fibra
producto, fibra_total, fecha

### Reporte de Saborizantes
producto, sab_natural, sab_identico, sab_total, fecha
