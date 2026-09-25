import io
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Catálogo de datos configurables por Inversionista
DATOS_INVERSIONISTAS = {
    "AMERIS": {
        "asegurado": "Ameris Financiamiento para Acceso a la Vivienda II Fondo de Inversión",
        "rut_asegurado": "76.629.842-7",
        "beneficiario": "Ameris Financiamiento para Acceso a la Vivienda II Fondo de Inversión",
        "rut_beneficiario": "76.629.842-7",
        "banco": "BCI",
        "cta_cte": "35194669",
        "email": "tesoreria@ameris.cl",
        "telefono": "+56 2 2499 7600",
        "firmantes": [
            {"nombre": "Ignacio Montané Yunge", "rut": "13.333.514-5"},
            {"nombre": "Rodrigo Guzmán Mohr", "rut": "10.931.176-6"}
        ]
    },
    "4LIFE": {
        "asegurado": "4 LIFE SEGUROS DE VIDA SA",
        "rut_asegurado": "76.418.751-2",
        "beneficiario": "4 LIFE SEGUROS DE VIDA SA",
        "rut_beneficiario": "76.418.751-2",
        "banco": "Banco de Chile",
        "cta_cte": "8000927607",
        "email": "control.inversiones@4lifeseguros.cl, tesoreria@4lifeseguros.cl, riesgo.credito@4lifeseguros.cl",
        "telefono": "",
        "firmantes": [
            {"nombre": "Mauricio Balbontín O’Ryan", "rut": ""}
        ]
    },
    "BTG": {
        "asegurado": "BTG Pactual Financiamiento con Garantías Inmobiliarias Fondo de Inversión.",
        "rut_asegurado": "76.777.679-9",
        "beneficiario": "BTG PACTUAL CHILE S.A. ADMINISTRADORA GENERAL",
        "rut_beneficiario": "96.966.250-7",
        "banco": "Banco de Chile",
        "cta_cte": "8002305802",
        "email": "BackFondosRegulados@btgpactual.com, trini.werth@btgpactual.com, OL-Activosalternativos@btgpactual.com",
        "telefono": "+562 2587 5124",
        "firmantes": [
            {"nombre": "José Miguel Correa Tagle", "rut": ""},
            {"nombre": "Trini Werth Ahumada", "rut": ""}
        ]
    },
    "BTG MHE HABITAC": {
        "asegurado": "BTG PACTUAL MHE HABITACIONAL FONDO DE INVERSIÓN",
        "rut_asegurado": "77.081.709-9",
        "beneficiario": "BTG PACTUAL MHE HABITACIONAL FONDO DE INVERSIÓN",
        "rut_beneficiario": "96.966.250-7",
        "banco": "Banco de Chile",
        "cta_cte": "8004170003",
        "email": "BackFondosRegulados@btgpactual.com, trini.werth@btgpactual.com, Ol-ActivosAlternativos@btgpactual.com",
        "telefono": "+562 2587 5124",
        "firmantes": [
            {"nombre": "José Miguel Correa Tagle", "rut": ""},
            {"nombre": "Trini Werth Ahumada", "rut": ""}
        ]
    },
    "BTG SUBSIDIO": {
        "asegurado": "BTG Pactual Deuda Habitacional con Subsidio Fondo de Inversión",
        "rut_asegurado": "76.627.616-4",
        "beneficiario": "BTG PACTUAL CHILE S.A. ADMINISTRADORA GENERAL",
        "rut_beneficiario": "96.966.250-7",
        "banco": "Banco de Chile",
        "cta_cte": "8007484401",
        "email": "BackFondosRegulados@btgpactual.com, trini.werth@btgpactual.com, Ol-ActivosAlternativos@btgpactual.com",
        "telefono": "+562 2587 5124",
        "firmantes": [
            {"nombre": "José Miguel Correa Tagle", "rut": ""},
            {"nombre": "Trini Werth Ahumada", "rut": ""}
        ]
    },
    "CONSORCIO": {
        "asegurado": "Compañía de Seguros de Vida Consorcio Nacional de Seguros S.A.",
        "rut_asegurado": "99.012.000-5",
        "beneficiario": "Compañía de Seguros de Vida Consorcio Nacional de Seguros S.A.",
        "rut_beneficiario": "99.012.000-5",
        "banco": "Santander",
        "cta_cte": "1059939",
        "email": "andrea.farfan@consorcio.cl, mariajose.maldonado@consorcio.cl",
        "telefono": "+562 2449 2370",
        "firmantes": [
            {"nombre": "", "rut": ""}
        ]
    },
    "DELPHOS": {
        "asegurado": "Compartment Lagorta - Exsilio Investments S.A.R.L (DELPHOS) ",
        "rut_asegurado": "GB54930064",
        "beneficiario": "AVLA SERVICIOS SPA",
        "rut_beneficiario": "76.255.149-7",
        "banco": "BCI",
        "cta_cte": "35428759",
        "email": "vicente.orueta@avla.com",
        "telefono": "",
        "firmantes": [
            {"nombre": "", "rut": ""}
        ]
    },
    "AVLA": {
        "asegurado": "Compartment Lagorta - Exsilio Investments S.A.R.L (DELPHOS) ",
        "rut_asegurado": "GB54930064",
        "beneficiario": "AVLA SERVICIOS SPA",
        "rut_beneficiario": "76.255.149-7",
        "banco": "BCI",
        "cta_cte": "35428759",
        "email": "vicente.orueta@avla.com",
        "telefono": "",
        "firmantes": [
            {"nombre": "", "rut": ""}
        ]
    },
    "MONEDA": {
        "asegurado": "Fondo de Inversión Privado SGR",
        "rut_asegurado": "76.736.411-3",
        "beneficiario": "Fondo de Inversión Privado SGR",
        "rut_beneficiario": "76.736.411-3",
        "banco": "Banco Bice",
        "cta_cte": "07-02886-5",
        "email": "ops@finixservices.com",
        "telefono": "+562 2599 4677",
        "firmantes": [
            {"nombre": "", "rut": ""}
        ]
    },
    "SURA": {
        "asegurado": "Seguros de Vida Sura S.A.",
        "rut_asegurado": "96.549.050-7",
        "beneficiario": "Seguros de Vida Sura S.A.",
        "rut_beneficiario": "96.549.050-7",
        "banco": "Banco de Chile",
        "cta_cte": "00-800-81516-05",
        "email": "tesoreriaSV.CL@surainvestments.com, legalchile@surainvestments.com",
        "telefono": "",
        "firmantes": [
            {"nombre": "Sergio Mirabelli", "rut": ""},
            {"nombre": "Rafal Szybisz", "rut": ""}
        ]
    },
    "WEG": {
        "asegurado": "Fondo de Inversión WEG-2",
        "rut_asegurado": "76.754.249-6",
        "beneficiario": "Fondo de Inversión WEG-2",
        "rut_beneficiario": "76.754.249-6",
        "banco": "Banco Security",
        "cta_cte": "917677045",
        "email": "meugenin@wegcapital.cl",
        "telefono": "+562 26057760",
        "firmantes": [
            {"nombre": "Fuad Eduardo Escaffi Johnson", "rut": ""}
        ]
    },
    "WEG WAREHOUSE": {
        "asegurado": "Fondo de Inversión WEG-2",
        "rut_asegurado": "76.754.249-6",
        "beneficiario": "Fondo de Inversión WEG-2",
        "rut_beneficiario": "76.754.249-6",
        "banco": "Banco Security",
        "cta_cte": "917677045",
        "email": "meugenin@wegcapital.cl",
        "telefono": "+562 26057760",
        "firmantes": [
            {"nombre": "Fuad Eduardo Escaffi Johnson", "rut": ""}
        ]
    },
    # Se pueden agregar más inversionistas según sea necesario
    "DEFAULT": {
        "asegurado": "",
        "rut_asegurado": "",
        "beneficiario": "",
        "rut_beneficiario": "",
        "banco": "",
        "cta_cte": "",
        "email": "",
        "telefono": "",
        "firmantes": [
            {"nombre": "Representante1", "rut": ""},
            {"nombre": "Representante2", "rut": ""}
        ]
    }
}

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

hoy = date.today()
def obtener_fecha_texto(hoy: date) -> str:
    """Convierte un objeto fecha a formato texto en español (ej. '10 de agosto de 2026')."""
    return f"{hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"

def agregar_borde_superior_celda(cell, sz="8", val="single", color="000000"):
    """Agrega una línea superior a una celda (sz='8' equivale a 1 pto)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), val)
    top.set(qn('w:sz'), sz)
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), color)
    tcBorders.append(top)
    tcPr.append(tcBorders)

def quitar_bordes_tabla(table):
    """Remueve los bordes visibles de una tabla para usarla únicamente como alineador."""
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def generar_word(
    id_operacion: str,
    nombre_cliente: str,
    rut_cliente: str,
    inversionista_clave: str,
    monto_siniestro: float
) -> io.BytesIO:
    
    """
    Genera el documento Word (.docx) de Aviso de Siniestro dirigido a AVLA.
    """

    # Buscar datos del inversionista o usar el perfil predeterminado
    key_inv = inversionista_clave.upper().strip() if inversionista_clave else "DEFAULT"
    datos_inv = DATOS_INVERSIONISTAS.get(key_inv, DATOS_INVERSIONISTAS["DEFAULT"])

    doc = Document()

    # Configurar márgenes estándar (2.5 cm)
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # Configuración del Encabezado con Imagen
    section = doc.sections[0]
    header = section.header
    p_header = header.paragraphs[0]
    p_header.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Agregar la imagen
    run_header = p_header.add_run()
    run_header.add_picture("Avla.jpg", width=Cm(3))  # Ajusta la ruta y el ancho deseado

    # Estilo base
    style = doc.styles['Normal']
    style.paragraph_format.line_spacing = 1.0
    style.paragraph_format.space_after = Pt(2)
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)

    # 1. Fecha (Alineada a la derecha)
    p_fecha = doc.add_paragraph(obtener_fecha_texto(hoy))
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # 2. Título principal
    p_titulo = doc.add_paragraph()
    r_t1 = p_titulo.add_run("AVISO DE SINIESTRO\n")
    r_t1.bold = True
    r_t1.underline = True
    r_t1.font.size = Pt(13)
    r_t2 = p_titulo.add_run("PÓLIZA DE CRÉDITO")
    r_t2.bold = True
    r_t2.font.size = Pt(12)
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_titulo.paragraph_format.space_after = Pt(14)

    # 3. Encabezado de Destinatario (AVLA)
    p_dest = doc.add_paragraph()
    saludo = p_dest.add_run(
        "Señores\n"
        "AVLA Seguros de Crédito y Garantía S.A.\n"
        "Cerro El Plomo 5420, Oficina 701, Las Condes, Santiago\n")
    saludo.bold = True
    presentes= p_dest.add_run("Presentes")
    presentes.bold = True
    presentes.underline = True
    p_dest.paragraph_format.space_after = Pt(14)

    # 4. Cuerpo / Texto introductorio
    p_cuerpo = doc.add_paragraph()
    p_cuerpo.add_run("De nuestra consideración:\n\n")
    p_cuerpo.add_run(
        "Informamos a ustedes que por medio de la presente venimos en dar aviso de siniestro "
        "y solicitar el pago de la indemnización de la póliza de crédito que se detalla a continuación:"
    )
    p_cuerpo.paragraph_format.space_after = Pt(10)

    # 5. TABLA 2x7 DE DATOS DE LA OPERACIÓN (Alineación uniforme)
    t_op = doc.add_table(rows=7, cols=2)
    quitar_bordes_tabla(t_op)
    
    t_op.autofit = False
    for row in t_op.rows:
        row.cells[0].width = Cm(2.5)
        row.cells[1].width = Cm(14)

    monto_fmt = f"UF    {monto_siniestro:,.4f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # Fila 0: Operacional N°
    p00 = t_op.cell(0, 0).paragraphs[0]
    p00.add_run("Operacional")
    p01 = t_op.cell(0, 1).paragraphs[0]
    p01.add_run(f"N° {id_operacion}").bold = True

    # Fila 1: Asegurado
    p10 = t_op.cell(1, 0).paragraphs[0]
    p10.add_run("Asegurado")
    p11 = t_op.cell(1, 1).paragraphs[0]
    p11.add_run(datos_inv['asegurado']).bold = True

    # Fila 2: RUT Asegurado
    p20 = t_op.cell(2, 0).paragraphs[0]
    p20.add_run("RUT N°")
    p21 = t_op.cell(2, 1).paragraphs[0]
    p21.add_run(datos_inv['rut_asegurado']).bold = True

    # Fila 3: Beneficiario
    p30 = t_op.cell(3, 0).paragraphs[0]
    p30.add_run("Beneficiario")
    p31 = t_op.cell(3, 1).paragraphs[0]
    p31.add_run(datos_inv['asegurado']).bold = True

    # Fila 4: RUT Beneficiario
    p40 = t_op.cell(4, 0).paragraphs[0]
    p40.add_run("RUT N°")
    p41 = t_op.cell(4, 1).paragraphs[0]
    p41.add_run(datos_inv['rut_asegurado']).bold = True

    # Fila 5: Deudor
    p50 = t_op.cell(5, 0).paragraphs[0]
    p50.add_run("Deudor")
    p51 = t_op.cell(5, 1).paragraphs[0]
    p51.add_run(nombre_cliente).bold = True

    # Fila 6: RUT Deudor
    p60 = t_op.cell(6, 0).paragraphs[0]
    p60.add_run("RUT N°")
    p61 = t_op.cell(6, 1).paragraphs[0]
    p61.add_run(rut_cliente).bold = True

    monto = doc.add_paragraph()
    monto.paragraph_format.space_before = Pt(8)
    monto.add_run(f"Monto de Indemnización Requerido: {monto_fmt}").bold = True

    # 6. DATOS DE TRANSFERENCIA
    p_banco_intro = doc.add_paragraph()
    p_banco_intro.paragraph_format.space_before = Pt(8)
    p_banco_intro.add_run("Por Favor realizar depósito/transferencia o emitir cheque/vale vista a nombre de:")

    # Nombre del Inversionista justo arriba en negrita y subrayado
    p_nom_inv = doc.add_paragraph()
    p_nom_inv.paragraph_format.space_before = Pt(6)
    p_nom_inv.add_run("Nombre Asegurado/Beneficiario ")
    r_inv = p_nom_inv.add_run(datos_inv["beneficiario"])
    r_inv.bold = True
    r_inv.font.all_caps = True

    # Tabla 2x6 para los datos bancarios alineados
    t_bank = doc.add_table(rows=6, cols=2)
    quitar_bordes_tabla(t_bank)

    t_bank.autofit = False
    for row in t_bank.rows:
        row.cells[0].width = Cm(6)
        row.cells[1].width = Cm(8)

    datos_bancarios = [
        ("RUT N°", datos_inv["rut_beneficiario"]),
        ("Banco", datos_inv["banco"]),
        ("N° Cta. Cte.", datos_inv["cta_cte"]),
        ("Correo electrónico", datos_inv["email"]),
        ("Teléfono", datos_inv["telefono"]),
        ("Observaciones:", ""),
    ]

    for idx, (label, val) in enumerate(datos_bancarios):
        p_lbl = t_bank.cell(idx, 0).paragraphs[0]
        p_lbl.add_run(label)
        p_val = t_bank.cell(idx, 1).paragraphs[0]
        p_val.add_run(val).bold = True

    # 7. Cierre y Tabla de Firmas Dinámica
    p_cierre = doc.add_paragraph("Sin otro particular, saluda cordialmente,")
    p_cierre.paragraph_format.space_before = Pt(6)
    p_cierre.paragraph_format.space_after = Pt(60)  # Espacio para firmar

    firmantes = datos_inv["firmantes"]
    num_firmantes = len(firmantes)

    # Calcular columnas: 1 si es un solo firmante, o 2 * N - 1 para incluir separadores
    num_cols = 1 if num_firmantes <= 1 else (2 * num_firmantes - 1)

    t_firmas = doc.add_table(rows=2, cols=num_cols)
    quitar_bordes_tabla(t_firmas)
    t_firmas.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_firmas.autofit = False

    # Insertar datos de firmantes en las columnas principales (0, 2, 4...)
    for idx, firmante in enumerate(firmantes):
        col_idx = idx * 2
        
        cell_nombre = t_firmas.cell(0, col_idx)
        cell_rut = t_firmas.cell(1, col_idx)
        
        # Borde superior de 1 pto solo sobre la celda del firmante
        agregar_borde_superior_celda(cell_nombre, sz="8")
        
        p_nom = cell_nombre.paragraphs[0]
        p_nom.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_nom.add_run(firmante["nombre"]).bold = True
        
        p_rut = cell_rut.paragraphs[0]
        p_rut.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_rut.add_run(firmante["rut"]).bold = True

    # Asignar anchos: 6.5 cm para firmantes y 2.0 cm para columnas separadoras
    ancho_firmante = Cm(6.5)
    ancho_separador = Cm(2.0)

    for row in t_firmas.rows:
        for c_idx in range(num_cols):
            if c_idx % 2 == 0:
                row.cells[c_idx].width = ancho_firmante
            else:
                row.cells[c_idx].width = ancho_separador

    # Nombre del Asegurado centrado debajo de la tabla
    p_pie_asegurado = doc.add_paragraph()
    p_pie_asegurado.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if inversionista_clave == "AVLA SERVICIOS / DELPHOS":
        r_asegurado = p_pie_asegurado.add_run(
            f"pp. {datos_inv["beneficiario"]} y este a su vez\n"
            "pp. EXILIO INVESTMENTS S.A.R.L., actuando únicamente en relación "
            "con su compartimento Lagorta"
            )
    else:
        p_pie_asegurado.paragraph_format.space_before = Pt(12)
        r_asegurado = p_pie_asegurado.add_run(datos_inv["asegurado"])
        r_asegurado.font.all_caps = True
    r_asegurado.bold = True

    # Guardar en memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer