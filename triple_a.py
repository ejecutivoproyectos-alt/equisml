import os
from copy import deepcopy
from datetime import date
from io import BytesIO

import streamlit as st
from dotenv import load_dotenv
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from openai import OpenAI

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_estilo_word import EmpresaEstiloWord
from app.models.empresa_plantilla_asignacion import EmpresaPlantillaAsignacion


load_dotenv()

MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

COLOR_PRINCIPAL = (31, 58, 95)
COLOR_SECUNDARIO = (232, 238, 245)
COLOR_GRIS = (242, 244, 247)
COLOR_BORDE = "9AA4B2"
COLOR_TEXTO = (0, 0, 0)
FUENTE_BASE = "Calibri"
TAMANIO_BASE = 11


MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


def fecha_larga_es(fecha):
    return f"{fecha.day} de {MESES_ES[fecha.month]} del {fecha.year}"


def convertir_color_bd_a_rgb(color):
    if not color:
        return COLOR_PRINCIPAL

    color = color.strip().lstrip("#")

    if len(color) != 6:
        return COLOR_PRINCIPAL

    try:
        return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return COLOR_PRINCIPAL


def obtener_empresas_triple_a():
    db = SessionLocal()

    try:
        empresas = (
            db.query(Empresa)
            .filter(Empresa.tipo_empresa == "TRIPLE_AAA")
            .order_by(Empresa.nombre.asc())
            .all()
        )

        return [
            {
                "id": empresa.id,
                "nombre": empresa.nombre,
                "razon_social": empresa.razon_social or empresa.nombre,
                "color_primario": empresa.color_primario,
                "color_secundario": empresa.color_secundario,
                "color_acento": empresa.color_acento,
            }
            for empresa in empresas
        ]
    finally:
        db.close()


def obtener_estilos_word_empresa(empresa_id):
    db = SessionLocal()

    try:
        estilos = (
            db.query(EmpresaEstiloWord)
            .filter(EmpresaEstiloWord.empresa_id == empresa_id)
            .all()
        )

        return {
            estilo.clave_estilo: {
                "tipografia": estilo.tipografia,
                "tamanio_letra": estilo.tamanio_letra,
                "negrita": estilo.negrita,
                "cursiva": estilo.cursiva,
                "alineacion": estilo.alineacion,
            }
            for estilo in estilos
        }
    finally:
        db.close()


def obtener_membrete_empresa(empresa_id):
    db = SessionLocal()

    try:
        asignacion = (
            db.query(EmpresaPlantillaAsignacion)
            .filter(EmpresaPlantillaAsignacion.empresa_externa_id == empresa_id)
            .first()
        )

        if not asignacion or not asignacion.membrete_path:
            return None

        return asignacion.membrete_path
    finally:
        db.close()


def obtener_cliente_openai():
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("No se encontro OPENAI_API_KEY en el archivo .env.")

    return OpenAI(api_key=api_key)


def llamar_openai_texto(prompt, modelo="gpt-4.1-mini"):
    client = obtener_cliente_openai()
    respuesta = client.responses.create(
        model=modelo,
        input=prompt,
    )
    return respuesta.output_text.strip()


def generar_textos_ia(nombre_empresa, nombre_cliente, nombre_servicio, periodo, actividades):
    actividades_texto = "\n".join(f"- {actividad}" for actividad in actividades)

    prompt = f"""
Actua como consultor empresarial senior en Mexico.

Redacta textos ejecutivos para un documento de seguimiento de servicios especializados.

EMPRESA PRESTADORA:
{nombre_empresa}

CLIENTE:
{nombre_cliente}

SERVICIO:
{nombre_servicio}

PERIODO:
{periodo}

ACTIVIDADES BASE:
{actividades_texto}

Entrega exclusivamente un JSON valido con estas llaves:
{{
  "introduccion": "un parrafo formal de 90 a 130 palabras",
  "objetivo": "un parrafo formal de 60 a 90 palabras",
  "alcance": "un parrafo formal de 70 a 100 palabras",
  "cierre": "un parrafo formal de 60 a 90 palabras"
}}

No incluyas markdown, explicaciones ni texto fuera del JSON.
"""

    texto = llamar_openai_texto(prompt)

    try:
        import json

        data = json.loads(texto)
        return {
            "introduccion": data.get("introduccion", "").strip(),
            "objetivo": data.get("objetivo", "").strip(),
            "alcance": data.get("alcance", "").strip(),
            "cierre": data.get("cierre", "").strip(),
        }
    except Exception:
        return {
            "introduccion": texto,
            "objetivo": "",
            "alcance": "",
            "cierre": "",
        }


def textos_predeterminados(nombre_empresa, nombre_cliente, nombre_servicio, periodo):
    return {
        "introduccion": (
            f"El presente documento integra el seguimiento de las actividades realizadas por "
            f"{nombre_empresa.upper()} en favor de {nombre_cliente.upper()}, correspondientes "
            f"al servicio de {nombre_servicio.upper()}. Su finalidad es dejar constancia "
            f"ordenada del avance, desarrollo y cumplimiento de las acciones ejecutadas "
            f"durante el periodo {periodo}."
        ),
        "objetivo": (
            "Documentar de forma clara y verificable las actividades desarrolladas, los "
            "entregables generados y el estado general del servicio, con el proposito de "
            "facilitar el control administrativo, la trazabilidad operativa y la revision "
            "del cumplimiento pactado entre las partes."
        ),
        "alcance": (
            "El alcance comprende la revision, organizacion y registro de las actividades "
            "asociadas al servicio, asi como la integracion de evidencias y comentarios "
            "relevantes para el seguimiento interno del proyecto."
        ),
        "cierre": (
            f"Con base en la informacion registrada, {nombre_empresa.upper()} mantiene el "
            f"seguimiento profesional del servicio prestado a {nombre_cliente.upper()}, "
            "procurando que cada actividad cuente con soporte documental y continuidad "
            "administrativa."
        ),
    }


def limpiar_lista(texto):
    return [linea.strip(" -\t") for linea in texto.splitlines() if linea.strip(" -\t")]


def limpiar_cuerpo_documento(doc):
    body = doc._body._element
    sect_pr = body.sectPr

    for elemento in list(body):
        body.remove(elemento)

    if sect_pr is not None:
        body.append(deepcopy(sect_pr))


def crear_documento_base(membrete_docx=None, membrete_path=None):
    if membrete_docx is not None:
        doc = Document(membrete_docx)
        limpiar_cuerpo_documento(doc)
    elif membrete_path and os.path.exists(membrete_path):
        doc = Document(membrete_path)
        limpiar_cuerpo_documento(doc)
    else:
        doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.49)
    section.footer_distance = Inches(0.49)

    return doc


def configurar_estilos(doc):
    styles = doc.styles

    try:
        normal = styles["Normal"]
        normal.font.name = FUENTE_BASE
        normal.font.size = Pt(TAMANIO_BASE)
        normal.font.color.rgb = RGBColor(*COLOR_TEXTO)
        normal.paragraph_format.space_after = Pt(6)
        normal.paragraph_format.line_spacing = 1.1
    except KeyError:
        pass

    for style_name, size, color in [
        ("Heading 1", 16, COLOR_PRINCIPAL),
        ("Heading 2", 13, COLOR_PRINCIPAL),
        ("Heading 3", 12, COLOR_PRINCIPAL),
    ]:
        try:
            style = styles[style_name]
            style.font.name = FUENTE_BASE
            style.font.size = Pt(size)
            style.font.bold = True
            style.font.color.rgb = RGBColor(*color)
            style.paragraph_format.space_before = Pt(10)
            style.paragraph_format.space_after = Pt(6)
        except KeyError:
            continue


def agregar_membrete_simple(doc, nombre_empresa, nombre_cliente):
    section = doc.sections[0]

    header = section.header
    if not header.paragraphs:
        header.add_paragraph()

    p = header.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(nombre_empresa.upper())
    run.font.name = FUENTE_BASE
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor(*COLOR_PRINCIPAL)

    p2 = header.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run("Documento de seguimiento de servicios especializados")
    run2.font.name = FUENTE_BASE
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(90, 90, 90)

    footer = section.footer
    p_footer = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p_footer.text = ""
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_footer = p_footer.add_run(f"{nombre_empresa.upper()} | {nombre_cliente.upper()}")
    run_footer.font.name = FUENTE_BASE
    run_footer.font.size = Pt(8)
    run_footer.font.color.rgb = RGBColor(90, 90, 90)


def aplicar_bordes_tabla(tabla, color=COLOR_BORDE):
    tbl = tabla._tbl
    tbl_pr = tbl.tblPr

    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)

    borders = OxmlElement("w:tblBorders")

    for nombre in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{nombre}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), color)
        border.set(qn("w:space"), "0")
        borders.append(border)

    existente = tbl_pr.find(qn("w:tblBorders"))
    if existente is not None:
        tbl_pr.remove(existente)

    tbl_pr.append(borders)


def fijar_ancho_tabla(tabla, anchos_cm):
    for fila in tabla.rows:
        for idx, ancho in enumerate(anchos_cm):
            if idx >= len(fila.cells):
                continue
            celda = fila.cells[idx]
            celda.width = Cm(ancho)
            tc_pr = celda._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(int(Cm(ancho).twips)))
            tc_w.set(qn("w:type"), "dxa")


def colorear_celda(celda, color_rgb):
    tc_pr = celda._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), f"{color_rgb[0]:02X}{color_rgb[1]:02X}{color_rgb[2]:02X}")
    shading.set(qn("w:val"), "clear")
    tc_pr.append(shading)


def configurar_margenes_celda(celda, top=80, bottom=80, start=120, end=120):
    tc_pr = celda._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)

    for nombre, valor in {
        "top": top,
        "bottom": bottom,
        "start": start,
        "end": end,
    }.items():
        margen = tc_mar.find(qn(f"w:{nombre}"))
        if margen is None:
            margen = OxmlElement(f"w:{nombre}")
            tc_mar.append(margen)
        margen.set(qn("w:w"), str(valor))
        margen.set(qn("w:type"), "dxa")


def escribir_celda(celda, texto, bold=False, color=COLOR_TEXTO, align=WD_ALIGN_PARAGRAPH.LEFT):
    celda.text = ""
    celda.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    configurar_margenes_celda(celda)

    p = celda.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(str(texto))
    run.font.name = FUENTE_BASE
    run.font.size = Pt(9)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)


def aplicar_alineacion(parrafo, alineacion):
    alineaciones = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    parrafo.alignment = alineaciones.get(
        (alineacion or "center").lower(),
        WD_ALIGN_PARAGRAPH.CENTER,
    )


def agregar_titulo(doc, titulo, estilo_titulo=None, color_principal=COLOR_PRINCIPAL):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)

    if estilo_titulo:
        aplicar_alineacion(p, estilo_titulo.get("alineacion") or "center")
        fuente = estilo_titulo.get("tipografia") or FUENTE_BASE
        tamanio = estilo_titulo.get("tamanio_letra") or 18
        negrita = estilo_titulo.get("negrita")
        cursiva = estilo_titulo.get("cursiva")
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fuente = FUENTE_BASE
        tamanio = 18
        negrita = True
        cursiva = False

    run = p.add_run(titulo.upper())
    run.font.name = fuente
    run.font.size = Pt(tamanio)
    run.font.bold = negrita
    run.font.italic = cursiva
    run.font.color.rgb = RGBColor(*color_principal)


def agregar_parrafo(doc, texto, alineacion=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = alineacion
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(texto)
    run.font.name = FUENTE_BASE
    run.font.size = Pt(TAMANIO_BASE)
    run.font.color.rgb = RGBColor(*COLOR_TEXTO)
    return p


def agregar_encabezado(doc, texto, nivel=1):
    tamanios = {
        1: 16,
        2: 13,
        3: 12,
    }

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)

    run = p.add_run(texto)
    run.font.name = FUENTE_BASE
    run.font.size = Pt(tamanios.get(nivel, 12))
    run.font.bold = True
    run.font.color.rgb = RGBColor(*COLOR_PRINCIPAL)

    return p


def agregar_tabla_datos_generales(doc, datos):
    agregar_encabezado(doc, "Datos generales", nivel=1)

    tabla = doc.add_table(rows=0, cols=2)
    tabla.autofit = False

    for etiqueta, valor in datos:
        fila = tabla.add_row()
        escribir_celda(fila.cells[0], etiqueta, bold=True, color=COLOR_PRINCIPAL)
        escribir_celda(fila.cells[1], valor)
        colorear_celda(fila.cells[0], COLOR_GRIS)

    aplicar_bordes_tabla(tabla)
    fijar_ancho_tabla(tabla, [5.0, 11.0])
    doc.add_paragraph()


def agregar_tabla_actividades(doc, actividades):
    agregar_encabezado(doc, "Seguimiento de actividades", nivel=1)

    tabla = doc.add_table(rows=1, cols=5)
    tabla.autofit = False
    encabezados = ["No.", "Actividad", "Estatus", "Evidencia", "Comentarios"]

    for idx, encabezado in enumerate(encabezados):
        escribir_celda(
            tabla.rows[0].cells[idx],
            encabezado,
            bold=True,
            color=(255, 255, 255),
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        colorear_celda(tabla.rows[0].cells[idx], COLOR_PRINCIPAL)

    for idx, actividad in enumerate(actividades, start=1):
        fila = tabla.add_row()
        escribir_celda(fila.cells[0], idx, align=WD_ALIGN_PARAGRAPH.CENTER)
        escribir_celda(fila.cells[1], actividad)
        escribir_celda(fila.cells[2], "Realizada", align=WD_ALIGN_PARAGRAPH.CENTER)
        escribir_celda(fila.cells[3], "Registro documental", align=WD_ALIGN_PARAGRAPH.CENTER)
        escribir_celda(fila.cells[4], "Sin observaciones")

    aplicar_bordes_tabla(tabla)
    fijar_ancho_tabla(tabla, [1.2, 6.5, 2.4, 3.4, 3.0])
    doc.add_paragraph()


def agregar_tabla_entregables(doc, entregables):
    agregar_encabezado(doc, "Entregables", nivel=1)

    tabla = doc.add_table(rows=1, cols=4)
    tabla.autofit = False
    encabezados = ["No.", "Entregable", "Responsable", "Estado"]

    for idx, encabezado in enumerate(encabezados):
        escribir_celda(
            tabla.rows[0].cells[idx],
            encabezado,
            bold=True,
            color=(255, 255, 255),
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        colorear_celda(tabla.rows[0].cells[idx], COLOR_PRINCIPAL)

    for idx, entregable in enumerate(entregables, start=1):
        fila = tabla.add_row()
        escribir_celda(fila.cells[0], idx, align=WD_ALIGN_PARAGRAPH.CENTER)
        escribir_celda(fila.cells[1], entregable)
        escribir_celda(fila.cells[2], "Empresa prestadora", align=WD_ALIGN_PARAGRAPH.CENTER)
        escribir_celda(fila.cells[3], "Integrado", align=WD_ALIGN_PARAGRAPH.CENTER)

    aplicar_bordes_tabla(tabla)
    fijar_ancho_tabla(tabla, [1.2, 8.0, 4.0, 3.0])
    doc.add_paragraph()


def agregar_firmas(doc, nombre_empresa, nombre_cliente):
    agregar_encabezado(doc, "Firmas de conformidad", nivel=1)
    doc.add_paragraph()

    tabla = doc.add_table(rows=2, cols=2)
    tabla.autofit = False

    escribir_celda(tabla.rows[0].cells[0], "______________________________", align=WD_ALIGN_PARAGRAPH.CENTER)
    escribir_celda(tabla.rows[0].cells[1], "______________________________", align=WD_ALIGN_PARAGRAPH.CENTER)
    escribir_celda(tabla.rows[1].cells[0], nombre_empresa.upper(), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    escribir_celda(tabla.rows[1].cells[1], nombre_cliente.upper(), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    fijar_ancho_tabla(tabla, [8.0, 8.0])


def crear_word_seguimiento(
    nombre_empresa,
    nombre_cliente,
    nombre_servicio,
    periodo,
    fecha_documento,
    responsable,
    actividades,
    entregables,
    textos,
    membrete_docx=None,
    membrete_path=None,
    estilos_bd=None,
    paleta=None,
):
    paleta = paleta or {"principal": COLOR_PRINCIPAL}
    estilos_bd = estilos_bd or {}

    doc = crear_documento_base(membrete_docx, membrete_path)
    configurar_estilos(doc)

    if membrete_docx is None and not membrete_path:
        agregar_membrete_simple(doc, nombre_empresa, nombre_cliente)

    agregar_titulo(
        doc,
        "Seguimiento bimestral",
        estilo_titulo=estilos_bd.get("titulo_1"),
        color_principal=paleta["principal"],
    )

    agregar_tabla_datos_generales(
        doc,
        [
            ("Empresa prestadora", nombre_empresa.upper()),
            ("Cliente", nombre_cliente.upper()),
            ("Servicio", nombre_servicio.upper()),
            ("Periodo", periodo),
            ("Fecha del documento", fecha_documento),
            ("Responsable", responsable),
        ],
    )

    agregar_encabezado(doc, "Introduccion", nivel=1)
    agregar_parrafo(doc, textos["introduccion"])

    agregar_encabezado(doc, "Objetivo", nivel=1)
    agregar_parrafo(doc, textos["objetivo"])

    agregar_encabezado(doc, "Alcance del seguimiento", nivel=1)
    agregar_parrafo(doc, textos["alcance"])

    agregar_tabla_actividades(doc, actividades)
    agregar_tabla_entregables(doc, entregables)

    agregar_encabezado(doc, "Cierre", nivel=1)
    agregar_parrafo(doc, textos["cierre"])

    agregar_firmas(doc, nombre_empresa, nombre_cliente)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def mostrar_modulo_triple_a():
    st.title("Generador de seguimiento AAA")

    st.write(
        "Genera un documento Word de seguimiento con membrete, textos estaticos, "
        "tablas codificadas en Python y textos opcionales generados con IA."
    )

    empresas = obtener_empresas_triple_a()

    if not empresas:
        st.error("No hay empresas registradas en la BD.")
        return

    empresas_por_nombre = {
        empresa["nombre"]: empresa
        for empresa in empresas
    }

    with st.form("form_triple_a"):
        col1, col2 = st.columns(2)

        with col1:
            empresa_seleccionada = st.selectbox(
                "Empresa prestadora",
                list(empresas_por_nombre.keys()),
            )
            nombre_servicio = st.text_input(
                "Nombre del servicio",
                value="Seguimiento de servicios especializados",
            )
            fecha_doc = st.date_input("Fecha del documento", value=date.today())

        with col2:
            nombre_cliente = st.text_input(
                "Cliente",
                value="HABITEC ECO SA DE C.V",
            )
            periodo = st.text_input(
                "Periodo",
                value="enero - febrero 2026",
            )
            responsable = st.text_input("Responsable", value="Area de seguimiento")

        membrete_docx = st.file_uploader(
            "Membrete Word opcional (.docx)",
            type=["docx"],
            help=(
                "Si lo subes, se usa como base del documento. "
                "No se modifica el archivo original."
            ),
        )

        actividades_texto = st.text_area(
            "Actividades (una por linea)",
            value=(
                "Revision documental del servicio\n"
                "Integracion de informacion de seguimiento\n"
                "Validacion de actividades realizadas\n"
                "Elaboracion de tablas de control\n"
                "Cierre administrativo del periodo"
            ),
            height=150,
        )

        entregables_texto = st.text_area(
            "Entregables (uno por linea)",
            value=(
                "Documento de seguimiento del periodo\n"
                "Tabla de actividades realizadas\n"
                "Resumen de entregables y evidencias"
            ),
            height=110,
        )

        usar_ia = st.checkbox(
            "Generar introduccion, objetivo, alcance y cierre con IA",
            value=False,
        )

        generar = st.form_submit_button("Generar documento")

    if not generar:
        return

    empresa_data = empresas_por_nombre[empresa_seleccionada]
    nombre_empresa = empresa_data["razon_social"]
    estilos_bd = obtener_estilos_word_empresa(empresa_data["id"])
    membrete_path = obtener_membrete_empresa(empresa_data["id"])
    paleta = {
        "principal": convertir_color_bd_a_rgb(empresa_data["color_primario"]),
        "secundario": convertir_color_bd_a_rgb(empresa_data["color_secundario"]),
        "acento": convertir_color_bd_a_rgb(empresa_data["color_acento"]),
        "texto": COLOR_TEXTO,
    }

    if not nombre_cliente.strip():
        st.warning("Escribe el cliente.")
        return

    actividades = limpiar_lista(actividades_texto)
    entregables = limpiar_lista(entregables_texto)

    if not actividades:
        st.warning("Agrega al menos una actividad.")
        return

    if not entregables:
        st.warning("Agrega al menos un entregable.")
        return

    try:
        if usar_ia:
            with st.spinner("Generando textos con IA..."):
                textos = generar_textos_ia(
                    nombre_empresa=nombre_empresa,
                    nombre_cliente=nombre_cliente,
                    nombre_servicio=nombre_servicio,
                    periodo=periodo,
                    actividades=actividades,
                )
        else:
            textos = textos_predeterminados(
                nombre_empresa=nombre_empresa,
                nombre_cliente=nombre_cliente,
                nombre_servicio=nombre_servicio,
                periodo=periodo,
            )

        word = crear_word_seguimiento(
            nombre_empresa=nombre_empresa,
            nombre_cliente=nombre_cliente,
            nombre_servicio=nombre_servicio,
            periodo=periodo,
            fecha_documento=fecha_larga_es(fecha_doc),
            responsable=responsable,
            actividades=actividades,
            entregables=entregables,
            textos=textos,
            membrete_docx=membrete_docx,
            membrete_path=membrete_path,
            estilos_bd=estilos_bd,
            paleta=paleta,
        )

        st.success("Documento generado correctamente.")
        st.download_button(
            label="Descargar seguimiento Word",
            data=word,
            file_name="seguimiento_servicios_especializados.docx",
            mime=MIME_DOCX,
            key="descargar_triple_a",
            on_click="ignore",
        )

    except ValueError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Ocurrio un error al generar el documento: {e}")


if __name__ == "__main__":
    mostrar_modulo_triple_a()
