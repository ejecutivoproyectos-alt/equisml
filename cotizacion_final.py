import re
from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from utils.calendario_mexico import calcular_fecha_larga_habil_mexico
from utils.moneda import (
    formatear_moneda,
    convertir_numero_a_letras_mxn
)
from utils.word_table import (
    agregar_parrafo_con_negritas,
    poner_bordes_tabla,
    colorear_celda,
    configurar_texto_celda
)

EMPRESAS = {
    "WAVELENS S.A. DE C.V.": {
        "slug": "wavelens",
        "nombre": "WAVELENS S.A. DE C.V.",
    },
    "EMPRESA DEMO S.A. DE C.V.": {
        "slug": "empresa_demo",
        "nombre": "EMPRESA DEMO S.A. DE C.V.",
    },
}

PALETAS = {
    "Naranja y rosa pastel": {
        "principal": (216, 75, 32),
        "secundario": (255, 205, 197),
        "texto": (0, 0, 0),
    },
    "Azul marino y gris": {
        "principal": (31, 56, 100),
        "secundario": (217, 217, 217),
        "texto": (0, 0, 0),
    },
    "Vino y rosa pastel": {
        "principal": (112, 48, 48),
        "secundario": (255, 205, 197),
        "texto": (0, 0, 0),
    },
}


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def extraer_entre_parentesis(texto):
    texto = limpiar_texto(texto)
    encontrados = re.findall(r"\((.*?)\)", texto)
    if encontrados:
        return encontrados[-1].strip()
    return texto


def extraer_datos_cotizacion_excel(archivo_excel):
    df = pd.read_excel(archivo_excel, header=None)

    fila_encabezados = None
    for i in range(len(df)):
        fila = [limpiar_texto(x).lower() for x in df.iloc[i].tolist()]
        if "concepto" in fila and "total factura ($)" in fila and "fecha" in fila:
            fila_encabezados = i
            break

    if fila_encabezados is None:
        raise ValueError(
            "No se encontraron los encabezados 'Concepto', "
            "'Total Factura ($)' y 'Fecha' en el Excel."
        )

    encabezados = [limpiar_texto(x) for x in df.iloc[fila_encabezados].tolist()]

    col_concepto = encabezados.index("Concepto")
    col_monto = encabezados.index("Total Factura ($)")
    col_fecha = encabezados.index("Fecha")

    registros = []
    ultimo_concepto = ""

    for i in range(fila_encabezados + 1, len(df)):
        concepto_raw = limpiar_texto(df.iloc[i, col_concepto])
        monto = df.iloc[i, col_monto]
        fecha = df.iloc[i, col_fecha]

        if concepto_raw:
            ultimo_concepto = extraer_entre_parentesis(concepto_raw)

        if pd.notna(monto):
            fecha_convertida = pd.to_datetime(fecha, errors="coerce", dayfirst=True)

            if pd.isna(fecha_convertida):
                fecha_formateada = ""
            else:
                fecha_formateada = fecha_convertida.strftime("%d/%m/%Y")

            registros.append({
                "Concepto": ultimo_concepto,
                "Monto": float(monto),
                "Fecha": fecha_formateada,
            })

    if not registros:
        raise ValueError("No se encontraron registros con monto en el Excel.")

    return registros


def agrupar_conceptos_por_monto(registros):
    grupos = {}
    for r in registros:
        concepto = r["Concepto"]
        monto = r["Monto"]
        if concepto not in grupos:
            grupos[concepto] = []
        grupos[concepto].append(monto)
    return grupos

def obtener_fecha_mas_antigua_registros(registros):
    fechas = []

    for r in registros:
        fecha = pd.to_datetime(r.get("Fecha"), errors="coerce", dayfirst=True)

        if pd.notna(fecha):
            fechas.append(fecha)

    if not fechas:
        raise ValueError("No se encontraron fechas válidas en el Excel.")

    fecha_mas_antigua = min(fechas)

    return fecha_mas_antigua.strftime("%d/%m/%Y")


def crear_word_cotizacion(
    empresa_nombre, slug, nombre_cliente, registros, fuente, size_letra, paleta
):
    ruta_plantilla = (
        Path("membretes") / slug / f"plantilla_{slug}.docx"
    )

    if not ruta_plantilla.exists():
        raise FileNotFoundError(
            "No existe la plantilla de cotización para esta empresa."
        )

    doc = Document(str(ruta_plantilla))

    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    for _ in range(3):
        doc.add_paragraph()

    # Fecha calculada desde la factura más antigua
    fecha_mas_antigua = obtener_fecha_mas_antigua_registros(registros)

    fecha_documento = calcular_fecha_larga_habil_mexico(
        fecha_str=fecha_mas_antigua,
        dias_restar=15,
        dias_sumar=4
    )

    fecha_str = f"Mérida, Yucatán, México a {fecha_documento}"

    p_fecha = doc.add_paragraph(fecha_str)
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Título
    titulo = doc.add_paragraph(style="Heading 1")
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = titulo.add_run("COTIZACIÓN FINAL")
    run_titulo.font.name = fuente
    run_titulo.font.size = Pt(size_letra + 6)
    run_titulo.font.bold = True
    run_titulo.font.color.rgb = RGBColor(*paleta["principal"])

    # Saludo
    agregar_parrafo_con_negritas(
        doc,
        [
            ("Estimado ", False),
            (nombre_cliente, True),
        ],
        fuente,
        size_letra,
        paleta["texto"]
    )

    #Parrafo introductorio
    agregar_parrafo_con_negritas(
        doc,
        [
            ("En ", False),
            (empresa_nombre, True),
            (", agradecemos su preferencia y confianza. A continuación, presentamos la cotización del servicio:",
             False),
        ],
        fuente,
        size_letra,
        paleta["texto"]
    )

    # Agrupar conceptos
    grupos = agrupar_conceptos_por_monto(registros)

    # Crear tabla
    tabla = doc.add_table(rows=1, cols=2)
    poner_bordes_tabla(tabla)

    # Encabezados
    header_row = tabla.rows[0]
    configurar_texto_celda(
        header_row.cells[0], "CONCEPTOS", fuente, size_letra,
        True, paleta["texto"], WD_ALIGN_PARAGRAPH.LEFT,
    )
    configurar_texto_celda(
        header_row.cells[1], "MONTO", fuente, size_letra,
        True, paleta["texto"], WD_ALIGN_PARAGRAPH.RIGHT,
    )
    colorear_celda(header_row.cells[0], paleta["principal"])
    colorear_celda(header_row.cells[1], paleta["principal"])
    header_row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    header_row.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    # Filas de datos
    total_general = 0
    indice_concepto = 0

    for concepto, montos in grupos.items():
        filas_concepto = []
        aplicar_relleno = indice_concepto % 2 == 0

        for idx, monto in enumerate(montos):
            row = tabla.add_row()
            filas_concepto.append(row)

            if idx == 0:
                configurar_texto_celda(
                    row.cells[0], concepto.upper(), fuente, size_letra,
                    False, paleta["texto"], WD_ALIGN_PARAGRAPH.LEFT,
                )
            else:
                configurar_texto_celda(
                    row.cells[0], "", fuente, size_letra,
                    False, paleta["texto"], WD_ALIGN_PARAGRAPH.LEFT,
                )

            configurar_texto_celda(
                row.cells[1], formatear_moneda(monto), fuente, size_letra,
                False, paleta["texto"], WD_ALIGN_PARAGRAPH.RIGHT,
            )

            if aplicar_relleno:
                colorear_celda(row.cells[0], paleta["secundario"])
                colorear_celda(row.cells[1], paleta["secundario"])

            total_general += monto

        if len(filas_concepto) > 1:
            celda_inicio = filas_concepto[0].cells[0]
            celda_fin = filas_concepto[-1].cells[0]
            celda_inicio.merge(celda_fin)

            configurar_texto_celda(
                celda_inicio, concepto.upper(), fuente, size_letra,
                False, paleta["texto"], WD_ALIGN_PARAGRAPH.LEFT,
            )

            if aplicar_relleno:
                colorear_celda(celda_inicio, paleta["secundario"])

        indice_concepto += 1

    # Fila de total sin relleno
    total_row = tabla.add_row()

    configurar_texto_celda(
        total_row.cells[0],
        "TOTAL",
        fuente,
        size_letra,
        True,
        paleta["texto"],
        WD_ALIGN_PARAGRAPH.CENTER,
    )

    configurar_texto_celda(
        total_row.cells[1],
        formatear_moneda(total_general),
        fuente,
        size_letra,
        True,
        paleta["texto"],
        WD_ALIGN_PARAGRAPH.RIGHT,
    )

    # Fila de cantidad en letras dentro de la tabla
    letras_row = tabla.add_row()
    celda_letras = letras_row.cells[0].merge(letras_row.cells[1])

    configurar_texto_celda(
        celda_letras,
        convertir_numero_a_letras_mxn(total_general),
        fuente,
        size_letra,
        True,
        paleta["texto"],
        WD_ALIGN_PARAGRAPH.CENTER,
    )

    # Textos condicionales del servicio
    doc.add_paragraph(
        "Los precios que figuran son en moneda nacional IVA incluido. "
        "El servicio de consultoría presentado será documentado y suministrado "
        "al cliente en documento PDF AL TÉRMINO DEL MISMO."
    )
    doc.add_paragraph(
        "El trabajo será cobrado y en su caso pagado por el cliente en su totalidad "
        "cuando se entregue el 100% del servicio acordado con el cliente de conformidad "
        "con el contrato de prestación de servicio."
    )
    doc.add_paragraph(
        "Espero que esta propuesta cumpla con las expectativas que usted se ha fijado "
        "y quedo a sus órdenes para disponer de cualquier duda que pudiera surgir "
        "en torno a la presente."
    )

    # Firma
    # Espacio antes de firma
    doc.add_paragraph()

    # Atentamente centrado
    p_atte = doc.add_paragraph()
    p_atte.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_atte = p_atte.add_run("Atentamente")
    run_atte.font.name = fuente
    run_atte.font.size = Pt(size_letra)
    run_atte.font.color.rgb = RGBColor(*paleta["texto"])

    # Espacio
    doc.add_paragraph()

    # Línea firma centrada
    p_linea = doc.add_paragraph()
    p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_linea = p_linea.add_run("__________________________________")
    run_linea.font.name = fuente
    run_linea.font.size = Pt(size_letra)
    run_linea.font.color.rgb = RGBColor(*paleta["texto"])

    # Representante legal centrado
    p_rep = doc.add_paragraph()
    p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_rep = p_rep.add_run("REPRESENTANTE LEGAL")
    run_rep.font.name = fuente
    run_rep.font.size = Pt(size_letra)
    run_rep.font.color.rgb = RGBColor(*paleta["texto"])

    # Empresa centrada
    p_emp = doc.add_paragraph()
    p_emp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_emp = p_emp.add_run(empresa_nombre)
    run_emp.font.name = fuente
    run_emp.font.size = Pt(size_letra)
    run_emp.font.bold = True
    run_emp.font.color.rgb = RGBColor(*paleta["texto"])

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def mostrar_modulo_cotizacion_final():
    st.title("Cotización final")

    empresa_seleccionada = st.selectbox(
        "Selecciona la empresa prestadora del servicio",
        list(EMPRESAS.keys()),
    )
    slug = EMPRESAS[empresa_seleccionada]["slug"]
    nombre_empresa = EMPRESAS[empresa_seleccionada]["nombre"]

    nombre_cliente = st.text_input("Nombre del cliente")

    fuente = st.selectbox(
        "Selecciona la tipografía",
        ["Arial", "Calibri", "Times New Roman", "Verdana", "Segoe UI", "Tahoma"],
    )

    size_letra = st.number_input(
        "Tamaño de letra",
        min_value=8,
        max_value=16,
        value=11,
        step=1
    )

    paleta_nombre = st.selectbox(
        "Selecciona la paleta de colores",
        list(PALETAS.keys()),
    )
    paleta = PALETAS[paleta_nombre]

    archivo = st.file_uploader(
        "Sube el archivo Excel con los conceptos",
        type=["xlsx", "xls"],
    )

    if archivo is None:
        st.warning("Sube un archivo Excel para continuar.")
        return

    if not nombre_cliente:
        st.warning("Escribe el nombre del cliente.")
        return

    ruta_plantilla = (
        Path("membretes") / slug / f"plantilla_{slug}.docx"
    )
    if not ruta_plantilla.exists():
        st.error("No existe la plantilla de cotización para esta empresa.")
        return

    try:
        registros = extraer_datos_cotizacion_excel(archivo)

        st.subheader("Vista previa de conceptos y montos")
        df_preview = pd.DataFrame(registros)
        df_preview["Monto"] = df_preview["Monto"].map(formatear_moneda)
        st.dataframe(df_preview, use_container_width=True)

        word = crear_word_cotizacion(
            nombre_empresa, slug, nombre_cliente, registros,
            fuente, size_letra, paleta,
        )

        st.download_button(
            label="Descargar cotización final (Word)",
            data=word,
            file_name="cotizacion_final.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except ValueError as e:
        st.error(str(e))
    except FileNotFoundError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
