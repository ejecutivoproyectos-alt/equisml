import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from utils.calendario_mexico import calcular_fecha_larga_habil_mexico
from utils.moneda import formatear_moneda, convertir_numero_a_letras_mxn
from utils.word_table import (
    poner_bordes_tabla,
    colorear_celda,
    configurar_texto_celda
)

# Ajusta estos imports según tu proyecto
from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord
from app.models.empresa_estilo_word import EmpresaEstiloWord

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

def obtener_estilos_word_por_plantilla(plantilla_id):
    db = SessionLocal()

    try:
        estilos = (
            db.query(EmpresaEstiloWord)
            .filter(EmpresaEstiloWord.plantilla_id == plantilla_id)
            .all()
        )

        return {
            estilo.clave_estilo: {
                "tipografia": estilo.tipografia,
                "tamanio_letra": estilo.tamanio_letra,
                "color_letra": estilo.color_letra,
                "color_fondo": estilo.color_fondo,
                "negrita": estilo.negrita,
                "cursiva": estilo.cursiva,
                "alineacion": estilo.alineacion,
            }
            for estilo in estilos
        }

    finally:
        db.close()

def convertir_color_bd_a_rgb(color):
    """
    Convierte colores guardados en BD como:
    '#D84B20'
    'D84B20'
    '216,75,32'
    a tupla RGB: (216, 75, 32)
    """
    if not color:
        return (0, 0, 0)

    color = str(color).strip()

    if "," in color:
        partes = color.split(",")
        return tuple(int(p.strip()) for p in partes)

    color = color.replace("#", "")

    if len(color) != 6:
        raise ValueError(f"Color inválido en BD: {color}")

    return tuple(int(color[i:i + 2], 16) for i in (0, 2, 4))


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

            fecha_formateada = ""
            if pd.notna(fecha_convertida):
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


def obtener_empresas_con_plantilla():
    db = SessionLocal()

    try:
        resultados = (
            db.query(Empresa, EmpresaPlantillaWord)
            .join(EmpresaPlantillaWord, EmpresaPlantillaWord.empresa_id == Empresa.id)
            .filter(EmpresaPlantillaWord.membrete_path.isnot(None))
            .all()
        )

        empresas = {}

        for empresa, plantilla in resultados:
            empresas[empresa.nombre] = {
                "empresa_id": empresa.id,
                "plantilla_id": plantilla.id,
                "nombre": empresa.nombre,
                "membrete_path": plantilla.membrete_path,
                "color_texto_base": plantilla.color_texto_base,
                "color_primario": plantilla.color_primario,
                "color_secundario": plantilla.color_secundario,
                "tipografia_base": plantilla.tipografia_base,
                "tamanio_base": plantilla.tamanio_base,
            }

        return empresas

    finally:
        db.close()


def reemplazar_texto_en_parrafo(parrafo, reemplazos):
    texto_original = parrafo.text

    texto_nuevo = texto_original

    for clave, valor in reemplazos.items():
        texto_nuevo = texto_nuevo.replace(clave, str(valor))

    if texto_nuevo == texto_original:
        return

    # borrar runs existentes
    for run in parrafo.runs:
        run.text = ""

    # crear nuevo run limpio
    nuevo_run = parrafo.runs[0] if parrafo.runs else parrafo.add_run()

    nuevo_run.text = texto_nuevo


def reemplazar_parametros_documento(doc, reemplazos):
    for parrafo in doc.paragraphs:
        reemplazar_texto_en_parrafo(parrafo, reemplazos)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    reemplazar_texto_en_parrafo(parrafo, reemplazos)


def aplicar_negrita_a_texto(doc, texto_objetivo, estilo_base=None):
    if not texto_objetivo:
        return

    def aplicar_formato_run(run, negrita=False):
        if estilo_base:
            run.font.name = estilo_base["tipografia"]
            run.font.size = Pt(int(estilo_base["tamanio_letra"]))
            run.font.color.rgb = RGBColor(
                *convertir_color_bd_a_rgb(estilo_base["color_letra"])
            )
            run.font.italic = bool(estilo_base["cursiva"])

        run.font.bold = negrita

    def procesar_parrafo(parrafo):
        if texto_objetivo not in parrafo.text:
            return

        texto_completo = parrafo.text

        for run in parrafo.runs:
            run.text = ""

        partes = texto_completo.split(texto_objetivo)

        for i, parte in enumerate(partes):
            if parte:
                run_normal = parrafo.add_run(parte)
                aplicar_formato_run(run_normal, negrita=False)

            if i < len(partes) - 1:
                run_negrita = parrafo.add_run(texto_objetivo)
                aplicar_formato_run(run_negrita, negrita=True)

    for parrafo in doc.paragraphs:
        procesar_parrafo(parrafo)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    procesar_parrafo(parrafo)


def insertar_tabla_despues_de_parrafo(parrafo, tabla):
    parrafo._p.addnext(tabla._tbl)


def eliminar_parrafo(parrafo):
    elemento = parrafo._element
    elemento.getparent().remove(elemento)
    parrafo._p = parrafo._element = None


def buscar_parrafo_con_texto(doc, texto_buscado):
    for parrafo in doc.paragraphs:
        if texto_buscado in parrafo.text:
            return parrafo

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if texto_buscado in parrafo.text:
                        return parrafo

    return None


def crear_tabla_cotizacion(doc, registros, fuente, size_letra, paleta):
    grupos = agrupar_conceptos_por_monto(registros)

    tabla = doc.add_table(rows=1, cols=2)
    poner_bordes_tabla(tabla)

    header_row = tabla.rows[0]

    configurar_texto_celda(
        header_row.cells[0],
        "CONCEPTOS",
        fuente,
        size_letra,
        True,
        paleta["texto"],
        WD_ALIGN_PARAGRAPH.LEFT,
    )

    configurar_texto_celda(
        header_row.cells[1],
        "MONTO",
        fuente,
        size_letra,
        True,
        paleta["texto"],
        WD_ALIGN_PARAGRAPH.RIGHT,
    )

    colorear_celda(header_row.cells[0], paleta["principal"])
    colorear_celda(header_row.cells[1], paleta["principal"])

    header_row.cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    header_row.cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

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
                    row.cells[0],
                    concepto.upper(),
                    fuente,
                    size_letra,
                    False,
                    paleta["texto"],
                    WD_ALIGN_PARAGRAPH.LEFT,
                )
            else:
                configurar_texto_celda(
                    row.cells[0],
                    "",
                    fuente,
                    size_letra,
                    False,
                    paleta["texto"],
                    WD_ALIGN_PARAGRAPH.LEFT,
                )

            configurar_texto_celda(
                row.cells[1],
                formatear_moneda(monto),
                fuente,
                size_letra,
                False,
                paleta["texto"],
                WD_ALIGN_PARAGRAPH.RIGHT,
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
                celda_inicio,
                concepto.upper(),
                fuente,
                size_letra,
                False,
                paleta["texto"],
                WD_ALIGN_PARAGRAPH.LEFT,
            )

            if aplicar_relleno:
                colorear_celda(celda_inicio, paleta["secundario"])

        indice_concepto += 1

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

    return tabla, total_general


def aplicar_estilo_base_a_todo(doc, estilo_base):
    fuente = estilo_base["tipografia"]
    size_letra = int(estilo_base["tamanio_letra"])
    color_texto = convertir_color_bd_a_rgb(estilo_base["color_letra"])

    for parrafo in doc.paragraphs:
        for run in parrafo.runs:
            run.font.name = fuente
            run.font.size = Pt(size_letra)
            run.font.bold = bool(estilo_base["negrita"])
            run.font.italic = bool(estilo_base["cursiva"])
            run.font.color.rgb = RGBColor(*color_texto)

        if estilo_base["alineacion"]:
            aplicar_alineacion_parrafo(parrafo, estilo_base["alineacion"])

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        run.font.name = fuente
                        run.font.size = Pt(size_letra)
                        run.font.bold = bool(estilo_base["negrita"])
                        run.font.italic = bool(estilo_base["cursiva"])
                        run.font.color.rgb = RGBColor(*color_texto)

                    if estilo_base["alineacion"]:
                        aplicar_alineacion_parrafo(parrafo, estilo_base["alineacion"])

def aplicar_estilo_a_texto(doc, texto_objetivo, estilo):
    if not texto_objetivo:
        return

    fuente = estilo["tipografia"]
    size_letra = int(estilo["tamanio_letra"])
    color_texto = convertir_color_bd_a_rgb(estilo["color_letra"])

    for parrafo in doc.paragraphs:
        if texto_objetivo in parrafo.text:
            for run in parrafo.runs:
                run.font.name = fuente
                run.font.size = Pt(size_letra)
                run.font.bold = bool(estilo["negrita"])
                run.font.italic = bool(estilo["cursiva"])
                run.font.color.rgb = RGBColor(*color_texto)

                if estilo["color_fondo"]:
                    colorear_fondo_run(run, estilo["color_fondo"])

            if estilo["alineacion"]:
                aplicar_alineacion_parrafo(parrafo, estilo["alineacion"])

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if texto_objetivo in parrafo.text:
                        for run in parrafo.runs:
                            run.font.name = fuente
                            run.font.size = Pt(size_letra)
                            run.font.bold = bool(estilo["negrita"])
                            run.font.italic = bool(estilo["cursiva"])
                            run.font.color.rgb = RGBColor(*color_texto)

                            if estilo["color_fondo"]:
                                colorear_fondo_run(run, estilo["color_fondo"])

                        if estilo["alineacion"]:
                            aplicar_alineacion_parrafo(parrafo, estilo["alineacion"])


def crear_word_cotizacion(
    empresa_nombre,
    ruta_plantilla_word,
    nombre_cliente,
    nombre_programa,
    registros,
    fuente,
    size_letra,
    paleta,
    estilos_bd
):
    ruta_plantilla = Path(ruta_plantilla_word) / "3.COTIZACION-FINAL.docx"

    if not ruta_plantilla.exists():
        raise FileNotFoundError(
            f"No existe la plantilla Word en la ruta: {ruta_plantilla}"
        )

    doc = Document(str(ruta_plantilla))

    fecha_mas_antigua = obtener_fecha_mas_antigua_registros(registros)

    fecha_documento = calcular_fecha_larga_habil_mexico(
        fecha_str=fecha_mas_antigua,
        dias_restar=15,
        dias_sumar=4
    )

    fecha_str = f"Mérida, Yucatán, México a {fecha_documento}"

    tabla, total_general = crear_tabla_cotizacion(
        doc,
        registros,
        fuente,
        size_letra,
        paleta
    )

    parrafo_tabla = buscar_parrafo_con_texto(doc, "{{TABLA_COTIZACION}}")

    if parrafo_tabla is None:
        raise ValueError(
            "La plantilla Word no tiene el parámetro {{TABLA_COTIZACION}}."
        )

    insertar_tabla_despues_de_parrafo(parrafo_tabla, tabla)
    eliminar_parrafo(parrafo_tabla)
    titulo_1 = "COTIZACIÓN FINAL"

    reemplazos = {
        "{{FECHA}}": fecha_str,
        "{{TITULO_1}}": titulo_1,
        "{{EMPRESA_RECIBE}}": nombre_cliente,
        "{{EMPRESA_BRINDA}}": empresa_nombre,
        "{{NOMBRE_PROGRAMA}}": nombre_programa,
        "{{TOTAL}}": formatear_moneda(total_general),
        "{{TOTAL_LETRA}}": convertir_numero_a_letras_mxn(total_general),
        "{{FIRMA_EMPRESA}}": empresa_nombre,
    }

    reemplazar_parametros_documento(doc, reemplazos)

    if "texto_normal" in estilos_bd:
        aplicar_estilo_base_a_todo(doc, estilos_bd["texto_normal"])
    else:
        aplicar_fuente_base_a_todo(doc, fuente, size_letra, paleta["texto"])

    estilo_base = estilos_bd.get("texto_normal")

    aplicar_negrita_a_texto(doc, nombre_cliente, estilo_base)
    aplicar_negrita_a_texto(doc, empresa_nombre, estilo_base)

    if "titulo_1" in estilos_bd:
        aplicar_estilo_a_texto(doc, titulo_1, estilos_bd["titulo_1"])

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    return output


def aplicar_fuente_base_a_todo(doc, fuente, size_letra, color_texto):
    for parrafo in doc.paragraphs:
        for run in parrafo.runs:
            run.font.name = fuente
            run.font.size = Pt(size_letra)
            run.font.color.rgb = RGBColor(*color_texto)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    for run in parrafo.runs:
                        run.font.name = fuente
                        run.font.size = Pt(size_letra)
                        run.font.color.rgb = RGBColor(*color_texto)


def aplicar_alineacion_parrafo(parrafo, alineacion):
    alineacion = alineacion.lower().strip()

    if alineacion in ["izquierda", "left"]:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif alineacion in ["centro", "center", "centrado"]:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alineacion in ["derecha", "right"]:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif alineacion in ["justificado", "justify"]:
        parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def colorear_fondo_run(run, color_hex):
    color_hex = color_hex.replace("#", "")

    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    rPr.append(shd)

def mostrar_modulo_cotizacion_final():
    st.title("Cotización final")

    empresas = obtener_empresas_con_plantilla()

    if not empresas:
        st.error(
            "No hay empresas con plantilla Word registrada en empresa_plantilla_word.membrete_path."
        )
        return

    empresa_seleccionada = st.selectbox(
        "Selecciona la empresa prestadora del servicio",
        list(empresas.keys()),
    )

    empresa_data = empresas[empresa_seleccionada]

    plantilla_id = empresa_data["plantilla_id"]

    fuente = empresa_data["tipografia_base"]
    size_letra = int(empresa_data["tamanio_base"])

    color_texto_base = convertir_color_bd_a_rgb(
        empresa_data["color_texto_base"]
    )

    color_primario = convertir_color_bd_a_rgb(
        empresa_data["color_primario"]
    )

    color_secundario = convertir_color_bd_a_rgb(
        empresa_data["color_secundario"]
    )

    paleta = {
        "principal": color_primario,
        "secundario": color_secundario,
        "texto": color_texto_base,
    }

    estilos_bd = obtener_estilos_word_por_plantilla(plantilla_id)

    nombre_empresa = empresa_data["nombre"]
    ruta_plantilla_word = empresa_data["membrete_path"]

    nombre_cliente = st.text_input("Nombre del cliente / empresa que recibe")

    nombre_programa = st.text_input("Nombre del programa")

    color_primario = convertir_color_bd_a_rgb(empresa_data["color_primario"])
    color_secundario = convertir_color_bd_a_rgb(empresa_data["color_secundario"])

    paleta = {
        "principal": color_primario,
        "secundario": color_secundario,
        "texto": (0, 0, 0),
    }

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

    if not nombre_programa:
        st.warning("Escribe el nombre del programa.")
        return

    try:
        registros = extraer_datos_cotizacion_excel(archivo)

        st.subheader("Vista previa de conceptos y montos")

        df_preview = pd.DataFrame(registros)
        df_preview["Monto"] = df_preview["Monto"].map(formatear_moneda)

        st.dataframe(df_preview, use_container_width=True)

        word = crear_word_cotizacion(
            empresa_nombre=nombre_empresa,
            ruta_plantilla_word=ruta_plantilla_word,
            nombre_cliente=nombre_cliente,
            nombre_programa=nombre_programa,
            registros=registros,
            fuente=fuente,
            size_letra=size_letra,
            paleta=paleta,
            estilos_bd=estilos_bd,
        )

        st.download_button(
            label="Descargar cotización final (Word)",
            data=word,
            file_name="cotizacion_final.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_descargar_cotizacion_final"
        )

    except ValueError as e:
        st.error(str(e))

    except FileNotFoundError as e:
        st.error(str(e))

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")