import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Font


def limpiar_concepto(texto):
    if not texto:
        return ""
    texto = texto.strip()
    if texto.startswith("-"):
        texto = texto.lstrip("-").strip()
    return texto


def formatear_fecha(fecha_raw):
    if not fecha_raw:
        return ""
    try:
        fecha_dt = datetime.fromisoformat(fecha_raw)
        return f"{fecha_dt.day}/{fecha_dt.month}/{fecha_dt.year}"
    except Exception:
        try:
            solo_fecha = fecha_raw.split("T")[0]
            anio, mes, dia = solo_fecha.split("-")
            return f"{int(dia)}/{int(mes)}/{anio}"
        except Exception:
            return fecha_raw


def parse_xml(file):
    registros = []

    try:
        tree = ET.parse(file)
        root = tree.getroot()

        ns = {
            "cfdi": "http://www.sat.gob.mx/cfd/4",
            "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital"
        }

        fecha_raw = root.attrib.get("Fecha", "")
        total = root.attrib.get("Total", "0")
        serie = root.attrib.get("Serie", "")
        folio = root.attrib.get("Folio", "")
        serie = str(serie).strip()
        folio = str(folio).strip()
        serie_folio = f"{serie}-{folio}" if serie and folio else (serie or folio)

        fecha = formatear_fecha(fecha_raw)

        try:
            importe_total = float(total)
        except Exception:
            importe_total = 0.0

        timbre = root.find(".//tfd:TimbreFiscalDigital", ns)
        uuid = ""
        if timbre is not None:
            uuid = timbre.attrib.get("UUID", "")

        conceptos = root.findall(".//cfdi:Concepto", ns)

        for concepto in conceptos:
            descripcion = limpiar_concepto(concepto.attrib.get("Descripcion", ""))

            registros.append({
                "CONCEPTO": descripcion,
                "FECHA": fecha,
                "SERIE Y FOLIO": serie_folio,
                "UUID": uuid,
                "IMPORTE TOTAL": importe_total
            })

    except Exception as e:
        st.error(f"Error al leer el archivo {getattr(file, 'name', 'XML')}: {e}")

    return registros


def generar_excel_xml(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")

        worksheet = writer.sheets["Datos"]

        for cell in worksheet[1]:
            cell.font = Font(bold=True)

        for col in worksheet.iter_cols(min_col=5, max_col=5, min_row=2):
            for cell in col:
                cell.number_format = '$#,##0.00'

        worksheet.column_dimensions["A"].width = 90
        worksheet.column_dimensions["B"].width = 15
        worksheet.column_dimensions["C"].width = 22
        worksheet.column_dimensions["D"].width = 40
        worksheet.column_dimensions["E"].width = 18

    output.seek(0)
    return output


def mostrar_modulo_xml():
    st.title("Lector de XML CFDI a Excel")
    st.write("Sube uno o varios archivos XML para extraer CONCEPTO, FECHA, SERIE Y FOLIO, UUID e IMPORTE TOTAL.")

    files = st.file_uploader(
        "Selecciona uno o varios archivos XML",
        type=["xml"],
        accept_multiple_files=True,
        key="xml_uploader"
    )

    if files:
        all_data = []

        for file in files:
            all_data.extend(parse_xml(file))

        if all_data:
            df = pd.DataFrame(all_data)

            st.subheader("Vista previa")
            df_vista = df.copy()
            df_vista["IMPORTE TOTAL"] = df_vista["IMPORTE TOTAL"].apply(lambda x: f"${x:,.2f}")
            st.dataframe(df_vista, use_container_width=True)

            output = generar_excel_xml(df)

            st.download_button(
                label="Descargar Excel",
                data=output,
                file_name="resultado_xml.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No se encontraron datos válidos en los archivos XML.")
    else:
        st.info("Sube uno o varios archivos XML para comenzar.")
