import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Font

st.set_page_config(page_title="Lector de XML CFDI", layout="wide")

st.title("Lector de XML CFDI a Excel")
st.write("Sube uno o varios archivos XML para extraer CONCEPTO, FECHA, FOLIO e IMPORTE.")

files = st.file_uploader(
    "Selecciona uno o varios archivos XML",
    type=["xml"],
    accept_multiple_files=True
)

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

        fecha = formatear_fecha(fecha_raw)

        try:
            importe_total = float(total)
        except Exception:
            importe_total = 0.0

        # Tomar UUID como FOLIO
        timbre = root.find(".//tfd:TimbreFiscalDigital", ns)
        folio_uuid = ""
        if timbre is not None:
            folio_uuid = timbre.attrib.get("UUID", "")

        conceptos = root.findall(".//cfdi:Concepto", ns)

        for concepto in conceptos:
            descripcion = limpiar_concepto(concepto.attrib.get("Descripcion", ""))

            registros.append({
                "CONCEPTO": descripcion,
                "FECHA": fecha,
                "FOLIO": folio_uuid,
                "IMPORTE": importe_total
            })

    except Exception as e:
        st.error(f"Error al leer el archivo {file.name}: {e}")

    return registros

if files:
    all_data = []

    for file in files:
        all_data.extend(parse_xml(file))

    if all_data:
        df = pd.DataFrame(all_data)

        st.subheader("Vista previa")
        df_vista = df.copy()
        df_vista["IMPORTE"] = df_vista["IMPORTE"].apply(lambda x: f"${x:,.2f}")
        st.dataframe(df_vista, use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Datos")

            worksheet = writer.sheets["Datos"]

            for cell in worksheet[1]:
                cell.font = Font(bold=True)

            for col in worksheet.iter_cols(min_col=4, max_col=4, min_row=2):
                for cell in col:
                    cell.number_format = '$#,##0.00'

            worksheet.column_dimensions["A"].width = 90
            worksheet.column_dimensions["B"].width = 15
            worksheet.column_dimensions["C"].width = 40
            worksheet.column_dimensions["D"].width = 18

        output.seek(0)

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