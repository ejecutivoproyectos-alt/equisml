import random
from io import BytesIO

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


THIN_BLACK = Side(style="thin", color="000000")


def normalizar_texto(valor):
    if valor is None:
        return ""
    return str(valor).strip()


def centrar(cell, bold=False, size=12):
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.font = Font(bold=bold, size=size)


def aplicar_borde_rango(ws, r1, c1, r2, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = Border(
                left=THIN_BLACK,
                right=THIN_BLACK,
                top=THIN_BLACK,
                bottom=THIN_BLACK
            )


def ajustar_ancho(ws, col_idx, width):
    ws.column_dimensions[get_column_letter(col_idx)].width = width


def limpiar_hoja_si_existe(wb, nombre_hoja):
    if nombre_hoja in wb.sheetnames:
        del wb[nombre_hoja]


def safe_sheet_name(name):
    invalid = ['\\', '/', '*', '[', ']', ':', '?']
    for ch in invalid:
        name = name.replace(ch, "")
    return name[:31]


def generar_desempeno():
    return random.choices(
        population=[5, 4, 3],
        weights=[60, 35, 5],
        k=1
    )[0]


def primer_no_vacio_en_fila(ws, row_num):
    for cell in ws[row_num]:
        if normalizar_texto(cell.value):
            return normalizar_texto(cell.value)
    return ""


def color_hex_desde_fill(cell):
    fill = cell.fill
    if fill is None:
        return "D9D9D9"

    fg = fill.fgColor
    if fg is None:
        return "D9D9D9"

    rgb = getattr(fg, "rgb", None)
    if rgb:
        rgb = str(rgb)
        if len(rgb) == 8:
            return rgb[2:]
        if len(rgb) == 6:
            return rgb

    return "D9D9D9"


def es_participa(cell):
    valor = cell.value
    s = normalizar_texto(valor).strip().lower()
    return s == "a"


def leer_tabla_actividades(file_bytes):
    wb = load_workbook(BytesIO(file_bytes))
    ws = wb.active

    titulo = primer_no_vacio_en_fila(ws, 1) or "TABLA DE ACTIVIDADES"
    puesto = primer_no_vacio_en_fila(ws, 2) or "PUESTO SIN DEFINIR"

    header_row = 3
    start_data_row = 5
    col_num = 1
    col_act = 2
    start_emp_col = 3

    empleados = []
    col = start_emp_col

    while col <= ws.max_column:
        nombre = normalizar_texto(ws.cell(header_row, col).value)
        if nombre == "":
            break

        color = color_hex_desde_fill(ws.cell(header_row, col))

        empleados.append({
            "nombre": nombre,
            "col": col,
            "color": color
        })
        col += 1

    actividades = []
    row = start_data_row

    while row <= ws.max_row:
        num = ws.cell(row, col_num).value
        actividad = normalizar_texto(ws.cell(row, col_act).value)

        if actividad == "":
            break

        participantes = []
        for emp in empleados:
            cell_part = ws.cell(row, emp["col"])
            if es_participa(cell_part):
                participantes.append({
                    "nombre": emp["nombre"],
                    "color": emp["color"]
                })

        actividades.append({
            "numero": num,
            "actividad": actividad,
            "participantes": participantes
        })

        row += 1

    return {
        "titulo": titulo,
        "puesto": puesto,
        "empleados": empleados,
        "actividades": actividades
    }


def generar_rendimiento(actividades, semanas=4):
    resultado = []

    for act in actividades:
        semanas_data = []
        for semana in range(1, semanas + 1):
            bloques = []
            for participante in act["participantes"]:
                bloques.append({
                    "nombre": participante["nombre"],
                    "color": participante["color"],
                    "valor": generar_desempeno()
                })

            semanas_data.append({
                "semana": semana,
                "bloques": bloques
            })

        resultado.append({
            "numero": act["numero"],
            "actividad": act["actividad"],
            "semanas": semanas_data
        })

    return resultado


def render_preview_html(puesto, mes_anio, rendimiento):
    html = f"""
    <div style="font-family:Arial,sans-serif; overflow-x:auto; margin-bottom:30px;">
      <table style="border-collapse:collapse; min-width:1100px; width:100%;">
        <tr>
          <th colspan="5" style="border:1px solid #000; background:#f2be00; font-size:22px; padding:8px;">
            TABLA DE RENDIMIENTO
          </th>
        </tr>
        <tr>
          <th colspan="5" style="border:1px solid #000; background:#f2be00; font-size:20px; padding:8px;">
            {puesto}
          </th>
        </tr>
        <tr>
          <th colspan="5" style="border:1px solid #000; background:#f2be00; font-size:20px; padding:8px;">
            {mes_anio}
          </th>
        </tr>
        <tr>
          <th style="border:1px solid #000; background:#eeeeee; width:100px; font-size:18px; padding:6px;">
            #ACT.
          </th>
          <th style="border:1px solid #000; background:#eeeeee; width:240px; font-size:18px;">
            SEMANA 1
          </th>
          <th style="border:1px solid #000; background:#eeeeee; width:240px; font-size:18px;">
            SEMANA 2
          </th>
          <th style="border:1px solid #000; background:#eeeeee; width:240px; font-size:18px;">
            SEMANA 3
          </th>
          <th style="border:1px solid #000; background:#eeeeee; width:240px; font-size:18px;">
            SEMANA 4
          </th>
        </tr>
    """

    for item in rendimiento:
        html += f"""
        <tr>
          <td style="border:1px solid #000; background:#f7f7f7; text-align:center; font-size:18px; padding:6px;">
            {item['numero']}
          </td>
        """

        for semana in item["semanas"]:
            bloques = semana["bloques"]

            if not bloques:
                html += '<td style="border:1px solid #000; height:38px; background:#fff;"></td>'
                continue

            html += '<td style="border:1px solid #000; padding:0; height:38px;">'
            html += '<div style="display:flex; width:100%; height:38px;">'

            for i, bloque in enumerate(bloques):
                borde = "border-right:1px solid #000;" if i < len(bloques) - 1 else ""
                html += f"""
                <div style="
                    flex:1;
                    background:#{bloque['color']};
                    {borde}
                    display:flex;
                    justify-content:center;
                    align-items:center;
                    font-size:18px;
                ">
                    {bloque['valor']}
                </div>
                """

            html += '</div></td>'

        html += "</tr>"

    html += "</table></div>"
    return html


def escribir_hoja_rendimiento(wb, nombre_hoja, puesto, mes_anio, rendimiento):
    limpiar_hoja_si_existe(wb, nombre_hoja)
    ws = wb.create_sheet(title=safe_sheet_name(nombre_hoja))

    yellow_fill = PatternFill(fill_type="solid", fgColor="F2BE00")
    gray_fill = PatternFill(fill_type="solid", fgColor="EFEFEF")
    white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")

    bloques_semana = {
        1: (2, 7),
        2: (8, 13),
        3: (14, 19),
        4: (20, 25)
    }

    ws.merge_cells("A1:Y1")
    ws["A1"] = "TABLA DE RENDIMIENTO"
    ws["A1"].fill = yellow_fill
    centrar(ws["A1"], bold=True, size=16)

    ws.merge_cells("A2:Y2")
    ws["A2"] = puesto
    ws["A2"].fill = yellow_fill
    centrar(ws["A2"], bold=True, size=15)

    ws.merge_cells("A3:Y3")
    ws["A3"] = mes_anio
    ws["A3"].fill = yellow_fill
    centrar(ws["A3"], bold=True, size=15)

    ws.merge_cells("A4:A5")
    ws["A4"] = "#ACT."
    ws["A4"].fill = gray_fill
    centrar(ws["A4"], bold=True, size=14)

    for semana, (c1, c2) in bloques_semana.items():
        ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
        cell = ws.cell(4, c1)
        cell.value = f"SEMANA {semana}"
        cell.fill = gray_fill
        centrar(cell, bold=True, size=14)

        for col in range(c1, c2 + 1):
            ws.cell(5, col).fill = white_fill

    aplicar_borde_rango(ws, 1, 1, 5, 25)

    ajustar_ancho(ws, 1, 10)
    for col in range(2, 26):
        ajustar_ancho(ws, col, 5)

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 28
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 12

    start_row = 6

    for idx, item in enumerate(rendimiento, start=0):
        excel_row = start_row + idx
        ws.row_dimensions[excel_row].height = 28

        c = ws.cell(excel_row, 1)
        c.value = item["numero"]
        c.fill = white_fill
        centrar(c, size=14)
        c.border = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)

        for sem_idx, semana in enumerate(item["semanas"], start=1):
            c1, c2 = bloques_semana[sem_idx]
            bloques = semana["bloques"]

            if len(bloques) == 0:
                ws.merge_cells(start_row=excel_row, start_column=c1, end_row=excel_row, end_column=c2)
                mc = ws.cell(excel_row, c1)
                mc.value = ""
                mc.fill = white_fill
                centrar(mc, size=14)
                aplicar_borde_rango(ws, excel_row, c1, excel_row, c2)

            elif len(bloques) == 1:
                b = bloques[0]
                ws.merge_cells(start_row=excel_row, start_column=c1, end_row=excel_row, end_column=c2)
                mc = ws.cell(excel_row, c1)
                mc.value = b["valor"]
                mc.fill = PatternFill(fill_type="solid", fgColor=b["color"])
                centrar(mc, size=14)
                aplicar_borde_rango(ws, excel_row, c1, excel_row, c2)

            elif len(bloques) == 2:
                rangos = [(c1, c1 + 2), (c1 + 3, c2)]
                for b, (r1, r2) in zip(bloques, rangos):
                    ws.merge_cells(start_row=excel_row, start_column=r1, end_row=excel_row, end_column=r2)
                    mc = ws.cell(excel_row, r1)
                    mc.value = b["valor"]
                    mc.fill = PatternFill(fill_type="solid", fgColor=b["color"])
                    centrar(mc, size=14)
                    aplicar_borde_rango(ws, excel_row, r1, excel_row, r2)

            else:
                total_cols = c2 - c1 + 1
                n = len(bloques)
                ancho_base = total_cols // n
                resto = total_cols % n

                rangos = []
                actual = c1

                for i in range(n):
                    ancho = ancho_base + (1 if i < resto else 0)
                    fin = actual + ancho - 1
                    rangos.append((actual, fin))
                    actual = fin + 1

                for b, (r1, r2) in zip(bloques, rangos):
                    ws.merge_cells(start_row=excel_row, start_column=r1, end_row=excel_row, end_column=r2)
                    mc = ws.cell(excel_row, r1)
                    mc.value = b["valor"]
                    mc.fill = PatternFill(fill_type="solid", fgColor=b["color"])
                    centrar(mc, size=14)
                    aplicar_borde_rango(ws, excel_row, r1, excel_row, r2)

    return wb


def construir_excel_final(file_bytes, puesto, mes1, anio1, mes2, anio2, rendimiento1, rendimiento2):
    wb = load_workbook(BytesIO(file_bytes))

    hoja1 = f"RENDIMIENTO {mes1} {anio1}"
    hoja2 = f"RENDIMIENTO {mes2} {anio2}"

    escribir_hoja_rendimiento(
        wb=wb,
        nombre_hoja=hoja1,
        puesto=puesto,
        mes_anio=f"{mes1} {anio1}",
        rendimiento=rendimiento1
    )

    escribir_hoja_rendimiento(
        wb=wb,
        nombre_hoja=hoja2,
        puesto=puesto,
        mes_anio=f"{mes2} {anio2}",
        rendimiento=rendimiento2
    )

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.getvalue()


def mostrar_modulo_rendimiento():
    st.title("Generador de dos tablas de rendimiento desde Excel")

    uploaded_file = st.file_uploader(
        "Sube el Excel con la TABLA DE ACTIVIDADES",
        type=["xlsx"],
        key="rendimiento_uploader"
    )

    meses = [
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]

    st.subheader("Mes 1")
    c1, c2 = st.columns(2)
    with c1:
        mes1 = st.selectbox("Mes 1", meses, index=4, key="mes1")
    with c2:
        anio1 = st.number_input("Año 1", min_value=2024, max_value=2100, value=2025, step=1, key="anio1")

    st.subheader("Mes 2")
    c3, c4 = st.columns(2)
    with c3:
        mes2 = st.selectbox("Mes 2", meses, index=5, key="mes2")
    with c4:
        anio2 = st.number_input("Año 2", min_value=2024, max_value=2100, value=2025, step=1, key="anio2")

    if uploaded_file is not None:
        file_bytes = uploaded_file.read()

        try:
            data = leer_tabla_actividades(file_bytes)

            puesto = data["puesto"]
            empleados = data["empleados"]
            actividades = data["actividades"]

            st.success("Archivo leído correctamente.")
            st.write(f"**Puesto detectado:** {puesto}")
            st.write(f"**Empleados detectados:** {len(empleados)}")
            st.write(f"**Actividades detectadas:** {len(actividades)}")

            with st.expander("Ver empleados y colores detectados"):
                for emp in empleados:
                    st.markdown(
                        f"""
                        <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                            <div style="width:22px; height:22px; background:#{emp['color']}; border:1px solid #000;"></div>
                            <div>{emp['nombre']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

            with st.expander("Ver participación detectada por actividad"):
                for act in actividades:
                    nombres = [p["nombre"] for p in act["participantes"]]
                    st.write(f"{act['numero']}. {act['actividad']} → {', '.join(nombres) if nombres else 'Sin participantes'}")

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                if st.button("Generar dos meses", key="btn_generar"):
                    st.session_state["rendimiento_mes_1"] = generar_rendimiento(actividades, semanas=4)
                    st.session_state["rendimiento_mes_2"] = generar_rendimiento(actividades, semanas=4)
                    st.session_state["file_bytes_original"] = file_bytes
                    st.session_state["puesto_detectado"] = puesto

            with col_btn2:
                if st.button("Regenerar aleatoriamente", key="btn_regenerar"):
                    st.session_state["rendimiento_mes_1"] = generar_rendimiento(actividades, semanas=4)
                    st.session_state["rendimiento_mes_2"] = generar_rendimiento(actividades, semanas=4)

            if "rendimiento_mes_1" in st.session_state and "rendimiento_mes_2" in st.session_state:
                st.subheader("Vista previa: Mes 1")
                html1 = render_preview_html(
                    puesto=st.session_state["puesto_detectado"],
                    mes_anio=f"{mes1} {anio1}",
                    rendimiento=st.session_state["rendimiento_mes_1"]
                )
                st.components.v1.html(html1, height=700, scrolling=True)

                st.subheader("Vista previa: Mes 2")
                html2 = render_preview_html(
                    puesto=st.session_state["puesto_detectado"],
                    mes_anio=f"{mes2} {anio2}",
                    rendimiento=st.session_state["rendimiento_mes_2"]
                )
                st.components.v1.html(html2, height=700, scrolling=True)

                excel_final = construir_excel_final(
                    file_bytes=st.session_state["file_bytes_original"],
                    puesto=st.session_state["puesto_detectado"],
                    mes1=mes1,
                    anio1=anio1,
                    mes2=mes2,
                    anio2=anio2,
                    rendimiento1=st.session_state["rendimiento_mes_1"],
                    rendimiento2=st.session_state["rendimiento_mes_2"]
                )

                st.download_button(
                    label="Descargar Excel final con dos tablas",
                    data=excel_final,
                    file_name=f"rendimiento_{mes1.lower()}_{anio1}_{mes2.lower()}_{anio2}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el archivo: {e}")
    else:
        st.info("Sube el Excel de actividades para comenzar.")
