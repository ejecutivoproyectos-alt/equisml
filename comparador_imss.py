import re
import unicodedata
from calendar import monthrange
from io import BytesIO
import pandas as pd
import streamlit as st

def safe_str(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def normalizar_texto(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = s.replace("#", "N")
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = re.sub(r"^\s*[^-]+-\s*", "", s)
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parsear_fecha_imss(valor):
    if pd.isna(valor):
        return None

    s = str(valor).strip()
    if s in ("", "-"):
        return None

    meses = {
        "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
        "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12
    }

    m = re.match(r"(\d{1,2})/([A-ZÁÉÍÓÚÑ]{3})/(\d{4})", s.upper())
    if m:
        dia = int(m.group(1))
        mes_txt = unicodedata.normalize("NFKD", m.group(2)).encode("ASCII", "ignore").decode("ASCII")
        mes = meses.get(mes_txt[:3])
        anio = int(m.group(3))
        if mes:
            return pd.Timestamp(anio, mes, dia)

    fecha = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(fecha):
        return None
    return fecha


def obtener_hoja_imss_desde_excel_file(excel_file):
    hojas = excel_file.sheet_names

    for hoja in hojas:
        hoja_up = safe_str(hoja).upper()
        if "MOVIMIENTOS" in hoja_up and ("EMA" in hoja_up or "EBA" in hoja_up):
            return hoja

    for hoja in hojas:
        hoja_up = safe_str(hoja).upper()
        if "EMA" in hoja_up or "EBA" in hoja_up:
            return hoja

    raise ValueError("No se encontró una hoja tipo EMA o EBA en el archivo IMSS.")


def extraer_periodo_desde_hoja(nombre_hoja: str):
    nombre_hoja = safe_str(nombre_hoja)
    nombre_norm = unicodedata.normalize("NFKD", nombre_hoja).encode("ASCII", "ignore").decode("ASCII").upper()

    tipo = "EMA" if "EMA" in nombre_norm else "EBA"

    meses_map = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
        "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9,
        "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
    }

    anio_match = re.search(r"(\d{4})", nombre_norm)
    if not anio_match:
        raise ValueError("No se pudo detectar el año en el nombre de la hoja del IMSS.")

    anio = int(anio_match.group(1))

    meses_encontrados = []
    for nombre_mes, num_mes in meses_map.items():
        if nombre_mes in nombre_norm:
            meses_encontrados.append(num_mes)

    meses_ordenados = []
    for mes in meses_encontrados:
        if mes not in meses_ordenados:
            meses_ordenados.append(mes)

    if not meses_ordenados:
        raise ValueError("No se pudo detectar el mes en el nombre de la hoja del IMSS.")

    if tipo == "EMA":
        return tipo, anio, [meses_ordenados[0]]

    return tipo, anio, meses_ordenados[:2]


def detectar_fila_encabezado_imss(df_preview):
    for i in range(min(15, len(df_preview))):
        fila = [safe_str(x).upper() for x in df_preview.iloc[i].tolist()]
        fila_texto = " | ".join(fila)
        if "NOMBRE" in fila_texto and "TIPO DEL MOVIMIENTO" in fila_texto:
            return i
    return 4


def leer_imss(imss_file):
    contenido = imss_file.read()
    bio = BytesIO(contenido)

    excel_file = pd.ExcelFile(bio)
    hoja = obtener_hoja_imss_desde_excel_file(excel_file)
    tipo_periodo, anio, meses = extraer_periodo_desde_hoja(hoja)

    bio.seek(0)
    df_preview = pd.read_excel(bio, sheet_name=hoja, header=None)
    fila_header = detectar_fila_encabezado_imss(df_preview)

    bio.seek(0)
    df = pd.read_excel(bio, sheet_name=hoja, header=fila_header)
    df.columns = [safe_str(c) for c in df.columns]

    registros = []

    for _, row in df.iterrows():
        nombre = safe_str(row.get("Nombre"))
        if not nombre:
            continue

        mov = safe_str(row.get("Tipo del Movimiento"))
        fecha_mov = parsear_fecha_imss(row.get("Fecha del Movimiento"))
        dias_col = row.get("Días")

        try:
            mov = int(mov)
        except:
            continue

        try:
            dias_col = int(dias_col)
        except:
            dias_col = 0

        dias_esperados = dias_col
        estatus_laboral = "Otro"

        if tipo_periodo == "EMA":
            mes = meses[0]
            ultimo_dia = monthrange(anio, mes)[1]
            inicio_mes = pd.Timestamp(anio, mes, 1)
            fin_mes = pd.Timestamp(anio, mes, ultimo_dia)

            if mov == 9:
                dias_esperados = dias_col
                estatus_laboral = "Laboró"
            elif mov == 8:
                dias_esperados = (fin_mes - fecha_mov).days + 1 if fecha_mov is not None else dias_col
                estatus_laboral = "Alta"
            elif mov == 2:
                dias_esperados = (fecha_mov - inicio_mes).days if fecha_mov is not None else dias_col
                estatus_laboral = "Baja"
            elif mov == 7:
                dias_esperados = dias_col
                estatus_laboral = "Cambio salarial"
            elif mov == 11:
                dias_esperados = 0
                estatus_laboral = "Ausentismo"
            elif mov == 12:
                dias_esperados = 0
                estatus_laboral = "Incapacidad"
            else:
                dias_esperados = dias_col
                estatus_laboral = f"Movimiento {mov}"

        else:
            if mov == 9:
                estatus_laboral = "Laboró"
            elif mov == 8:
                estatus_laboral = "Alta"
            elif mov == 2:
                estatus_laboral = "Baja"
            elif mov == 7:
                estatus_laboral = "Cambio salarial"
            elif mov == 11:
                dias_esperados = 0
                estatus_laboral = "Ausentismo"
            elif mov == 12:
                dias_esperados = 0
                estatus_laboral = "Incapacidad"
            else:
                estatus_laboral = f"Movimiento {mov}"

        registros.append({
            "empleado_imss": str(nombre).strip(),
            "nombre_norm": normalizar_texto(nombre),
            "tipo_movimiento": mov,
            "fecha_movimiento": fecha_mov,
            "dias_imss_columna": dias_col,
            "dias_esperados_imss": dias_esperados,
            "estatus_laboral": estatus_laboral,
            "hoja_imss": hoja,
            "tipo_periodo": tipo_periodo,
            "anio": anio,
            "meses": meses
        })

    df_imss = pd.DataFrame(registros)
    return df_imss, hoja, tipo_periodo

def leer_reporte(reporte_file):
    contenido = reporte_file.read()
    bio = BytesIO(contenido)

    df = pd.read_excel(bio, sheet_name=0, header=None)

    fechas = pd.to_datetime(df.iloc[5, 3:], dayfirst=True, errors="coerce")

    resumen = []

    for i in range(6, len(df)):
        empleado = safe_str(df.iloc[i, 1])
        if not empleado:
            continue

        nombre_norm = normalizar_texto(empleado)
        fila_estados = df.iloc[i, 3:3 + len(fechas)].tolist()

        laborado = 0
        no_habia_ingresado = 0
        dado_baja = 0

        for valor in fila_estados:
            txt = safe_str(valor).upper()
            txt_norm = unicodedata.normalize("NFKD", txt).encode("ASCII", "ignore").decode("ASCII")

            if txt_norm == "LABORADO":
                laborado += 1
            elif "NO HABIA INGRESADO" in txt_norm:
                no_habia_ingresado += 1
            elif "BAJA" in txt_norm or "YA NO TRABAJA" in txt_norm or "DADO DE BAJA" in txt_norm:
                dado_baja += 1

        resumen.append({
            "empleado_reporte": str(empleado).strip(),
            "nombre_norm": nombre_norm,
            "dias_laborado_reporte": laborado,
            "dias_no_ingresado_reporte": no_habia_ingresado,
            "dias_baja_reporte": dado_baja
        })

    return pd.DataFrame(resumen)


def resumir_movimientos_por_empleado(df_imss):
    if df_imss.empty:
        return pd.DataFrame()

    resumen = []

    for nombre_norm, grupo in df_imss.groupby("nombre_norm", dropna=False):
        empleado_imss = safe_str(grupo["empleado_imss"].iloc[0])

        movimientos_raw = grupo["tipo_movimiento"].dropna().tolist()
        movimientos = sorted(set(int(x) for x in movimientos_raw if pd.notna(x)))

        fechas_mov = [x for x in grupo["fecha_movimiento"].dropna().tolist() if pd.notna(x)]

        tiene_9 = 9 in movimientos
        tiene_8 = 8 in movimientos
        tiene_2 = 2 in movimientos
        tiene_7 = 7 in movimientos
        tiene_11 = 11 in movimientos
        tiene_12 = 12 in movimientos

        dias_numericos = pd.to_numeric(
            grupo["dias_esperados_imss"], errors="coerce"
        ).fillna(0)

        if tiene_11 or tiene_12:
            dias_esperados = 0
        else:
            dias_esperados = dias_numericos.max()

        estado_resumen = []
        if tiene_8:
            estado_resumen.append("Alta")
        if tiene_2:
            estado_resumen.append("Baja")
        if tiene_9:
            estado_resumen.append("Laboró")
        if tiene_7:
            estado_resumen.append("Cambio salarial")
        if tiene_11:
            estado_resumen.append("Ausentismo")
        if tiene_12:
            estado_resumen.append("Incapacidad")

        resumen.append({
            "empleado_imss": empleado_imss,
            "nombre_norm": safe_str(nombre_norm),
            "movimientos_detectados": ", ".join([str(x) for x in movimientos]),
            "detalle_movimientos": ", ".join([safe_str(x) for x in estado_resumen]),
            "dias_esperados_imss": int(dias_esperados) if pd.notna(dias_esperados) else 0,
            "fecha_primer_movimiento": min(fechas_mov) if fechas_mov else None,
            "fecha_ultimo_movimiento": max(fechas_mov) if fechas_mov else None,
            "duplicado_en_imss": len(grupo) > 1
        })

    return pd.DataFrame(resumen)


def comparar_archivos(imss_file, reporte_file):
    df_imss_raw, hoja, tipo_periodo = leer_imss(imss_file)
    df_reporte = leer_reporte(reporte_file)

    df_imss = resumir_movimientos_por_empleado(df_imss_raw)

    comparacion = df_imss.merge(
        df_reporte,
        on="nombre_norm",
        how="left",
        indicator=True
    )

    comparacion["encontrado_en_reporte"] = comparacion["_merge"] == "both"
    comparacion["dias_coinciden"] = comparacion["dias_esperados_imss"] == comparacion["dias_laborado_reporte"]

    no_encontrados = comparacion[~comparacion["encontrado_en_reporte"]].copy()
    correctos = comparacion[
        comparacion["encontrado_en_reporte"] & comparacion["dias_coinciden"]
    ].copy()
    con_diferencia = comparacion[
        comparacion["encontrado_en_reporte"] & ~comparacion["dias_coinciden"]
    ].copy()

    duplicados = df_imss[df_imss["duplicado_en_imss"]].copy()

    altas = comparacion[comparacion["detalle_movimientos"].str.contains("Alta", na=False)].copy()
    bajas = comparacion[comparacion["detalle_movimientos"].str.contains("Baja", na=False)].copy()
    laboraron = comparacion[comparacion["detalle_movimientos"].str.contains("Laboró", na=False)].copy()
    salario = comparacion[comparacion["detalle_movimientos"].str.contains("Cambio salarial", na=False)].copy()

    # AQUI VA EL CAMBIO NUEVO
    ausentismo = comparacion[comparacion["detalle_movimientos"].str.contains("Ausentismo", na=False)].copy()
    incapacidad = comparacion[comparacion["detalle_movimientos"].str.contains("Incapacidad", na=False)].copy()

    return {
        "hoja_imss": hoja,
        "tipo_periodo": tipo_periodo,
        "imss_raw": df_imss_raw,
        "imss_resumen": df_imss,
        "reporte": df_reporte,
        "comparacion": comparacion,
        "no_encontrados": no_encontrados,
        "correctos": correctos,
        "con_diferencia": con_diferencia,
        "duplicados": duplicados,
        "altas": altas,
        "bajas": bajas,
        "laboraron": laboraron,
        "salario": salario,
        "ausentismo": ausentismo,      # NUEVO
        "incapacidad": incapacidad     # NUEVO
    }

def mostrar_resumen(resultado):
    st.subheader("Resumen general")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Registros IMSS", len(resultado["imss_raw"]))
    col2.metric("Empleados únicos IMSS", len(resultado["imss_resumen"]))
    col3.metric("Empleados en reporte", len(resultado["correctos"]) + len(resultado["con_diferencia"]))
    col4.metric("No encontrados", len(resultado["no_encontrados"]))

    st.write(f"**Hoja IMSS detectada:** {resultado['hoja_imss']}")
    st.write(f"**Tipo de periodo:** {resultado['tipo_periodo']}")

    st.markdown("---")

    st.subheader("Movimientos encontrados en IMSS")

    col1, col2, col3 = st.columns(3)
    col1.metric("Laboró (9)", len(resultado["laboraron"]))
    col2.metric("Altas (8)", len(resultado["altas"]))
    col3.metric("Bajas (2)", len(resultado["bajas"]))

    col4, col5, col6 = st.columns(3)
    col4.metric("Cambio salarial (7)", len(resultado["salario"]))
    col5.metric("Ausentismo (11)", len(resultado["ausentismo"]))
    col6.metric("Incapacidad (12)", len(resultado["incapacidad"]))

    st.markdown("---")

    st.subheader("Resultado de validación")
    c1, c2, c3 = st.columns(3)
    c1.metric("Coinciden días", len(resultado["correctos"]))
    c2.metric("Con diferencia", len(resultado["con_diferencia"]))
    c3.metric("Duplicados en IMSS", len(resultado["duplicados"]))

def mostrar_detalle_tablas(resultado):
    st.subheader("Empleados duplicados en IMSS")
    if resultado["duplicados"].empty:
        st.success("No se detectaron empleados duplicados en IMSS.")
    else:
        st.dataframe(
            resultado["duplicados"][[
                "empleado_imss",
                "movimientos_detectados",
                "detalle_movimientos",
                "dias_esperados_imss",
                "fecha_primer_movimiento",
                "fecha_ultimo_movimiento"
            ]],
            use_container_width=True
        )

    st.subheader("Altas detectadas")
    if resultado["altas"].empty:
        st.info("No se detectaron altas.")
    else:
        st.dataframe(
            resultado["altas"][[
                "empleado_imss",
                "detalle_movimientos",
                "fecha_primer_movimiento",
                "dias_esperados_imss"
            ]],
            use_container_width=True
        )

    st.subheader("Bajas detectadas")
    if resultado["bajas"].empty:
        st.info("No se detectaron bajas.")
    else:
        st.dataframe(
            resultado["bajas"][[
                "empleado_imss",
                "detalle_movimientos",
                "fecha_primer_movimiento",
                "dias_esperados_imss"
            ]],
            use_container_width=True
        )

    st.subheader("No encontrados en reporte")
    if resultado["no_encontrados"].empty:
        st.success("Todos los empleados del IMSS sí aparecen en el reporte.")
    else:
        st.dataframe(
            resultado["no_encontrados"][[
                "empleado_imss",
                "detalle_movimientos",
                "dias_esperados_imss"
            ]],
            use_container_width=True
        )

    st.subheader("Empleados con diferencia en días")
    if resultado["con_diferencia"].empty:
        st.success("No se detectaron diferencias en días laborados.")
    else:
        st.dataframe(
            resultado["con_diferencia"][[
                "empleado_imss",
                "empleado_reporte",
                "detalle_movimientos",
                "dias_esperados_imss",
                "dias_laborado_reporte"
            ]],
            use_container_width=True
        )

    st.subheader("Empleados correctos")
    if resultado["correctos"].empty:
        st.warning("No hubo coincidencias correctas.")
    else:
        st.dataframe(
            resultado["correctos"][[
                "empleado_imss",
                "empleado_reporte",
                "detalle_movimientos",
                "dias_esperados_imss",
                "dias_laborado_reporte"
            ]],
            use_container_width=True
        )

    st.subheader("Ausentismo detectado")
    if resultado["ausentismo"].empty:
        st.info("No se detectó ausentismo.")
    else:
        st.dataframe(
            resultado["ausentismo"][[
                "empleado_imss",
                "detalle_movimientos",
                "fecha_primer_movimiento",
                "dias_esperados_imss"
            ]],
            use_container_width=True
        )

    st.subheader("Incapacidades detectadas")
    if resultado["incapacidad"].empty:
        st.info("No se detectaron incapacidades.")
    else:
        st.dataframe(
            resultado["incapacidad"][[
                "empleado_imss",
                "detalle_movimientos",
                "fecha_primer_movimiento",
                "dias_esperados_imss"
            ]],
            use_container_width=True
        )


def mostrar_conclusion(resultado):
    st.subheader("Conclusión")

    if len(resultado["no_encontrados"]) == 0 and len(resultado["con_diferencia"]) == 0:
        st.success(
            "El reporte está correctamente armado: todos los empleados del IMSS aparecen en el reporte y los días laborados coinciden."
        )
    else:
        mensajes = []

        if len(resultado["no_encontrados"]) > 0:
            mensajes.append(f"hay {len(resultado['no_encontrados'])} empleados del IMSS que no aparecen en el reporte")

        if len(resultado["con_diferencia"]) > 0:
            mensajes.append(f"hay {len(resultado['con_diferencia'])} empleados con diferencia en días laborados")

        st.error("Se detectaron diferencias: " + " y ".join(mensajes) + ".")


def mostrar_modulo_comparador_imss():
    st.title("Comparar IMSS vs Reporte de trabajadores")
    st.write(
        "Sube el archivo del IMSS y el reporte de días trabajados para validar si coinciden los empleados y sus días laborados."
    )

    col1, col2 = st.columns(2)

    with col1:
        imss_file = st.file_uploader(
            "Sube el archivo IMSS",
            type=["xls", "xlsx"],
            key="imss_file"
        )

    with col2:
        reporte_file = st.file_uploader(
            "Sube el reporte de trabajadores",
            type=["xlsx", "xls"],
            key="reporte_file"
        )

    if imss_file and reporte_file:
        try:
            resultado = comparar_archivos(imss_file, reporte_file)

            mostrar_resumen(resultado)
            mostrar_detalle_tablas(resultado)
            mostrar_conclusion(resultado)

        except Exception as e:
            st.error(f"Ocurrió un error al comparar los archivos: {e}")
    else:
        st.info("Carga ambos archivos para comenzar la comparación.")