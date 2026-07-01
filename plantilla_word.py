import os
import re

import matplotlib.font_manager as fm
import streamlit as st

from app.db.database import SessionLocal
from app.models.empresa_plantilla_word import EmpresaPlantillaWord


TIPOS_PLANTILLA = ["DOBLE_AA", "TRIPLE_AAA"]

DOCUMENTOS_PLANTILLA_DOBLE_AA = [
    "1.PROPUESTA.docx",
    "2.COTIZACION-INICIAL.docx",
    "3.COTIZACION-FINAL.docx",
    "4.CALENDARIO-DE-TRABAJO.docx",
    "5.ENTREGABLE.docx",
    "6.RESUMEN-EJECUTIVO.docx",
    "7.ACUSE.docx",
]

DOCUMENTOS_PLANTILLA_TRIPLE_AAA = [
    "1. Seguimiento.docx",
    "2.Calendario.docx",
    "3.Entregable.docx",
    "4.Acuse.docx",
]


def obtener_documentos_por_tipo(tipo_plantilla):
    if tipo_plantilla == "TRIPLE_AAA":
        return DOCUMENTOS_PLANTILLA_TRIPLE_AAA

    return DOCUMENTOS_PLANTILLA_DOBLE_AA


def limpiar_nombre_carpeta(nombre):
    nombre = nombre.lower().strip()
    nombre = re.sub(r"\s+", "_", nombre)
    nombre = re.sub(r"[^a-z0-9áéíóúñü_-]", "", nombre)
    return nombre or "plantilla"


def crear_carpeta_plantilla(nombre_disenio, tipo_plantilla):
    nombre_carpeta = limpiar_nombre_carpeta(nombre_disenio)
    subcarpeta_tipo = "triple_a" if tipo_plantilla == "TRIPLE_AAA" else "doble_a"
    ruta_carpeta = os.path.join("plantillas", subcarpeta_tipo, nombre_carpeta)

    os.makedirs(ruta_carpeta, exist_ok=True)

    return ruta_carpeta.replace("\\", "/")


def guardar_documentos_plantilla(archivos_subidos, ruta_carpeta):
    for nombre_documento, archivo in archivos_subidos.items():
        if archivo is None:
            continue

        ruta_archivo = os.path.join(ruta_carpeta, nombre_documento)

        with open(ruta_archivo, "wb") as f:
            f.write(archivo.getbuffer())


def obtener_fuentes_sistema():
    return sorted(set(f.name for f in fm.fontManager.ttflist))


def listar_plantillas(db):
    return (
        db.query(EmpresaPlantillaWord)
        .order_by(EmpresaPlantillaWord.tipo_plantilla.asc())
        .order_by(EmpresaPlantillaWord.nombre_disenio.asc())
        .all()
    )


def crear_plantilla(db, datos_plantilla):
    plantilla = EmpresaPlantillaWord(**datos_plantilla)
    db.add(plantilla)
    db.commit()
    db.refresh(plantilla)
    return plantilla


def actualizar_plantilla(db, plantilla, datos_plantilla):
    for campo, valor in datos_plantilla.items():
        setattr(plantilla, campo, valor)

    db.commit()
    db.refresh(plantilla)
    return plantilla


def mostrar_estado_documentos(ruta_plantilla, tipo_plantilla):
    if not ruta_plantilla:
        return

    for nombre_documento in obtener_documentos_por_tipo(tipo_plantilla):
        ruta_documento = os.path.join(ruta_plantilla, nombre_documento)

        if os.path.exists(ruta_documento):
            st.caption(f"{nombre_documento} actual: {ruta_documento}")
        else:
            st.caption(f"{nombre_documento} pendiente")


def mostrar_formulario_plantilla(db, plantilla=None, modo_creacion=False):
    st.subheader("Datos generales de la plantilla")

    fuentes = obtener_fuentes_sistema()
    plantilla_id = plantilla.id if plantilla else "nuevo"

    col1, col2, col3 = st.columns(3)

    with col1:
        nombre_disenio = st.text_input(
            "Nombre del diseno",
            value=plantilla.nombre_disenio if plantilla else "",
            key=f"nombre_plantilla_{plantilla_id}",
        )

        tamanio_base = st.number_input(
            "Tamano base",
            min_value=6,
            max_value=80,
            value=plantilla.tamanio_base if plantilla else 11,
            key=f"tamanio_plantilla_{plantilla_id}",
        )

    with col2:
        tipo_actual = plantilla.tipo_plantilla if plantilla else "DOBLE_AA"
        tipo_plantilla = st.selectbox(
            "Tipo de plantilla",
            TIPOS_PLANTILLA,
            index=TIPOS_PLANTILLA.index(tipo_actual)
            if tipo_actual in TIPOS_PLANTILLA else 0,
            key=f"tipo_plantilla_{plantilla_id}",
        )

    with col3:
        tipografia_actual = plantilla.tipografia_base if plantilla else None
        tipografia_base = st.selectbox(
            "Tipografia base",
            fuentes,
            index=fuentes.index(tipografia_actual)
            if tipografia_actual in fuentes
            else 0,
            key=f"tipografia_plantilla_{plantilla_id}",
        )

    ruta_plantilla = crear_carpeta_plantilla(nombre_disenio, tipo_plantilla)

    st.subheader("Documentos Word de la plantilla")

    if plantilla:
        mostrar_estado_documentos(plantilla.plantilla_path, tipo_plantilla)

    archivos_subidos = {}
    documentos_plantilla = obtener_documentos_por_tipo(tipo_plantilla)
    columnas = st.columns(4)

    for i, nombre_documento in enumerate(documentos_plantilla):
        with columnas[i % 4]:
            archivos_subidos[nombre_documento] = st.file_uploader(
                nombre_documento,
                type=["docx"],
                key=f"upload_{nombre_documento}_{plantilla_id}_{tipo_plantilla}",
            )

    if st.button("Guardar plantilla", use_container_width=True, key=f"guardar_plantilla_{plantilla_id}"):
        if not nombre_disenio.strip():
            st.error("Debes escribir el nombre del diseno.")
            return

        ruta_plantilla = crear_carpeta_plantilla(nombre_disenio, tipo_plantilla)

        datos_plantilla = {
            "nombre_disenio": nombre_disenio.strip(),
            "tipo_plantilla": tipo_plantilla,
            "tipografia_base": tipografia_base,
            "tamanio_base": int(tamanio_base),
            "plantilla_path": ruta_plantilla,
        }

        if plantilla:
            plantilla_guardada = actualizar_plantilla(db, plantilla, datos_plantilla)
        else:
            plantilla_guardada = crear_plantilla(db, datos_plantilla)

        guardar_documentos_plantilla(archivos_subidos, ruta_plantilla)

        st.session_state["plantilla_word_id"] = plantilla_guardada.id
        st.success("Plantilla guardada correctamente.")
        st.rerun()


def mostrar_modulo_plantilla_word():
    st.title("Catalogo de plantillas Word")

    db = SessionLocal()

    try:
        opcion = st.radio(
            "Seleccione la accion",
            [
                "Crear plantilla",
                "Editar plantilla existente",
            ],
            horizontal=True,
        )

        if opcion == "Crear plantilla":
            mostrar_formulario_plantilla(
                db=db,
                plantilla=None,
                modo_creacion=True,
            )
            return

        plantillas = listar_plantillas(db)

        if not plantillas:
            st.info("Todavia no hay plantillas registradas.")
            return

        plantilla = st.selectbox(
            "Selecciona una plantilla",
            plantillas,
            format_func=lambda p: f"{p.nombre_disenio} ({p.tipo_plantilla})",
        )

        mostrar_formulario_plantilla(
            db=db,
            plantilla=plantilla,
            modo_creacion=False,
        )

    finally:
        db.close()
