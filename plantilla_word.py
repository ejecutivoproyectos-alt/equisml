import streamlit as st
import matplotlib.font_manager as fm

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord
import os
import re


DOCUMENTOS_PLANTILLA = [
    "1.PROPUESTA.docx",
    "2.COTIZACION-INICIAL.docx",
    "3.COTIZACION-FINAL.docx",
    "4.CALENDARIO-DE-TRABAJO.docx",
    "5.ENTREGABLE.docx",
    "6.RESUMEN-EJECUTIVO.docx",
    "7.ACUSE.docx",
]


def limpiar_nombre_carpeta(nombre):
    nombre = nombre.lower().strip()
    nombre = re.sub(r"\s+", "_", nombre)
    nombre = re.sub(r"[^a-z0-9áéíóúñü_-]", "", nombre)
    return nombre


def crear_carpeta_plantilla(nombre_disenio):
    nombre_carpeta = limpiar_nombre_carpeta(nombre_disenio)

    carpeta_padre = "plantillas"
    ruta_carpeta = os.path.join(carpeta_padre, nombre_carpeta)

    os.makedirs(ruta_carpeta, exist_ok=True)

    return ruta_carpeta.replace("\\", "/")


def guardar_documentos_plantilla(archivos_subidos, ruta_carpeta):
    for nombre_documento, archivo in archivos_subidos.items():
        if archivo is not None:
            ruta_archivo = os.path.join(ruta_carpeta, nombre_documento)

            with open(ruta_archivo, "wb") as f:
                f.write(archivo.getbuffer())


def obtener_fuentes_sistema():
    return sorted(set(f.name for f in fm.fontManager.ttflist))


def listar_plantillas(db):
    return (
        db.query(EmpresaPlantillaWord)
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


def mostrar_formulario_plantilla(
    db,
    plantilla=None,
    modo_creacion=False
):
    st.subheader("Datos generales de la plantilla")

    fuentes = obtener_fuentes_sistema()

    col1, col2 = st.columns(2)

    with col1:
        nombre_disenio = st.text_input(
            "Nombre del diseño",
            value=plantilla.nombre_disenio if plantilla else ""
        )

        tamanio_base = st.number_input(
            "Tamaño base",
            min_value=6,
            max_value=80,
            value=plantilla.tamanio_base if plantilla else 11
        )

    with col2:
        tipografia_base = st.selectbox(
            "Tipografía base",
            fuentes,
            index=fuentes.index(plantilla.tipografia_base)
            if plantilla and plantilla.tipografia_base in fuentes
            else 0
        )

    ruta_plantilla = crear_carpeta_plantilla(nombre_disenio)

    st.subheader("Documentos Word de la plantilla")

    archivos_subidos = {}

    columnas = st.columns(4)

    for i, nombre_documento in enumerate(DOCUMENTOS_PLANTILLA):
        with columnas[i % 4]:
            archivos_subidos[nombre_documento] = st.file_uploader(
                nombre_documento,
                type=["docx"],
                key=f"upload_{nombre_documento}_{plantilla.id if plantilla else 'nuevo'}"
            )

    if st.button("Guardar plantilla", use_container_width=True):
        datos_plantilla = {
            "nombre_disenio": nombre_disenio,
            "tipografia_base": tipografia_base,
            "tamanio_base": tamanio_base,
            "plantilla_path": ruta_plantilla,
        }

        if plantilla:
            plantilla_guardada = actualizar_plantilla(
                db,
                plantilla,
                datos_plantilla
            )
        else:
            plantilla_guardada = crear_plantilla(
                db,
                datos_plantilla
            )

        guardar_documentos_plantilla(
            archivos_subidos,
            ruta_plantilla
        )
        st.session_state["plantilla_word_id"] = plantilla_guardada.id
        st.success("Plantilla guardada correctamente.")
        st.rerun()


def mostrar_modulo_plantilla_word():
    st.title("Catálogo de plantillas Word")

    db = SessionLocal()

    try:
        opcion = st.radio(
            "Seleccione la acción",
            [
                "Crear plantilla",
                "Editar plantilla existente"
            ],
            horizontal=True
        )

        if opcion == "Crear plantilla":
            mostrar_formulario_plantilla(
                db=db,
                plantilla=None,
                modo_creacion=True
            )

        else:
            plantillas = listar_plantillas(db)

            if not plantillas:
                st.info("Todavía no hay plantillas registradas.")
                return

            plantilla = st.selectbox(
                "Selecciona una plantilla",
                plantillas,
                format_func=lambda p: p.nombre_disenio
            )

            mostrar_formulario_plantilla(
                db=db,
                plantilla=plantilla,
                modo_creacion=False
            )

    finally:
        db.close()