import os
import re

import matplotlib.font_manager as fm
import streamlit as st

from app.db.database import SessionLocal
from app.models.empresa_plantilla_word import EmpresaPlantillaWord


TIPO_PLANTILLA_AAA = "TRIPLE_AAA"

DOCUMENTOS_PLANTILLA_AAA = {
    "seguimiento": "1. Seguimiento.docx",
    "calendario": "2.Calendario.docx",
    "entregable": "3.Entregable.docx",
    "acuse": "4.Acuse.docx",
}


def limpiar_nombre_carpeta(nombre):
    nombre = nombre.lower().strip()
    nombre = re.sub(r"\s+", "_", nombre)
    nombre = re.sub(r"[^a-z0-9áéíóúñü_-]", "", nombre)
    return nombre or "plantilla_aaa"


def crear_carpeta_plantilla_aaa(nombre_disenio):
    nombre_carpeta = limpiar_nombre_carpeta(nombre_disenio)
    ruta_carpeta = os.path.join("plantillas", "triple_a", nombre_carpeta)
    os.makedirs(ruta_carpeta, exist_ok=True)
    return ruta_carpeta.replace("\\", "/")


def obtener_fuentes_sistema():
    return sorted(set(f.name for f in fm.fontManager.ttflist))


def listar_plantillas_triple_a(db):
    return (
        db.query(EmpresaPlantillaWord)
        .filter(EmpresaPlantillaWord.tipo_plantilla == TIPO_PLANTILLA_AAA)
        .order_by(EmpresaPlantillaWord.nombre_disenio.asc())
        .all()
    )


def crear_plantilla_triple_a(db, datos_plantilla):
    plantilla = EmpresaPlantillaWord(**datos_plantilla)
    db.add(plantilla)
    db.commit()
    db.refresh(plantilla)
    return plantilla


def actualizar_plantilla_triple_a(db, plantilla, datos_plantilla):
    for campo, valor in datos_plantilla.items():
        setattr(plantilla, campo, valor)

    db.commit()
    db.refresh(plantilla)
    return plantilla


def guardar_documentos_plantilla_aaa(archivos_subidos, ruta_carpeta):
    os.makedirs(ruta_carpeta, exist_ok=True)

    for clave, nombre_documento in DOCUMENTOS_PLANTILLA_AAA.items():
        archivo = archivos_subidos.get(clave)

        if archivo is None:
            continue

        ruta_archivo = os.path.join(ruta_carpeta, nombre_documento)

        with open(ruta_archivo, "wb") as f:
            f.write(archivo.getbuffer())


def obtener_rutas_documentos_aaa(ruta_carpeta):
    return {
        clave: os.path.join(ruta_carpeta, nombre_documento).replace("\\", "/")
        for clave, nombre_documento in DOCUMENTOS_PLANTILLA_AAA.items()
    }


def mostrar_estado_documentos(ruta_carpeta):
    rutas = obtener_rutas_documentos_aaa(ruta_carpeta)

    for clave, ruta in rutas.items():
        if os.path.exists(ruta):
            st.caption(f"{DOCUMENTOS_PLANTILLA_AAA[clave]} actual: {ruta}")
        else:
            st.caption(f"{DOCUMENTOS_PLANTILLA_AAA[clave]} pendiente")


def mostrar_formulario_plantilla_aaa(db, plantilla=None):
    st.subheader("Datos generales de la plantilla AAA")

    fuentes = obtener_fuentes_sistema()
    plantilla_id = plantilla.id if plantilla else "nueva"

    col1, col2 = st.columns(2)

    with col1:
        nombre_disenio = st.text_input(
            "Nombre del diseno",
            value=plantilla.nombre_disenio if plantilla else "",
            key=f"nombre_plantilla_aaa_{plantilla_id}",
        )

        tamanio_base = st.number_input(
            "Tamano base",
            min_value=6,
            max_value=80,
            value=int(plantilla.tamanio_base) if plantilla else 11,
            key=f"tamanio_plantilla_aaa_{plantilla_id}",
        )

    with col2:
        tipografia_actual = plantilla.tipografia_base if plantilla else None
        tipografia_base = st.selectbox(
            "Tipografia base",
            fuentes,
            index=fuentes.index(tipografia_actual)
            if tipografia_actual in fuentes
            else 0,
            key=f"tipografia_plantilla_aaa_{plantilla_id}",
        )

    ruta_plantilla = (
        plantilla.plantilla_path
        if plantilla and plantilla.plantilla_path
        else crear_carpeta_plantilla_aaa(nombre_disenio)
    )

    st.subheader("Documentos Word de la plantilla AAA")

    if plantilla:
        mostrar_estado_documentos(ruta_plantilla)

    archivos_subidos = {}
    columnas = st.columns(4)

    for i, (clave, nombre_documento) in enumerate(DOCUMENTOS_PLANTILLA_AAA.items()):
        with columnas[i % 4]:
            archivos_subidos[clave] = st.file_uploader(
                nombre_documento,
                type=["docx"],
                key=f"upload_aaa_{clave}_{plantilla_id}",
            )

    if st.button(
        "Guardar plantilla AAA",
        use_container_width=True,
        key=f"guardar_plantilla_aaa_{plantilla_id}",
    ):
        if not nombre_disenio.strip():
            st.error("Debes escribir el nombre del diseno.")
            return

        ruta_plantilla = crear_carpeta_plantilla_aaa(nombre_disenio)

        datos_plantilla = {
            "nombre_disenio": nombre_disenio.strip(),
            "tipo_plantilla": TIPO_PLANTILLA_AAA,
            "tipografia_base": tipografia_base,
            "tamanio_base": int(tamanio_base),
            "plantilla_path": ruta_plantilla,
        }

        if plantilla:
            actualizar_plantilla_triple_a(db, plantilla, datos_plantilla)
        else:
            crear_plantilla_triple_a(db, datos_plantilla)

        guardar_documentos_plantilla_aaa(archivos_subidos, ruta_plantilla)

        st.success("Plantilla AAA guardada correctamente.")
        st.rerun()


def mostrar_modulo_plantilla_triple_a():
    st.title("Catalogo de plantillas AAA")

    db = SessionLocal()

    try:
        opcion = st.radio(
            "Seleccione la accion",
            [
                "Crear plantilla AAA",
                "Editar plantilla AAA existente",
            ],
            horizontal=True,
        )

        if opcion == "Crear plantilla AAA":
            mostrar_formulario_plantilla_aaa(db)
            return

        plantillas = listar_plantillas_triple_a(db)

        if not plantillas:
            st.info("Todavia no hay plantillas AAA registradas.")
            return

        plantilla = st.selectbox(
            "Selecciona una plantilla AAA",
            plantillas,
            format_func=lambda p: p.nombre_disenio,
        )

        mostrar_formulario_plantilla_aaa(db, plantilla)

    finally:
        db.close()
