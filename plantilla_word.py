import streamlit as st
import matplotlib.font_manager as fm

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord
from app.models.empresa_estilo_word import EmpresaEstiloWord
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


def obtener_estilos_plantilla(db, plantilla_id):
    return (
        db.query(EmpresaEstiloWord)
        .filter(EmpresaEstiloWord.plantilla_id == plantilla_id)
        .order_by(EmpresaEstiloWord.id.asc())
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


def normalizar_clave_estilo(texto):
    texto = texto.lower().strip()
    texto = re.sub(r"\s+", "_", texto)
    texto = re.sub(r"[^a-z0-9áéíóúñü_]", "", texto)
    return texto


def inicializar_estilos_temporales():
    if "estilos_temporales" not in st.session_state:
        st.session_state["estilos_temporales"] = [0]


def agregar_estilo_temporal():
    nuevo_id = max(st.session_state["estilos_temporales"]) + 1
    st.session_state["estilos_temporales"].append(nuevo_id)


def capturar_estilos_desde_inputs():
    inicializar_estilos_temporales()

    st.subheader("Estilos de la plantilla")

    if st.button("➕ Agregar estilo"):
        agregar_estilo_temporal()
        st.rerun()

    fuentes = obtener_fuentes_sistema()
    estilos = []

    for estilo_id in st.session_state["estilos_temporales"]:
        with st.expander(f"Estilo {estilo_id + 1}", expanded=True):
            clave_estilo_input = st.text_input(
                "Nombre del estilo",
                key=f"clave_estilo_{estilo_id}",
                placeholder="Ejemplo: titulo 1"
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                tipografia = st.selectbox(
                    "Tipografía",
                    fuentes,
                    key=f"tipografia_estilo_{estilo_id}"
                )

            with col2:
                tamanio_letra = st.number_input(
                    "Tamaño letra",
                    min_value=6,
                    max_value=80,
                    value=11,
                    key=f"tamanio_estilo_{estilo_id}"
                )

            with col3:
                alineacion = st.selectbox(
                    "Alineación",
                    ["left", "center", "right", "justify"],
                    key=f"alineacion_estilo_{estilo_id}"
                )

            col4, col5, col6 = st.columns(3)

            with col4:
                color_letra = st.color_picker(
                    "Color letra",
                    "#000000",
                    key=f"color_letra_estilo_{estilo_id}"
                )

            with col5:
                color_fondo = st.color_picker(
                    "Color fondo",
                    "#FFFFFF",
                    key=f"color_fondo_estilo_{estilo_id}"
                )

            with col6:
                negrita = st.checkbox(
                    "Negrita",
                    key=f"negrita_estilo_{estilo_id}"
                )

                cursiva = st.checkbox(
                    "Cursiva",
                    key=f"cursiva_estilo_{estilo_id}"
                )

            if clave_estilo_input.strip():
                estilos.append({
                    "clave_estilo": normalizar_clave_estilo(clave_estilo_input),
                    "tipografia": tipografia,
                    "tamanio_letra": tamanio_letra,
                    "color_letra": color_letra,
                    "color_fondo": color_fondo,
                    "negrita": negrita,
                    "cursiva": cursiva,
                    "alineacion": alineacion,
                })

    return estilos


def guardar_estilo(db, plantilla_id, datos_estilo):
    estilo_id = datos_estilo.get("id")

    if estilo_id:
        estilo = (
            db.query(EmpresaEstiloWord)
            .filter(EmpresaEstiloWord.id == estilo_id)
            .first()
        )

        if estilo:
            for campo, valor in datos_estilo.items():
                if campo != "id":
                    setattr(estilo, campo, valor)

            db.commit()
            db.refresh(estilo)
            return estilo

    datos_estilo.pop("id", None)

    estilo = EmpresaEstiloWord(
        plantilla_id=plantilla_id,
        **datos_estilo
    )

    db.add(estilo)
    db.commit()
    db.refresh(estilo)

    return estilo


def eliminar_estilo(db, estilo_id):
    estilo = (
        db.query(EmpresaEstiloWord)
        .filter(EmpresaEstiloWord.id == estilo_id)
        .first()
    )

    if estilo:
        db.delete(estilo)
        db.commit()


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

    col4, col5, col6, col7 = st.columns(4)

    with col4:
        color_texto_base = st.color_picker(
            "Color texto base",
            value=plantilla.color_texto_base if plantilla else "#000000"
        )

    with col5:
        color_primario = st.color_picker(
            "Color primario",
            value=plantilla.color_primario if plantilla else "#000000"
        )

    with col6:
        color_secundario = st.color_picker(
            "Color secundario",
            value=plantilla.color_secundario if plantilla and plantilla.color_secundario else "#FFFFFF"
        )

    with col7:
        color_acento = st.color_picker(
            "Color acento",
            value=plantilla.color_acento if plantilla and plantilla.color_acento else "#FFFFFF"
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

    estilos_nuevos = []

    if modo_creacion:
        st.markdown("---")
        estilos_nuevos = capturar_estilos_desde_inputs()

    if st.button("Guardar plantilla", use_container_width=True):
        datos_plantilla = {
            "nombre_disenio": nombre_disenio,
            "tipografia_base": tipografia_base,
            "tamanio_base": tamanio_base,
            "color_texto_base": color_texto_base,
            "color_primario": color_primario,
            "color_secundario": color_secundario,
            "color_acento": color_acento,
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

            if estilos_nuevos:
                for estilo in estilos_nuevos:
                    guardar_estilo(
                        db=db,
                        plantilla_id=plantilla_guardada.id,
                        datos_estilo=estilo
                    )

        guardar_documentos_plantilla(
            archivos_subidos,
            ruta_plantilla
        )
        st.session_state["plantilla_word_id"] = plantilla_guardada.id
        st.success("Plantilla guardada correctamente.")
        st.rerun()


def mostrar_formulario_estilos(db, plantilla_id):
    st.subheader("Estilos de la plantilla")

    estilos = obtener_estilos_plantilla(db, plantilla_id)
    fuentes = obtener_fuentes_sistema()

    if estilos:
        for estilo in estilos:
            with st.expander(f"Editar estilo: {estilo.clave_estilo}", expanded=False):
                clave_estilo = st.text_input(
                    "Clave estilo",
                    value=estilo.clave_estilo,
                    key=f"clave_{estilo.id}"
                )

                col1, col2, col3 = st.columns(3)

                with col1:
                    tipografia = st.selectbox(
                        "Tipografía",
                        fuentes,
                        index=fuentes.index(estilo.tipografia)
                        if estilo.tipografia in fuentes
                        else 0,
                        key=f"tipografia_{estilo.id}"
                    )

                with col2:
                    tamanio_letra = st.number_input(
                        "Tamaño letra",
                        min_value=6,
                        max_value=80,
                        value=estilo.tamanio_letra,
                        key=f"tamanio_{estilo.id}"
                    )

                with col3:
                    alineacion = st.selectbox(
                        "Alineación",
                        ["left", "center", "right", "justify"],
                        index=["left", "center", "right", "justify"].index(estilo.alineacion)
                        if estilo.alineacion in ["left", "center", "right", "justify"]
                        else 0,
                        key=f"alineacion_{estilo.id}"
                    )

                col4, col5, col6 = st.columns(3)

                with col4:
                    color_letra = st.color_picker(
                        "Color letra",
                        value=estilo.color_letra,
                        key=f"color_letra_{estilo.id}"
                    )

                with col5:
                    color_fondo = st.color_picker(
                        "Color fondo",
                        value=estilo.color_fondo if estilo.color_fondo else "#FFFFFF",
                        key=f"color_fondo_{estilo.id}"
                    )

                with col6:
                    negrita = st.checkbox(
                        "Negrita",
                        value=estilo.negrita,
                        key=f"negrita_{estilo.id}"
                    )

                    cursiva = st.checkbox(
                        "Cursiva",
                        value=estilo.cursiva,
                        key=f"cursiva_{estilo.id}"
                    )

                col1, col2 = st.columns(2)

                with col1:
                    if st.button("Actualizar estilo", key=f"guardar_estilo_{estilo.id}"):
                        guardar_estilo(
                            db,
                            plantilla_id,
                            {
                                "id": estilo.id,
                                "clave_estilo": clave_estilo,
                                "tipografia": tipografia,
                                "tamanio_letra": tamanio_letra,
                                "color_letra": color_letra,
                                "color_fondo": color_fondo,
                                "negrita": negrita,
                                "cursiva": cursiva,
                                "alineacion": alineacion,
                            }
                        )

                        st.success("Estilo actualizado correctamente.")
                        st.rerun()

                with col2:
                    if st.button("Eliminar estilo", key=f"eliminar_estilo_{estilo.id}"):
                        eliminar_estilo(db, estilo.id)
                        st.warning("Estilo eliminado.")
                        st.rerun()

    st.markdown("---")
    st.subheader("Agregar nuevo estilo")

    nueva_clave = st.text_input(
        "Nueva clave de estilo"
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        nueva_tipografia = st.selectbox(
            "Tipografía",
            fuentes,
            key="nueva_tipografia"
        )

    with col2:
        nuevo_tamanio = st.number_input(
            "Tamaño",
            min_value=6,
            max_value=80,
            value=11,
            key="nuevo_tamanio"
        )

    with col3:
        nueva_alineacion = st.selectbox(
            "Alineación",
            ["left", "center", "right", "justify"],
            key="nueva_alineacion"
        )

    col4, col5, col6 = st.columns(3)

    with col4:
        nuevo_color_letra = st.color_picker(
            "Color letra",
            "#000000",
            key="nuevo_color_letra"
        )

    with col5:
        nuevo_color_fondo = st.color_picker(
            "Color fondo",
            "#FFFFFF",
            key="nuevo_color_fondo"
        )

    with col6:
        nueva_negrita = st.checkbox(
            "Negrita",
            value=False,
            key="nueva_negrita"
        )

        nueva_cursiva = st.checkbox(
            "Cursiva",
            value=False,
            key="nueva_cursiva"
        )

    if st.button("Agregar estilo", use_container_width=True):
        if not nueva_clave.strip():
            st.error("Debes escribir una clave de estilo.")
            return

        guardar_estilo(
            db,
            plantilla_id,
            {
                "clave_estilo": normalizar_clave_estilo(nueva_clave),
                "tipografia": nueva_tipografia,
                "tamanio_letra": nuevo_tamanio,
                "color_letra": nuevo_color_letra,
                "color_fondo": nuevo_color_fondo,
                "negrita": nueva_negrita,
                "cursiva": nueva_cursiva,
                "alineacion": nueva_alineacion,
            }
        )

        st.success("Estilo agregado correctamente.")
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

            st.markdown("---")

            mostrar_formulario_estilos(
                db=db,
                plantilla_id=plantilla.id
            )

    finally:
        db.close()