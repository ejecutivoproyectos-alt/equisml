import streamlit as st
import os
import re

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord
from app.models.empresa_plantilla_asignacion import EmpresaPlantillaAsignacion
from app.models.empresa_estilo_word import EmpresaEstiloWord
import matplotlib.font_manager as fm

def limpiar_nombre_archivo(nombre):
    nombre = nombre.lower().strip()
    nombre = re.sub(r"\s+", "_", nombre)
    nombre = re.sub(r"[^a-z0-9áéíóúñü_-]", "", nombre)
    return nombre


def guardar_membrete_empresa(archivo_membrete, nombre_empresa, tipo_empresa):
    if archivo_membrete is None:
        return None

    subcarpeta_tipo = "triple_a" if tipo_empresa == "TRIPLE_AAA" else "doble_a"
    carpeta_membretes = os.path.join("membretes", subcarpeta_tipo)
    os.makedirs(carpeta_membretes, exist_ok=True)

    nombre_limpio = limpiar_nombre_archivo(nombre_empresa)
    ruta_archivo = os.path.join(carpeta_membretes, f"{nombre_limpio}.docx")

    with open(ruta_archivo, "wb") as f:
        f.write(archivo_membrete.getbuffer())

    return ruta_archivo.replace("\\", "/")


def listar_empresas(db):
    return (
        db.query(Empresa)
        .order_by(Empresa.nombre.asc())
        .all()
    )


def listar_plantillas(db, tipo_plantilla=None):
    query = db.query(EmpresaPlantillaWord)

    if tipo_plantilla:
        query = query.filter(EmpresaPlantillaWord.tipo_plantilla == tipo_plantilla)

    return query.order_by(EmpresaPlantillaWord.nombre_disenio.asc()).all()


def obtener_asignacion_empresa(db, empresa_id):
    return (
        db.query(EmpresaPlantillaAsignacion)
        .filter(EmpresaPlantillaAsignacion.empresa_externa_id == empresa_id)
        .first()
    )


def guardar_asignacion_plantilla(db, empresa_id, plantilla_id=None, membrete_path=None):
    asignacion = obtener_asignacion_empresa(db, empresa_id)

    if asignacion:
        asignacion.plantilla_id = plantilla_id
        asignacion.activo = True

        if membrete_path:
            asignacion.membrete_path = membrete_path

    else:
        asignacion = EmpresaPlantillaAsignacion(
            empresa_externa_id=empresa_id,
            plantilla_id=plantilla_id,
            membrete_path=membrete_path,
            activo=True
        )
        db.add(asignacion)

    db.commit()
    db.refresh(asignacion)

    return asignacion


def obtener_fuentes_sistema():
    return sorted(set(f.name for f in fm.fontManager.ttflist))


def normalizar_clave_estilo(texto):
    texto = texto.lower().strip()
    texto = re.sub(r"\s+", "_", texto)
    texto = re.sub(r"[^a-z0-9áéíóúñü_]", "", texto)
    return texto


def obtener_estilos_empresa(db, empresa_id):
    return (
        db.query(EmpresaEstiloWord)
        .filter(EmpresaEstiloWord.empresa_id == empresa_id)
        .order_by(EmpresaEstiloWord.id.asc())
        .all()
    )


def guardar_estilo_empresa(db, empresa_id, datos_estilo):
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
        empresa_id=empresa_id,
        **datos_estilo
    )

    db.add(estilo)
    db.commit()
    db.refresh(estilo)

    return estilo


def eliminar_estilo_empresa(db, estilo_id):
    estilo = (
        db.query(EmpresaEstiloWord)
        .filter(EmpresaEstiloWord.id == estilo_id)
        .first()
    )

    if estilo:
        db.delete(estilo)
        db.commit()


def mostrar_formulario_estilos_empresa(db, empresa_id):
    st.markdown("**Estilos Word de la empresa**")

    estilos = obtener_estilos_empresa(db, empresa_id)
    fuentes = obtener_fuentes_sistema()

    for estilo in estilos:
        with st.expander(f"Editar estilo: {estilo.clave_estilo}", expanded=False):
            clave_estilo = st.text_input(
                "Clave estilo",
                value=estilo.clave_estilo,
                key=f"clave_estilo_empresa_{estilo.id}"
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                tipografia = st.selectbox(
                    "Tipografía",
                    fuentes,
                    index=fuentes.index(estilo.tipografia)
                    if estilo.tipografia in fuentes else 0,
                    key=f"tipografia_empresa_{estilo.id}"
                )

            with col2:
                tamanio_letra = st.number_input(
                    "Tamaño letra",
                    min_value=6,
                    max_value=80,
                    value=estilo.tamanio_letra,
                    key=f"tamanio_empresa_{estilo.id}"
                )

            with col3:
                alineacion = st.selectbox(
                    "Alineación",
                    ["left", "center", "right", "justify"],
                    index=["left", "center", "right", "justify"].index(estilo.alineacion)
                    if estilo.alineacion in ["left", "center", "right", "justify"] else 0,
                    key=f"alineacion_empresa_{estilo.id}"
                )

            negrita = st.checkbox(
                "Negrita",
                value=estilo.negrita,
                key=f"negrita_empresa_{estilo.id}"
            )

            cursiva = st.checkbox(
                "Cursiva",
                value=estilo.cursiva,
                key=f"cursiva_empresa_{estilo.id}"
            )

            col_guardar, col_eliminar = st.columns(2)

            with col_guardar:
                if st.button("Actualizar estilo", key=f"actualizar_estilo_empresa_{estilo.id}"):
                    guardar_estilo_empresa(
                        db,
                        empresa_id,
                        {
                            "id": estilo.id,
                            "clave_estilo": normalizar_clave_estilo(clave_estilo),
                            "tipografia": tipografia,
                            "tamanio_letra": tamanio_letra,
                            "negrita": negrita,
                            "cursiva": cursiva,
                            "alineacion": alineacion,
                        }
                    )

                    st.success("Estilo actualizado correctamente.")
                    st.rerun()

            with col_eliminar:
                if st.button("Eliminar estilo", key=f"eliminar_estilo_empresa_{estilo.id}"):
                    eliminar_estilo_empresa(db, estilo.id)
                    st.warning("Estilo eliminado.")
                    st.rerun()

    st.markdown("**Agregar nuevo estilo**")

    nueva_clave = st.text_input(
        "Nueva clave de estilo",
        key=f"nueva_clave_estilo_empresa_{empresa_id}"
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        nueva_tipografia = st.selectbox(
            "Tipografía",
            fuentes,
            key=f"nueva_tipografia_empresa_{empresa_id}"
        )

    with col2:
        nuevo_tamanio = st.number_input(
            "Tamaño",
            min_value=6,
            max_value=80,
            value=11,
            key=f"nuevo_tamanio_empresa_{empresa_id}"
        )

    with col3:
        nueva_alineacion = st.selectbox(
            "Alineación",
            ["left", "center", "right", "justify"],
            key=f"nueva_alineacion_empresa_{empresa_id}"
        )

    nueva_negrita = st.checkbox(
        "Negrita",
        value=False,
        key=f"nueva_negrita_empresa_{empresa_id}"
    )

    nueva_cursiva = st.checkbox(
        "Cursiva",
        value=False,
        key=f"nueva_cursiva_empresa_{empresa_id}"
    )

    if st.button("Agregar estilo", key=f"agregar_estilo_empresa_{empresa_id}", use_container_width=True):
        if not nueva_clave.strip():
            st.error("Debes escribir una clave de estilo.")
            return

        guardar_estilo_empresa(
            db,
            empresa_id,
            {
                "clave_estilo": normalizar_clave_estilo(nueva_clave),
                "tipografia": nueva_tipografia,
                "tamanio_letra": nuevo_tamanio,
                "negrita": nueva_negrita,
                "cursiva": nueva_cursiva,
                "alineacion": nueva_alineacion,
            }
        )

        st.success("Estilo agregado correctamente.")
        st.rerun()


def capturar_estilos_nueva_empresa():
    fuentes = obtener_fuentes_sistema()
    estilos = []

    if "estilos_nueva_empresa" not in st.session_state:
        st.session_state["estilos_nueva_empresa"] = [0]

    if st.button("➕ Agregar estilo", key="agregar_estilo_nueva_empresa"):
        nuevo_id = max(st.session_state["estilos_nueva_empresa"]) + 1
        st.session_state["estilos_nueva_empresa"].append(nuevo_id)
        st.rerun()

    for estilo_id in st.session_state["estilos_nueva_empresa"]:
        with st.expander(f"Estilo {estilo_id + 1}", expanded=True):
            clave = st.text_input(
                "Nombre del estilo",
                key=f"clave_nueva_empresa_{estilo_id}",
                placeholder="Ejemplo: titulo 1"
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                tipografia = st.selectbox(
                    "Tipografía",
                    fuentes,
                    key=f"tipografia_nueva_empresa_{estilo_id}"
                )

            with col2:
                tamanio = st.number_input(
                    "Tamaño letra",
                    min_value=6,
                    max_value=80,
                    value=11,
                    key=f"tamanio_nueva_empresa_{estilo_id}"
                )

            with col3:
                alineacion = st.selectbox(
                    "Alineación",
                    ["left", "center", "right", "justify"],
                    key=f"alineacion_nueva_empresa_{estilo_id}"
                )

            negrita = st.checkbox(
                "Negrita",
                key=f"negrita_nueva_empresa_{estilo_id}"
            )

            cursiva = st.checkbox(
                "Cursiva",
                key=f"cursiva_nueva_empresa_{estilo_id}"
            )

            if clave.strip():
                estilos.append({
                    "clave_estilo": normalizar_clave_estilo(clave),
                    "tipografia": tipografia,
                    "tamanio_letra": tamanio,
                    "negrita": negrita,
                    "cursiva": cursiva,
                    "alineacion": alineacion,
                })

    return estilos


def mostrar_modulo_empresas():
    st.title("Empresas")

    db = SessionLocal()

    try:
        st.subheader("Agregar empresa")

        col_nombre, col_razon, col_tipo = st.columns(3)

        with col_nombre:
            nombre = st.text_input("Nombre comercial")

        with col_razon:
            razon_social = st.text_input("Razón social")

        with col_tipo:
            tipo_empresa = st.selectbox(
                "Tipo de empresa",
                ["DOBLE_AA", "TRIPLE_AAA"],
            )

        st.subheader("Paleta de colores")

        col_color1, col_color2, col_color3 = st.columns(3)

        with col_color1:
            color_primario = st.color_picker("Color primario", "#000000")

        with col_color2:
            color_secundario = st.color_picker("Color secundario", "#FFFFFF")

        with col_color3:
            color_acento = st.color_picker("Color acento", "#FFFFFF")

        plantillas = listar_plantillas(db, tipo_empresa)

        opciones_plantilla = [None] + plantillas

        col_membrete, col_plantilla = st.columns(2)

        with col_membrete:
            archivo_membrete = st.file_uploader(
                "Membrete Word de la empresa",
                type=["docx"],
                key="membrete_nueva_empresa"
            )

        with col_plantilla:
            plantilla_seleccionada = st.selectbox(
                "Plantilla Word asignada opcional",
                opciones_plantilla,
                format_func=lambda p: "Sin plantilla asignada" if p is None else p.nombre_disenio
            )

        st.markdown("---")
        st.subheader("Estilos Word de la empresa")

        estilos_nueva_empresa = capturar_estilos_nueva_empresa()

        if st.button("Guardar empresa", use_container_width=True):
            if not nombre.strip():
                st.error("Debes escribir el nombre de la empresa.")
                return

            empresa = Empresa(
                nombre=nombre.strip(),
                razon_social=razon_social.strip() if razon_social else None,
                tipo_empresa=tipo_empresa,
                color_primario=color_primario,
                color_secundario=color_secundario,
                color_acento=color_acento,
            )

            db.add(empresa)
            db.commit()
            db.refresh(empresa)

            membrete_path = guardar_membrete_empresa(
                archivo_membrete,
                nombre,
                tipo_empresa
            )

            guardar_asignacion_plantilla(
                db=db,
                empresa_id=empresa.id,
                plantilla_id=plantilla_seleccionada.id if plantilla_seleccionada else None,
                membrete_path=membrete_path
            )

            for estilo in estilos_nueva_empresa:
                guardar_estilo_empresa(
                    db=db,
                    empresa_id=empresa.id,
                    datos_estilo=estilo
                )

            st.success("Empresa guardada correctamente.")
            st.rerun()

        st.markdown("---")
        st.subheader("Empresas registradas")

        empresas = listar_empresas(db)

        for i in range(0, len(empresas), 2):
            col1, col2 = st.columns(2)

            fila_empresas = empresas[i:i + 2]

            for col, empresa in zip([col1, col2], fila_empresas):
                with col:
                    asignacion = obtener_asignacion_empresa(db, empresa.id)

                    with st.expander(empresa.nombre):
                        nuevo_nombre = st.text_input(
                            "Nombre comercial",
                            value=empresa.nombre,
                            key=f"nombre_empresa_{empresa.id}"
                        )

                        nueva_razon_social = st.text_input(
                            "Razón social",
                            value=empresa.razon_social or "",
                            key=f"razon_social_empresa_{empresa.id}"
                        )

                        nuevo_tipo_empresa = st.selectbox(
                            "Tipo de empresa",
                            ["DOBLE_AA", "TRIPLE_AAA"],
                            index=["DOBLE_AA", "TRIPLE_AAA"].index(empresa.tipo_empresa)
                            if empresa.tipo_empresa in ["DOBLE_AA", "TRIPLE_AAA"] else 0,
                            key=f"tipo_empresa_{empresa.id}"
                        )

                        st.markdown("**Paleta de colores**")

                        col_color1, col_color2, col_color3 = st.columns(3)

                        with col_color1:
                            nuevo_color_primario = st.color_picker(
                                "Color primario",
                                value=empresa.color_primario or "#000000",
                                key=f"color_primario_empresa_{empresa.id}"
                            )

                        with col_color2:
                            nuevo_color_secundario = st.color_picker(
                                "Color secundario",
                                value=empresa.color_secundario or "#FFFFFF",
                                key=f"color_secundario_empresa_{empresa.id}"
                            )

                        with col_color3:
                            nuevo_color_acento = st.color_picker(
                                "Color acento",
                                value=empresa.color_acento or "#FFFFFF",
                                key=f"color_acento_empresa_{empresa.id}"
                            )

                        plantillas_empresa = listar_plantillas(db, nuevo_tipo_empresa)
                        opciones_plantilla_empresa = [None] + plantillas_empresa
                        plantilla_actual_id = asignacion.plantilla_id if asignacion else None

                        index_actual = 0

                        for i, plantilla in enumerate(opciones_plantilla_empresa):
                            if plantilla and plantilla.id == plantilla_actual_id:
                                index_actual = i
                                break

                        nueva_plantilla = st.selectbox(
                            "Plantilla Word asignada",
                            opciones_plantilla_empresa,
                            index=index_actual,
                            format_func=lambda p: "Sin plantilla asignada" if p is None else p.nombre_disenio,
                            key=f"plantilla_empresa_{empresa.id}"
                        )

                        if asignacion and asignacion.membrete_path:
                            st.caption(f"Membrete actual: {asignacion.membrete_path}")

                        nuevo_membrete = st.file_uploader(
                            "Actualizar membrete Word",
                            type=["docx"],
                            key=f"membrete_empresa_{empresa.id}"
                        )

                        if st.button("Actualizar empresa", key=f"actualizar_empresa_{empresa.id}"):
                            empresa.nombre = nuevo_nombre.strip()
                            empresa.razon_social = nueva_razon_social.strip() if nueva_razon_social else None
                            empresa.tipo_empresa = nuevo_tipo_empresa
                            empresa.color_primario = nuevo_color_primario
                            empresa.color_secundario = nuevo_color_secundario
                            empresa.color_acento = nuevo_color_acento

                            membrete_path = guardar_membrete_empresa(
                                nuevo_membrete,
                                nuevo_nombre,
                                nuevo_tipo_empresa
                            )

                            guardar_asignacion_plantilla(
                                db=db,
                                empresa_id=empresa.id,
                                plantilla_id=nueva_plantilla.id if nueva_plantilla else None,
                                membrete_path=membrete_path
                            )

                            db.commit()

                            st.success("Empresa actualizada correctamente.")
                            st.rerun()

                        st.markdown("---")
                        mostrar_formulario_estilos_empresa(db, empresa.id)
    finally:
        db.close()
