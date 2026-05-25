import streamlit as st
import os
import re

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord
from app.models.empresa_plantilla_asignacion import EmpresaPlantillaAsignacion


def limpiar_nombre_archivo(nombre):
    nombre = nombre.lower().strip()
    nombre = re.sub(r"\s+", "_", nombre)
    nombre = re.sub(r"[^a-z0-9áéíóúñü_-]", "", nombre)
    return nombre


def guardar_membrete_empresa(archivo_membrete, nombre_empresa):
    if archivo_membrete is None:
        return None

    os.makedirs("membretes", exist_ok=True)

    nombre_limpio = limpiar_nombre_archivo(nombre_empresa)
    ruta_archivo = os.path.join("membretes", f"{nombre_limpio}.docx")

    with open(ruta_archivo, "wb") as f:
        f.write(archivo_membrete.getbuffer())

    return ruta_archivo.replace("\\", "/")


def listar_empresas(db):
    return (
        db.query(Empresa)
        .order_by(Empresa.nombre.asc())
        .all()
    )


def listar_plantillas(db):
    return (
        db.query(EmpresaPlantillaWord)
        .order_by(EmpresaPlantillaWord.nombre_disenio.asc())
        .all()
    )


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


def mostrar_modulo_empresas():
    st.title("Empresas")

    db = SessionLocal()

    try:
        st.subheader("Agregar empresa")

        col_nombre, col_razon = st.columns(2)

        with col_nombre:
            nombre = st.text_input("Nombre comercial")

        with col_razon:
            razon_social = st.text_input("Razón social")

        plantillas = listar_plantillas(db)

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

        if st.button("Guardar empresa", use_container_width=True):
            if not nombre.strip():
                st.error("Debes escribir el nombre de la empresa.")
                return

            empresa = Empresa(
                nombre=nombre.strip(),
                razon_social=razon_social.strip() if razon_social else None
            )

            db.add(empresa)
            db.commit()
            db.refresh(empresa)

            membrete_path = guardar_membrete_empresa(
                archivo_membrete,
                nombre
            )

            guardar_asignacion_plantilla(
                db=db,
                empresa_id=empresa.id,
                plantilla_id=plantilla_seleccionada.id if plantilla_seleccionada else None,
                membrete_path=membrete_path
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

                        plantilla_actual_id = asignacion.plantilla_id if asignacion else None

                        index_actual = 0

                        for i, plantilla in enumerate(opciones_plantilla):
                            if plantilla and plantilla.id == plantilla_actual_id:
                                index_actual = i
                                break

                        nueva_plantilla = st.selectbox(
                            "Plantilla Word asignada",
                            opciones_plantilla,
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

                            membrete_path = guardar_membrete_empresa(
                                nuevo_membrete,
                                nuevo_nombre
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

    finally:
        db.close()