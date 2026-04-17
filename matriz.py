import streamlit as st
from xml_cfdi import mostrar_modulo_xml
from rendimiento import mostrar_modulo_rendimiento
from bitacora import mostrar_modulo_bitacora

st.set_page_config(page_title="Herramientas de Excel y XML", layout="wide")

st.sidebar.title("Menú")
opcion = st.sidebar.radio(
    "¿Qué quieres hacer?",
    [
        "XML CFDI a Excel",
        "Tablas de rendimiento",
        "Bitácora de asistencia"
    ]
)

st.sidebar.markdown("---")
st.sidebar.write("Selecciona una opción para trabajar.")

if opcion == "XML CFDI a Excel":
    mostrar_modulo_xml()
elif opcion == "Tablas de rendimiento":
    mostrar_modulo_rendimiento()
elif opcion == "Bitácora de asistencia":
    mostrar_modulo_bitacora()
