import streamlit as st
from xml_cfdi import mostrar_modulo_xml
from rendimiento import mostrar_modulo_rendimiento
from bitacora import mostrar_modulo_bitacora
from comparador_imss import mostrar_modulo_comparador_imss
from propuesta_ia import mostrar_modulo_propuesta_ia
from propuesta_excel import mostrar_modulo_propuesta_excel
from cotizacion_final import mostrar_modulo_cotizacion_final
from ppp import mostrar_modulo_acuse
from pppp import mostrar_modulo_resumen

st.set_page_config(page_title="Herramientas de Excel y XML", layout="wide")

st.sidebar.title("Menú")
opcion = st.sidebar.radio(
    "¿Qué quieres hacer?",
    [
        "XML CFDI a Excel",
        "Tablas de rendimiento",
        "Bitácora de asistencia",
        "Comparar IMSS vs Reporte",
        #"Propuesta con IA",
        "Extraer propuesta Excel",
        "Cotización final",
        "Acuse",
        "Resumen"
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
elif opcion == "Comparar IMSS vs Reporte":
    mostrar_modulo_comparador_imss()
#elif opcion == "Propuesta con IA":
    #mostrar_modulo_propuesta_ia()
elif opcion == "Extraer propuesta Excel":
    mostrar_modulo_propuesta_excel()
elif opcion == "Cotización final":
    mostrar_modulo_cotizacion_final()
elif opcion == "Acuse":
    mostrar_modulo_acuse()
elif opcion == "Resumen":
    mostrar_modulo_resumen()