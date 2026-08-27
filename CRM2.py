RUN

import os
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
 
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "produccion.db")
SEED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "seed_marzo.csv")
 
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
MESES_INV = {v.lower(): k for k, v in MESES_ES.items()}
LINEAS = {"al01": "AL-01 (25,000 un)", "al02": "AL-02 (23,000 un)", "al04": "AL-04 (60,000 un)"}
 
st.set_page_config(page_title="Seguimiento de Producción · Planta Tubos", layout="wide", page_icon="🏭")

init_db()
 
# ---------------------------------------------------------------------------
# Sidebar / navegación
# ---------------------------------------------------------------------------
st.sidebar.title("🏭 Planta Tubos")
st.sidebar.caption("Seguimiento de producción")
 
df_all = load_data()
 
if df_all.empty:
    st.sidebar.info("Todavía no hay datos cargados.")
    with st.sidebar.expander("Cargar datos de ejemplo (marzo)"):
        demo_year = st.number_input("Año", value=date.today().year, step=1, key="demo_year")
        if st.button("Cargar marzo de ejemplo"):
            seed_marzo_demo(int(demo_year))
            st.rerun()
 
page = st.sidebar.radio(
    "Navegación",
    ["📊 Dashboard", "➕ Registrar producción", "🗂️ Datos históricos", "📁 Importar Excel"],
)
 
with st.sidebar.expander("⚙️ Opciones avanzadas"):
    if st.button("🗑️ Borrar todos los datos"):
        reset_db()
        st.rerun()
 
st.sidebar.caption(f"Registros totales: {len(df_all)}")