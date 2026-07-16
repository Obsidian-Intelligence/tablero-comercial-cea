"""Tablero Comercial CEA — entrypoint: login + navegación entre tabs."""
import streamlit as st

from src import auth, etl
from tabs import (
    tab1_mercado,
    tab2_avance,
    tab3_competencia,
    tab4_participacion,
    tab5_finanzas,
    tab6_okrs,
)

st.set_page_config(
    page_title="Tablero Comercial CEA",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

auth.exigir_login()

with st.sidebar:
    st.markdown("### 🎓 CEA Universidad")
    st.caption("Tablero Comercial")
    if st.button("Cerrar sesión"):
        auth.cerrar_sesion()

st.title("Tablero Comercial CEA")

hojas, es_demo = etl.cargar_datos()
con = etl.construir_duckdb(hojas)

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📈 Mercado", "🎯 Avance vs Meta", "⚔️ Competencia", "🥧 Participación", "💰 Finanzas", "🧭 OKRs y Plan"]
)

with tab1:
    tab1_mercado.render(con, hojas, es_demo)
with tab2:
    tab2_avance.render(con, hojas, es_demo)
with tab3:
    tab3_competencia.render(con, hojas, es_demo)
with tab4:
    tab4_participacion.render(con, hojas, es_demo)
with tab5:
    tab5_finanzas.render(con, hojas, es_demo)
with tab6:
    tab6_okrs.render(con, hojas, es_demo)
