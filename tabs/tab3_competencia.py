"""Tab 3 · Competencia — ¿Por qué CEA y no la de enfrente?"""
import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src import ui
from src.constants import PLAZAS

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿Por qué CEA y no la de enfrente?")

    with open(DATA_DIR / "competencia.json", encoding="utf-8") as f:
        competencia = json.load(f)

    plaza_filtro = st.selectbox("Plaza", ["Todas"] + PLAZAS)
    competidores = competencia["competidores"]
    if plaza_filtro != "Todas":
        competidores = [
            c for c in competidores
            if plaza_filtro in c["plazas"] or "Nacional" in c["plazas"]
        ]

    cea = competencia["cea"]

    directa = [c for c in competidores if c["tipo_competencia"] == "Directa"]
    indirecta = [c for c in competidores if c["tipo_competencia"] == "Indirecta"]

    st.markdown("**Competencia directa**")
    if directa:
        st.dataframe(
            pd.DataFrame(directa)[["nombre", "mensualidad_lic", "flexibilidad", "amenaza", "razon", "marca"]]
            .rename(columns={"nombre": "Nombre", "mensualidad_lic": "Mensualidad", "flexibilidad": "Flexibilidad (1-5)",
                              "amenaza": "Amenaza", "razon": "Por qué compite", "marca": ""}),
            use_container_width=True, hide_index=True,
        )
    else:
        st.caption("Sin competencia directa registrada para esta plaza.")

    st.markdown("**Competencia indirecta**")
    if indirecta:
        st.dataframe(
            pd.DataFrame(indirecta)[["nombre", "mensualidad_lic", "flexibilidad", "amenaza", "razon", "marca"]]
            .rename(columns={"nombre": "Nombre", "mensualidad_lic": "Mensualidad", "flexibilidad": "Flexibilidad (1-5)",
                              "amenaza": "Amenaza", "razon": "Por qué compite", "marca": ""}),
            use_container_width=True, hide_index=True,
        )
    else:
        st.caption("Sin competencia indirecta registrada para esta plaza.")

    st.divider()

    st.markdown("**Precio vs flexibilidad/servicio**")
    fig = go.Figure()
    if directa:
        df_d = pd.DataFrame(directa)
        fig.add_trace(go.Scatter(
            x=df_d["mensualidad_lic"], y=df_d["flexibilidad"], mode="markers+text",
            text=df_d["nombre"], textposition="top center", name="Directa",
            marker=dict(size=14, symbol="circle", color="#C62828"),
        ))
    if indirecta:
        df_i = pd.DataFrame(indirecta)
        fig.add_trace(go.Scatter(
            x=df_i["mensualidad_lic"], y=df_i["flexibilidad"], mode="markers+text",
            text=df_i["nombre"], textposition="top center", name="Indirecta",
            marker=dict(size=14, symbol="diamond", color="#F9A825"),
        ))
    fig.add_trace(go.Scatter(
        x=[cea["mensualidad_lic"]], y=[cea["flexibilidad"]], mode="markers+text",
        text=["CEA Universidad"], textposition="top center", name="CEA",
        marker=dict(size=20, symbol="star", color="#2E7D32"),
    ))
    fig.update_layout(
        height=420, margin=dict(t=20, b=10, l=10, r=10),
        xaxis_title="Mensualidad (MXN)", yaxis_title="Flexibilidad / servicio (1-5)",
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown(f"**CEA Universidad** — {' · '.join(cea['diferenciadores'])}")

    st.divider()
    st.info(competencia["lectura_estrategica"])
    st.caption(f"Fecha de corte: {competencia['fecha_corte']}")
