"""Tab 4 · Participación de mercado — ¿Qué rebanada tenemos?"""
import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src import ui

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿Qué rebanada tenemos?")

    with open(DATA_DIR / "participacion.json", encoding="utf-8") as f:
        participacion = json.load(f)

    st.warning(participacion["advertencia"], icon="⚠️")

    st.markdown("**Participación por plaza**")
    cols = st.columns(len(participacion["por_plaza"]))
    for col, fila in zip(cols, participacion["por_plaza"]):
        with col:
            matricula = fila["matricula_segmento_municipal"]
            valor = f"Año {fila['anio_dato']}" if fila["anio_dato"] else "Por llenar"
            ui.tarjeta_semaforo(
                f"{fila['plaza']} {fila.get('marca', '')}",
                ui.fmt(matricula),
                "#F9A825" if matricula is None else "#2E7D32",
                f"{valor} · {fila.get('nota', '')}",
            )

    st.divider()

    st.markdown("**Alumnos CEA — serie histórica**")
    df_serie = pd.DataFrame(participacion["serie_cea"])
    df_serie["alumnos_altura"] = df_serie["alumnos"].fillna(0)
    df_serie["alumnos_texto"] = df_serie["alumnos"].apply(ui.fmt)
    fig = px.bar(df_serie, x="anio", y="alumnos_altura", text="alumnos_texto")
    fig.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10), xaxis_title="", yaxis_title="Alumnos")
    st.plotly_chart(fig, use_container_width=True)
    for _, r in df_serie.iterrows():
        if pd.isna(r["alumnos"]):
            st.caption(f"{r['marca']} {int(r['anio'])}: {r.get('nota', 'sin dato')}")

    st.divider()
    st.caption(f"Fecha de corte: {participacion['fecha_corte']}")
