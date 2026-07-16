"""Tab 1 · Mercado — ¿Qué tan grande es el pastel y dónde está?"""
import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from src import ui

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿Qué tan grande es el pastel y dónde está?")

    with open(DATA_DIR / "mercado.json", encoding="utf-8") as f:
        mercado = json.load(f)

    st.markdown("**TAM · SAM · SOM por plaza**")
    st.caption(mercado["tam_sam_som"]["nota_metodologica"])
    st.caption(f"_TAM — {ui.definicion('TAM')}_")
    st.caption(f"_SAM — {ui.definicion('SAM')}_")
    st.caption(f"_SOM — {ui.definicion('SOM')}_")
    cols = st.columns(len(mercado["tam_sam_som"]["por_plaza"]))
    for col, fila in zip(cols, mercado["tam_sam_som"]["por_plaza"]):
        with col:
            ui.tarjeta_semaforo(
                f"{fila['plaza']} {fila.get('marca', '')}",
                f"SOM: {fila['som']}",
                "#2E7D32",
                f"TAM: {ui.fmt(fila['tam'])} · SAM: {ui.fmt(fila['sam'])}",
            )
            if fila.get("nota"):
                st.caption(fila["nota"])

    st.divider()

    st.markdown("**Contexto estatal**")
    df_ctx = pd.DataFrame(mercado["contexto_estatal"])
    df_ctx["valor_mostrado"] = df_ctx.apply(
        lambda r: r["valor_pct"] if pd.notna(r.get("valor_pct")) else r.get("valor"), axis=1
    )
    fig = px.bar(df_ctx, x="indicador", y="valor_mostrado", text="valor_mostrado")
    fig.update_layout(height=380, margin=dict(t=20, b=100, l=10, r=10), xaxis_title="", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)
    for _, r in df_ctx.iterrows():
        st.caption(f"{r['marca']} {r['indicador']} — fuente: {r['fuente']}")

    st.divider()

    st.markdown("**Mercado de regreso (win-back)**")
    df_wb = pd.DataFrame(mercado["mercado_winback"])
    for _, r in df_wb.iterrows():
        valor = r["valor"] if pd.notna(r.get("valor")) else r.get("valor_pct")
        valor_txt = ui.fmt(valor)
        ui.tarjeta_semaforo(f"{r['marca']} {r['indicador']}", valor_txt,
                             "#F9A825" if pd.isna(valor) else "#2E7D32",
                             r.get("nota", f"Fuente: {r.get('fuente', '')}"))

    st.divider()
    st.caption(f"Fecha de corte de las cifras de esta tab: {mercado['fecha_corte']} · Refresco trimestral manual (editar `data/mercado.json`).")
