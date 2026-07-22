"""Tab 2 · Avance vs Meta — ¿Llegamos o no esta semana, y si no, por qué?"""
import datetime as dt
import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src import etl, kpis, ui
from src.constants import (
    ALUMNOS_HOY_DEMO,
    CAC_E6_UMBRAL,
    CIERRE_PROMOCIONES,
    COLOR_GRIS,
    COLOR_ROJO,
    COLOR_VERDE,
    META_ALUMNOS,
    NOMBRES_ESTRATEGIAS,
    PLAZAS,
    PROGRAMAS,
    PUNTO_EQUILIBRIO,
    color_semaforo,
)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def _semanas_disponibles(con) -> list[str]:
    df = con.execute("SELECT DISTINCT semana_iso FROM metas_t ORDER BY semana_iso DESC").fetchdf()
    return df["semana_iso"].tolist()


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿Llegamos o no esta semana, y si no, por qué?")

    if es_demo:
        ui.banner_demo()

    col_sel, col_btn = st.columns([3, 1])
    semanas = _semanas_disponibles(con)
    with col_sel:
        semana_actual = st.selectbox("Semana ISO", semanas, index=0 if semanas else None)
    with col_btn:
        st.write("")
        if st.button("🔄 Actualizar ahora", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    # ------------------------------------------------------------------
    # 1. Héroe
    # ------------------------------------------------------------------
    alumnos_activos = hojas.get("Alumnos_Activos", pd.DataFrame())
    if not alumnos_activos.empty:
        alumnos_activos = alumnos_activos.copy()
        alumnos_activos["fecha_corte"] = pd.to_datetime(alumnos_activos["fecha_corte"])
        ultimo_corte = alumnos_activos["fecha_corte"].max()
        alumnos_hoy = int(alumnos_activos[alumnos_activos["fecha_corte"] == ultimo_corte]["total_activos"].sum())
        serie_semanal = alumnos_activos.groupby("fecha_corte", as_index=False)["total_activos"].sum()
    else:
        alumnos_hoy = ALUMNOS_HOY_DEMO
        serie_semanal = pd.DataFrame()

    fecha_cruce_207 = kpis.proyeccion_cruce(serie_semanal, PUNTO_EQUILIBRIO, dt.date(2026, 7, 13)) if not serie_semanal.empty else None
    ae_portafolio = kpis.ae_portafolio(alumnos_activos) if not alumnos_activos.empty else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Alumnos hoy", alumnos_hoy)
    c2.metric("Punto de equilibrio", PUNTO_EQUILIBRIO)
    c3.metric("Meta", META_ALUMNOS)
    c4.metric("AE del portafolio", f"{ae_portafolio:.2f}",
              help=f"{ui.definicion('AE')} Meta real: 250 alumnos con AE ≥ 0.85")

    progreso = min(alumnos_hoy / META_ALUMNOS, 1.0)
    st.progress(progreso, text=f"{alumnos_hoy} / {META_ALUMNOS} alumnos ({progreso:.0%})")

    if fecha_cruce_207:
        st.caption(f"A ritmo actual, cruzamos {PUNTO_EQUILIBRIO} el **{fecha_cruce_207.strftime('%d-%b-%Y')}** 📊")
    else:
        st.caption("Faltan datos suficientes de cortes semanales para proyectar la fecha de cruce 📊")

    st.divider()

    # ------------------------------------------------------------------
    # 1b. ¿Dónde están los alumnos?
    # ------------------------------------------------------------------
    st.markdown("**¿Dónde están los alumnos?**")
    if not alumnos_activos.empty:
        fechas_disponibles = sorted(alumnos_activos["fecha_corte"].unique())
        corte_actual = pd.Timestamp(fechas_disponibles[-1])
        corte_anterior = pd.Timestamp(fechas_disponibles[-2]) if len(fechas_disponibles) > 1 else None

        df_actual = alumnos_activos[alumnos_activos["fecha_corte"] == corte_actual]

        pivote = (
            df_actual.pivot_table(index="plaza", columns="programa", values="total_activos", aggfunc="sum", fill_value=0)
            .reindex(index=PLAZAS, columns=PROGRAMAS, fill_value=0)
        )
        pivote["Total"] = pivote.sum(axis=1)

        def _variacion(diff: int) -> str:
            if diff > 0:
                return f"▲ +{diff}"
            if diff < 0:
                return f"▼ {diff}"
            return "→ 0"

        if corte_anterior is not None:
            df_anterior = alumnos_activos[alumnos_activos["fecha_corte"] == corte_anterior]
            totales_anterior = df_anterior.groupby("plaza")["total_activos"].sum().reindex(PLAZAS, fill_value=0)
            pivote["Variación"] = [
                _variacion(int(pivote.loc[p, "Total"] - totales_anterior[p])) for p in PLAZAS
            ]
        else:
            pivote["Variación"] = "—"

        fila_total = pivote.drop(columns="Variación").sum(numeric_only=True)
        fila_total["Variación"] = (
            _variacion(int(pivote["Total"].sum() - totales_anterior.sum())) if corte_anterior is not None else "—"
        )
        fila_total.name = "Total"
        pivote = pd.concat([pivote, fila_total.to_frame().T])

        st.dataframe(pivote, use_container_width=True)

        fig_plaza = go.Figure()
        df_barras = df_actual.groupby(["plaza", "programa"], as_index=False)["total_activos"].sum()
        for programa in PROGRAMAS:
            sub = df_barras[df_barras["programa"] == programa].set_index("plaza").reindex(PLAZAS, fill_value=0)
            fig_plaza.add_bar(name=programa, x=PLAZAS, y=sub["total_activos"])
        fig_plaza.update_layout(
            barmode="stack", height=340, margin=dict(t=20, b=10, l=10, r=10), yaxis_title="Alumnos activos"
        )
        st.plotly_chart(fig_plaza, use_container_width=True)

        st.caption(
            f"_Fuente única: hoja Alumnos_Activos del Excel de Ventas, corte del {corte_actual.strftime('%d-%b-%Y')}. "
            "Se estará actualizando con los datos que vayamos recopilando todos._"
        )
    else:
        st.info("Sin datos de Alumnos_Activos todavía.")

    st.divider()

    # ------------------------------------------------------------------
    # 2. Semana actual — semáforo por estrategia
    # ------------------------------------------------------------------
    st.markdown("**Semana actual — semáforo por estrategia**")
    if semana_actual:
        avance = kpis.avance_semanal_por_estrategia(con, semana_actual)
        idx_actual = semanas.index(semana_actual)
        semana_anterior = semanas[idx_actual + 1] if idx_actual + 1 < len(semanas) else None
        avance_anterior = (
            kpis.avance_semanal_por_estrategia(con, semana_anterior).set_index("estrategia")["real"]
            if semana_anterior
            else pd.Series(dtype=float)
        )

        filas = []
        for _, r in avance.iterrows():
            anterior_val = avance_anterior.get(r["estrategia"], 0)
            filas.append(
                {
                    "Estrategia": f"{r['estrategia']} · {NOMBRES_ESTRATEGIAS.get(r['estrategia'], '')}",
                    "Meta": int(r["meta_alumnos"]),
                    "Real": int(r["real"]),
                    "% Cumplimiento": f"{r['pct_cumplimiento']:.0%}",
                    "Tendencia": ui.flecha_tendencia(r["real"], anterior_val),
                    "": ui.punto_semaforo(r["pct_cumplimiento"]),
                }
            )
        tabla = pd.DataFrame(filas)
        st.dataframe(tabla, use_container_width=True, hide_index=True)
    else:
        st.info("No hay metas registradas todavía.")

    st.divider()

    # ------------------------------------------------------------------
    # 3. Embudo
    # ------------------------------------------------------------------
    st.markdown("**Embudo comercial**")
    st.caption(f"_{ui.definicion('Embudo')}_")
    fc1, fc2 = st.columns(2)
    with fc1:
        plaza_filtro = st.selectbox("Plaza", ["Todas"] + PLAZAS, key="embudo_plaza")
    with fc2:
        canales_df = con.execute("SELECT DISTINCT canal_origen FROM prospectos_t ORDER BY 1").fetchdf()
        canal_filtro = st.selectbox("Canal", ["Todos"] + canales_df["canal_origen"].tolist(), key="embudo_canal")

    df_embudo = kpis.embudo(con, plaza_filtro, canal_filtro)
    df_dias = kpis.dias_promedio_por_etapa(con)
    df_embudo = df_embudo.merge(df_dias, on="etapa", how="left")

    if not df_embudo.empty:
        fig = go.Figure(
            go.Funnel(
                y=df_embudo["etapa"],
                x=df_embudo["conteo"],
                textinfo="value+percent initial",
            )
        )
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=380)
        st.plotly_chart(fig, use_container_width=True)

        cols_dias = st.columns(len(df_embudo))
        for col, (_, row) in zip(cols_dias, df_embudo.iterrows()):
            dias = row.get("dias_promedio")
            col.metric(row["etapa"], f"{row['conteo']}", f"{dias:.1f} días" if pd.notna(dias) else "—")
    else:
        st.info("No hay prospectos que coincidan con el filtro.")

    resultado_contacto = kpis.perdidos_sin_contactar(con)
    if resultado_contacto is None:
        ui.tarjeta_semaforo("⚪ Perdidos sin contactar", "Por llenar", COLOR_GRIS, "")
    else:
        pct_sin_contactar = resultado_contacto["pct"]
        en_rojo = pct_sin_contactar > 0.15
        ui.tarjeta_semaforo(
            f"{'🔴' if en_rojo else '🟢'} Perdidos sin contactar",
            f"{pct_sin_contactar:.0%}",
            COLOR_ROJO if en_rojo else COLOR_VERDE,
            "🎯 Meta: por debajo de 15%" if en_rojo else "",
        )
    st.caption(
        "_De los prospectos marcados 'Perdido — No responde', qué % en realidad nunca llegamos "
        "a buscar 3 veces. Si sale en rojo, probablemente los estamos perdiendo por no insistir, "
        "no porque de verdad ya no quisieran._"
    )

    st.divider()

    # ------------------------------------------------------------------
    # 4. Reglas de corte
    # ------------------------------------------------------------------
    st.markdown("**Reglas de corte**")
    with open(DATA_DIR / "costos_fijos.json", encoding="utf-8") as f:
        costos = json.load(f)
    gasto_pauta_manual = kpis.valor_kpi_manual(hojas, "Gasto pauta MXN")
    if gasto_pauta_manual is not None:
        gasto_pauta = gasto_pauta_manual
    else:
        gasto_pauta = next(
            (r["monto"] for r in costos["renglones"] if r["concepto"] == "Pauta publicitaria E6" and r["monto"]),
            7000,
        )
    historial_cac = kpis.cac_e6_mensual(con, gasto_pauta)
    resumen_cac = kpis.cac_e6(con, gasto_pauta)
    excede_dos_meses = kpis.excede_dos_meses_seguidos(historial_cac)

    col_e6, col_dias = st.columns(2)
    with col_e6:
        cac_valor = resumen_cac["cac"]
        color = "🔴" if excede_dos_meses else ("🟡" if resumen_cac["excede_umbral"] else "🟢")
        texto_cac = f"${cac_valor:,.0f} MXN" if cac_valor is not None else "Sin alumnos E6 aún"
        ui.tarjeta_semaforo(
            f"{color} CAC de E6 vs umbral ${CAC_E6_UMBRAL:,} MXN",
            texto_cac,
            "#C62828" if excede_dos_meses else ("#F9A825" if resumen_cac["excede_umbral"] else "#2E7D32"),
            "Excede el umbral 2 meses seguidos — considerar matar/ajustar canal" if excede_dos_meses else "",
        )
        st.caption(f"_{ui.definicion('CAC')}_")
    with col_dias:
        cierre = dt.date.fromisoformat(CIERRE_PROMOCIONES)
        dias_restantes = (cierre - dt.date(2026, 7, 13)).days
        ui.tarjeta_semaforo(
            "Ventana de inscripción",
            f"{dias_restantes} días restantes",
            "#F9A825" if dias_restantes < 30 else "#2E7D32",
            f"Cierre de promociones: {cierre.strftime('%d-%b-%Y')}",
        )

    st.divider()

    # ------------------------------------------------------------------
    # 5. ¿Por qué no llegamos?
    # ------------------------------------------------------------------
    st.markdown("**¿Por qué no llegamos?**")
    comentarios = con.execute(
        """
        SELECT semana_iso, estrategia, responsable, comentario_cierre
        FROM metas_t
        WHERE comentario_cierre IS NOT NULL AND comentario_cierre != ''
        ORDER BY semana_iso DESC
        LIMIT 8
        """
    ).fetchdf()
    if comentarios.empty:
        st.caption("Sin comentarios de cierre registrados.")
    else:
        for _, r in comentarios.iterrows():
            st.markdown(f"> *\"{r['comentario_cierre']}\"* — **{r['responsable']}**, semana {r['semana_iso']} ({r['estrategia']})")

    st.divider()

    # ------------------------------------------------------------------
    # 5b. Cobranza y KPIs manuales de la semana
    # ------------------------------------------------------------------
    st.markdown("**Cobranza y KPIs manuales**")
    cobranza = kpis.cobranza_pct(alumnos_activos)
    col_cob, col_nps, col_citas, col_bajas = st.columns(4)
    color_cob = "🟢" if cobranza["pct"] >= 0.90 else ("🟡" if cobranza["pct"] >= 0.70 else "🔴")
    col_cob.metric(f"{color_cob} Cobranza (KPI 10)", f"{cobranza['pct']:.0%}", help="cobrado ÷ facturable del corte más reciente. Meta ≥90%.")
    nps_valor = kpis.valor_kpi_manual(hojas, "NPS")
    col_nps.metric("NPS post-trámite", f"{nps_valor:.0f}" if nps_valor is not None else "—")
    citas_valor = kpis.valor_kpi_manual(hojas, "Citas agendadas", plaza="Puebla")
    col_citas.metric("Citas agendadas (Puebla)", f"{citas_valor:.0f}" if citas_valor is not None else "—")
    bajas_valor = kpis.valor_kpi_manual(hojas, "Bajas del mes")
    col_bajas.metric("Bajas del mes", f"{bajas_valor:.0f}" if bajas_valor is not None else "—")

    # ------------------------------------------------------------------
    # 6. Calidad de datos
    # ------------------------------------------------------------------
    with st.expander("🔍 Calidad de datos"):
        resultados = etl.validar_calidad(con)
        for regla, conteo in resultados.items():
            icono = "✅" if conteo == 0 else "⚠️"
            st.write(f"{icono} {regla}: **{conteo}** filas")
