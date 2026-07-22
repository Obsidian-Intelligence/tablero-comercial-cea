"""Tab 5 · Finanzas de la estrategia — ¿La meta de 250 nos deja dinero o nos lo quita?"""
import datetime as dt
import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src import kpis, ui
from src.constants import META_ALUMNOS, PUNTO_EQUILIBRIO, color_ae

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def _cargar_json(nombre: str) -> dict:
    with open(DATA_DIR / nombre, encoding="utf-8") as f:
        return json.load(f)


def _blended_cuota_plena(alumnos_activos: pd.DataFrame) -> float:
    tabla = kpis.alumno_equivalente(alumnos_activos)
    if tabla.empty:
        return 1500.0
    return float((tabla["cuota_plena"] * tabla["total_activos"]).sum() / tabla["total_activos"].sum())


def _tabla_compensacion_descuento() -> pd.DataFrame:
    """ceil(207 / (1 - d)) para cada nivel de descuento promedio de la spec."""
    filas = [
        {"descuento": "0%", "alumnos": 207, "nota": ""},
        {"descuento": "10%", "alumnos": 230, "nota": ""},
        {"descuento": "15%", "alumnos": 244, "nota": ""},
        {"descuento": "17%", "alumnos": 250, "nota": "← nuestra meta"},
        {"descuento": "20%", "alumnos": 259, "nota": "ya rebasa la meta"},
    ]
    return pd.DataFrame(filas)


def _bloque_precios_y_promociones(politica: dict) -> None:
    st.markdown("### Precios y promociones en una mirada")
    st.markdown("#### El descuento máximo por alumno es 20% del año. Nada se acumula.")
    st.markdown("#### La cartera completa aguanta hasta 17% de descuento promedio. Arriba de eso, 250 alumnos ya no alcanzan.")
    st.markdown("#### Cada punto de descuento se paga con más alumnos:")

    df_comp = _tabla_compensacion_descuento()
    st.dataframe(
        df_comp.rename(columns={
            "descuento": "Descuento promedio de la cartera",
            "alumnos": "Alumnos necesarios para salir tablas",
            "nota": "",
        }),
        use_container_width=True, hide_index=True,
    )
    st.caption(
        "_Lectura: si regalamos 20% en promedio, necesitamos 259 alumnos, 9 más que "
        "la meta. Por eso los descuentos tienen tope y no se acumulan._ 📊"
    )

    estrategias_con_tope = ["E2", "E3", "E4", "E5", "E9"]
    filas_promo = [
        r for r in politica["por_estrategia"]
        if any(r["estrategia"].startswith(e) for e in estrategias_con_tope)
    ]
    df_promo = pd.DataFrame(filas_promo).rename(
        columns={"estrategia": "Estrategia", "descuento_max": "Tope", "logica": "Lógica"}
    )
    df_promo["Vence"] = "28-ago-2026"
    st.markdown("**Promociones vigentes**")
    st.dataframe(df_promo[["Estrategia", "Tope", "Vence"]], use_container_width=True, hide_index=True)

    st.caption(
        "_Basado en el punto de equilibrio de 207 alumnos a precio pleno 📊; se "
        "recalcula en pesos cuando se completen los costos fijos._"
    )
    st.divider()


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿La meta de 250 nos deja dinero o nos lo quita?")
    if es_demo:
        ui.banner_demo()

    _bloque_precios_y_promociones(_cargar_json("politica_descuentos.json"))

    alumnos_activos = hojas.get("Alumnos_Activos", pd.DataFrame()).copy()
    ultimo_corte_df = pd.DataFrame()
    if not alumnos_activos.empty:
        alumnos_activos["fecha_corte"] = pd.to_datetime(alumnos_activos["fecha_corte"])
        ultimo_corte = alumnos_activos["fecha_corte"].max()
        ultimo_corte_df = alumnos_activos[alumnos_activos["fecha_corte"] == ultimo_corte]

    costos = _cargar_json("costos_fijos.json")
    costos_definidos = [r["monto"] for r in costos["renglones"] if r["monto"] is not None]
    renglones_faltantes = len(costos["renglones"]) - len(costos_definidos)
    imr_equilibrio = sum(costos_definidos)

    imr_real = float(ultimo_corte_df["ingreso_mensual_real_mxn"].sum()) if not ultimo_corte_df.empty else 0.0
    ae_portafolio = kpis.ae_portafolio(alumnos_activos) if not alumnos_activos.empty else 0.0
    alumnos_hoy = int(ultimo_corte_df["total_activos"].sum()) if not ultimo_corte_df.empty else 0
    blended_cuota = _blended_cuota_plena(alumnos_activos)

    # ------------------------------------------------------------------
    # 1. Héroe
    # ------------------------------------------------------------------
    st.caption(f"_IMR — {ui.definicion('IMR')}_")
    c1, c2, c3 = st.columns(3)
    c1.metric("IMR (Ingreso Mensual Recurrente Real)", f"${imr_real:,.0f}")
    c2.metric("IMR de equilibrio (costos fijos)", f"${imr_equilibrio:,.0f}",
              help="Suma de costos_fijos.json — incompleto hasta que Administración y Finanzas llene los placeholders")
    if renglones_faltantes:
        st.caption(f"Equilibrio parcial: faltan {renglones_faltantes} renglones por llenar 📊")
    color = color_ae(ae_portafolio)
    punto = "🟢" if color == "#2E7D32" else ("🟡" if color == "#F9A825" else "🔴")
    c3.metric(f"{punto} AE del portafolio", f"{ae_portafolio:.2f}",
              help=f"{ui.definicion('AE')} Verde ≥0.90 · Amarillo 0.85-0.89 · Rojo <0.85")

    st.divider()

    # ------------------------------------------------------------------
    # ¿Salimos tablas a 250?
    # ------------------------------------------------------------------
    st.markdown("**¿Salimos tablas a 250?**")
    st.caption("Simulador: equivalentes = 250 × (1 − descuento promedio), contra la línea de equilibrio de 207.")

    descuento_pct = st.slider("Descuento promedio aplicado (%)", 0, 25, 15, key="slider_descuento")
    equivalentes_slider = int(META_ALUMNOS * (1 - descuento_pct / 100))

    escenarios = pd.DataFrame(
        [
            {"escenario": "Tope 20% parejo", "descuento": 20, "equivalentes": 200},
            {"escenario": "15% parejo", "descuento": 15, "equivalentes": 212},
            {"escenario": "Realista: 110 plena + 140 al 15%", "descuento": None, "equivalentes": 229},
            {"escenario": f"Tu escenario ({descuento_pct}%)", "descuento": descuento_pct, "equivalentes": equivalentes_slider},
        ]
    )
    escenarios["color"] = escenarios["equivalentes"].apply(
        lambda v: "#2E7D32" if v >= PUNTO_EQUILIBRIO else ("#F9A825" if v >= PUNTO_EQUILIBRIO * 0.85 else "#C62828")
    )

    fig = go.Figure()
    fig.add_bar(x=escenarios["escenario"], y=escenarios["equivalentes"], marker_color=escenarios["color"])
    fig.add_hline(y=PUNTO_EQUILIBRIO, line_dash="dash", line_color="gray",
                  annotation_text=f"Equilibrio ({PUNTO_EQUILIBRIO})")
    fig.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10), yaxis_title="Alumnos equivalentes")
    st.plotly_chart(fig, use_container_width=True)
    st.caption("El portafolio aguanta hasta ~17% de descuento promedio 📊 (con estructura de costos actual; se recalcula al llenar costos fijos).")

    st.divider()

    # ------------------------------------------------------------------
    # 2. Puente alumnos → dinero (waterfall)
    # ------------------------------------------------------------------
    st.markdown("**Puente alumnos → dinero**")
    fuga = kpis.fuga_por_descuentos(con)
    factor_escala = (META_ALUMNOS / alumnos_hoy) if alumnos_hoy else 1.0
    fuga_dict = dict(zip(fuga["estrategia"], fuga["fuga_mxn"])) if not fuga.empty else {}

    becas_e3 = fuga_dict.get("E3", 0) * factor_escala
    referidos_e4 = fuga_dict.get("E4", 0) * factor_escala
    convenios_e5 = fuga_dict.get("E5", 0) * factor_escala
    condonaciones = sum(v for k, v in fuga_dict.items() if k not in ("E3", "E4", "E5")) * factor_escala

    base = META_ALUMNOS * blended_cuota
    imr_proyectado = base - becas_e3 - referidos_e4 - convenios_e5 - condonaciones

    fig_wf = go.Figure(
        go.Waterfall(
            orientation="v",
            measure=["absolute", "relative", "relative", "relative", "relative", "total"],
            x=["250 × cuota plena", "− Becas E3", "− Referidos E4", "− Convenios E5", "− Condonaciones", "IMR proyectado"],
            y=[base, -becas_e3, -referidos_e4, -convenios_e5, -condonaciones, imr_proyectado],
            text=[f"${v:,.0f}" for v in [base, -becas_e3, -referidos_e4, -convenios_e5, -condonaciones, imr_proyectado]],
            connector={"line": {"color": "rgba(127,127,127,0.4)"}},
        )
    )
    fig_wf.update_layout(height=380, margin=dict(t=20, b=10, l=10, r=10))
    st.plotly_chart(fig_wf, use_container_width=True)
    st.caption("Proyección con la mezcla actual de canales, escalada a la meta de 250 alumnos 📊")

    st.divider()

    # ------------------------------------------------------------------
    # 3. Ingresos futuros vs costos (6 meses)
    # ------------------------------------------------------------------
    st.markdown("**Ingresos futuros planeados vs costos (6 meses)**")
    metas_totales = con.execute(
        "SELECT AVG(meta_alumnos) FROM metas_t WHERE estrategia = 'Total'"
    ).fetchone()[0] or 0
    alumnos_nuevos_mes = float(metas_totales) * 4.33

    meses = list(range(0, 7))
    imr_proy = [imr_real + m * alumnos_nuevos_mes * blended_cuota * ae_portafolio for m in meses]
    fig_seis = go.Figure()
    fig_seis.add_trace(go.Scatter(x=meses, y=imr_proy, mode="lines+markers", name="IMR proyectado"))
    fig_seis.add_hline(y=imr_equilibrio, line_dash="dash", line_color="#C62828", annotation_text="Costos fijos (equilibrio)")
    fig_seis.update_layout(height=340, margin=dict(t=20, b=10, l=10, r=10),
                            xaxis_title="Meses desde hoy", yaxis_title="MXN/mes")
    st.plotly_chart(fig_seis, use_container_width=True)

    cruce = next((m for m, v in zip(meses, imr_proy) if v >= imr_equilibrio), None)
    if cruce is not None:
        st.caption(f"Cruce proyectado con costos fijos en el mes **{cruce}** 📊")
    else:
        st.caption("No se proyecta cruce con costos fijos dentro de los próximos 6 meses con el ritmo actual 📊")

    st.divider()

    # ------------------------------------------------------------------
    # 4. Tabla de costos fijos
    # ------------------------------------------------------------------
    st.markdown("**Costos fijos (editable en `data/costos_fijos.json`)**")
    df_costos = pd.DataFrame(costos["renglones"])
    df_costos["Monto"] = df_costos["monto"].apply(
        lambda v: f"${v:,.0f}" if v is not None else ui.fmt(v)
    )
    st.dataframe(df_costos[["concepto", "Monto"]].rename(columns={"concepto": "Concepto"}),
                 use_container_width=True, hide_index=True)

    st.divider()

    # ------------------------------------------------------------------
    # 5. Política de descuentos
    # ------------------------------------------------------------------
    st.markdown("**Política de descuentos por estrategia**")
    politica = _cargar_json("politica_descuentos.json")
    st.warning(
        f"{politica['estatus']} — los descuentos NO son apilables; tope absoluto "
        "20% del valor anual por alumno; toda condonación requiere aprobación "
        "registrada en la hoja Descuentos_Otorgados."
    )
    for regla in politica["reglas_globales"]:
        st.caption(f"• {regla}")
    df_politica = pd.DataFrame(politica["por_estrategia"]).rename(
        columns={"estrategia": "Estrategia", "descuento_max": "Descuento máximo", "logica": "Lógica"}
    )
    st.dataframe(df_politica, use_container_width=True, hide_index=True)

    st.divider()

    # ------------------------------------------------------------------
    # 6. Fuga por descuentos del mes
    # ------------------------------------------------------------------
    st.markdown("**Fuga por descuentos del mes**")
    comparativo = kpis.fuga_mes_actual_vs_anterior(con, dt.date(2026, 7, 13))
    fc1, fc2 = st.columns(2)
    fc1.metric("Mes actual", f"${comparativo['mes_actual_mxn']:,.0f}")
    fc2.metric("Mes anterior", f"${comparativo['mes_anterior_mxn']:,.0f}",
               delta=f"{comparativo['mes_actual_mxn'] - comparativo['mes_anterior_mxn']:,.0f}",
               delta_color="inverse")

    if not fuga.empty:
        fuga_mostrar = fuga.rename(columns={"estrategia": "Estrategia", "fuga_mxn": "Fuga MXN/mes"})
        st.dataframe(fuga_mostrar, use_container_width=True, hide_index=True)
    else:
        st.caption("Sin descuentos registrados todavía.")
