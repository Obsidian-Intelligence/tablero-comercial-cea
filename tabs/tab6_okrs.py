"""Tab 6 · OKRs y Plan — ¿Qué apostamos este trimestre y cómo vamos?"""
import json
from pathlib import Path

import pandas as pd
import streamlit as st

from src import kpis, ui
from src.constants import PLAZAS

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def _avance_desde_sheet(hojas: dict) -> dict[str, float]:
    """Lee la hoja OKR_Avance si existe y trae datos; regresa {kr_id: avance}."""
    df = hojas.get("OKR_Avance", pd.DataFrame())
    if df.empty or "kr_id" not in df.columns or "avance" not in df.columns:
        return {}
    out = {}
    for _, r in df.iterrows():
        try:
            out[r["kr_id"]] = float(r["avance"])
        except (TypeError, ValueError):
            continue
    return out


def _fila_kr(kr: dict, avance_override: dict[str, float]) -> None:
    avance = avance_override.get(kr["id"], kr["avance"])
    col_marca, col_texto, col_barra = st.columns([0.6, 4, 2])
    with col_marca:
        st.markdown(kr.get("marca", "") or "&nbsp;", unsafe_allow_html=True)
    with col_texto:
        st.markdown(f"**{kr['id']}** {kr['texto']}")
    with col_barra:
        if avance is None:
            st.progress(0)
            st.caption("⚪ Sin registro")
        else:
            st.progress(min(max(avance, 0.0), 1.0))
            st.caption(f"{ui.punto_score_okr(avance)} {avance:.0%}")


def render(con, hojas, es_demo: bool) -> None:
    ui.encabezado_tab("¿Qué apostamos este trimestre y cómo vamos?")

    with open(DATA_DIR / "okrs.json", encoding="utf-8") as f:
        okrs = json.load(f)

    avance_override = _avance_desde_sheet(hojas)
    if not avance_override:
        st.caption("📊 Modo estático — sin hoja `OKR_Avance` en el Excel de Ventas todavía; los KR muestran 'Sin registro'.")

    # ------------------------------------------------------------------
    # 1. Tarjeta Norte
    # ------------------------------------------------------------------
    st.markdown(
        f"""
        <div style="border-left: 8px solid #2E7D32; padding: 1rem 1.4rem;
                    background: rgba(46,125,50,0.08); border-radius: 8px; margin-bottom: 0.8rem;">
            <div style="font-size: 0.85rem; opacity: 0.75; text-transform: uppercase; letter-spacing: 0.05em;">Norte (12 meses)</div>
            <div style="font-size: 1.3rem; font-weight: 700; margin: 0.3rem 0;">{okrs['norte']['titular']}</div>
            <div style="font-size: 0.9rem; opacity: 0.8;">Métrica estrella: {okrs['norte']['metrica_estrella']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ------------------------------------------------------------------
    # 2. Regla de lectura
    # ------------------------------------------------------------------
    with st.expander("📖 Regla de lectura"):
        st.markdown(f"**OKR** — {okrs['regla_lectura']['okr']}")
        st.markdown(f"**KR** — {okrs['regla_lectura']['kr']}")
        st.markdown(f"**KPI** — {okrs['regla_lectura']['kpi']}")
        st.markdown(f"**Iniciativa** — {okrs['regla_lectura']['iniciativa']}")
        st.divider()
        st.markdown(f"_{okrs['regla_lectura']['cierre']}_")

    st.divider()

    # ------------------------------------------------------------------
    # 3. Objetivos Q3-2026
    # ------------------------------------------------------------------
    st.markdown("### Objetivos Q3-2026 (13-jul → 30-sep)")
    st.caption(f"_Cohorte — {ui.definicion('Cohorte')}_")
    st.caption("Semáforo de score: 🟢 ≥0.7 verde · 🟡 0.3–0.69 amarillo · 🔴 <0.3 rojo · ⚪ sin registro. 0.7 es éxito; 1.0 sistemático = metas flojas.")

    for obj in okrs["objetivos_q3"]:
        with st.expander(f"{obj['id']} · {obj['titulo']}", expanded=(obj["id"] == "O1")):
            for kr in obj["krs"]:
                _fila_kr(kr, avance_override)

    st.divider()

    # ------------------------------------------------------------------
    # 4. OKRs direccionales
    # ------------------------------------------------------------------
    st.markdown("### OKRs direccionales")
    col_q4, col_q1 = st.columns(2)
    for col, clave in [(col_q4, "q4_2026"), (col_q1, "q1_2027")]:
        bloque = okrs["direccionales"][clave]
        with col:
            st.markdown(
                f"""
                <div style="border: 1px solid rgba(127,127,127,0.3); padding: 0.8rem 1rem;
                            background: rgba(127,127,127,0.06); border-radius: 8px;">
                    <div style="font-weight: 700;">{bloque['titulo']}</div>
                    <div style="font-size: 0.8rem; opacity: 0.65; margin-bottom: 0.5rem;">{bloque['etiqueta']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            for punto in bloque["puntos"]:
                st.markdown(f"- {punto}")

    st.divider()

    # ------------------------------------------------------------------
    # 5. Iniciativas
    # ------------------------------------------------------------------
    st.markdown("### Iniciativas")
    st.caption(f"_CPL — {ui.definicion('CPL')}_")
    df_ini = pd.DataFrame(okrs["iniciativas"])
    st.dataframe(
        df_ini[["id", "iniciativa", "dueno", "krs", "incluye", "costo"]].rename(
            columns={"id": "Iniciativa", "iniciativa": "Nombre", "dueno": "Dueño",
                     "krs": "KRs que mueve", "incluye": "Qué incluye", "costo": "Costo"}
        ),
        use_container_width=True, hide_index=True,
    )
    for ini in okrs["iniciativas"]:
        if ini.get("alerta"):
            st.error(f"⚠️ {ini['id']} — {ini['alerta']}")

    st.divider()

    # ------------------------------------------------------------------
    # 6. KPIs permanentes
    # ------------------------------------------------------------------
    st.markdown("### KPIs permanentes (tablero de los lunes)")
    st.caption(f"_NPS — {ui.definicion('NPS')}_")
    df_kpi = pd.DataFrame(okrs["kpis_permanentes"])

    mapa_kpi_manual = {
        "Mediana 1ª respuesta": "Mediana 1ª respuesta (min)",
        "Citas agendadas": "Citas agendadas",
        "Bajas del mes": "Bajas del mes",
        "NPS post-trámite": "NPS",
        "Reseñas Google (# y promedio)": "Reseñas Google #",
    }
    cobranza = kpis.cobranza_pct(hojas.get("Alumnos_Activos", pd.DataFrame()))

    # Estos KPIs de KPI_Manual se capturan por plaza (Puebla/Huauchinango/Zacatlán),
    # nunca con plaza == "General", así que hay que sumarlos por plaza en vez de
    # buscar una fila "General" que nunca existe.
    kpis_por_plaza = {"Citas agendadas", "Reseñas Google (# y promedio)"}

    def _valor_actual(nombre_kpi: str) -> str:
        if nombre_kpi == "Cobranza % del facturable":
            return f"{cobranza['pct']:.0%}"
        clave_manual = mapa_kpi_manual.get(nombre_kpi)
        if not clave_manual:
            return "—"
        if nombre_kpi in kpis_por_plaza:
            valores = [kpis.valor_kpi_manual(hojas, clave_manual, plaza=p) for p in PLAZAS]
            if all(v is None for v in valores):
                return "—"
            total = sum(v for v in valores if v is not None)
            return f"{total:g}"
        valor = kpis.valor_kpi_manual(hojas, clave_manual)
        return f"{valor:g}" if valor is not None else "—"

    df_kpi["valor_actual"] = df_kpi["kpi"].apply(_valor_actual)
    st.dataframe(
        df_kpi[["kpi", "valor_actual", "fuente", "frecuencia", "dueno", "meta", "contra"]].rename(
            columns={"kpi": "KPI", "valor_actual": "Valor actual", "fuente": "Fuente", "frecuencia": "Frecuencia",
                     "dueno": "Dueño", "meta": "Meta", "contra": "Contramétrica anti-Goodhart"}
        ),
        use_container_width=True, hide_index=True,
    )
    st.caption("_\"Valor actual\" se lee de KPI_Manual / Alumnos_Activos cuando el dato existe; \"—\" significa que aún no se ha capturado._")

    st.divider()

    # ------------------------------------------------------------------
    # 7. Cadencia
    # ------------------------------------------------------------------
    st.markdown("### Cadencia")
    for linea in okrs["cadencia"]:
        st.markdown(f"- {linea}")

    st.caption(f"Última actualización del contenido de esta tab: {okrs['fecha_actualizacion']}.")
