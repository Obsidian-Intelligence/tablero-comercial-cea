"""Componentes de UI reutilizables: tarjetas, semáforos, etiquetas de marca."""
import math

import pandas as pd
import streamlit as st

from src.constants import (
    COLOR_AMARILLO,
    COLOR_GRIS,
    COLOR_ROJO,
    COLOR_VERDE,
    color_ae,
    color_score_okr,
    color_semaforo,
)


def badge_marca(marca: str) -> str:
    """✅ dato real · 📊 inferido · 🎯 hipótesis — para anexar junto a una cifra."""
    return marca or ""


def tarjeta_semaforo(titulo: str, valor: str, color: str, subtitulo: str = "") -> None:
    st.markdown(
        f"""
        <div style="border-left: 6px solid {color}; padding: 0.6rem 1rem;
                    background: rgba(127,127,127,0.06); border-radius: 6px; margin-bottom: 0.5rem;">
            <div style="font-size: 0.85rem; opacity: 0.75;">{titulo}</div>
            <div style="font-size: 1.6rem; font-weight: 700;">{valor}</div>
            <div style="font-size: 0.8rem; opacity: 0.65;">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def punto_semaforo(pct: float) -> str:
    color = color_semaforo(pct)
    if color == COLOR_VERDE:
        return "🟢"
    if color == COLOR_AMARILLO:
        return "🟡"
    return "🔴"


def punto_ae(ae: float) -> str:
    color = color_ae(ae)
    if color == COLOR_VERDE:
        return "🟢"
    if color == COLOR_AMARILLO:
        return "🟡"
    return "🔴"


def flecha_tendencia(actual: float, anterior: float) -> str:
    if anterior == 0:
        return "→"
    if actual > anterior:
        return "↑"
    if actual < anterior:
        return "↓"
    return "→"


def encabezado_tab(pregunta: str) -> None:
    st.markdown(f"##### {pregunta}")
    st.divider()


def banner_demo() -> None:
    st.warning("⚠️ **MODO DEMO** — sin conexión al Excel de Ventas. Los datos que ves son sintéticos.", icon="⚠️")


GLOSARIO: dict[str, str] = {
    "TAM": "Mercado total: todas las personas de la zona que podrían estudiar con "
           "nosotros si nada las limitara. Sirve para saber qué tan grande es el pastel.",
    "SAM": "Mercado alcanzable: la parte del pastel que sí busca lo que ofrecemos "
           "(sabatino/mixto, a nuestro precio). Es a quien realmente le vendemos.",
    "SOM": "Mercado que vamos a capturar: nuestra meta realista de alumnos. Aquí "
           "vive el 250.",
    "AE": "Cuánto vale un alumno en dinero comparado con uno que paga completo. Un "
          "alumno con beca del 10% vale 0.90. Sirve para que 250 alumnos con "
          "descuento no nos engañen: la meta es 250 alumnos que valgan al menos "
          "212 completos.",
    "CAC": "Cuánto nos cuesta en pesos conseguir UN alumno nuevo por cada canal. Si "
           "cuesta más de lo que deja, el canal se mata.",
    "OKR": "Apuesta del trimestre con fecha: se logra o no se logra.",
    "KR": "El número que demuestra si la apuesta se logró.",
    "KPI": "Signo vital permanente del negocio: se revisa cada lunes.",
    "NPS": "Del 0 al 10, ¿qué tanto nos recomendaría el alumno? Mide si el "
           "servicio genera recomendación.",
    "CPL": "Cuánto cuesta en pauta que una persona interesada nos deje sus datos.",
    "IMR": "El dinero que entra cada mes por colegiaturas, ya con descuentos "
           "descontados.",
    "Cohorte": "Grupo de prospectos que entró el mismo mes; se le sigue la pista "
               "60 días para ver cuántos se inscriben.",
    "Embudo": "El camino del interesado: pregunta → cita → inscripción. Muestra "
              "dónde se nos caen.",
}


def definicion(clave: str) -> str:
    """Regresa el texto de glosario para `clave`, o cadena vacía si no existe."""
    return GLOSARIO.get(clave, "")


def nota_glosario(clave: str) -> None:
    """Imprime la definición de `clave` en itálicas/letra chica, si existe."""
    texto = GLOSARIO.get(clave)
    if texto:
        st.caption(f"_{texto}_")


def fmt(valor, sufijo: str = "") -> str:
    """Formatea `valor` para mostrar en UI; nunca deja pasar 'nan'/'None'/vacío.

    None, NaN (float o pandas) -> 'Por llenar'. Números -> con separador de
    miles. Cualquier otro valor -> str(valor) + sufijo.
    """
    if valor is None:
        return "Por llenar"
    if isinstance(valor, float) and math.isnan(valor):
        return "Por llenar"
    try:
        if pd.isna(valor):
            return "Por llenar"
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (int, float)):
        return f"{valor:,.0f}{sufijo}"
    return f"{valor}{sufijo}"


def punto_score_okr(avance: float | None) -> str:
    """Semáforo de score de KR: ⚪ sin registro, 🟢≥0.7, 🟡0.3-0.69, 🔴<0.3."""
    if avance is None:
        return "⚪"
    color = color_score_okr(avance)
    if color == COLOR_VERDE:
        return "🟢"
    if color == COLOR_AMARILLO:
        return "🟡"
    return "🔴"
