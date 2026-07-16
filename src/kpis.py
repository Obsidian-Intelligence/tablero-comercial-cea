"""Cálculos de KPI: embudo, días por etapa, AE, CAC, proyecciones, fuga por descuentos."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import duckdb
import numpy as np
import pandas as pd

from src.constants import CAC_E6_UMBRAL, ESTRATEGIAS, ETAPAS

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# Mapea el programa/plaza tal como se captura en Prospectos/Alumnos_Activos
# a la fila correspondiente de precios_cea.json.
_PROGRAMAS_LICENCIATURA = {"Lic. Administración", "Lic. Contaduría", "Lic. Derecho"}


def _cargar_json(nombre: str) -> dict:
    with open(DATA_DIR / nombre, encoding="utf-8") as f:
        return json.load(f)


def cargar_precios() -> dict:
    return _cargar_json("precios_cea.json")


def cuota_plena(plaza: str, programa: str) -> float:
    """Regresa la mensualidad plena aplicable a un (plaza, programa) según
    la tabla oficial de precios. Las licenciaturas y la maestría solo se
    ofrecen oficialmente en Puebla; se usa ese precio como referencia."""
    precios = cargar_precios()
    programas = precios["programas"]

    if programa in _PROGRAMAS_LICENCIATURA:
        fila = next(p for p in programas if p["programa"].startswith("Licenciatura"))
        return fila["mensualidad_plena"]

    if programa == "Maestría Derecho Penal":
        fila = next(p for p in programas if p["programa"].startswith("Maestría"))
        return fila["mensualidad_plena"]

    if programa == "Bachillerato":
        fila = next(
            (p for p in programas if p["programa"] == "Bachillerato Sabatino" and p["plaza"] == plaza),
            None,
        )
        if fila is None:
            # Puebla también tiene Entresemana; si no hay Sabatino para la plaza, usar Puebla
            fila = next(p for p in programas if p["programa"] == "Bachillerato Sabatino" and p["plaza"] == "Puebla")
        return fila["mensualidad_plena"]

    return 1500.0  # fallback conservador


# ------------------------------------------------------------------
# Embudo
# ------------------------------------------------------------------

def embudo(con: duckdb.DuckDBPyConnection, plaza: str | None = None, canal: str | None = None) -> pd.DataFrame:
    filtros = []
    if plaza and plaza != "Todas":
        filtros.append(f"plaza = '{plaza}'")
    if canal and canal != "Todos":
        filtros.append(f"canal_origen = '{canal}'")
    where = f"WHERE {' AND '.join(filtros)}" if filtros else ""

    df = con.execute(
        f"SELECT etapa, COUNT(*) AS conteo FROM prospectos_t {where} GROUP BY etapa"
    ).fetchdf()

    orden = {e: i for i, e in enumerate(ETAPAS)}
    df["orden"] = df["etapa"].map(orden)
    df = df.sort_values("orden").drop(columns="orden").reset_index(drop=True)

    df["tasa_conversion"] = df["conteo"] / df["conteo"].iloc[0] if len(df) else 1.0
    return df


def dias_promedio_por_etapa(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    """Días promedio que tarda cada transición de etapa, según Eventos."""
    df = con.execute(
        """
        SELECT id_prospecto, etapa_anterior, etapa_nueva, CAST(fecha_hora AS TIMESTAMP) AS fecha_hora
        FROM eventos_t ORDER BY id_prospecto, fecha_hora
        """
    ).fetchdf()
    if df.empty:
        return pd.DataFrame(columns=["etapa", "dias_promedio"])

    df["fecha_anterior"] = df.groupby("id_prospecto")["fecha_hora"].shift(1)
    df = df.dropna(subset=["fecha_anterior"])
    df["dias"] = (df["fecha_hora"] - df["fecha_anterior"]).dt.total_seconds() / 86400

    resumen = df.groupby("etapa_nueva")["dias"].mean().reset_index()
    resumen.columns = ["etapa", "dias_promedio"]
    return resumen


# ------------------------------------------------------------------
# Alumnos nuevos por semana / estrategia
# ------------------------------------------------------------------

def alumnos_nuevos_por_semana(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    df = con.execute(
        """
        SELECT p.canal_origen, CAST(e.fecha_hora AS TIMESTAMP) AS fecha_hora
        FROM eventos_t e
        JOIN prospectos_t p ON p.id = e.id_prospecto
        WHERE e.etapa_nueva = 'Alumno'
        """
    ).fetchdf()
    if df.empty:
        return pd.DataFrame(columns=["semana_iso", "canal_origen", "alumnos_nuevos"])

    iso = df["fecha_hora"].dt.isocalendar()
    df["semana_iso"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
    resumen = df.groupby(["semana_iso", "canal_origen"]).size().reset_index(name="alumnos_nuevos")
    return resumen


def avance_semanal_por_estrategia(con: duckdb.DuckDBPyConnection, semana_iso: str) -> pd.DataFrame:
    """Junta meta_alumnos de la hoja Metas con los alumnos reales logrados esa semana."""
    metas = con.execute(
        "SELECT estrategia, meta_alumnos FROM metas_t WHERE semana_iso = ? AND estrategia != 'Total'",
        [semana_iso],
    ).fetchdf()

    reales = alumnos_nuevos_por_semana(con)
    reales_semana = reales[reales["semana_iso"] == semana_iso].set_index("canal_origen")["alumnos_nuevos"]

    metas["real"] = metas["estrategia"].map(reales_semana).fillna(0).astype(int)
    metas["pct_cumplimiento"] = metas.apply(
        lambda r: (r["real"] / r["meta_alumnos"]) if r["meta_alumnos"] else 0.0, axis=1
    )
    return metas


# ------------------------------------------------------------------
# Proyección lineal de alumnos activos
# ------------------------------------------------------------------

def proyeccion_cruce(
    serie: pd.DataFrame, objetivo: int, hoy: dt.date
) -> dt.date | None:
    """serie: columnas fecha_corte (date), total_activos (int), ya agregada por semana.
    Usa la pendiente de las últimas 4 semanas para proyectar la fecha de cruce."""
    serie = serie.sort_values("fecha_corte").tail(4)
    if len(serie) < 2:
        return None

    x = (serie["fecha_corte"] - serie["fecha_corte"].min()).dt.days.to_numpy(dtype=float)
    y = serie["total_activos"].to_numpy(dtype=float)
    pendiente, intercepto = np.polyfit(x, y, 1) if len(x) > 1 else (0, y[-1])

    if pendiente <= 0:
        return None

    ultimo_x = x[-1]
    ultimo_y = y[-1]
    if ultimo_y >= objetivo:
        return serie["fecha_corte"].iloc[-1].date() if hasattr(serie["fecha_corte"].iloc[-1], "date") else serie["fecha_corte"].iloc[-1]

    dias_faltantes = (objetivo - ultimo_y) / pendiente
    fecha_base = serie["fecha_corte"].iloc[-1]
    if hasattr(fecha_base, "date"):
        fecha_base = fecha_base.date()
    return fecha_base + dt.timedelta(days=float(dias_faltantes))


# ------------------------------------------------------------------
# CAC de E6
# ------------------------------------------------------------------

def cac_e6(con: duckdb.DuckDBPyConnection, gasto_pauta_mxn: float) -> dict:
    alumnos_e6 = con.execute(
        "SELECT COUNT(*) FROM prospectos_t WHERE canal_origen = 'E6' AND etapa = 'Alumno'"
    ).fetchone()[0]
    cac = gasto_pauta_mxn / alumnos_e6 if alumnos_e6 else None
    return {
        "alumnos_e6": alumnos_e6,
        "gasto_pauta_mxn": gasto_pauta_mxn,
        "cac": cac,
        "excede_umbral": (cac is not None and cac > CAC_E6_UMBRAL),
    }


def cac_e6_mensual(con: duckdb.DuckDBPyConnection, gasto_pauta_mensual_mxn: float) -> pd.DataFrame:
    """CAC de E6 mes a mes, para detectar si excede el umbral 2 meses seguidos."""
    df = con.execute(
        """
        SELECT CAST(e.fecha_hora AS TIMESTAMP) AS fecha_hora
        FROM eventos_t e
        JOIN prospectos_t p ON p.id = e.id_prospecto
        WHERE e.etapa_nueva = 'Alumno' AND p.canal_origen = 'E6'
        """
    ).fetchdf()
    if df.empty:
        return pd.DataFrame(columns=["mes", "alumnos_e6", "cac"])

    df["mes"] = df["fecha_hora"].dt.to_period("M").astype(str)
    resumen = df.groupby("mes").size().reset_index(name="alumnos_e6")
    resumen["cac"] = resumen["alumnos_e6"].apply(
        lambda n: gasto_pauta_mensual_mxn / n if n else None
    )
    return resumen.sort_values("mes")


def excede_dos_meses_seguidos(historial_cac: pd.DataFrame, umbral: float = CAC_E6_UMBRAL) -> bool:
    if len(historial_cac) < 2:
        return False
    ultimos_dos = historial_cac.tail(2)["cac"]
    return bool((ultimos_dos > umbral).all())


# ------------------------------------------------------------------
# Alumno Equivalente (AE)
# ------------------------------------------------------------------

def alumno_equivalente(alumnos_activos: pd.DataFrame) -> pd.DataFrame:
    """AE por plaza/programa = ingreso_real / (activos * cuota_plena)."""
    if alumnos_activos.empty:
        return pd.DataFrame(columns=["plaza", "programa", "total_activos", "ingreso_mensual_real_mxn", "ae"])

    ultimo_corte = alumnos_activos["fecha_corte"].max()
    df = alumnos_activos[alumnos_activos["fecha_corte"] == ultimo_corte].copy()
    df = df.groupby(["plaza", "programa"], as_index=False).agg(
        total_activos=("total_activos", "sum"),
        ingreso_mensual_real_mxn=("ingreso_mensual_real_mxn", "sum"),
    )
    df["cuota_plena"] = df.apply(lambda r: cuota_plena(r["plaza"], r["programa"]), axis=1)
    df["ae"] = df["ingreso_mensual_real_mxn"] / (df["total_activos"] * df["cuota_plena"])
    return df


def ae_portafolio(alumnos_activos: pd.DataFrame) -> float:
    df = alumno_equivalente(alumnos_activos)
    if df.empty:
        return 0.0
    ingreso_total = df["ingreso_mensual_real_mxn"].sum()
    potencial_total = (df["total_activos"] * df["cuota_plena"]).sum()
    return ingreso_total / potencial_total if potencial_total else 0.0


# ------------------------------------------------------------------
# Fuga por descuentos
# ------------------------------------------------------------------

def fuga_por_descuentos(con: duckdb.DuckDBPyConnection) -> pd.DataFrame:
    return con.execute(
        """
        SELECT estrategia_asociada AS estrategia, SUM(monto_mensual_mxn) AS fuga_mxn
        FROM descuentos_otorgados_t
        GROUP BY estrategia_asociada
        ORDER BY fuga_mxn DESC
        """
    ).fetchdf()


def fuga_mes_actual_vs_anterior(con: duckdb.DuckDBPyConnection, hoy: dt.date) -> dict:
    inicio_mes = hoy.replace(day=1)
    mes_anterior_fin = inicio_mes - dt.timedelta(days=1)
    inicio_mes_anterior = mes_anterior_fin.replace(day=1)

    actual = con.execute(
        "SELECT COALESCE(SUM(monto_mensual_mxn), 0) FROM descuentos_otorgados_t WHERE CAST(fecha AS DATE) >= ?",
        [inicio_mes],
    ).fetchone()[0]
    anterior = con.execute(
        "SELECT COALESCE(SUM(monto_mensual_mxn), 0) FROM descuentos_otorgados_t WHERE CAST(fecha AS DATE) >= ? AND CAST(fecha AS DATE) < ?",
        [inicio_mes_anterior, inicio_mes],
    ).fetchone()[0]
    return {"mes_actual_mxn": actual, "mes_anterior_mxn": anterior}


# ------------------------------------------------------------------
# Cobranza (KPI 10) y lectura de KPI_Manual
# ------------------------------------------------------------------

def cobranza_pct(alumnos_activos: pd.DataFrame) -> dict:
    """KPI 10: cobrado / facturable del corte más reciente. Meta >=90%."""
    if alumnos_activos.empty or "facturable_mxn" not in alumnos_activos.columns:
        return {"facturable_mxn": 0.0, "cobrado_mxn": 0.0, "pct": 0.0}

    df = alumnos_activos.copy()
    df["fecha_corte"] = pd.to_datetime(df["fecha_corte"])
    ultimo_corte = df["fecha_corte"].max()
    corte = df[df["fecha_corte"] == ultimo_corte]

    facturable = float(corte["facturable_mxn"].sum())
    cobrado = float(corte["cobrado_mxn"].sum())
    pct = cobrado / facturable if facturable else 0.0
    return {"facturable_mxn": facturable, "cobrado_mxn": cobrado, "pct": pct}


def valor_kpi_manual(hojas: dict, kpi: str, plaza: str = "General", semana_iso: str | None = None) -> float | None:
    """Regresa el valor más reciente capturado en KPI_Manual para (kpi, plaza).
    Si semana_iso es None, usa la semana más reciente disponible para ese kpi/plaza."""
    df = hojas.get("KPI_Manual", pd.DataFrame())
    if df.empty:
        return None
    filtro = df[(df["kpi"] == kpi) & (df["plaza"] == plaza)]
    if filtro.empty:
        return None
    if semana_iso:
        filtro = filtro[filtro["semana_iso"] == semana_iso]
        if filtro.empty:
            return None
    else:
        filtro = filtro.sort_values("semana_iso")
    try:
        return float(filtro.iloc[-1]["valor"])
    except (TypeError, ValueError):
        return None
