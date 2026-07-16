"""ETL: DuckDB en memoria — limpieza, validación y datos sintéticos de demo."""
from __future__ import annotations

import datetime as dt

import duckdb
import numpy as np
import pandas as pd
import streamlit as st

from src import sheets
from src.constants import (
    APROBADORES,
    CANALES,
    ESTRATEGIAS,
    ETAPAS,
    KPIS_MANUALES,
    KR_IDS,
    MOTIVOS_PERDIDA,
    PLAZAS,
    PROGRAMAS,
    RESPONSABLES,
    TIPOS_DESCUENTO,
    VIGENCIA_DESCUENTO,
)

# Columnas con PII: se descartan antes de cualquier otra operación.
COLUMNAS_PII = ["telefono", "notas"]


# ------------------------------------------------------------------
# Carga de datos: Sheet real o demo sintético
# ------------------------------------------------------------------

def cargar_datos() -> tuple[dict[str, pd.DataFrame], bool]:
    """Regresa (hojas, es_demo). Si no hay credenciales de Google, genera demo."""
    if sheets.hay_credenciales():
        try:
            return sheets.leer_hojas(), False
        except Exception as exc:  # credenciales inválidas, sheet no compartido, etc.
            st.warning(f"No se pudo leer el Excel de Ventas real ({exc}). Usando modo demo.")
    return generar_datos_demo(), True


@st.cache_data(ttl=600)
def generar_datos_demo(seed: int = 42) -> dict[str, pd.DataFrame]:
    """Genera ~80 prospectos, 12 semanas de metas y cortes de alumnos activos
    100→112, con distribución realista entre canales, plazas y etapas."""
    rng = np.random.default_rng(seed)
    hoy = dt.date(2026, 7, 13)
    semanas = [hoy - dt.timedelta(weeks=w) for w in range(11, -1, -1)]  # 12 semanas

    n = 80
    ids = [f"P-{i:04d}" for i in range(1, n + 1)]
    fechas_registro = [
        semanas[rng.integers(0, len(semanas))] - dt.timedelta(days=int(rng.integers(0, 7)))
        for _ in range(n)
    ]
    plazas = rng.choice(PLAZAS, size=n, p=[0.60, 0.25, 0.15])
    programas = rng.choice(PROGRAMAS, size=n)
    canal_pesos = np.array([0.10, 0.06, 0.05, 0.07, 0.05, 0.22, 0.10, 0.04, 0.06, 0.05, 0.14, 0.06])
    canal_pesos = canal_pesos / canal_pesos.sum()
    canales_origen = rng.choice(CANALES, size=n, p=canal_pesos)

    etapa_pesos = [0.28, 0.22, 0.16, 0.10, 0.15, 0.09]  # suma con Perdido al final
    etapas = rng.choice(ETAPAS, size=n, p=etapa_pesos)

    responsables = rng.choice(RESPONSABLES, size=n)
    motivos = [
        MOTIVOS_PERDIDA[rng.integers(0, len(MOTIVOS_PERDIDA))] if e == "Perdido" else ""
        for e in etapas
    ]
    referidos_por = [
        ids[rng.integers(0, i)] if canales_origen[i] == "E4" and i > 0 else ""
        for i in range(n)
    ]

    prospectos = pd.DataFrame(
        {
            "id": ids,
            "fecha_registro": fechas_registro,
            "plaza": plazas,
            "programa": programas,
            "canal_origen": canales_origen,
            "referido_por": referidos_por,
            "etapa": etapas,
            "motivo_perdida": motivos,
            "responsable": responsables,
            "telefono": [f"222{rng.integers(1000000, 9999999)}" for _ in range(n)],
            "notas": ["" for _ in range(n)],
        }
    )

    # --- Eventos: transición de etapas para prospectos que avanzaron ---
    orden_etapas = ["Prospecto", "Contactado", "Cita agendada", "Inscripción provisional", "Alumno"]
    eventos_rows = []
    for _, row in prospectos.iterrows():
        etapa_final = row["etapa"]
        if etapa_final == "Perdido":
            # se perdió en algún punto intermedio
            corte = rng.integers(1, 4)
            secuencia = orden_etapas[:corte] + ["Perdido"]
        elif etapa_final in orden_etapas:
            corte = orden_etapas.index(etapa_final) + 1
            secuencia = orden_etapas[:corte]
        else:
            secuencia = ["Prospecto"]

        fecha_base = dt.datetime.combine(row["fecha_registro"], dt.time(9, 0))
        for i in range(1, len(secuencia)):
            fecha_base = fecha_base + dt.timedelta(days=int(rng.integers(1, 5)))
            eventos_rows.append(
                {
                    "id_prospecto": row["id"],
                    "etapa_anterior": secuencia[i - 1],
                    "etapa_nueva": secuencia[i],
                    "fecha_hora": fecha_base,
                    "usuario": row["responsable"],
                }
            )
    eventos = pd.DataFrame(eventos_rows)

    # --- Metas: 12 semanas x (E1..E10 + Total) ---
    metas_rows = []
    for idx, semana in enumerate(semanas):
        semana_iso = f"{semana.isocalendar().year}-W{semana.isocalendar().week:02d}"
        total_meta_alumnos = 0
        for estrategia in ESTRATEGIAS:
            meta_alumnos = int(rng.integers(1, 4))
            meta_prospectos = meta_alumnos * int(rng.integers(3, 6))
            total_meta_alumnos += meta_alumnos
            comentario = ""
            if idx < len(semanas) - 1 and rng.random() < 0.3:
                comentario = "Semana corta por puente; agenda de citas se recorrió."
            metas_rows.append(
                {
                    "semana_iso": semana_iso,
                    "estrategia": estrategia,
                    "meta_alumnos": meta_alumnos,
                    "meta_prospectos": meta_prospectos,
                    "responsable": RESPONSABLES[idx % len(RESPONSABLES)],
                    "comentario_cierre": comentario,
                }
            )
        metas_rows.append(
            {
                "semana_iso": semana_iso,
                "estrategia": "Total",
                "meta_alumnos": total_meta_alumnos,
                "meta_prospectos": sum(r["meta_prospectos"] for r in metas_rows if r["semana_iso"] == semana_iso),
                "responsable": "Eduardo",
                "comentario_cierre": "",
            }
        )
    metas = pd.DataFrame(metas_rows)

    # --- Alumnos_Activos: cortes semanales 100 -> 112 ---
    from src.kpis import cuota_plena as _cuota_plena

    activos_rows = []
    total_inicial = 100
    total_final = 112
    # Distribución fija entre plaza/programa: se calcula una sola vez para que
    # el total semanal siga exactamente la rampa 100->112 (sin ruido de
    # redondeo que arruine la pendiente de la proyección lineal).
    combos = [(p, prog) for p in PLAZAS for prog in PROGRAMAS]
    pesos = rng.dirichlet(np.ones(len(combos)))
    orden_pesos = np.argsort(-pesos)

    for idx, semana in enumerate(semanas):
        avance = total_inicial + round((total_final - total_inicial) * idx / (len(semanas) - 1))
        cuentas = np.floor(avance * pesos).astype(int)
        residual = avance - int(cuentas.sum())
        for j in range(residual):
            cuentas[orden_pesos[j % len(combos)]] += 1

        for i, (plaza, programa) in enumerate(combos):
            cuenta = int(cuentas[i])
            if cuenta <= 0:
                continue
            cuota_plena_valor = _cuota_plena(plaza, programa)
            factor_real = rng.uniform(0.82, 1.0)
            ingreso_real = round(cuenta * cuota_plena_valor * factor_real, 2)
            # facturable = lo que debería entrar según lista de precios vigente
            # y descuentos otorgados (aprox. 85-95% de precio pleno, en promedio).
            factor_descuento_promedio = rng.uniform(0.85, 0.95)
            facturable = round(cuenta * cuota_plena_valor * factor_descuento_promedio, 2)
            # cobrado = lo que sí entró de lo facturable (meta de cobranza >=90%).
            factor_cobranza = rng.uniform(0.80, 1.0)
            cobrado = round(facturable * factor_cobranza, 2)
            activos_rows.append(
                {
                    "fecha_corte": semana,
                    "plaza": plaza,
                    "programa": programa,
                    "total_activos": cuenta,
                    "ingreso_mensual_real_mxn": ingreso_real,
                    "facturable_mxn": facturable,
                    "cobrado_mxn": cobrado,
                }
            )
    alumnos_activos = pd.DataFrame(activos_rows)

    # --- Descuentos_Otorgados ---
    estrategias_con_descuento = ["E2", "E3", "E4", "E5", "E9"]
    m = 18
    alumnos_ids = prospectos.loc[prospectos["etapa"] == "Alumno", "id"].tolist()
    vigencia_por_tipo = {
        "Beca %": "Permanente",
        "Condonación adeudo": "Única vez",
        "Mensualidad 50% referente": "Única vez",
        "Inscripción 50% referido": "Única vez",
        "Inscripción gratis campaña": "Única vez",
    }
    tipos_elegidos = rng.choice(TIPOS_DESCUENTO, size=m)
    folios = (
        rng.choice(alumnos_ids, size=m) if alumnos_ids else [f"P-{rng.integers(1, 80):04d}" for _ in range(m)]
    )
    descuentos = pd.DataFrame(
        {
            "fecha": [
                hoy - dt.timedelta(days=int(rng.integers(0, 60))) for _ in range(m)
            ],
            "folio_alumno": folios,
            "plaza": rng.choice(PLAZAS, size=m),
            "programa": rng.choice(PROGRAMAS, size=m),
            "estrategia_asociada": rng.choice(estrategias_con_descuento, size=m),
            "tipo": tipos_elegidos,
            "monto_mensual_mxn": rng.integers(150, 900, size=m).astype(float),
            "vigencia": [vigencia_por_tipo[t] for t in tipos_elegidos],
            "aprobado_por": rng.choice(APROBADORES, size=m),
            "evidencia": [f"folio-{i:03d}" for i in range(m)],
        }
    )

    # --- KPI_Manual: captura semanal de los 9 KPIs que no salen de otra hoja ---
    kpis_por_plaza = {"Reseñas Google #", "Reseñas Google promedio", "Citas agendadas", "Citas que asistieron"}
    rangos_kpi = {
        "Reseñas Google #": (5, 30),
        "Reseñas Google promedio": (4.2, 4.8),
        "NPS": (30, 65),
        "NPS tasa de respuesta %": (40, 75),
        "Mediana 1ª respuesta (min)": (2, 12),
        "Citas agendadas": (3, 12),
        "Citas que asistieron": (2, 10),
        "Gasto pauta MXN": (5000, 9000),
        "Bajas del mes": (0, 5),
    }
    kpi_manual_rows = []
    for semana in semanas[-8:]:
        semana_iso = f"{semana.isocalendar().year}-W{semana.isocalendar().week:02d}"
        for kpi in KPIS_MANUALES:
            lo, hi = rangos_kpi[kpi]
            plazas_kpi = PLAZAS if kpi in kpis_por_plaza else ["General"]
            for plaza in plazas_kpi:
                valor = round(rng.uniform(lo, hi), 1) if isinstance(lo, float) else int(rng.integers(lo, hi + 1))
                kpi_manual_rows.append(
                    {
                        "semana_iso": semana_iso,
                        "kpi": kpi,
                        "plaza": plaza,
                        "valor": valor,
                        "capturado_por": RESPONSABLES[rng.integers(0, len(RESPONSABLES))],
                    }
                )
    kpi_manual = pd.DataFrame(kpi_manual_rows)

    # --- OKR_Avance ---
    okr_avance = pd.DataFrame(
        {
            "kr_id": KR_IDS,
            "avance": [round(float(rng.uniform(0.1, 0.8)), 2) for _ in KR_IDS],
            "fecha_actualizacion": [hoy for _ in KR_IDS],
            "comentario": ["" for _ in KR_IDS],
        }
    )

    # --- Comité_Lunes ---
    lunes_recientes = [s for s in semanas[-6:]]
    kpis_rojos_pool = [
        "CAC de E6 sobre umbral", "Cobranza bajo 90%", "Mediana 1ª respuesta sobre 5 min",
        "Bajas del mes sobre meta", "NPS bajo 50",
    ]
    comite_rows = []
    for idx, lunes in enumerate(lunes_recientes):
        semana_iso = f"{lunes.isocalendar().year}-W{lunes.isocalendar().week:02d}"
        alumnos_ese_dia = 100 + idx * 2
        rojos = list(rng.choice(kpis_rojos_pool, size=2, replace=False))
        comite_rows.append(
            {
                "fecha": lunes,
                "semana_iso": semana_iso,
                "asistentes": "Eduardo, Paola, C. Gutiérrez, Obsidian",
                "alumnos_activos_hoy": alumnos_ese_dia,
                "kpis_en_rojo": ", ".join(rojos),
                "krs_movidos": "KR1.2 subió por inscripciones de la semana.",
                "decision_de_la_semana": "Reforzar seguimiento de citas agendadas sin confirmar.",
                "responsable_decision": RESPONSABLES[rng.integers(0, len(RESPONSABLES))],
                "fecha_limite": lunes + dt.timedelta(days=7),
                "decision_semana_pasada_cumplida": ["Sí", "Parcial", "No"][rng.integers(0, 3)],
                "nota_para_direccion": "",
            }
        )
    comite_lunes = pd.DataFrame(comite_rows)

    config = pd.DataFrame(
        {
            "lista": ["etapas"] * len(ETAPAS)
            + ["canales"] * len(CANALES)
            + ["plazas"] * len(PLAZAS)
            + ["programas"] * len(PROGRAMAS)
            + ["responsables"] * len(RESPONSABLES)
            + ["aprobadores"] * len(APROBADORES)
            + ["tipos_descuento"] * len(TIPOS_DESCUENTO)
            + ["motivos_perdida"] * len(MOTIVOS_PERDIDA)
            + ["vigencia_descuento"] * len(VIGENCIA_DESCUENTO)
            + ["estrategias"] * len(ESTRATEGIAS)
            + ["kr_ids"] * len(KR_IDS)
            + ["kpis_manuales"] * len(KPIS_MANUALES),
            "valor": (
                ETAPAS + CANALES + PLAZAS + PROGRAMAS + RESPONSABLES + APROBADORES
                + TIPOS_DESCUENTO + MOTIVOS_PERDIDA + VIGENCIA_DESCUENTO + ESTRATEGIAS
                + KR_IDS + KPIS_MANUALES
            ),
        }
    )

    return {
        "Prospectos": prospectos,
        "Eventos": eventos,
        "Metas": metas,
        "Alumnos_Activos": alumnos_activos,
        "Descuentos_Otorgados": descuentos,
        "Config": config,
        "OKR_Avance": okr_avance,
        "KPI_Manual": kpi_manual,
        "Comité_Lunes": comite_lunes,
    }


# ------------------------------------------------------------------
# DuckDB: registro, limpieza y validación
# ------------------------------------------------------------------

def construir_duckdb(hojas: dict[str, pd.DataFrame]) -> duckdb.DuckDBPyConnection:
    """Registra las hojas en DuckDB, descartando PII antes de cualquier otra cosa."""
    con = duckdb.connect(database=":memory:")

    prospectos = hojas.get("Prospectos", pd.DataFrame()).copy()
    for col in COLUMNAS_PII:
        if col in prospectos.columns:
            prospectos = prospectos.drop(columns=[col])

    con.register("prospectos", prospectos)
    con.register("eventos", hojas.get("Eventos", pd.DataFrame()))
    con.register("metas", hojas.get("Metas", pd.DataFrame()))
    con.register("alumnos_activos", hojas.get("Alumnos_Activos", pd.DataFrame()))
    con.register("descuentos_otorgados", hojas.get("Descuentos_Otorgados", pd.DataFrame()))
    con.register("config", hojas.get("Config", pd.DataFrame()))
    con.register("okr_avance", hojas.get("OKR_Avance", pd.DataFrame()))
    con.register("kpi_manual", hojas.get("KPI_Manual", pd.DataFrame()))
    con.register("comite_lunes", hojas.get("Comité_Lunes", pd.DataFrame()))

    # Materializar como tablas propias (register solo crea una vista sobre el DF)
    tablas = [
        "prospectos", "eventos", "metas", "alumnos_activos", "descuentos_otorgados",
        "config", "okr_avance", "kpi_manual", "comite_lunes",
    ]
    for tabla in tablas:
        con.execute(f"CREATE OR REPLACE TABLE {tabla}_t AS SELECT * FROM {tabla}")

    return con


def validar_calidad(con: duckdb.DuckDBPyConnection) -> dict[str, int]:
    """Corre las reglas de calidad de datos y regresa {regla: filas_rechazadas}."""
    etapas_sql = ", ".join(f"'{e}'" for e in ETAPAS)
    plazas_sql = ", ".join(f"'{p}'" for p in PLAZAS)
    hoy = dt.date(2026, 7, 13).isoformat()

    resultados = {}

    resultados["Etapas fuera de catálogo"] = con.execute(
        f"SELECT COUNT(*) FROM prospectos_t WHERE etapa NOT IN ({etapas_sql})"
    ).fetchone()[0]

    resultados["Plazas fuera de catálogo"] = con.execute(
        f"SELECT COUNT(*) FROM prospectos_t WHERE plaza NOT IN ({plazas_sql})"
    ).fetchone()[0]

    resultados["Fechas de registro futuras"] = con.execute(
        f"SELECT COUNT(*) FROM prospectos_t WHERE CAST(fecha_registro AS DATE) > DATE '{hoy}'"
    ).fetchone()[0]

    resultados["IDs duplicados"] = con.execute(
        "SELECT COUNT(*) FROM (SELECT id, COUNT(*) c FROM prospectos_t GROUP BY id HAVING COUNT(*) > 1)"
    ).fetchone()[0]

    resultados["Alumnos sin evento que lo respalde"] = con.execute(
        """
        SELECT COUNT(*) FROM prospectos_t p
        WHERE p.etapa = 'Alumno'
        AND NOT EXISTS (
            SELECT 1 FROM eventos_t e
            WHERE e.id_prospecto = p.id AND e.etapa_nueva = 'Alumno'
        )
        """
    ).fetchone()[0]

    resultados["Descuentos sobre el tope de 20%"] = con.execute(
        """
        WITH ultimo_corte AS (
            SELECT plaza, programa, facturable_mxn, total_activos,
                   ROW_NUMBER() OVER (PARTITION BY plaza, programa ORDER BY fecha_corte DESC) AS rn
            FROM alumnos_activos_t
        )
        SELECT COUNT(*) FROM descuentos_otorgados_t d
        JOIN ultimo_corte a ON a.plaza = d.plaza AND a.programa = d.programa AND a.rn = 1
        WHERE d.monto_mensual_mxn > 0.20 * (a.facturable_mxn / NULLIF(a.total_activos, 0))
        """
    ).fetchone()[0]

    return resultados
