# Excel de Ventas CEA — Plantilla + Ritual Semanal Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate a reproducible `scripts/generar_excel_ventas.py` (openpyxl) that produces `Excel_de_Ventas_CEA.xlsx` — a 10-sheet Google-Sheets-ready workbook with dropdowns, header comments, conditional formatting and example rows — and extend the existing Streamlit dashboard (`src/`, `tabs/`) so it actually reads the two sheets the spec adds (`KPI_Manual`, `Comité_Lunes`) and the new columns on existing sheets, then update the README with the operational rollout steps.

**Architecture:** This repo (`/home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards/`) already has a working Streamlit tablero reading a Google Sheet called "Excel de Ventas CEA" through `src/sheets.py` → `src/etl.py` → DuckDB → `tabs/*.py`. The workbook spec introduces 3 columns/renames on existing sheets and 2 wholly new sheets (`KPI_Manual`, `Comité_Lunes`) plus `INSTRUCCIONES`. Per the spec's own rule ("si detectas discrepancia con el código existente, ajusta el código, no los nombres de este documento"), the dashboard code is the thing that moves to match the spec's column names — not the other way around. The `.xlsx` generator is a standalone, declarative script: a list of sheet-spec dicts (columns, comments, validation, example rows) consumed by shared helper functions, so every sheet is built the same way instead of 10 near-duplicate blocks.

**Tech Stack:** Python 3.12, openpyxl (new dependency), pandas, duckdb, streamlit (existing).

## Global Constraints

- Sheet/column names must match exactly what `src/sheets.py` / `src/etl.py` read — this plan changes both in lockstep, never one without the other.
- Dates: DD/MM/AAAA display format. Currency: MXN, zero decimals (`#,##0 "MXN"`).
- Zero anglicisms in any user-visible sheet/column/comment text. Never write "CRM" or "firma" (use "acuerdo"/"aprobación").
- Header row (row 1) frozen and `auto_filter` enabled on every data-capture sheet (not `INSTRUCCIONES`).
- Every categorical column is a dropdown sourced from a named range on `Config` — never a hardcoded inline list.
- 3 example rows on every capture sheet (`Prospectos`, `Metas`, `Alumnos_Activos`, `Descuentos_Otorgados`, `KPI_Manual`, `OKR_Avance`, `Comité_Lunes`), gray fill, `id`/key prefixed `EJEMPLO — `.
- `Eventos` is never hand-edited — no example rows there, just the header comment saying so.
- Discount tipos are exactly: `Beca %`, `Condonación adeudo`, `Mensualidad 50% referente`, `Inscripción 50% referido`, `Inscripción gratis campaña` (this supersedes the older demo list `Beca % / Condonación adeudo / Mes gratis / Inscripción gratis` in `src/etl.py`).
- `Descuentos_Otorgados.folio_alumno` (renamed from `id_alumno_anonimo`) holds the alumno's `P-xxxx` id, never a name.

---

### Task 1: Extend `src/constants.py` with the catalog lists the workbook and the demo data both need

**Files:**
- Modify: `src/constants.py`

**Interfaces:**
- Produces: `RESPONSABLES`, `APROBADORES`, `TIPOS_DESCUENTO`, `VIGENCIA_DESCUENTO`, `KR_IDS`, `KPIS_MANUALES` — all consumed by Task 4 (`src/etl.py` demo data), Task 8 (Config sheet + named ranges), and Tasks 9–13 (per-sheet dropdowns).

- [ ] **Step 1: Add the new constants**

Append to `src/constants.py` (after `MOTIVOS_PERDIDA`):

```python
RESPONSABLES = ["Angélica", "Paola", "Marisol", "Diego", "C. Gutiérrez", "Eduardo", "Obsidian", "Control Escolar"]

APROBADORES = ["Eduardo", "Paola"]

TIPOS_DESCUENTO = [
    "Beca %",
    "Condonación adeudo",
    "Mensualidad 50% referente",
    "Inscripción 50% referido",
    "Inscripción gratis campaña",
]

VIGENCIA_DESCUENTO = ["Única vez", "Permanente"]

KR_IDS = [
    "KR1.1", "KR1.2", "KR1.3", "KR1.4",
    "KR2.1", "KR2.2", "KR2.3", "KR2.4", "KR2.5",
    "KR3.1", "KR3.2", "KR3.3", "KR3.4", "KR3.5",
    "KR4.1", "KR4.2", "KR4.3",
]

KPIS_MANUALES = [
    "Reseñas Google #",
    "Reseñas Google promedio",
    "NPS",
    "NPS tasa de respuesta %",
    "Mediana 1ª respuesta (min)",
    "Citas agendadas",
    "Citas que asistieron",
    "Gasto pauta MXN",
    "Bajas del mes",
]
```

- [ ] **Step 2: Verify it imports cleanly**

Run: `cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards && .venv/bin/python -c "from src.constants import RESPONSABLES, APROBADORES, TIPOS_DESCUENTO, VIGENCIA_DESCUENTO, KR_IDS, KPIS_MANUALES; print(len(RESPONSABLES), len(KR_IDS), len(KPIS_MANUALES))"`
Expected: `8 17 9`

- [ ] **Step 3: Commit**

```bash
git add src/constants.py
git commit -m "feat: add discount/responsible/KPI catalog constants for Excel de Ventas"
```

---

### Task 2: Add `openpyxl` to `requirements.txt` and install it

**Files:**
- Modify: `requirements.txt`

- [ ] **Step 1: Add the dependency**

Append `openpyxl==3.1.5` to `requirements.txt` (keep alphabetical-ish grouping consistent with the existing file — just add as a new line at the end).

- [ ] **Step 2: Install into the project venv**

Run: `cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards && .venv/bin/pip install openpyxl==3.1.5`
Expected: `Successfully installed openpyxl-3.1.5`

- [ ] **Step 3: Verify**

Run: `.venv/bin/python -c "import openpyxl; print(openpyxl.__version__)"`
Expected: `3.1.5`

- [ ] **Step 4: Commit**

```bash
git add requirements.txt
git commit -m "chore: add openpyxl dependency for the Excel de Ventas generator"
```

---

### Task 3: Extend `src/sheets.py` to read the two new sheets

**Files:**
- Modify: `src/sheets.py:9`

**Interfaces:**
- Consumes: nothing new.
- Produces: `sheets.HOJAS` now includes `"KPI_Manual"` and `"Comité_Lunes"` — Task 4's `construir_duckdb` and the tabs in Tasks 6–7 rely on these keys existing in the `hojas` dict returned by `cargar_datos()`.

- [ ] **Step 1: Edit the `HOJAS` list**

In `src/sheets.py`, change:

```python
HOJAS = ["Prospectos", "Eventos", "Metas", "Alumnos_Activos", "Descuentos_Otorgados", "Config", "OKR_Avance"]
```

to:

```python
HOJAS = [
    "Prospectos",
    "Eventos",
    "Metas",
    "Alumnos_Activos",
    "Descuentos_Otorgados",
    "Config",
    "OKR_Avance",
    "KPI_Manual",
    "Comité_Lunes",
]
```

- [ ] **Step 2: Verify demo mode still boots (sheets.py's `leer_hojas` only runs with real credentials, so this is just a syntax/import check)**

Run: `.venv/bin/python -c "from src.sheets import HOJAS; assert 'KPI_Manual' in HOJAS and 'Comité_Lunes' in HOJAS; print(HOJAS)"`
Expected: prints the 9-item list with no error.

- [ ] **Step 3: Commit**

```bash
git add src/sheets.py
git commit -m "feat: read KPI_Manual and Comité_Lunes sheets from the Excel de Ventas"
```

---

### Task 4: Align `src/etl.py` demo data with the spec's column names, and add the two new sheets to the demo generator + DuckDB

This is the task that fixes every discrepancy called out in the spec's verification checklist: `referido_por` on Prospectos, `folio_alumno`/`vigencia`/updated `tipo` list on Descuentos_Otorgados, `facturable_mxn`/`cobrado_mxn` on Alumnos_Activos, and brand-new demo rows for `KPI_Manual`, `OKR_Avance`, `Comité_Lunes`.

**Files:**
- Modify: `src/etl.py`

**Interfaces:**
- Consumes: `RESPONSABLES, APROBADORES, TIPOS_DESCUENTO, VIGENCIA_DESCUENTO, KR_IDS, KPIS_MANUALES` from `src.constants` (Task 1).
- Produces: `generar_datos_demo()` returns a dict with keys `Prospectos, Eventos, Metas, Alumnos_Activos, Descuentos_Otorgados, Config, OKR_Avance, KPI_Manual, Comité_Lunes`. `construir_duckdb()` registers all of them as `{tabla}_t` DuckDB tables. `validar_calidad()` gains one new rule. Task 5 (`kpis.py`) and Tasks 6–7 (tabs) consume the new columns/tables by these exact names.

- [ ] **Step 1: Update the import line**

In `src/etl.py`, change:

```python
from src.constants import CANALES, ESTRATEGIAS, ETAPAS, MOTIVOS_PERDIDA, PLAZAS, PROGRAMAS
```

to:

```python
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
```

Delete the now-redundant local `RESPONSABLES_DEMO = ["Angélica", "Paola", "Marisol", "Diego"]` line and replace every use of `RESPONSABLES_DEMO` in the file with `RESPONSABLES` (there are uses in the `responsables = rng.choice(RESPONSABLES_DEMO, ...)` line, the `Metas` loop's `RESPONSABLES_DEMO[idx % len(RESPONSABLES_DEMO)]`, and the `Config` dataframe's `RESPONSABLES_DEMO` reference).

- [ ] **Step 2: Add `referido_por` to `Prospectos`**

Right after the existing `responsables = rng.choice(RESPONSABLES, size=n)` line (which computes `etapas`/`motivos` above it), add:

```python
referidos_por = [
    ids[rng.integers(0, i)] if canales_origen[i] == "E4" and i > 0 else ""
    for i in range(n)
]
```

Then add `"referido_por": referidos_por,` to the `prospectos = pd.DataFrame({...})` dict, placed right after `"canal_origen": canales_origen,` and before `"etapa": etapas,` (matches the spec's column order: canal_origen → referido_por → etapa).

- [ ] **Step 3: Rewrite the `Alumnos_Activos` block to add `facturable_mxn` / `cobrado_mxn`**

Find the loop building `activos_rows` (inside the `for i, (plaza, programa) in enumerate(combos):` block). Replace:

```python
            cuota_plena_valor = _cuota_plena(plaza, programa)
            factor_real = rng.uniform(0.82, 1.0)
            ingreso_real = round(cuenta * cuota_plena_valor * factor_real, 2)
            activos_rows.append(
                {
                    "fecha_corte": semana,
                    "plaza": plaza,
                    "programa": programa,
                    "total_activos": cuenta,
                    "ingreso_mensual_real_mxn": ingreso_real,
                }
            )
```

with:

```python
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
```

- [ ] **Step 4: Rewrite the `Descuentos_Otorgados` block — rename to `folio_alumno`, update `tipo` list, add `vigencia`**

Replace the whole block:

```python
    # --- Descuentos_Otorgados ---
    tipos_descuento = ["Beca %", "Condonación adeudo", "Mes gratis", "Inscripción gratis"]
    estrategias_con_descuento = ["E2", "E3", "E4", "E5", "E9"]
    m = 18
    descuentos = pd.DataFrame(
        {
            "fecha": [
                hoy - dt.timedelta(days=int(rng.integers(0, 60))) for _ in range(m)
            ],
            "id_alumno_anonimo": [f"A-{rng.integers(1000, 9999)}" for _ in range(m)],
            "plaza": rng.choice(PLAZAS, size=m),
            "programa": rng.choice(PROGRAMAS, size=m),
            "estrategia_asociada": rng.choice(estrategias_con_descuento, size=m),
            "tipo": rng.choice(tipos_descuento, size=m),
            "monto_mensual_mxn": rng.integers(150, 900, size=m).astype(float),
            "aprobado_por": rng.choice(["Eduardo", "Paola"], size=m),
            "evidencia": [f"folio-{i:03d}" for i in range(m)],
        }
    )
```

with:

```python
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
```

- [ ] **Step 5: Add `KPI_Manual` demo generation**

Right after the `descuentos = pd.DataFrame(...)` block (still inside `generar_datos_demo`), add:

```python
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
```

- [ ] **Step 6: Add the new list columns to the demo `Config` sheet, and return the new sheets**

Replace the `config = pd.DataFrame({...})` block with:

```python
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
```

This is a single, direct replacement of the old `config = pd.DataFrame(...)` assignment and the old `return {...}` — the file must end up with exactly one `config = pd.DataFrame(...)` line (the one above) and one `return` at the end of `generar_datos_demo`.

- [ ] **Step 7: Register the new tables in `construir_duckdb`**

Replace:

```python
    con.register("prospectos", prospectos)
    con.register("eventos", hojas.get("Eventos", pd.DataFrame()))
    con.register("metas", hojas.get("Metas", pd.DataFrame()))
    con.register("alumnos_activos", hojas.get("Alumnos_Activos", pd.DataFrame()))
    con.register("descuentos_otorgados", hojas.get("Descuentos_Otorgados", pd.DataFrame()))
    con.register("config", hojas.get("Config", pd.DataFrame()))

    # Materializar como tablas propias (register solo crea una vista sobre el DF)
    for tabla in ["prospectos", "eventos", "metas", "alumnos_activos", "descuentos_otorgados", "config"]:
        con.execute(f"CREATE OR REPLACE TABLE {tabla}_t AS SELECT * FROM {tabla}")
```

with:

```python
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
```

- [ ] **Step 8: Add a data-quality rule for discounts over the 20% tope**

In `validar_calidad`, after the `"Alumnos sin evento que lo respalde"` rule, add:

```python
    resultados["Descuentos sobre el tope de 20%"] = con.execute(
        """
        SELECT COUNT(*) FROM descuentos_otorgados_t d
        JOIN alumnos_activos_t a ON a.plaza = d.plaza AND a.programa = d.programa
        WHERE d.monto_mensual_mxn > 0.20 * (a.facturable_mxn / NULLIF(a.total_activos, 0))
        """
    ).fetchone()[0]
```

(This mirrors the workbook's own conditional-formatting rule so the dashboard's "Calidad de datos" expander in Tab 2 catches the same violations Google Sheets highlights in red.)

- [ ] **Step 9: Run the demo generator end-to-end**

Run:
```bash
cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards
.venv/bin/python -c "
from src import etl
hojas = etl.generar_datos_demo()
for k, v in hojas.items():
    print(k, list(v.columns))
con = etl.construir_duckdb(hojas)
print(etl.validar_calidad(con))
"
```
Expected: prints 9 sheet names with columns including `referido_por` (Prospectos), `folio_alumno`/`vigencia` (Descuentos_Otorgados), `facturable_mxn`/`cobrado_mxn` (Alumnos_Activos), and a `validar_calidad` dict with a `"Descuentos sobre el tope de 20%"` key, no traceback.

- [ ] **Step 10: Commit**

```bash
git add src/etl.py
git commit -m "feat: align demo ETL with Excel de Ventas spec (referido_por, folio_alumno, facturable/cobrado, KPI_Manual, OKR_Avance, Comité_Lunes)"
```

---

### Task 5: Add the cobranza KPI (KPI 10) and a KPI_Manual lookup helper to `src/kpis.py`

**Files:**
- Modify: `src/kpis.py`

**Interfaces:**
- Produces: `cobranza_pct(alumnos_activos: pd.DataFrame) -> dict` (keys: `facturable_mxn`, `cobrado_mxn`, `pct`), `valor_kpi_manual(hojas: dict, kpi: str, plaza: str = "General", semana_iso: str | None = None) -> float | None`. Both consumed by Tasks 6–7.

- [ ] **Step 1: Add the functions**

Append to `src/kpis.py`:

```python
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
```

- [ ] **Step 2: Verify against the demo data from Task 4**

Run:
```bash
.venv/bin/python -c "
from src import etl, kpis
hojas = etl.generar_datos_demo()
print(kpis.cobranza_pct(hojas['Alumnos_Activos']))
print(kpis.valor_kpi_manual(hojas, 'Gasto pauta MXN'))
print(kpis.valor_kpi_manual(hojas, 'Citas agendadas', plaza='Puebla'))
"
```
Expected: a dict with a `pct` between 0 and 1, then two numeric (or `None` only if that combination genuinely has no rows — should not happen for these two given Task 4's generator) values, no traceback.

- [ ] **Step 3: Commit**

```bash
git add src/kpis.py
git commit -m "feat: add cobranza KPI (KPI 10) and KPI_Manual lookup helper"
```

---

### Task 6: Wire `KPI_Manual` and cobranza into Tab 2 (Avance vs Meta)

**Files:**
- Modify: `tabs/tab2_avance.py`

**Interfaces:**
- Consumes: `kpis.cobranza_pct`, `kpis.valor_kpi_manual` (Task 5).

- [ ] **Step 1: Prefer `KPI_Manual`'s "Gasto pauta MXN" over the `costos_fijos.json` fallback**

In the "Reglas de corte" section, replace:

```python
    with open(DATA_DIR / "costos_fijos.json", encoding="utf-8") as f:
        costos = json.load(f)
    gasto_pauta = next(
        (r["monto"] for r in costos["renglones"] if r["concepto"] == "Pauta publicitaria E6" and r["monto"]),
        7000,
    )
```

with:

```python
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
```

- [ ] **Step 2: Add a "KPIs manuales y cobranza" section**

After the "¿Por qué no llegamos?" section (right before the "6. Calidad de datos" `with st.expander` block), add:

```python
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
```

- [ ] **Step 3: Manual smoke check in the running app**

Run: `cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards && .venv/bin/streamlit run app.py --server.headless true &` then open the app (or use the `run` skill) and confirm Tab 2 shows the new "Cobranza y KPIs manuales" row with real numbers (demo mode) and no exceptions in the terminal. Stop the server after checking.

- [ ] **Step 4: Commit**

```bash
git add tabs/tab2_avance.py
git commit -m "feat: surface cobranza (KPI 10) and manual KPIs in Tab 2"
```

---

### Task 7: Wire `KPI_Manual` and cobranza into Tab 6 (OKRs y Plan)

**Files:**
- Modify: `tabs/tab6_okrs.py`

**Interfaces:**
- Consumes: `kpis.cobranza_pct`, `kpis.valor_kpi_manual` (Task 5). Needs `con` and `hojas` which `render(con, hojas, es_demo)` already receives — currently only `hojas` is used; `con` is accepted but unused, which is fine, no signature change needed. Needs `from src import kpis` and `from src.etl import ...` — actually just needs `from src import kpis`.

- [ ] **Step 1: Import `kpis`**

Change:

```python
from src import ui
```

to:

```python
from src import kpis, ui
```

- [ ] **Step 2: Add a "Valor actual" column to the KPIs permanentes table sourced from KPI_Manual / cobranza / alumnos activos**

Replace the "6. KPIs permanentes" block:

```python
    # ------------------------------------------------------------------
    # 6. KPIs permanentes
    # ------------------------------------------------------------------
    st.markdown("### KPIs permanentes (tablero de los lunes)")
    st.caption(f"_NPS — {ui.definicion('NPS')}_")
    df_kpi = pd.DataFrame(okrs["kpis_permanentes"])
    st.dataframe(
        df_kpi[["kpi", "fuente", "frecuencia", "dueno", "meta", "contra"]].rename(
            columns={"kpi": "KPI", "fuente": "Fuente", "frecuencia": "Frecuencia",
                     "dueno": "Dueño", "meta": "Meta", "contra": "Contramétrica anti-Goodhart"}
        ),
        use_container_width=True, hide_index=True,
    )
```

with:

```python
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

    def _valor_actual(nombre_kpi: str) -> str:
        if nombre_kpi == "Cobranza % del facturable":
            return f"{cobranza['pct']:.0%}"
        clave_manual = mapa_kpi_manual.get(nombre_kpi)
        if clave_manual:
            valor = kpis.valor_kpi_manual(hojas, clave_manual)
            return f"{valor:g}" if valor is not None else "—"
        return "—"

    df_kpi["valor_actual"] = df_kpi["kpi"].apply(_valor_actual)
    st.dataframe(
        df_kpi[["kpi", "valor_actual", "fuente", "frecuencia", "dueno", "meta", "contra"]].rename(
            columns={"kpi": "KPI", "valor_actual": "Valor actual", "fuente": "Fuente", "frecuencia": "Frecuencia",
                     "dueno": "Dueño", "meta": "Meta", "contra": "Contramétrica anti-Goodhart"}
        ),
        use_container_width=True, hide_index=True,
    )
    st.caption("_\"Valor actual\" se lee de KPI_Manual / Alumnos_Activos cuando el dato existe; \"—\" significa que aún no se ha capturado._")
```

- [ ] **Step 3: Manual smoke check**

With the app running (Task 6 Step 3's server, or restart it), open Tab 6 and confirm the "KPIs permanentes" table now shows a populated "Valor actual" column for Cobranza, Mediana 1ª respuesta, Citas agendadas, Bajas del mes, NPS post-trámite, and Reseñas Google, and "—" for the rest, with no exceptions.

- [ ] **Step 4: Commit**

```bash
git add tabs/tab6_okrs.py
git commit -m "feat: surface KPI_Manual and cobranza values in Tab 6's KPIs permanentes table"
```

---

### Task 8: Start `scripts/generar_excel_ventas.py` — shared helpers, `INSTRUCCIONES`, `Config`

This is the first of six tasks building the openpyxl generator. This task lays down every shared helper the later sheet-builders call, plus the two sheets that don't fit the generic "columns + example rows" mold.

**Files:**
- Create: `scripts/generar_excel_ventas.py`

**Interfaces:**
- Produces: `AUTOR_COMENTARIOS = "Obsidian"`, `GRIS_CALCULADO`, `GRIS_EJEMPLO`, `VERDE_ALUMNO`, `GRIS_PERDIDO`, `FMT_FECHA`, `FMT_MXN`; helper functions `nueva_hoja(wb, nombre, *, con_grid=True)`, `encabezados(ws, columnas)`, `comentario(ws, celda, texto)`, `agregar_validacion_lista(wb, ws, columna_letra, nombre_rango, fila_inicio, fila_fin)`, `agregar_validacion_fecha(ws, columna_letra, fila_inicio, fila_fin)`, `agregar_validacion_decimal(ws, columna_letra, fila_inicio, fila_fin, minimo, maximo)`, `pintar_filas_ejemplo(ws, fila_inicio, fila_fin, num_columnas)`, `congelar_y_filtrar(ws, num_columnas)`, `ajustar_anchos(ws, anchos)`; `construir_config(wb) -> dict[str, tuple[str, int]]` (returns `{nombre_lista: (columna_letra, cantidad_filas)}` so later tasks can build validations against `Config!$X$2:$X$N`); `construir_instrucciones(wb)`. Tasks 9–13 import and call all of these from this same module (they add functions to the same file, not new files).

- [ ] **Step 1: Write the module header, imports, and style constants**

```python
"""Genera Excel_de_Ventas_CEA.xlsx: la plantilla operativa del Tablero Comercial CEA.

Reproducible: correr `python scripts/generar_excel_ventas.py` desde la raíz del
repo regenera el archivo desde cero. Los nombres de hoja y columna DEBEN
coincidir con lo que leen src/sheets.py y src/etl.py — si cambias algo aquí,
cambia también esos dos archivos.
"""
from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

from src.constants import (  # noqa: E402
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

AUTOR_COMENTARIOS = "Obsidian"
SALIDA = REPO_ROOT / "Excel_de_Ventas_CEA.xlsx"

GRIS_CALCULADO = PatternFill("solid", fgColor="EAEAEA")
GRIS_EJEMPLO = PatternFill("solid", fgColor="D9D9D9")
VERDE_ALUMNO = PatternFill("solid", fgColor="C6E0B4")
GRIS_PERDIDO = PatternFill("solid", fgColor="D9D9D9")
ROJO_TOPE = PatternFill("solid", fgColor="F8CBAD")

FUENTE_ENCABEZADO = Font(bold=True)
FMT_FECHA = "DD/MM/AAAA"
FMT_MXN = '#,##0 "MXN"'
```

- [ ] **Step 2: Write the shared helper functions**

```python
def nueva_hoja(wb: Workbook, nombre: str, *, con_grid: bool = True):
    ws = wb.create_sheet(nombre)
    if not con_grid:
        ws.sheet_view.showGridLines = False
    return ws


def encabezados(ws, columnas: list[str]) -> None:
    for i, nombre in enumerate(columnas, start=1):
        celda = ws.cell(row=1, column=i, value=nombre)
        celda.font = FUENTE_ENCABEZADO
    ws.freeze_panes = "A2"


def comentario(ws, celda: str, texto: str) -> None:
    c = Comment(texto, AUTOR_COMENTARIOS)
    c.width = 320
    c.height = 140
    ws[celda].comment = c


def agregar_validacion_lista(wb: Workbook, ws, columna_letra: str, nombre_rango: str, fila_inicio: int, fila_fin: int) -> None:
    dv = DataValidation(type="list", formula1=f"={nombre_rango}", allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")


def agregar_validacion_fecha(ws, columna_letra: str, fila_inicio: int, fila_fin: int) -> None:
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
    dv.error = "Escribe una fecha válida (DD/MM/AAAA)."
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")
    for fila in range(fila_inicio, fila_fin + 1):
        ws[f"{columna_letra}{fila}"].number_format = FMT_FECHA


def agregar_validacion_decimal(ws, columna_letra: str, fila_inicio: int, fila_fin: int, minimo: float, maximo: float) -> None:
    dv = DataValidation(type="decimal", operator="between", formula1=str(minimo), formula2=str(maximo), allow_blank=True)
    dv.error = f"Debe estar entre {minimo} y {maximo}."
    ws.add_data_validation(dv)
    dv.add(f"{columna_letra}{fila_inicio}:{columna_letra}{fila_fin}")


def pintar_filas_ejemplo(ws, fila_inicio: int, fila_fin: int, num_columnas: int) -> None:
    for fila in range(fila_inicio, fila_fin + 1):
        for col in range(1, num_columnas + 1):
            ws.cell(row=fila, column=col).fill = GRIS_EJEMPLO


def congelar_y_filtrar(ws, num_columnas: int) -> None:
    ultima_col = get_column_letter(num_columnas)
    ws.auto_filter.ref = f"A1:{ultima_col}1"


def ajustar_anchos(ws, anchos: dict[str, int]) -> None:
    for columna, ancho in anchos.items():
        ws.column_dimensions[columna].width = ancho
```

- [ ] **Step 3: Write `construir_config`**

```python
def construir_config(wb: Workbook) -> dict[str, tuple[str, int]]:
    """Crea la hoja Config con una columna por lista, y un rango con nombre
    por cada una (usado por las validaciones de lista de las demás hojas).
    Regresa {nombre_lista: (columna_letra, cantidad_de_valores)}."""
    ws = nueva_hoja(wb, "Config")
    comentario(ws, "A1", "Solo Obsidian edita esta hoja. Agregar aquí = aparece en las listas de todas las hojas.")

    listas = {
        "etapas": ETAPAS,
        "canales": CANALES,
        "plazas": PLAZAS,
        "programas": PROGRAMAS,
        "responsables": RESPONSABLES,
        "aprobadores": APROBADORES,
        "tipos_descuento": TIPOS_DESCUENTO,
        "motivos_perdida": MOTIVOS_PERDIDA,
        "vigencia_descuento": VIGENCIA_DESCUENTO,
        "estrategias": ESTRATEGIAS + ["Total"],
        "kr_ids": KR_IDS,
        "kpis_manuales": KPIS_MANUALES,
    }

    rangos: dict[str, tuple[str, int]] = {}
    for i, (nombre_lista, valores) in enumerate(listas.items(), start=1):
        col_letra = get_column_letter(i)
        ws.cell(row=1, column=i, value=nombre_lista).font = FUENTE_ENCABEZADO
        for j, valor in enumerate(valores, start=2):
            ws.cell(row=j, column=i, value=valor)
        rango_nombre = f"lista_{nombre_lista}"
        wb.defined_names[rango_nombre] = f"'Config'!${col_letra}$2:${col_letra}${len(valores) + 1}"
        rangos[nombre_lista] = (col_letra, len(valores))

    # Tabla programa -> mensualidad plena, para el formato condicional de Descuentos.
    with open(REPO_ROOT / "data" / "precios_cea.json", encoding="utf-8") as f:
        precios = json.load(f)["programas"]
    mensualidad_por_programa = {
        "Lic. Administración": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Lic. Contaduría": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Lic. Derecho": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Licenciatura")),
        "Bachillerato": next(p["mensualidad_plena"] for p in precios if p["programa"] == "Bachillerato Sabatino" and p["plaza"] == "Puebla"),
        "Maestría Derecho Penal": next(p["mensualidad_plena"] for p in precios if p["programa"].startswith("Maestría")),
    }
    col_prog = get_column_letter(len(listas) + 2)
    col_mens = get_column_letter(len(listas) + 3)
    ws.cell(row=1, column=len(listas) + 2, value="programa_precio").font = FUENTE_ENCABEZADO
    ws.cell(row=1, column=len(listas) + 3, value="mensualidad_plena").font = FUENTE_ENCABEZADO
    for j, (programa, monto) in enumerate(mensualidad_por_programa.items(), start=2):
        ws.cell(row=j, column=len(listas) + 2, value=programa)
        ws.cell(row=j, column=len(listas) + 3, value=monto)
    wb.defined_names["tabla_mensualidad_plena"] = f"'Config'!${col_prog}$2:${col_mens}${len(mensualidad_por_programa) + 1}"

    ajustar_anchos(ws, {get_column_letter(i): 22 for i in range(1, len(listas) + 4)})
    return rangos
```

- [ ] **Step 4: Write `construir_instrucciones`**

```python
def construir_instrucciones(wb: Workbook) -> None:
    ws = nueva_hoja(wb, "INSTRUCCIONES", con_grid=False)
    ws.column_dimensions["A"].width = 100

    lineas = [
        ("titulo", "Excel de Ventas CEA — cómo se llena"),
        ("texto", ""),
        ("sub", "¿Qué es esto?"),
        ("texto", "Este archivo alimenta el Tablero Comercial CEA que se revisa cada lunes. "
                   "Cada hoja corresponde a una parte del ritual semanal: prospección, avance de "
                   "metas, alumnos activos, descuentos y los KPIs que no salen de ningún otro lado."),
        ("texto", ""),
        ("sub", "Regla de oro"),
        ("texto", "Si no está aquí, no existió. El tablero solo cree lo que esté capturado antes "
                   "del lunes 8:00 am."),
        ("texto", ""),
        ("sub", "Quién llena qué"),
        ("texto", "• Asesoras → hoja Prospectos."),
        ("texto", "• Control Escolar → hoja Alumnos_Activos y parte de KPI_Manual (NPS, bajas)."),
        ("texto", "• Paola → hoja Alumnos_Activos (ingreso/facturable/cobrado), Descuentos_Otorgados y cobranza."),
        ("texto", "• C. Gutiérrez → KPI_Manual (reseñas, prospectos por canal de pauta)."),
        ("texto", "• Obsidian → hojas Metas y OKR_Avance, en vivo durante el comité de los lunes."),
        ("texto", ""),
        ("sub", "Los 3 errores que rompen el tablero"),
        ("texto", "1. Escribir un valor a mano en vez de elegir de la lista desplegable."),
        ("texto", "2. Borrar o reciclar el id de un prospecto."),
        ("texto", "3. Capturar el teléfono en la columna de notas."),
        ("texto", ""),
        ("sub", "Nunca tocar la hoja Eventos"),
        ("texto", "Se llena sola con el Apps Script instalado (ver apps_script/registro_eventos.gs). "
                   "Escribir ahí a mano rompe las métricas de velocidad del embudo."),
        ("texto", ""),
        ("sub", "Aviso de privacidad"),
        ("texto", "Los datos de contacto (teléfono, notas) solo se usan para dar seguimiento a la "
                   "inscripción, conforme a la LFPDPPP. El tablero público nunca muestra nombres ni "
                   "teléfonos — esas columnas se descartan antes de llegar a cualquier gráfica."),
        ("texto", ""),
        ("sub", "Antes de arrancar"),
        ("texto", "Borra las 3 filas que empiezan con \"EJEMPLO —\" en cada hoja de captura. Son "
                   "solo para mostrar cómo se llena, no son datos reales."),
    ]

    fila = 1
    for tipo, texto in lineas:
        celda = ws.cell(row=fila, column=1, value=texto)
        if tipo == "titulo":
            celda.font = Font(bold=True, size=16)
        elif tipo == "sub":
            celda.font = Font(bold=True, size=12)
        else:
            celda.alignment = Alignment(wrap_text=True, vertical="top")
        fila += 1
```

- [ ] **Step 5: Add a temporary `main()` to verify these two sheets build, run it, then check output**

Temporarily append (this `main` gets replaced wholesale in Task 13 — it exists here only so this task is independently testable):

```python
def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)
    construir_instrucciones(wb)
    construir_config(wb)
    wb.save(SALIDA)
    print(f"Escrito {SALIDA} con hojas: {wb.sheetnames}")


if __name__ == "__main__":
    main()
```

Run: `cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards && .venv/bin/python scripts/generar_excel_ventas.py`
Expected: `Escrito .../Excel_de_Ventas_CEA.xlsx con hojas: ['INSTRUCCIONES', 'Config']`, no traceback.

- [ ] **Step 6: Commit**

```bash
git add scripts/generar_excel_ventas.py requirements.txt
git commit -m "feat: start Excel de Ventas generator — shared helpers, INSTRUCCIONES, Config"
```

---

### Task 9: Add `Prospectos` and `Eventos` sheet builders

**Files:**
- Modify: `scripts/generar_excel_ventas.py`

**Interfaces:**
- Consumes: everything from Task 8 (`nueva_hoja`, `encabezados`, `comentario`, `agregar_validacion_lista`, `agregar_validacion_fecha`, `pintar_filas_ejemplo`, `congelar_y_filtrar`, `ajustar_anchos`, `construir_config`'s returned `rangos` dict).
- Produces: `construir_prospectos(wb, rangos)`, `construir_eventos(wb)`. Called from `main()` in Task 13.

- [ ] **Step 1: Write `construir_prospectos`**

```python
def construir_prospectos(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["id", "fecha_registro", "plaza", "programa", "canal_origen", "referido_por",
                "etapa", "motivo_perdida", "responsable", "telefono", "notas"]
    ws = nueva_hoja(wb, "Prospectos")
    encabezados(ws, columnas)

    comentario(ws, "A1", "NO editar ni reciclar. Siguiente número libre. Es la huella del prospecto en todo el sistema.")
    comentario(ws, "B1", "Día en que la persona preguntó por primera vez.")
    comentario(ws, "C1", "Sede donde quiere estudiar, no donde preguntó.")
    comentario(ws, "E1", "¿Cómo nos encontró? Si te dice 'me recomendó X', es E4 Referidos y llenas la col. F. "
                         "Esta columna decide dónde invertimos: no adivines, pregunta.")
    comentario(ws, "F1", "Solo si canal = E4. Sin esto no se paga el premio.")
    comentario(ws, "G1", "Actualízala EN EL MOMENTO. Cada cambio queda registrado automáticamente con fecha (hoja Eventos).")
    comentario(ws, "H1", "Solo si etapa = Perdido. Sé honesta: 'Precio' nos sirve más que 'Otro'.")
    comentario(ws, "J1", "Solo para seguimiento. JAMÁS va al tablero.")
    comentario(ws, "K1", "Contexto útil: horario que le acomoda, empresa donde trabaja, etc.")

    fila_inicio, fila_fin = 2, 4  # 3 filas EJEMPLO
    ejemplos = [
        ["EJEMPLO — P-0001", dt.date(2026, 7, 1), "Puebla", "Lic. Administración", "E6", "",
         "Contactado", "", "Marisol", "2221234567", "Prefiere sábados"],
        ["EJEMPLO — P-0002", dt.date(2026, 7, 2), "Huauchinango", "Bachillerato", "E4", "P-0001",
         "Inscripción provisional", "", "Diego", "2227654321", "Refirió su hermano"],
        ["EJEMPLO — P-0003", dt.date(2026, 6, 20), "Zacatlán", "Lic. Derecho", "E6", "",
         "Perdido", "Precio", "Angélica", "2229998888", ""],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 500  # rango de captura + validaciones
    agregar_validacion_fecha(ws, "B", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "C", "lista_plazas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "D", "lista_programas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "E", "lista_canales", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "G", "lista_etapas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "H", "lista_motivos_perdida", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "I", "lista_responsables", fila_inicio, fila_datos_fin)

    # Formato condicional: etapa=Alumno -> verde, Perdido -> gris.
    from openpyxl.formatting.rule import FormulaRule
    rango_completo = f"A{fila_inicio}:K{fila_datos_fin}"
    ws.conditional_formatting.add(rango_completo, FormulaRule(formula=[f"$G{fila_inicio}=\"Alumno\""], fill=VERDE_ALUMNO))
    ws.conditional_formatting.add(rango_completo, FormulaRule(formula=[f"$G{fila_inicio}=\"Perdido\""], fill=GRIS_PERDIDO))

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 14, "C": 14, "D": 20, "E": 12, "F": 14, "G": 20, "H": 16, "I": 14, "J": 14, "K": 30})
```

- [ ] **Step 2: Write `construir_eventos`**

```python
def construir_eventos(wb: Workbook) -> None:
    columnas = ["id_prospecto", "etapa_anterior", "etapa_nueva", "fecha_hora", "usuario"]
    ws = nueva_hoja(wb, "Eventos")
    encabezados(ws, columnas)
    comentario(ws, "A1", "Historial automático. Si escribes aquí, rompes las métricas de velocidad del embudo.")
    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 20, "C": 20, "D": 20, "E": 20})
```

- [ ] **Step 3: Temporarily wire into `main()` and run**

Add the two calls into the temporary `main()` from Task 8 (after `construir_config(wb)`):

```python
    rangos = construir_config(wb)
    construir_prospectos(wb, rangos)
    construir_eventos(wb)
```

Run: `.venv/bin/python scripts/generar_excel_ventas.py`
Expected: `Escrito ... con hojas: ['INSTRUCCIONES', 'Config', 'Prospectos', 'Eventos']`, no traceback.

- [ ] **Step 4: Spot-check the generated file**

Run:
```bash
.venv/bin/python -c "
import openpyxl
wb = openpyxl.load_workbook('Excel_de_Ventas_CEA.xlsx')
ws = wb['Prospectos']
print([c.value for c in ws[1]])
print(ws['A1'].comment.text[:30])
print([ws.cell(row=r, column=1).value for r in (2,3,4)])
print(len(ws.conditional_formatting._cf_rules))
"
```
Expected: header list starting with `id`, a non-empty comment string, the 3 `EJEMPLO —` ids, and `2` (two conditional formatting rule ranges registered — verifying at least one rule list is non-empty is enough; exact internal count may vary by openpyxl version, so treat "no exception + non-zero" as pass).

- [ ] **Step 5: Commit**

```bash
git add scripts/generar_excel_ventas.py
git commit -m "feat: add Prospectos and Eventos sheet builders to the Excel de Ventas generator"
```

---

### Task 10: Add `Metas` and `Alumnos_Activos` sheet builders

**Files:**
- Modify: `scripts/generar_excel_ventas.py`

**Interfaces:**
- Produces: `construir_metas(wb, rangos)`, `construir_alumnos_activos(wb, rangos)`.

- [ ] **Step 1: Write `construir_metas`**

```python
def construir_metas(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["semana_iso", "estrategia", "meta_alumnos", "meta_prospectos", "responsable", "comentario_cierre"]
    ws = nueva_hoja(wb, "Metas")
    encabezados(ws, columnas)
    comentario(ws, "F1", "Se llena al cierre: si no llegamos, ¿por qué? Una frase honesta. Esto se lee LITERAL "
                         "en el tablero y en el comité.")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        ["EJEMPLO — 2026-W29", "E4", 2, 8, "Diego", ""],
        ["EJEMPLO — 2026-W29", "E6", 1, 6, "Marisol", "Se retrasó la campaña por ajustes de presupuesto."],
        ["EJEMPLO — 2026-W29", "Total", 10, 42, "Eduardo", ""],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 500
    agregar_validacion_lista(wb, ws, "B", "lista_estrategias", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "E", "lista_responsables", fila_inicio, fila_datos_fin)

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 10, "C": 14, "D": 16, "E": 14, "F": 50})


def construir_alumnos_activos(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["fecha_corte", "plaza", "programa", "total_activos", "ingreso_mensual_real_mxn",
                "facturable_mxn", "cobrado_mxn"]
    ws = nueva_hoja(wb, "Alumnos_Activos")
    encabezados(ws, columnas)
    comentario(ws, "D1", "Cuenta de alumnos vigentes al viernes. 5 minutos, una fila por plaza-programa.")
    comentario(ws, "F1", "Lo que DEBERÍA entrar este mes según lista de precios y descuentos vigentes.")
    comentario(ws, "G1", "Lo que SÍ entró. cobrado ÷ facturable = KPI 10 de cobranza (meta ≥90%).")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        ["EJEMPLO — 2026-07-10", "Puebla", "Lic. Administración", 40, 64000, 68000, 61200],
        ["EJEMPLO — 2026-07-10", "Huauchinango", "Bachillerato", 25, 18000, 19000, 17500],
        ["EJEMPLO — 2026-07-10", "Zacatlán", "Bachillerato", 15, 9000, 9500, 8500],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))
    for fila in range(fila_inicio, fila_fin + 1):
        for col_letra in ("E", "F", "G"):
            ws[f"{col_letra}{fila}"].number_format = FMT_MXN

    fila_datos_fin = 500
    agregar_validacion_fecha(ws, "A", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "B", "lista_plazas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "C", "lista_programas", fila_inicio, fila_datos_fin)
    for fila in range(fila_inicio, fila_datos_fin + 1):
        for col_letra in ("E", "F", "G"):
            ws[f"{col_letra}{fila}"].number_format = FMT_MXN

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 14, "C": 20, "D": 14, "E": 20, "F": 16, "G": 16})
```

- [ ] **Step 2: Wire into temporary `main()` and run**

Add after `construir_eventos(wb)`:

```python
    construir_metas(wb, rangos)
    construir_alumnos_activos(wb, rangos)
```

Run: `.venv/bin/python scripts/generar_excel_ventas.py`
Expected: sheet list now includes `Metas` and `Alumnos_Activos`, no traceback.

- [ ] **Step 3: Spot-check**

Run:
```bash
.venv/bin/python -c "
import openpyxl
wb = openpyxl.load_workbook('Excel_de_Ventas_CEA.xlsx')
ws = wb['Alumnos_Activos']
print([c.value for c in ws[1]])
print(ws['G1'].comment.text[:20])
print(ws['F2'].number_format)
"
```
Expected: 7 headers ending in `cobrado_mxn`, a comment string, and `'#,##0 "MXN"'`.

- [ ] **Step 4: Commit**

```bash
git add scripts/generar_excel_ventas.py
git commit -m "feat: add Metas and Alumnos_Activos sheet builders to the Excel de Ventas generator"
```

---

### Task 11: Add the `Descuentos_Otorgados` sheet builder (with the 20%-tope conditional formatting)

**Files:**
- Modify: `scripts/generar_excel_ventas.py`

**Interfaces:**
- Produces: `construir_descuentos(wb, rangos)`.

- [ ] **Step 1: Write `construir_descuentos`**

```python
def construir_descuentos(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["fecha", "folio_alumno", "plaza", "programa", "estrategia_asociada", "tipo",
                "monto_mensual_mxn", "vigencia", "aprobado_por", "evidencia"]
    ws = nueva_hoja(wb, "Descuentos_Otorgados")
    encabezados(ws, columnas)
    comentario(ws, "A1", "Todo descuento vive aquí o no existe. Tope por alumno: 20% del valor anual. "
                         "Nada se acumula. Promociones vencen 28-ago-2026.")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        [dt.date(2026, 7, 5), "EJEMPLO — P-0002", "Huauchinango", "Bachillerato", "E4",
         "Inscripción 50% referido", 300, "Única vez", "Paola", "folio-102"],
        [dt.date(2026, 7, 8), "EJEMPLO — P-0015", "Puebla", "Lic. Contaduría", "E3",
         "Beca %", 180, "Permanente", "Eduardo", "correo-2026-07-08"],
        [dt.date(2026, 7, 10), "EJEMPLO — P-0022", "Zacatlán", "Bachillerato", "E9",
         "Inscripción gratis campaña", 0, "Única vez", "Paola", "acuerdo-tianguis-03"],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 500
    agregar_validacion_fecha(ws, "A", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "C", "lista_plazas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "D", "lista_programas", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "E", "lista_estrategias", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "F", "lista_tipos_descuento", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "H", "lista_vigencia_descuento", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "I", "lista_aprobadores", fila_inicio, fila_datos_fin)
    for fila in range(fila_inicio, fila_datos_fin + 1):
        ws[f"G{fila}"].number_format = FMT_MXN

    # Formato condicional: monto_mensual_mxn > 20% de la mensualidad plena del programa -> rojo.
    from openpyxl.formatting.rule import FormulaRule
    rango_completo = f"A{fila_inicio}:J{fila_datos_fin}"
    formula_tope = (
        f'$G{fila_inicio}>0.2*IFERROR(VLOOKUP($D{fila_inicio},tabla_mensualidad_plena,2,FALSE),9999999)'
    )
    ws.conditional_formatting.add(rango_completo, FormulaRule(formula=[formula_tope], fill=ROJO_TOPE))

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 14, "B": 18, "C": 14, "D": 20, "E": 14, "F": 24, "G": 16, "H": 14, "I": 14, "J": 20})
```

- [ ] **Step 2: Wire into temporary `main()` and run**

Add after `construir_alumnos_activos(wb, rangos)`:

```python
    construir_descuentos(wb, rangos)
```

Run: `.venv/bin/python scripts/generar_excel_ventas.py`
Expected: sheet list now includes `Descuentos_Otorgados`, no traceback.

- [ ] **Step 3: Spot-check**

```bash
.venv/bin/python -c "
import openpyxl
wb = openpyxl.load_workbook('Excel_de_Ventas_CEA.xlsx')
ws = wb['Descuentos_Otorgados']
print([c.value for c in ws[1]])
print(ws['B2'].value, ws['H2'].value)
"
```
Expected: 10 headers starting with `fecha, folio_alumno` and ending `evidencia`; `B2` = `EJEMPLO — P-0002`, `H2` = `Única vez`.

- [ ] **Step 4: Commit**

```bash
git add scripts/generar_excel_ventas.py
git commit -m "feat: add Descuentos_Otorgados sheet builder with 20%-tope conditional formatting"
```

---

### Task 12: Add `KPI_Manual` and `OKR_Avance` sheet builders

**Files:**
- Modify: `scripts/generar_excel_ventas.py`

**Interfaces:**
- Produces: `construir_kpi_manual(wb, rangos)`, `construir_okr_avance(wb, rangos)`.

- [ ] **Step 1: Write `construir_kpi_manual`**

```python
def construir_kpi_manual(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["semana_iso", "kpi", "plaza", "valor", "capturado_por"]
    ws = nueva_hoja(wb, "KPI_Manual")
    encabezados(ws, columnas)
    comentario(ws, "A1", "Solo estos 9 se capturan a mano; todo lo demás lo calcula el tablero. "
                         "Viernes antes de las 6 pm.")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        ["EJEMPLO — 2026-W28", "NPS", "General", 52, "Control Escolar"],
        ["EJEMPLO — 2026-W28", "Citas agendadas", "Puebla", 9, "Marisol"],
        ["EJEMPLO — 2026-W28", "Gasto pauta MXN", "General", 7200, "C. Gutiérrez"],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 500
    agregar_validacion_lista(wb, ws, "B", "lista_kpis_manuales", fila_inicio, fila_datos_fin)

    # KPI_Manual.plaza acepta "General" además de las 3 plazas — eso no vive en
    # ninguna lista de Config todavía, así que se agrega una columna propia
    # ("plazas_o_general") en una columna nueva al final de Config.
    ws_config = wb["Config"]
    col_nueva = ws_config.max_column + 1
    col_letra = get_column_letter(col_nueva)
    ws_config.cell(row=1, column=col_nueva, value="plazas_o_general").font = FUENTE_ENCABEZADO
    valores_plaza_general = PLAZAS + ["General"]
    for i, valor in enumerate(valores_plaza_general, start=2):
        ws_config.cell(row=i, column=col_nueva, value=valor)
    wb.defined_names["lista_plazas_o_general"] = f"'Config'!${col_letra}$2:${col_letra}${len(valores_plaza_general) + 1}"

    agregar_validacion_lista(wb, ws, "C", "lista_plazas_o_general", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "E", "lista_responsables", fila_inicio, fila_datos_fin)

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 26, "C": 16, "D": 12, "E": 16})


def construir_okr_avance(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["kr_id", "avance", "fecha_actualizacion", "comentario"]
    ws = nueva_hoja(wb, "OKR_Avance")
    encabezados(ws, columnas)
    comentario(ws, "B1", "0.0 = nada · 0.5 = a medio camino · 0.7 = éxito · 1.0 sistemático = la meta "
                         "estaba floja. Se actualiza en vivo cada lunes.")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        ["EJEMPLO — KR1.2", 0.35, dt.date(2026, 7, 13), "Vamos a ritmo para llegar a 145 en septiembre."],
        ["EJEMPLO — KR2.5", 0.20, dt.date(2026, 7, 13), "Falta arrancar la encuesta NPS formal."],
        ["EJEMPLO — KR4.1", 0.60, dt.date(2026, 7, 13), "Cobranza mejorando pero aún bajo 90%."],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 500
    agregar_validacion_lista(wb, ws, "A", "lista_kr_ids", fila_inicio, fila_datos_fin)
    agregar_validacion_decimal(ws, "B", fila_inicio, fila_datos_fin, 0.0, 1.0)
    agregar_validacion_fecha(ws, "C", fila_inicio, fila_datos_fin)

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 12, "B": 12, "C": 18, "D": 50})
```

- [ ] **Step 2: Wire into temporary `main()` and run**

Add after `construir_descuentos(wb, rangos)`:

```python
    construir_kpi_manual(wb, rangos)
    construir_okr_avance(wb, rangos)
```

Run: `.venv/bin/python scripts/generar_excel_ventas.py`
Expected: sheet list now includes `KPI_Manual` and `OKR_Avance`, no traceback.

- [ ] **Step 3: Spot-check**

```bash
.venv/bin/python -c "
import openpyxl
wb = openpyxl.load_workbook('Excel_de_Ventas_CEA.xlsx')
print([c.value for c in wb['KPI_Manual'][1]])
print([c.value for c in wb['OKR_Avance'][1]])
ws = wb['Config']
encabezados_config = [c.value for c in ws[1]]
print('plazas_o_general' in encabezados_config)
"
```
Expected: 5 headers for KPI_Manual, 4 for OKR_Avance, and `True` for the Config check (confirms the `plazas_o_general` column got appended regardless of exactly which column letter it landed on).

- [ ] **Step 4: Commit**

```bash
git add scripts/generar_excel_ventas.py
git commit -m "feat: add KPI_Manual and OKR_Avance sheet builders to the Excel de Ventas generator"
```

---

### Task 13: Add `Comité_Lunes`, finalize `main()`, generate the real file, and write a verification script

**Files:**
- Modify: `scripts/generar_excel_ventas.py`
- Create: `scripts/verificar_excel_ventas.py`

- [ ] **Step 1: Write `construir_comite_lunes`**

```python
def construir_comite_lunes(wb: Workbook, rangos: dict[str, tuple[str, int]]) -> None:
    columnas = ["fecha", "semana_iso", "asistentes", "alumnos_activos_hoy", "kpis_en_rojo",
                "krs_movidos", "decision_de_la_semana", "responsable_decision", "fecha_limite",
                "decision_semana_pasada_cumplida", "nota_para_direccion"]
    ws = nueva_hoja(wb, "Comité_Lunes")
    encabezados(ws, columnas)
    comentario(ws, "A1", "30 minutos: 1) ¿cumplimos la decisión pasada? 2) tablero: alumnos, rojos, "
                         "KRs 3) UNA decisión nueva con dueño y fecha. Este historial es la memoria de "
                         "la operación.")
    comentario(ws, "D1", "Se lee del tablero, se anota aquí como acta.")
    comentario(ws, "E1", "Máximo 3, los peores. Del tablero.")
    comentario(ws, "F1", "¿Qué KR subió o bajó esta semana y por qué (qué dato lo movió)?")
    comentario(ws, "G1", "UNA sola decisión ejecutable. Regla del comité: sin decisión no hay sesión.")
    comentario(ws, "J1", "Se revisa ANTES de tomar la nueva.")
    comentario(ws, "K1", "Una frase para Eduardo si hay algo que necesite su decisión o su cartera.")

    fila_inicio, fila_fin = 2, 4
    ejemplos = [
        ["EJEMPLO — 06/07/2026", "2026-W28", "Eduardo, Paola, C. Gutiérrez, Obsidian", 108,
         "Cobranza bajo 90%, CAC de E6 sobre umbral", "KR1.2 bajó por semana corta de puente.",
         "Confirmar convenio COPARMEX antes del viernes.", "Eduardo", dt.date(2026, 7, 10),
         "Sí", ""],
        ["EJEMPLO — 13/07/2026", "2026-W29", "Eduardo, Paola, Obsidian", 110,
         "Mediana 1ª respuesta sobre 5 min", "KR3.1 sin cambio, seguimos sin sostenerlo 4 semanas.",
         "Asesoras revisan WhatsApp cada 2 horas fijas.", "Marisol", dt.date(2026, 7, 17),
         "Parcial", ""],
        ["EJEMPLO — 20/07/2026", "2026-W30", "Eduardo, Paola, C. Gutiérrez", 111,
         "Bajas del mes sobre meta", "KR2.1 empeoró; se investigan las 2 bajas de julio.",
         "Control Escolar llama a las 2 bajas para entender motivo real.", "Control Escolar",
         dt.date(2026, 7, 24), "No", "Necesitamos que Eduardo apruebe presupuesto de retención."],
    ]
    for i, fila_valores in enumerate(ejemplos):
        for j, valor in enumerate(fila_valores, start=1):
            ws.cell(row=fila_inicio + i, column=j, value=valor)
    pintar_filas_ejemplo(ws, fila_inicio, fila_fin, len(columnas))

    fila_datos_fin = 260  # ~5 años de lunes es de sobra
    agregar_validacion_fecha(ws, "A", fila_inicio, fila_datos_fin)
    agregar_validacion_fecha(ws, "I", fila_inicio, fila_datos_fin)
    agregar_validacion_lista(wb, ws, "H", "lista_responsables", fila_inicio, fila_datos_fin)

    dv_cumplida = DataValidation(type="list", formula1='"Sí,No,Parcial"', allow_blank=True)
    ws.add_data_validation(dv_cumplida)
    dv_cumplida.add(f"J{fila_inicio}:J{fila_datos_fin}")

    congelar_y_filtrar(ws, len(columnas))
    ajustar_anchos(ws, {"A": 16, "B": 14, "C": 30, "D": 16, "E": 30, "F": 40, "G": 40, "H": 16, "I": 14, "J": 20, "K": 40})
```

- [ ] **Step 2: Replace the temporary `main()` with the final version, in sheet order matching the spec (`INSTRUCCIONES` first, `Eventos` and `Config` near the end since they're least-touched by day-to-day staff)**

```python
def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    construir_instrucciones(wb)
    rangos = construir_config(wb)  # Config debe existir antes que cualquier hoja con listas
    construir_prospectos(wb, rangos)
    construir_eventos(wb)
    construir_metas(wb, rangos)
    construir_alumnos_activos(wb, rangos)
    construir_descuentos(wb, rangos)
    construir_kpi_manual(wb, rangos)
    construir_okr_avance(wb, rangos)
    construir_comite_lunes(wb, rangos)

    wb.active = 0
    wb.save(SALIDA)
    print(f"Escrito {SALIDA} con hojas: {wb.sheetnames}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Run the full generator**

Run: `cd /home/poncholicious/obsidian-intelligence/ibarragalanis/cea/dashboards && .venv/bin/python scripts/generar_excel_ventas.py`
Expected: `Escrito .../Excel_de_Ventas_CEA.xlsx con hojas: ['INSTRUCCIONES', 'Config', 'Prospectos', 'Eventos', 'Metas', 'Alumnos_Activos', 'Descuentos_Otorgados', 'KPI_Manual', 'OKR_Avance', 'Comité_Lunes']`, no traceback.

- [ ] **Step 4: Write `scripts/verificar_excel_ventas.py`**

```python
"""Verificación estructural de Excel_de_Ventas_CEA.xlsx — corre después de
generar_excel_ventas.py y confirma que el archivo cumple el checklist del spec."""
from __future__ import annotations

from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
ARCHIVO = REPO_ROOT / "Excel_de_Ventas_CEA.xlsx"

HOJAS_ESPERADAS = [
    "INSTRUCCIONES", "Config", "Prospectos", "Eventos", "Metas", "Alumnos_Activos",
    "Descuentos_Otorgados", "KPI_Manual", "OKR_Avance", "Comité_Lunes",
]

HOJAS_CON_EJEMPLOS = [
    "Prospectos", "Metas", "Alumnos_Activos", "Descuentos_Otorgados", "KPI_Manual", "OKR_Avance", "Comité_Lunes",
]


def verificar() -> list[str]:
    errores = []
    wb = openpyxl.load_workbook(ARCHIVO)

    for hoja in HOJAS_ESPERADAS:
        if hoja not in wb.sheetnames:
            errores.append(f"Falta la hoja {hoja}")
    if errores:
        return errores  # sin hojas no tiene caso seguir

    for hoja in wb.sheetnames:
        ws = wb[hoja]
        if hoja == "INSTRUCCIONES":
            continue
        if ws.freeze_panes != "A2":
            errores.append(f"{hoja}: fila 1 no está congelada")
        if hoja != "INSTRUCCIONES" and not ws.auto_filter.ref:
            errores.append(f"{hoja}: no tiene auto_filter")
        primera_fila = [c.value for c in ws[1]]
        if not any(c.comment for c in ws[1]):
            errores.append(f"{hoja}: ningún encabezado tiene comentario")

    for hoja in HOJAS_CON_EJEMPLOS:
        ws = wb[hoja]
        col_a_fila2 = str(ws["A2"].value or "")
        if "EJEMPLO" not in col_a_fila2 and "EJEMPLO" not in str(ws["B2"].value or ""):
            errores.append(f"{hoja}: la fila 2 no parece ser un EJEMPLO")

    return errores


if __name__ == "__main__":
    errores = verificar()
    if errores:
        print("FALLÓ la verificación:")
        for e in errores:
            print(f"  - {e}")
        raise SystemExit(1)
    print("OK — Excel_de_Ventas_CEA.xlsx pasa la verificación estructural.")
```

- [ ] **Step 5: Run the verification script**

Run: `.venv/bin/python scripts/verificar_excel_ventas.py`
Expected: `OK — Excel_de_Ventas_CEA.xlsx pasa la verificación estructural.` If it fails, read the printed errors and fix the corresponding `construir_*` function from Tasks 9–13 before moving on.

- [ ] **Step 6: Commit**

```bash
git add scripts/generar_excel_ventas.py scripts/verificar_excel_ventas.py Excel_de_Ventas_CEA.xlsx
git commit -m "feat: finish Excel de Ventas generator (Comité_Lunes, final main()) and add structural verification script"
```

---

### Task 14: Verify `apps_script/registro_eventos.gs` needs no changes, and say so in its header comment

The script only watches column `etapa` on `Prospectos` and appends to `Eventos`. None of the new/changed sheets (`KPI_Manual`, `Comité_Lunes`, `OKR_Avance`, the new `Descuentos_Otorgados`/`Alumnos_Activos` columns) are event-sourced — they're filled directly by staff — so the trigger logic itself doesn't need to change. This task just confirms that reasoning and records it so a future reader doesn't re-litigate it.

**Files:**
- Modify: `apps_script/registro_eventos.gs`

- [ ] **Step 1: Add one clarifying line to the file's top comment block**

In `apps_script/registro_eventos.gs`, after the existing line:

```
 * IMPORTANTE: la hoja Eventos la llena SOLO este script. Nunca editarla a
 * mano — el ETL del dashboard usa Eventos para validar que cada alumno tenga
 * un evento que respalde su etapa actual.
 */
```

add, still inside the comment block, right before the closing `*/`:

```
 *
 * Este script NO necesita tocar las hojas KPI_Manual, OKR_Avance o
 * Comité_Lunes — esas se llenan directamente a mano por el equipo, no por
 * transición de etapa. Si en el futuro se agrega una hoja que también deba
 * quedar registrada automáticamente por edición, dale su propio bloque en
 * onEdit() en vez de reutilizar HOJA_ORIGEN/HOJA_DESTINO.
 */
```

- [ ] **Step 2: Sanity-check the file still parses as valid JS (no build step in this repo, so this is just an eyeball + syntax check)**

Run: `node --check apps_script/registro_eventos.gs 2>&1 || echo "node not available, skipping"`
Expected: either no output (valid syntax) or the `node not available` fallback line — either is acceptable since this repo has no Node toolchain; the real check is that you only added comment lines, so syntax cannot have broken.

- [ ] **Step 3: Commit**

```bash
git add apps_script/registro_eventos.gs
git commit -m "docs: clarify registro_eventos.gs does not need to cover the new manual-entry sheets"
```

---

### Task 15: Rewrite the README with the Google Sheets rollout steps

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Replace the "## 1. Crear el Google Sheet" section through "## 2. Instalar el Apps Script" section**

Replace everything from `## 1. Crear el Google Sheet "Excel de Ventas CEA"` down to (but not including) `## 3. Crear el Service Account de Google Cloud` with:

```markdown
## 1. Generar y revisar el Excel de Ventas

El archivo `Excel_de_Ventas_CEA.xlsx` se genera con un script reproducible —
nunca se edita el `.xlsx` a mano fuera de Google Sheets una vez importado:

```bash
pip install -r requirements.txt
python scripts/generar_excel_ventas.py
python scripts/verificar_excel_ventas.py
```

Esto crea/actualiza `Excel_de_Ventas_CEA.xlsx` en la raíz del repo con las 10
hojas (`INSTRUCCIONES`, `Config`, `Prospectos`, `Eventos`, `Metas`,
`Alumnos_Activos`, `Descuentos_Otorgados`, `KPI_Manual`, `OKR_Avance`,
`Comité_Lunes`), listas desplegables, comentarios de encabezado y 3 filas
`EJEMPLO —` por hoja de captura. El script de verificación confirma que la
estructura sobrevivió (fila 1 congelada, filtros, comentarios, ejemplos).

## 2. Importar a Google Sheets

1. Ir a [sheets.google.com](https://sheets.google.com) → Archivo en blanco.
2. Archivo → Importar → subir `Excel_de_Ventas_CEA.xlsx`.
3. En el diálogo de importación, elegir **"Reemplazar hoja de cálculo"** (así
   se traen las 10 hojas con sus nombres exactos, no una hoja suelta).
4. Verificar que las listas desplegables sobrevivieron: abrir `Prospectos`,
   columna `plaza`, confirmar que sigue mostrando flecha de selección. Google
   Sheets normalmente conserva las validaciones de lista de openpyxl sin
   problema, pero **los comentarios de celda SÍ se importan como notas** (ícono
   triangular naranja) — es normal que no se vean como "comentarios" con hilo
   de respuesta de Google, son notas de solo lectura.
5. Renombrar el archivo a **"Excel de Ventas CEA"** (Archivo → Cambiar
   nombre) — el tablero busca el Sheet por ID, no por nombre, pero mantener
   el nombre evita confusión del equipo.
6. **Borrar las 3 filas `EJEMPLO —` de cada hoja de captura antes de que el
   equipo empiece a usarlo.** Son solo para enseñar cómo se llena.

## 3. Proteger las hojas Eventos y Config

En Google Sheets: clic derecho en la pestaña de la hoja → **Proteger hoja**.

- **`Eventos`**: proteger la hoja completa, solo el dueño (o la cuenta del
  Apps Script) puede editar. El equipo nunca debe escribir aquí a mano.
- **`Config`**: proteger la hoja completa, solo Obsidian edita. Agregar un
  valor aquí es lo único que hace aparecer una opción nueva en los
  desplegables de las demás hojas.
- **Celdas calculadas** (fondo gris en el resto de las hojas, si las hay):
  seleccionar el rango → Datos → Rangos protegidos → dar acceso solo al
  dueño de esa hoja.

## 4. Compartir por rol

Compartir el Sheet (botón "Compartir", esquina superior derecha) con permiso
de **Editor** para quien captura, y **Lector** para quien solo consulta:

| Persona/rol | Hoja(s) que llena | Permiso sugerido |
|---|---|---|
| Asesoras (Angélica, Marisol, Diego) | `Prospectos` | Editor |
| Control Escolar | `Alumnos_Activos`, parte de `KPI_Manual` (NPS, bajas) | Editor |
| Paola | `Alumnos_Activos` (ingreso/facturable/cobrado), `Descuentos_Otorgados` | Editor |
| C. Gutiérrez | `KPI_Manual` (reseñas, prospectos por canal de pauta) | Editor |
| Obsidian | `Config`, `Metas`, `OKR_Avance`, `Comité_Lunes` (en vivo, comité) | Editor |
| Eduardo | Todo | Editor (o Lector si solo revisa el tablero) |

## 5. Instalar el Apps Script de registro de eventos

Ver instrucciones completas dentro de `apps_script/registro_eventos.gs`.
Resumen: Extensiones → Apps Script → pegar el archivo → guardar → correr
`onEdit` una vez para autorizar. A partir de ahí, cada cambio manual en la
columna `etapa` de `Prospectos` queda registrado en `Eventos` automáticamente.
Este script no necesita ningún cambio para las hojas nuevas (`KPI_Manual`,
`OKR_Avance`, `Comité_Lunes`) — esas se llenan directamente, no por
transición de etapa.
```

- [ ] **Step 2: Renumber the remaining sections and refresh the schema tables**

The old `## 3. Crear el Service Account` becomes `## 6.`, and `## 4. Deploy en Streamlit Community Cloud` becomes `## 7.`. Update both headers accordingly (just the numbers — the content of those two sections doesn't change).

Then replace the old `## 5. Editar los datos estáticos` section's number to `## 8.`, and immediately after it (before `## Estructura del repo`), insert the updated schema reference (the old inline `### Hoja X` tables scattered through the original `## 1.` section are gone now that Section 1 points at the generator script — this new section is their replacement, kept for anyone who wants the column reference without opening the `.xlsx`):

```markdown
## 9. Referencia de columnas por hoja

(Ver también `INSTRUCCIONES` dentro del propio Excel y los comentarios de
cada encabezado — esta tabla es solo un resumen para consulta rápida.)

| Hoja | Columnas | Quién la llena |
|---|---|---|
| `Prospectos` | id, fecha_registro, plaza, programa, canal_origen, referido_por, etapa, motivo_perdida, responsable, telefono, notas | Asesoras |
| `Eventos` | id_prospecto, etapa_anterior, etapa_nueva, fecha_hora, usuario | Nadie — Apps Script |
| `Metas` | semana_iso, estrategia, meta_alumnos, meta_prospectos, responsable, comentario_cierre | Obsidian (en comité) |
| `Alumnos_Activos` | fecha_corte, plaza, programa, total_activos, ingreso_mensual_real_mxn, facturable_mxn, cobrado_mxn | Control Escolar + Paola |
| `Descuentos_Otorgados` | fecha, folio_alumno, plaza, programa, estrategia_asociada, tipo, monto_mensual_mxn, vigencia, aprobado_por, evidencia | Paola |
| `KPI_Manual` | semana_iso, kpi, plaza, valor, capturado_por | Control Escolar + C. Gutiérrez |
| `OKR_Avance` | kr_id, avance, fecha_actualizacion, comentario | Obsidian (en comité) |
| `Comité_Lunes` | fecha, semana_iso, asistentes, alumnos_activos_hoy, kpis_en_rojo, krs_movidos, decision_de_la_semana, responsable_decision, fecha_limite, decision_semana_pasada_cumplida, nota_para_direccion | Obsidian (en comité) |
| `Config` | listas de catálogo (etapas, canales, plazas, programas, responsables, aprobadores, tipos_descuento, motivos_perdida, vigencia_descuento, estrategias, kr_ids, kpis_manuales) + tabla programa→mensualidad_plena | Solo Obsidian |
```

- [ ] **Step 3: Update the repo structure listing to include the new script files**

In the `## Estructura del repo` code block, add these two lines right after `app.py`:

```
scripts/generar_excel_ventas.py     # genera Excel_de_Ventas_CEA.xlsx (reproducible)
scripts/verificar_excel_ventas.py   # valida la estructura del .xlsx generado
```

- [ ] **Step 4: Re-read the whole file once to make sure section numbers are sequential and nothing dangles**

Run: `grep -n "^## " README.md`
Expected: a sequential list `## 1.` through `## 9.` (plus the un-numbered `## Estructura del repo`, `## Modo demo`, `## Notas de calibración` at the end, unchanged) with no duplicate or skipped numbers.

- [ ] **Step 5: Commit**

```bash
git add README.md
git commit -m "docs: rewrite README with Excel de Ventas rollout steps (import, protect, share by role)"
```

---

## Self-Review Notes (for whoever executes this plan)

- **Spec coverage:** every sheet/column in the user's spec is covered by Tasks 8–13 (generator) and Task 4 (ETL alignment). The `Descuentos_Otorgados` conditional formatting, the `Prospectos` etapa-based row coloring, the `OKR_Avance` decimal validation, and the `Comité_Lunes` narrative fields are all explicit steps, not left implicit.
- **Dashboard extension requirement:** the spec's verification checklist says "si etl.py no lee KPI_Manual, facturable/cobrado o Comité_Lunes, EXTENDER etl.py y el tablero" — Tasks 4–7 do exactly that (ETL + Tab 2 + Tab 6). `Comité_Lunes` itself is not surfaced in a dedicated tab in this plan (the spec doesn't ask for a new tab, only that the dashboard "extend" to cover the KPIs it drives — its data reaches the dashboard via `hojas["Comité_Lunes"]` and DuckDB's `comite_lunes_t`, available to any future tab without further plumbing).
- **Known risk:** Google Sheets' import of openpyxl `DataValidation` list-type validations sourced from named ranges (`=lista_plazas`) is generally reliable, but cross-sheet named-range validations have occasionally been reported flaky across Sheets versions. Task 15's README Step 1 explicitly tells the team to verify dropdowns survived import as its own checklist item — if a validation doesn't survive, the fix is re-creating that one dropdown in Sheets' own Data → Data validation UI pointing at `Config!$X$2:$X$N`, not a code change.
