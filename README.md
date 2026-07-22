# Tablero Comercial CEA

Dashboard comercial para CEA Universidad (Puebla, Huauchinango, Zacatlán).
Streamlit + Google Sheets (Excel de Ventas) + DuckDB. Soporta el ritual semanal: metas por
semana, status cada 3 días, decisión de matar/escalar canales con datos.

## Arranque rápido (modo demo)

Sin ninguna configuración, la app levanta con datos sintéticos:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Contraseña por defecto en modo local: crea `.streamlit/secrets.toml` (ver
plantilla abajo) o define `APP_PASSWORD` — sin ese secreto, el login nunca
deja pasar a nadie.

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
- **`Config`**: proteger la hoja completa, solo Alfonso edita. Agregar un
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
| Asesoras | `Prospectos` | Editor |
| Control Escolar | `Alumnos_Activos`, parte de `KPI_Manual` (NPS, bajas) | Editor |
| Paola | `Alumnos_Activos` (ingreso/facturable/cobrado), `Descuentos_Otorgados` | Editor |
| C. Gutiérrez | `KPI_Manual` (reseñas, prospectos por canal de pauta) | Editor |
| Alfonso | `Config`, `Metas`, `OKR_Avance`, `Comité_Lunes` (en vivo, comité) | Editor |
| Eduardo | Todo | Editor (o Lector si solo revisa el tablero) |

## 5. Instalar el Apps Script de registro de eventos

Ver instrucciones completas dentro de `apps_script/registro_eventos.gs`.
Resumen: Extensiones → Apps Script → pegar el archivo → guardar → correr
`onEdit` una vez para autorizar. A partir de ahí, cada cambio manual en la
columna `etapa` de `Prospectos` queda registrado en `Eventos` automáticamente.
Este script no necesita ningún cambio para las hojas nuevas (`KPI_Manual`,
`OKR_Avance`, `Comité_Lunes`) — esas se llenan directamente, no por
transición de etapa.

## 6. Crear el Service Account de Google Cloud

1. Google Cloud Console → crear proyecto (o usar uno existente).
2. Habilitar la API de Google Sheets y la API de Google Drive.
3. IAM y administración → Cuentas de servicio → Crear cuenta de servicio.
4. Crear una clave (JSON) para esa cuenta — se descarga un archivo `.json`.
5. Copiar el `client_email` de ese JSON y **compartir el Google Sheet** con
   esa dirección (permiso de Lector basta, ya que la app solo lee).
6. Copiar el contenido del JSON a los campos `[gcp_service_account]` de
   `.streamlit/secrets.toml` (ver plantilla en
   `.streamlit/secrets.toml.example`).
7. Copiar el ID del Sheet (de la URL) al campo `SHEET_ID`.

## 7. Deploy en Streamlit Community Cloud

El repo ya vive en `github.com/Alfonso-Intelligence/tablero-comercial-cea` (privado).

1. Entrar a https://share.streamlit.io con "Sign in with GitHub". Recomendado:
   usar la cuenta de GitHub de **Alfonso-Intelligence** (dueña del repo) —
   es la opción más limpia porque el deploy queda atado a la organización.
   La cuenta `AlfonsoAG`, como colaboradora del repo, también puede desplegar,
   pero entonces el deploy queda ligado a esa cuenta personal.
2. Si Streamlit no lista el repo, revisar en la pantalla de autorización
   OAuth de GitHub que se concedió acceso a la organización
   **Alfonso-Intelligence** (botón "Grant" junto al nombre de la org).
3. New app → repo `Alfonso-Intelligence/tablero-comercial-cea` → branch
   `main` → archivo principal `app.py`.
4. Advanced settings → Python version: **3.11**.
5. En Advanced settings → Secrets: pegar el contenido de `secrets.toml`
   (con las credenciales reales, siguiendo la plantilla de
   `.streamlit/secrets.toml.example`) — `APP_PASSWORD` como mínimo, y
   `SHEET_ID` + `[gcp_service_account]` cuando ya existan.
6. Deploy. Sin `SHEET_ID`/`gcp_service_account` configurados la app corre en
   **MODO DEMO** con datos sintéticos — es suficiente para la revisión del
   profe. Verificar que el banner amarillo de modo demo NO aparezca una vez
   que los secrets reales estén completos.

## 8. Editar los datos estáticos (JSON versionados)

Los archivos en `data/` (mercado, competencia, participación, precios,
política de descuentos, costos fijos) NO se leen de Sheets — se editan
directamente en el repo y se hace commit/push. Cada cifra trae un campo
`"marca"`: `✅` dato real, `📊` inferido, `🎯` hipótesis — la UI lo muestra
junto al número. Refresco recomendado:
- `mercado.json`: trimestral.
- `competencia.json` / `participacion.json`: cuando cambie el panorama.
- `costos_fijos.json`: en cuanto Paola entregue las cifras reales
  (reemplazar los `null` — cada renglón sin llenar se muestra como
  "Por llenar — Paola").
- `politica_descuentos.json`: solo tras la aprobación de Eduardo (quitar el
  `"estatus": "🎯 Borrador..."`).

## 9. Referencia de columnas por hoja

(Ver también `INSTRUCCIONES` dentro del propio Excel y los comentarios de
cada encabezado — esta tabla es solo un resumen para consulta rápida.)

| Hoja | Columnas | Quién la llena |
|---|---|---|
| `Prospectos` | id, fecha_registro, plaza, programa, canal_origen, referido_por, etapa, motivo_perdida, responsable, telefono, notas | Asesoras |
| `Eventos` | id_prospecto, etapa_anterior, etapa_nueva, fecha_hora, usuario | Nadie — Apps Script |
| `Metas` | semana_iso, estrategia, meta_alumnos, meta_prospectos, responsable, comentario_cierre | Alfonso (en comité) |
| `Alumnos_Activos` | fecha_corte, plaza, programa, total_activos, ingreso_mensual_real_mxn, facturable_mxn, cobrado_mxn | Control Escolar + Paola |
| `Descuentos_Otorgados` | fecha, folio_alumno, plaza, programa, estrategia_asociada, tipo, monto_mensual_mxn, vigencia, aprobado_por, evidencia | Paola |
| `KPI_Manual` | semana_iso, kpi, plaza, valor, capturado_por | Control Escolar + C. Gutiérrez |
| `OKR_Avance` | kr_id, avance, fecha_actualizacion, comentario | Alfonso (en comité) |
| `Comité_Lunes` | fecha, semana_iso, asistentes, alumnos_activos_hoy, kpis_en_rojo, krs_movidos, decision_de_la_semana, responsable_decision, fecha_limite, decision_semana_pasada_cumplida, nota_para_direccion | Alfonso (en comité) |
| `Config` | listas de catálogo (etapas, canales, plazas, programas, responsables, aprobadores, tipos_descuento, motivos_perdida, vigencia_descuento, estrategias, kr_ids, kpis_manuales) + tabla programa→mensualidad_plena | Solo Alfonso |

## Estructura del repo

```
app.py                      # entrypoint: login + navegación
scripts/generar_excel_ventas.py     # genera Excel_de_Ventas_CEA.xlsx (reproducible)
scripts/verificar_excel_ventas.py   # valida la estructura del .xlsx generado
src/auth.py                 # password gate (hmac.compare_digest)
src/sheets.py                # conexión y lectura de Google Sheets (gspread)
src/etl.py                   # DuckDB: limpieza, validación, datos demo
src/kpis.py                  # embudo, AE, CAC, proyecciones, fuga por descuentos
src/ui.py                    # tarjetas, semáforos, badges de marca
tabs/tab1_mercado.py
tabs/tab2_avance.py          # ⭐ tab principal
tabs/tab3_competencia.py
tabs/tab4_participacion.py
tabs/tab5_finanzas.py
data/*.json                  # datos estáticos versionados
apps_script/registro_eventos.gs
```

## Modo demo

Si `st.secrets` no trae `gcp_service_account` + `SHEET_ID`, la app genera
~80 prospectos sintéticos (12 semanas, distribuidos entre canales/plazas/
etapas), metas semanales y cortes de alumnos activos 100→112, con semilla
fija (reproducible). Aparece un banner amarillo "MODO DEMO — sin conexión al
Excel de Ventas" en cada tab.

## Notas de calibración

- El piso de AE del semáforo (0.85 rojo / 0.90 verde) es una propuesta:
  el peor caso razonable (portafolio completo a precio convenio) da AE
  ≈ 0.833, justo debajo del piso. Se ajusta en `src/constants.py` si Eduardo
  quiere más colchón.
- La hoja `Descuentos_Otorgados` (con `aprobado_por` + `evidencia`) no solo
  mide fuga: instala la gobernanza de condonaciones — toda excepción de
  precio debe quedar documentada ahí.
