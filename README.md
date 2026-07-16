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

## 1. Crear el Google Sheet "Excel de Ventas CEA"

Crear un Sheet nuevo con estas hojas y columnas EXACTAS (nombres de columnas
en la primera fila):

### Hoja `Prospectos`
| columna | tipo | validación |
|---|---|---|
| id | texto (P-0001…) | único |
| fecha_registro | fecha | |
| plaza | lista desplegable | Puebla, Huauchinango, Zacatlán |
| programa | lista desplegable | Lic. Administración, Lic. Contaduría, Lic. Derecho, Bachillerato, Maestría Derecho Penal |
| canal_origen | lista desplegable | E1…E10, Orgánico, Recomendación histórica |
| etapa | lista desplegable | Prospecto, Contactado, Cita agendada, Inscripción provisional, Alumno, Perdido |
| motivo_perdida | lista desplegable | Precio, Horario, Competencia, No responde, Otro |
| responsable | lista desplegable | (nombres del equipo, configurable en hoja Config) |
| telefono | texto | ⚠️ nunca se carga al dashboard |
| notas | texto | ⚠️ nunca se carga al dashboard |

Para las listas desplegables: seleccionar la columna → Datos → Validación de
datos → "Lista de un rango" apuntando a la hoja `Config`, o "Lista de
elementos" con los valores pegados directamente.

### Hoja `Eventos` (la llena SOLO el Apps Script — nunca a mano)
`id_prospecto, etapa_anterior, etapa_nueva, fecha_hora, usuario`

### Hoja `Metas`
`semana_iso` (ej. 2026-W33), `estrategia` (E1…E10, Total), `meta_alumnos`,
`meta_prospectos`, `responsable`, `comentario_cierre` (texto libre "¿por qué
no llegamos?")

### Hoja `Alumnos_Activos` (corte semanal manual de Control Escolar)
`fecha_corte, plaza, programa, total_activos, ingreso_mensual_real_mxn`

### Hoja `Descuentos_Otorgados` (gobernanza de condonaciones)
`fecha, id_alumno_anonimo, plaza, programa, estrategia_asociada, tipo (Beca %
/ Condonación adeudo / Mes gratis / Inscripción gratis), monto_mensual_mxn,
aprobado_por, evidencia (URL o folio de correo)`

### Hoja `Config`
Listas para las validaciones: etapas, canales, plazas, programas,
responsables (una columna `lista` con el nombre de la lista y otra `valor`
con cada elemento — ver `src/etl.py::generar_datos_demo` para el formato
exacto que la app espera al leerla).

## 2. Instalar el Apps Script de registro de eventos

Ver instrucciones completas dentro de `apps_script/registro_eventos.gs`.
Resumen: Extensiones → Apps Script → pegar el archivo → guardar → correr
`onEdit` una vez para autorizar. A partir de ahí, cada cambio manual en la
columna `etapa` de `Prospectos` queda registrado en `Eventos` automáticamente.

## 3. Crear el Service Account de Google Cloud

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

## 4. Deploy en Streamlit Community Cloud

El repo ya vive en `github.com/Obsidian-Intelligence/tablero-comercial-cea` (privado).

1. Entrar a https://share.streamlit.io con "Sign in with GitHub". Recomendado:
   usar la cuenta de GitHub de **Obsidian-Intelligence** (dueña del repo) —
   es la opción más limpia porque el deploy queda atado a la organización.
   La cuenta `AlfonsoAG`, como colaboradora del repo, también puede desplegar,
   pero entonces el deploy queda ligado a esa cuenta personal.
2. Si Streamlit no lista el repo, revisar en la pantalla de autorización
   OAuth de GitHub que se concedió acceso a la organización
   **Obsidian-Intelligence** (botón "Grant" junto al nombre de la org).
3. New app → repo `Obsidian-Intelligence/tablero-comercial-cea` → branch
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

## 5. Editar los datos estáticos (JSON versionados)

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
- `politica_descuentos.json`: solo tras la firma de Eduardo (quitar el
  `"estatus": "🎯 Borrador..."`).

## Estructura del repo

```
app.py                      # entrypoint: login + navegación
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
