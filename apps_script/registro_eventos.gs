/**
 * Registro automático de eventos — Excel de Ventas CEA
 *
 * Instala este script en el Google Sheet "Excel de Ventas CEA":
 *   1. Abrir el Sheet → Extensiones → Apps Script.
 *   2. Borrar el contenido de Código.gs y pegar este archivo completo.
 *   3. Guardar (icono de disco) y ponerle un nombre al proyecto, por ejemplo
 *      "Registro de eventos CEA".
 *   4. En el editor, correr la función `onEdit` una vez manualmente (▶) para
 *      forzar la pantalla de autorización. Elegir la cuenta de Google dueña
 *      del Sheet y aceptar los permisos ("Ver y administrar hojas de cálculo").
 *   5. A partir de aquí, cada vez que alguien cambie a mano la columna
 *      "etapa" en la hoja Prospectos, se agregará automáticamente una fila
 *      en la hoja Eventos con el before/after, fecha/hora y el correo de
 *      quien editó.
 *
 * IMPORTANTE: la hoja Eventos la llena SOLO este script. Nunca editarla a
 * mano — el ETL del dashboard usa Eventos para validar que cada alumno tenga
 * un evento que respalde su etapa actual.
 */

const HOJA_ORIGEN = 'Prospectos';
const HOJA_DESTINO = 'Eventos';
const COLUMNA_ID = 'id';
const COLUMNA_ETAPA = 'etapa';

function onEdit(e) {
  const hoja = e.range.getSheet();
  if (hoja.getName() !== HOJA_ORIGEN) return;

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const colEtapaIdx = encabezados.indexOf(COLUMNA_ETAPA) + 1;
  const colIdIdx = encabezados.indexOf(COLUMNA_ID) + 1;

  if (colEtapaIdx === 0 || e.range.getColumn() !== colEtapaIdx) return;
  if (e.range.getNumRows() > 1) return; // edición individual, no pegado masivo

  const fila = e.range.getRow();
  if (fila === 1) return; // encabezado

  const idProspecto = hoja.getRange(fila, colIdIdx).getValue();
  const etapaAnterior = e.oldValue || '';
  const etapaNueva = e.value || '';
  const usuario = Session.getActiveUser().getEmail() || 'desconocido';

  const libro = e.source;
  const hojaEventos = libro.getSheetByName(HOJA_DESTINO);
  if (!hojaEventos) return;

  hojaEventos.appendRow([idProspecto, etapaAnterior, etapaNueva, new Date(), usuario]);
}
