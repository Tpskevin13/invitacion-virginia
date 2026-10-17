/**
 * Guarda las confirmaciones de la invitación en una hoja de Google y
 * devuelve cuántas personas confirmaron que van.
 *
 * PASOS (una sola vez):
 * 1. Creá una planilla nueva en Google Sheets (sheets.new).
 * 2. Menú: Extensiones > Apps Script. Borrá lo que haya y pegá TODO este archivo. Guardá.
 * 3. Botón "Implementar" > "Nueva implementación" > engranaje > "Aplicación web".
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier persona
 *    Tocá "Implementar" y aceptá los permisos (si aparece "app no verificada": Avanzado > Ir a ... > Permitir).
 * 4. Copiá la "URL de la aplicación web" (termina en /exec).
 * 5. Pegala en la invitación, en la línea:  const SHEET_URL = "";
 *
 * Si después cambiás este código, hay que hacer "Implementar > Administrar implementaciones >
 * editar > Nueva versión" para que los cambios se apliquen.
 */

const HOJA = 'Confirmaciones';
const COLUMNAS = ['Fecha', 'Nombre', 'Asiste', 'Restricción alimentaria', 'Mensaje'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLUMNAS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
  }
  return sh;
}

// Evita que un texto empiece con = + - @ (se interpretaría como fórmula en la planilla)
function limpiar_(v) {
  v = String(v || '').trim().slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function norm_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
}

// Recibe una confirmación. Si la misma persona confirma de nuevo, se actualiza su fila (no se duplica).
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter || {};
    const nombre = limpiar_(p.nombre);
    if (!nombre) return salida_({ ok: false });
    const fila = [
      new Date(),
      nombre,
      p.asistencia === 'si' ? 'si' : 'no',
      limpiar_(p.dieta),
      limpiar_(p.mensaje)
    ];
    const sh = hoja_();
    const ultima = sh.getLastRow();
    let destino = 0;
    if (ultima > 1) {
      const nombres = sh.getRange(2, 2, ultima - 1, 1).getValues();
      for (let i = 0; i < nombres.length; i++) {
        if (norm_(nombres[i][0]) === norm_(nombre)) { destino = i + 2; break; }
      }
    }
    if (destino) sh.getRange(destino, 1, 1, fila.length).setValues([fila]);
    else sh.appendRow(fila);
    return salida_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Devuelve cuántas personas confirmaron que asisten: { "count": 12 }
function doGet() {
  const sh = hoja_();
  const ultima = sh.getLastRow();
  let count = 0;
  if (ultima > 1) {
    const asiste = sh.getRange(2, 3, ultima - 1, 1).getValues();
    asiste.forEach(r => { if (String(r[0]).toLowerCase() === 'si') count++; });
  }
  return salida_({ count: count });
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
