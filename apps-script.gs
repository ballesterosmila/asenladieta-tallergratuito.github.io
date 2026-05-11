/**
 * GOOGLE APPS SCRIPT — Backend del formulario del taller
 * As en la Dieta · Alexandra Serna
 *
 * QUÉ HACE:
 *   Recibe el email del formulario de la landing y lo guarda en una
 *   pestaña de tu Google Sheets, con la fecha y hora de inscripción.
 *   Evita duplicados (si el mismo email se inscribe dos veces, no se
 *   añade otra fila — sólo se actualiza la última fecha).
 *
 * CÓMO USARLO (5 minutos · una sola vez):
 *
 *   1) Crea una hoja de cálculo en https://sheets.new
 *      Renómbrala a "Inscripciones taller" (opcional).
 *
 *   2) En esa hoja: Extensiones → Apps Script.
 *      Borra el código de ejemplo y pega TODO este archivo.
 *
 *   3) (Opcional) Cambia abajo SHEET_NAME si tu pestaña no se llama "Hoja 1".
 *      Por defecto Google Sheets en español la nombra "Hoja 1".
 *
 *   4) Guarda (Ctrl/Cmd + S). Ponle nombre al proyecto, p. ej.
 *      "Inscripciones taller Alexandra".
 *
 *   5) Pulsa "Implementar" (arriba a la derecha) → "Nueva implementación".
 *        · Tipo:               Aplicación web
 *        · Descripción:        v1
 *        · Ejecutar como:      Yo (tu cuenta)
 *        · Quién tiene acceso: Cualquier usuario   ← IMPORTANTE
 *      Pulsa "Implementar".
 *
 *   6) Google te pedirá permisos la primera vez. Acepta:
 *        "Revisar permisos" → tu cuenta → "Avanzado" → "Ir a [proyecto]
 *         (no seguro)" → "Permitir".  (Es seguro: es tu propio script.)
 *
 *   7) Te dará una URL del tipo:
 *        https://script.google.com/macros/s/AKfycbx.../exec
 *      Cópiala y pégamela — yo la pongo en index.html.
 *
 *  PARA PROBARLO MANUALMENTE:
 *      En el editor de Apps Script: selecciona la función "test" en el
 *      desplegable y pulsa "Ejecutar". Verás aparecer una fila en tu
 *      hoja con un email de prueba.
 */

// ─────────────────────────────────────────────────────────────
// CONFIGURACIÓN
// ─────────────────────────────────────────────────────────────
const SHEET_NAME = 'Hoja 1';            // pestaña de Google Sheets
const TIMEZONE   = 'Europe/Madrid';     // zona horaria de España

// ─────────────────────────────────────────────────────────────
// ENDPOINT que recibe los POST desde la landing
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse({ ok: false, error: 'sheet_not_found' });
    }

    // El front envía text/plain con un JSON dentro (para evitar preflight CORS).
    const payload = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    const email = (payload.email || '').toString().trim().toLowerCase();

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return jsonResponse({ ok: false, error: 'invalid_email' });
    }

    ensureHeader_(sheet);

    const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    // Anti-duplicados: si el email ya existe, actualizamos su última inscripción.
    const data = sheet.getDataRange().getValues(); // [ [fecha, email], ... ]
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if ((data[i][1] || '').toString().trim().toLowerCase() === email) {
        existingRow = i + 1; // 1-based
        break;
      }
    }

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1).setValue(now); // sólo refresca la fecha
      return jsonResponse({ ok: true, status: 'already_subscribed' });
    }

    sheet.appendRow([now, email]);
    return jsonResponse({ ok: true, status: 'created' });

  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// Ping de salud: abrir la URL en el navegador devuelve {ok:true}
function doGet() {
  return jsonResponse({ ok: true, service: 'asenladieta-taller' });
}

// ─────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────
function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['fecha', 'email']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// TEST manual: ejecuta esta función desde el editor para probar.
// Después borra la fila de prueba si quieres.
// ─────────────────────────────────────────────────────────────
function test() {
  const fakeRequest = {
    postData: { contents: JSON.stringify({ email: 'test@asenladieta.es' }) }
  };
  const res = doPost(fakeRequest);
  Logger.log(res.getContent());
}
