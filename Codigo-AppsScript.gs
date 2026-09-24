/**
 * BACKEND del sistema de gestión (Google Apps Script)
 * ---------------------------------------------------------------
 * Este código va DENTRO del Google de cada cliente (pegado en el
 * Apps Script de UNA planilla vacía de Google Sheets).
 *
 * Guarda cada tipo de dato (clientes, catálogo, pedidos, etc.) en
 * su propia pestaña de esa planilla. Cada fila = un registro,
 * guardado como texto.
 *
 * No hay que modificar nada acá adentro: se pega tal cual y listo.
 */

// Cuando la app PIDE datos (leer)
function doGet(e) {
  var tipo = (e && e.parameter && e.parameter.tipo) ? e.parameter.tipo : "";
  var hoja = obtenerHoja(tipo);
  var valores = [];
  if (hoja && hoja.getLastRow() > 0) {
    valores = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues();
  }
  return ContentService
    .createTextOutput(JSON.stringify(valores))
    .setMimeType(ContentService.MimeType.JSON);
}

// Cuando la app GUARDA datos (escribir). Reemplaza toda la pestaña
// de ese tipo con lo que manda la app.
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita que dos guardados se pisen
  } catch (err) {
    return salida({ ok: false, error: "ocupado, probá de nuevo" });
  }
  try {
    var datos   = JSON.parse(e.postData.contents);
    var tipo    = datos.tipo;
    var payload = datos.payload || [];
    var hoja    = obtenerHoja(tipo);

    hoja.clearContents();
    if (payload.length > 0) {
      var filas = payload.map(function (obj) { return [JSON.stringify(obj)]; });
      hoja.getRange(1, 1, filas.length, 1).setValues(filas);
    }
    return salida({ ok: true, guardados: payload.length });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Busca la pestaña de ese tipo; si no existe, la crea.
function obtenerHoja(tipo) {
  var nombre = "datos_" + (tipo || "sin_tipo");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  return hoja;
}

function salida(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
