// PPL Tracker — Google Apps Script Endpoint
// ─────────────────────────────────────────
// INSTRUCCIONES:
// 1. Abrí script.google.com → Nuevo proyecto
// 2. Pegá todo este código
// 3. Click en "Implementar" → "Nueva implementación"
// 4. Tipo: Aplicación web
// 5. Ejecutar como: Yo
// 6. Quién tiene acceso: Cualquier usuario
// 7. Click "Implementar" → copiás la URL que te da
// 8. En la app, tocás ⚡ y pegás esa URL
// ─────────────────────────────────────────

const SHEET_SESIONES = "Sesiones";
const SHEET_NUTRICION = "Nutrición";

// ── PUNTO DE ENTRADA ──────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var tipo = payload.tipo;
    var result;

    if (tipo === "sesion") {
      result = guardarSesion(payload);
    } else if (tipo === "nutricion") {
      result = guardarNutricion(payload);
    } else {
      result = { ok: false, error: "tipo desconocido: " + tipo };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite CORS desde cualquier origen
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, status: "PPL Tracker activo" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GUARDAR SESIÓN ──────────────────────────
function guardarSesion(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESIONES);
  
  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SESIONES);
    var headers = ["Sesión","Tipo","Fecha","Semana","Sets completados","Sets totales","% Completado","Datos JSON","Timestamp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#c9a84c")
      .setFontColor("#0a0906")
      .setFontWeight("bold")
      .setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(8, 300);
  }
  
  var pct = data.sets_totales > 0 
    ? Math.round((data.sets_completados / data.sets_totales) * 100) + "%" 
    : "0%";
  
  sheet.appendRow([
    data.sesion || "",
    data.tipo || "",
    data.fecha || "",
    data.semana || "",
    data.sets_completados || 0,
    data.sets_totales || 0,
    pct,
    data.datos_json || "",
    new Date().toISOString()
  ]);
  
  return { ok: true, sheet: SHEET_SESIONES, row: sheet.getLastRow() };
}

// ── GUARDAR NUTRICIÓN ──────────────────────
function guardarNutricion(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NUTRICION);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NUTRICION);
    var headers = ["Fecha","Momento","Comida","Notas","Kcal total","Prot total","Timestamp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#33aa77")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 300);
  }
  
  var momentoMap = {
    des: "Desayuno", med: "Snack", alm: "Almuerzo",
    mer: "Snack", pos: "Post-entreno", cen: "Cena", ext: "Snack"
  };
  
  var comidas = data.comidas || [];
  comidas.forEach(function(c) {
    sheet.appendRow([
      data.fecha || "",
      momentoMap[c.id] || c.id || "",
      c.nombre || "",
      c.notas || "",
      c.kcal || 0,
      c.prot || 0,
      new Date().toISOString()
    ]);
  });
  
  return { ok: true, sheet: SHEET_NUTRICION, rows: comidas.length };
}
