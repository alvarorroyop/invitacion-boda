// Columnas de la hoja. Cada dato se escribe en la columna cuyo ENCABEZADO coincide
// (no por posición), así se pueden mover, borrar o añadir columnas sin romper nada.
var HEADERS = [
  "Fecha / Hora",
  "Asistencia",
  "Nombre",
  "Apellidos",
  "Alergias / Intolerancias",
  "Otras Necesidades",
  "Autobús / Transporte",
  "Mensaje / Comentarios"
];

// Pestaña donde se guardan las respuestas (si se renombra, se usa la primera pestaña)
var RESPONSES_SHEET = "Hoja 1";
var SUMMARY_SHEET = "Resumen";

function getResponsesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(RESPONSES_SHEET) || ss.getSheets()[0];
}

// Compara encabezados ignorando mayúsculas, tildes y espacios
function normalizeHeader(text) {
  return String(text)
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, "");
}

function doPost(e) {
  try {
    var sheet = getResponsesSheet();

    // Crear encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground("#425442");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Obtener datos soportando tanto JSON como form-urlencoded
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var values = {};
    values[normalizeHeader("Fecha / Hora")] = data.submittedAt || new Date().toLocaleString("es-ES");
    values[normalizeHeader("Asistencia")] = data.attendance === "yes" ? "SÍ ASISTE" : (data.attendance === "no" ? "NO ASISTE" : (data.attendance || ""));
    values[normalizeHeader("Nombre")] = data.firstName || "";
    values[normalizeHeader("Apellidos")] = data.lastName || "";
    values[normalizeHeader("Alergias / Intolerancias")] = data.allergies || "Ninguna";
    values[normalizeHeader("Otras Necesidades")] = data.specialNeeds || "";
    values[normalizeHeader("Autobús / Transporte")] = Array.isArray(data.transport) ? data.transport.join(", ") : (data.transport || "");
    values[normalizeHeader("Mensaje / Comentarios")] = data.comments || "";

    // Construir la fila siguiendo el orden real de los encabezados de la hoja
    var lastColumn = sheet.getLastColumn();
    var headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var row = headerRow.map(function (header) {
      var key = normalizeHeader(header);
      return values.hasOwnProperty(key) ? values[key] : "";
    });

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Guardado correctamente" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/*
 * EJECUTAR UNA SOLA VEZ desde el editor (desplegable de funciones → configurarHoja → Ejecutar).
 * - Pinta de verde las filas "SÍ ASISTE" y de rojo las "NO ASISTE" (también las que lleguen después).
 * - Crea la pestaña "Resumen" con el contador de asistentes, que se actualiza solo.
 * Se puede volver a ejecutar sin duplicar nada.
 */
function configurarHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getResponsesSheet();
  var sheetName = sheet.getName();

  // Localizar la columna "Asistencia" por su encabezado
  var lastColumn = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var attendanceIndex = -1;
  for (var i = 0; i < headerRow.length; i++) {
    if (normalizeHeader(headerRow[i]) === normalizeHeader("Asistencia")) { attendanceIndex = i; break; }
  }
  if (attendanceIndex === -1) throw new Error('No encuentro la columna "Asistencia" en la fila 1 de "' + sheetName + '".');
  var col = columnLetter(attendanceIndex + 1);

  // --- Colores por fila ---
  var dataRange = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), lastColumn);
  var yesFormula = '=$' + col + '2="SÍ ASISTE"';
  var noFormula = '=$' + col + '2="NO ASISTE"';

  // Conservar otras reglas que hubiera, quitando solo las nuestras para no duplicarlas
  var rules = sheet.getConditionalFormatRules().filter(function (rule) {
    var cond = rule.getBooleanCondition();
    if (!cond) return true;
    var values = cond.getCriteriaValues();
    return !(values.length && (values[0] === yesFormula || values[0] === noFormula));
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(yesFormula)
    .setBackground("#D9EAD3").setFontColor("#1E4620")
    .setRanges([dataRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(noFormula)
    .setBackground("#F4CCCC").setFontColor("#7A1010")
    .setRanges([dataRange]).build());
  sheet.setConditionalFormatRules(rules);

  // --- Pestaña Resumen con contadores ---
  var summary = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, ss.getSheets().length);
  summary.clear();
  var ref = "'" + sheetName + "'!" + col + "2:" + col;

  summary.getRange("A1:B1").setValues([["Resumen de confirmaciones", ""]]);
  summary.getRange("A1:B1").merge().setFontWeight("bold").setFontSize(14)
    .setBackground("#425442").setFontColor("#FFFFFF");

  summary.getRange("A3").setValue("✅ Sí asisten");
  summary.getRange("B3").setFormula('=COUNTIF(' + ref + ',"SÍ ASISTE")');
  summary.getRange("A3:B3").setBackground("#D9EAD3").setFontColor("#1E4620");

  summary.getRange("A4").setValue("❌ No asisten");
  summary.getRange("B4").setFormula('=COUNTIF(' + ref + ',"NO ASISTE")');
  summary.getRange("A4:B4").setBackground("#F4CCCC").setFontColor("#7A1010");

  summary.getRange("A6").setValue("Total respuestas");
  summary.getRange("B6").setFormula("=B3+B4");
  summary.getRange("A6:B6").setFontWeight("bold");

  summary.getRange("A3:B6").setFontSize(12);
  summary.getRange("B3:B6").setHorizontalAlignment("center").setFontWeight("bold");
  summary.setColumnWidth(1, 220);
  summary.setColumnWidth(2, 90);
}

function columnLetter(n) {
  var s = "";
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function doGet(e) {
  return ContentService.createTextOutput("El webhook de la boda Paula y Agustín está activo y listo para recibir datos.");
}
