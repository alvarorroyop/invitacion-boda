function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Crear encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Fecha / Hora",
        "Asistencia",
        "Nombre",
        "Apellidos",
        "Alergias / Intolerancias",
        "Menú Infantil",
        "Otras Necesidades",
        "Autobús / Transporte",
        "Horario Taxi",
        "Mensaje / Comentarios"
      ];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
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
    
    var row = [
      data.submittedAt || new Date().toLocaleString("es-ES"),
      data.attendance === "yes" ? "SÍ ASISTE" : (data.attendance === "no" ? "NO ASISTE" : (data.attendance || "")),
      data.firstName || "",
      data.lastName || "",
      data.allergies || "Ninguna",
      data.childMenu || "No",
      data.specialNeeds || "",
      Array.isArray(data.transport) ? data.transport.join(", ") : (data.transport || "No necesita"),
      data.taxiTime || "",
      data.comments || ""
    ];
    
    sheet.appendRow(row);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Guardado correctamente" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("El webhook de la boda Paula y Agustín está activo y listo para recibir datos.");
}
