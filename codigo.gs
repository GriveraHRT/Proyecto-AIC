/**
 * ========================================================
 * PROYECTO AIC - API DE GOOGLE SHEETS v3
 * ========================================================
 * Incluye: Errores, Pizarra (Muestras, Curvas, Urgentes, Recordatorios, Custom),
 * Centros, Maestro de Exámenes y Chat.
 */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = {
    "Errores": ["Fecha", "Día", "Acción", "N° Petición", "Examen", "Usuario"],
    "Pizarra_Muestras": ["ID", "Tipo", "N° Petición", "Examen", "Almacenada"],
    "Pizarra_Curvas": ["ID", "N° Petición", "Validada"],
    "Pizarra_Urgentes": ["ID", "N° Petición", "Validada"],
    "Pizarra_Recordatorios": ["ID", "Texto", "Usuario", "Fecha"],
    "Pizarra_Custom": ["Cuadrante", "ID", "N° Petición", "Detalle"],
    "Centros": ["Centro", "Estado", "Pdte Rev", "Pdte Val", "Responsable"],
    "Maestro_Examenes": ["Examen"],
    "Chat": ["ID", "Usuario", "Mensaje", "Fecha"]
  };

  for (var name in sheets) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, sheets[name].length).setValues([sheets[name]]);
    }
  }

  // Poblar Centros
  var centrosSheet = ss.getSheetByName("Centros");
  if (centrosSheet.getLastRow() <= 1) {
    var centros = [
      ["Cachureo", "NO REVISADO :(", "", "", ""],
      ["San Clemente", "NO REVISADO :(", "", "", ""],
      ["Maule", "NO REVISADO :(", "", "", ""],
      ["Externos", "NO REVISADO :(", "", "", ""],
      ["Hospitalizados", "NO REVISADO :(", "", "", ""],
      ["Ambulatorio", "NO REVISADO :(", "", "", ""],
      ["San Rafael", "NO REVISADO :(", "", "", ""],
      ["Rio Claro", "NO REVISADO :(", "", "", ""],
      ["Pencahue", "NO REVISADO :(", "", "", ""],
      ["Pelarco", "NO REVISADO :(", "", "", ""],
      ["TOMA DE MUESTRA", "NO REVISADO :(", "", "", ""]
    ];
    centrosSheet.getRange(2, 1, centros.length, 5).setValues(centros);
  }

  // Poblar Maestro Exámenes
  var maestroSheet = ss.getSheetByName("Maestro_Examenes");
  if (maestroSheet.getLastRow() <= 1) {
    var examenes = [
      "GLU", "LDL", "HDL", "COL", "TG", "TSH", "T3", "T4", "T4L",
      "HBA1C", "CRE", "URE", "AC. URICO", "BIL TOTAL", "BIL DIRECTA",
      "GOT", "GPT", "GGT", "FA", "LDH", "CK", "CK-MB", "TROPO",
      "NA", "K", "CL", "CA", "MG", "P", "FE", "FERRITINA", "VIT B12",
      "AC. FOLICO", "VIT D", "PTH", "PCR", "VHS", "HEMOGRAMA",
      "PROTROMBINA", "TTPA", "FIBRINOGENO", "INR", "ORINA COMP",
      "PSA", "CEA", "AFP", "CA 19-9", "CA 125", "INSULINA",
      "CORTISOL", "PROLACTINA", "FSH", "LH", "ESTRADIOL",
      "PROGESTERONA", "TESTOSTERONA", "DHEA-S", "AMILASA", "LIPASA",
      "PROTEINAS TOT", "ALBUMINA", "GLOBULINA", "TP", "ELECT. PROT"
    ];
    var rows = examenes.sort().map(function (e) { return [e]; });
    maestroSheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var getSheetData = function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j] ? data[i][j].toString() : "";
      }
      rows.push(obj);
    }
    return rows;
  };

  var getMaestro = function () {
    var sheet = ss.getSheetByName("Maestro_Examenes");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) list.push(data[i][0].toString());
    }
    return list.sort();
  };

  // Solo retornar los últimos 50 mensajes del chat
  var getChatRecent = function () {
    var sheet = ss.getSheetByName("Chat");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var rows = [];
    var start = Math.max(1, data.length - 50);
    for (var i = start; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j] ? data[i][j].toString() : "";
      }
      rows.push(obj);
    }
    return rows;
  };

  var response = {
    errores: getSheetData("Errores"),
    muestras: getSheetData("Pizarra_Muestras"),
    curvas: getSheetData("Pizarra_Curvas"),
    urgentes: getSheetData("Pizarra_Urgentes"),
    recordatorios: getSheetData("Pizarra_Recordatorios"),
    custom: getSheetData("Pizarra_Custom"),
    centros: getSheetData("Centros"),
    maestro_examenes: getMaestro(),
    chat: getChatRecent()
  };

  return ContentService.createTextOutput(JSON.stringify({ success: true, data: response }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var payload = JSON.parse(e.postData.contents);
    var result = handleAction(payload);
    output.setContent(JSON.stringify({ success: true, data: result }));
  } catch (error) {
    output.setContent(JSON.stringify({ success: false, error: error.toString() }));
  }
  return output;
}

function handleAction(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = payload.action;

  if (action === "cerrar_semana") return procesarCierreSemana();

  var sheet = ss.getSheetByName(payload.sheet);
  if (!sheet) throw new Error("No se encontró la hoja: " + payload.sheet);

  if (action === "insert") {
    sheet.appendRow(payload.row);
    return "OK";
  }

  if (action === "update") {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == payload.id) {
        if (payload.row) {
          sheet.getRange(i + 1, 1, 1, payload.row.length).setValues([payload.row]);
        } else if (payload.updates) {
          for (var key in payload.updates) {
            var col = headers.indexOf(key);
            if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(payload.updates[key]);
          }
        }
        return "OK";
      }
    }
    throw new Error("ID no encontrado");
  }

  if (action === "delete") {
    var data2 = sheet.getDataRange().getValues();
    for (var i = 1; i < data2.length; i++) {
      if (data2[i][0].toString() == payload.id) {
        sheet.deleteRow(i + 1);
        return "OK";
      }
    }
    throw new Error("ID no encontrado");
  }

  if (action === "delete_by_col") {
    var data3 = sheet.getDataRange().getValues();
    var headers3 = data3[0];
    var colIdx = headers3.indexOf(payload.column);
    if (colIdx === -1) throw new Error("Columna no encontrada");
    var count = 0;
    for (var i = data3.length - 1; i >= 1; i--) {
      if (data3[i][colIdx].toString() == payload.value) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    return "Eliminadas " + count + " filas";
  }

  throw new Error("Acción desconocida: " + action);
}

function procesarCierreSemana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetErrores = ss.getSheetByName("Errores");
  var data = sheetErrores.getDataRange().getValues();

  if (data.length <= 1) return "No hay datos semanales.";

  var csv = "\uFEFF";
  for (var i = 0; i < data.length; i++) {
    var row = data[i].map(function (item) {
      var str = typeof item === 'string' ? item.replace(/"/g, '""') : String(item);
      if (str.indexOf(',') > -1 || str.indexOf('\n') > -1 || str.indexOf('"') > -1) str = '"' + str + '"';
      return str;
    });
    csv += row.join(",") + "\n";
  }

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var blob = Utilities.newBlob(csv, "text/csv", "Reporte_Errores_" + fecha + ".csv");

  try {
    MailApp.sendEmail({
      to: "grivera@hospitaldetalca.cl",
      subject: "Reporte de Errores - Cierre Semanal " + fecha,
      body: "Adjunto el reporte de errores. El sistema ha limpiado la tabla.",
      attachments: [blob]
    });
  } catch (e) {
    throw new Error("Error correo: " + e.toString());
  }

  var backup = ss.insertSheet("Backup_" + fecha);
  backup.getRange(1, 1, data.length, data[0].length).setValues(data);

  if (sheetErrores.getMaxRows() > 1) {
    sheetErrores.getRange(2, 1, sheetErrores.getMaxRows() - 1, sheetErrores.getMaxColumns()).clearContent();
  }

  var pizarras = ["Pizarra_Muestras", "Pizarra_Curvas", "Pizarra_Urgentes", "Pizarra_Recordatorios", "Pizarra_Custom"];
  for (var p = 0; p < pizarras.length; p++) {
    var s = ss.getSheetByName(pizarras[p]);
    if (s && s.getMaxRows() > 1) {
      s.getRange(2, 1, s.getMaxRows() - 1, s.getMaxColumns()).clearContent();
    }
  }

  var centrosSheet = ss.getSheetByName("Centros");
  if (centrosSheet) {
    var cd = centrosSheet.getDataRange().getValues();
    for (var i = 1; i < cd.length; i++) {
      centrosSheet.getRange(i + 1, 2, 1, 4).setValues([["NO REVISADO :(", "", "", ""]]);
    }
  }

  return "Semana cerrada, correo enviado.";
}
