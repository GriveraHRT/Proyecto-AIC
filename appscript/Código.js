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
    "Pizarra_Muestras": ["ID", "Tipo", "N° Petición", "Examen", "Almacenada", "Fecha"],
    "Pizarra_Curvas": ["ID", "N° Petición", "Validada", "Fecha"],
    "Pizarra_Urgentes": ["ID", "N° Petición", "Validada", "Fecha"],
    "Pizarra_Recordatorios": ["ID", "Texto", "Usuario", "Fecha"],
    "Pizarra_Custom": ["Cuadrante", "ID", "N° Petición", "Detalle"],
    "Centros": ["Centro", "Estado", "Pdte Rev", "Pdte Val", "Rev Hasta", "Responsable"],
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
      ["Cachureo", "NO REVISADO :(", "", "", "", ""],
      ["San Clemente", "NO REVISADO :(", "", "", "", ""],
      ["Maule", "NO REVISADO :(", "", "", "", ""],
      ["Externos", "NO REVISADO :(", "", "", "", ""],
      ["Hospitalizados", "NO REVISADO :(", "", "", "", ""],
      ["Ambulatorio", "NO REVISADO :(", "", "", "", ""],
      ["San Rafael", "NO REVISADO :(", "", "", "", ""],
      ["Rio Claro", "NO REVISADO :(", "", "", "", ""],
      ["Pencahue", "NO REVISADO :(", "", "", "", ""],
      ["Pelarco", "NO REVISADO :(", "", "", "", ""],
      ["TOMA DE MUESTRA", "NO REVISADO :(", "", "", "", ""],
      ["Laboratorio", "NO REVISADO :(", "", "", "", ""]
    ];
    centrosSheet.getRange(2, 1, centros.length, 6).setValues(centros);
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
/**
 * Auto-migrar headers: agrega columnas faltantes a hojas existentes
 */
function migrateHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expected = {
    "Pizarra_Muestras": ["ID", "Tipo", "N° Petición", "Examen", "Almacenada", "Fecha"],
    "Pizarra_Curvas": ["ID", "N° Petición", "Validada", "Fecha"],
    "Pizarra_Urgentes": ["ID", "N° Petición", "Validada", "Fecha"],
    "Centros": ["Centro", "Estado", "Pdte Rev", "Pdte Val", "Rev Hasta", "Responsable"]
  };
  for (var name in expected) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0].filter(function (h) { return h !== ""; });
    var target = expected[name];
    for (var i = 0; i < target.length; i++) {
      if (currentHeaders.indexOf(target[i]) === -1) {
        var nextCol = currentHeaders.length + 1;
        sheet.getRange(1, nextCol).setValue(target[i]);
        currentHeaders.push(target[i]);
      }
    }
  }
}

function doGet(e) {
  migrateHeaders();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var getSheetData = function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      var hasData = false;
      for (var j = 0; j < headers.length; j++) {
        if (headers[j]) {
          obj[headers[j]] = data[i][j] !== undefined && data[i][j] !== null ? data[i][j].toString() : "";
          if (data[i][j]) hasData = true;
        }
      }
      if (hasData) rows.push(obj);
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
  migrateHeaders();
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

  if (action === "reset_centros") {
    var centrosSheet = ss.getSheetByName("Centros");
    if (!centrosSheet) throw new Error("No se encontró la hoja Centros");
    var cd = centrosSheet.getDataRange().getValues();
    for (var i = 1; i < cd.length; i++) {
      centrosSheet.getRange(i + 1, 2, 1, 5).setValues([["NO REVISADO :(", "", "", "", ""]]);
    }
    return "Centros reiniciados.";
  }

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

  // --- Calcular Estadísticas ---
  var headers = data[0];
  var colAccion = headers.indexOf("Acción");
  var colExamen = headers.indexOf("Examen");
  var colUsuario = headers.indexOf("Usuario");

  var totalErrores = data.length - 1;
  var totalAgregados = 0, totalEliminados = 0;
  var userCount = {}, exAgCount = {}, exElCount = {};

  for (var i = 1; i < data.length; i++) {
    var accion = data[i][colAccion] ? data[i][colAccion].toString() : "";
    var examen = data[i][colExamen] ? data[i][colExamen].toString() : "";
    var usuario = data[i][colUsuario] ? data[i][colUsuario].toString() : "";

    if (accion === "AGREGADO") { totalAgregados++; if (examen) exAgCount[examen] = (exAgCount[examen] || 0) + 1; }
    if (accion === "ELIMINADO") { totalEliminados++; if (examen) exElCount[examen] = (exElCount[examen] || 0) + 1; }
    if (usuario) userCount[usuario] = (userCount[usuario] || 0) + 1;
  }

  // Top 15 usuarios
  var topUsers = [];
  for (var u in userCount) topUsers.push({ name: u, count: userCount[u] });
  topUsers.sort(function (a, b) { return b.count - a.count; });
  topUsers = topUsers.slice(0, 15);

  // Top 15 exámenes agregados
  var topExAg = [];
  for (var e in exAgCount) topExAg.push({ name: e, count: exAgCount[e] });
  topExAg.sort(function (a, b) { return b.count - a.count; });
  topExAg = topExAg.slice(0, 15);

  // Top 15 exámenes eliminados
  var topExEl = [];
  for (var e2 in exElCount) topExEl.push({ name: e2, count: exElCount[e2] });
  topExEl.sort(function (a, b) { return b.count - a.count; });
  topExEl = topExEl.slice(0, 15);

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // --- Generar CSV de Errores (en memoria, sin DriveApp) ---
  var csvErrores = "\uFEFF";
  for (var ci = 0; ci < data.length; ci++) {
    var row = data[ci].map(function (cell) {
      var s = (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm") : String(cell).replace(/"/g, '""');
      return (s.indexOf(',') > -1 || s.indexOf('\n') > -1 || s.indexOf('"') > -1) ? '"' + s + '"' : s;
    });
    csvErrores += row.join(",") + "\n";
  }

  // --- Generar CSV de Estadísticas (en memoria) ---
  var csvStats = "\uFEFF";
  var statsRows = [
    ["RESUMEN SEMANAL", fecha],
    ["", ""],
    ["Total Errores", totalErrores],
    ["Agregados", totalAgregados],
    ["Eliminados", totalEliminados],
    ["", ""],
    ["TOP USUARIOS CON MÁS ERRORES", ""],
    ["Usuario", "Cantidad"]
  ];
  for (var ui2 = 0; ui2 < topUsers.length; ui2++) {
    statsRows.push([topUsers[ui2].name, topUsers[ui2].count]);
  }
  statsRows.push(["", ""]);
  statsRows.push(["TOP EXÁMENES AGREGADOS", ""]);
  statsRows.push(["Examen", "Cantidad"]);
  for (var ai2 = 0; ai2 < topExAg.length; ai2++) {
    statsRows.push([topExAg[ai2].name, topExAg[ai2].count]);
  }
  statsRows.push(["", ""]);
  statsRows.push(["TOP EXÁMENES ELIMINADOS", ""]);
  statsRows.push(["Examen", "Cantidad"]);
  for (var ei2 = 0; ei2 < topExEl.length; ei2++) {
    statsRows.push([topExEl[ei2].name, topExEl[ei2].count]);
  }
  for (var si = 0; si < statsRows.length; si++) {
    csvStats += statsRows[si].join(",") + "\n";
  }

  var blobErrores = Utilities.newBlob(csvErrores, "text/csv", "LabControl_Errores_" + fecha + ".csv");
  var blobStats = Utilities.newBlob(csvStats, "text/csv", "LabControl_Estadisticas_" + fecha + ".csv");

  // --- Construir HTML del correo ---
  var maxU = topUsers.length > 0 ? topUsers[0].count : 1;
  var maxEAg = topExAg.length > 0 ? topExAg[0].count : 1;
  var maxEEl = topExEl.length > 0 ? topExEl[0].count : 1;

  var usersHtml = topUsers.map(function (u) {
    var pct = Math.round(u.count / maxU * 100);
    return '<tr><td style="padding:4px 8px;font-weight:600;font-size:13px;">' + u.name + '</td>' +
      '<td style="padding:4px 8px;width:55%;"><div style="background:#fee2e2;border-radius:8px;overflow:hidden;height:18px;">' +
      '<div style="background:linear-gradient(90deg,#ef4444,#f97316);height:18px;width:' + pct + '%;border-radius:8px;"></div></div></td>' +
      '<td style="padding:4px 8px;text-align:right;font-weight:700;color:#ef4444;font-size:14px;">' + u.count + '</td></tr>';
  }).join("");

  var exAgHtml = topExAg.map(function (e) {
    var pct = Math.round(e.count / maxEAg * 100);
    return '<tr><td style="padding:3px 8px;font-size:11px;font-weight:600;">' + e.name + '</td>' +
      '<td style="padding:3px 8px;width:40%;"><div style="background:#d1fae5;border-radius:8px;overflow:hidden;height:14px;">' +
      '<div style="background:linear-gradient(90deg,#22c55e,#10b981);height:14px;width:' + pct + '%;border-radius:8px;"></div></div></td>' +
      '<td style="padding:3px 8px;text-align:right;font-weight:700;color:#10b981;font-size:12px;">' + e.count + '</td></tr>';
  }).join("");

  var exElHtml = topExEl.map(function (e) {
    var pct = Math.round(e.count / maxEEl * 100);
    return '<tr><td style="padding:3px 8px;font-size:11px;font-weight:600;">' + e.name + '</td>' +
      '<td style="padding:3px 8px;width:40%;"><div style="background:#fee2e2;border-radius:8px;overflow:hidden;height:14px;">' +
      '<div style="background:linear-gradient(90deg,#ef4444,#dc2626);height:14px;width:' + pct + '%;border-radius:8px;"></div></div></td>' +
      '<td style="padding:3px 8px;text-align:right;font-weight:700;color:#ef4444;font-size:12px;">' + e.count + '</td></tr>';
  }).join("");

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:650px;margin:0 auto;background:#f8fafc;padding:20px;border-radius:12px;">' +
    '<div style="text-align:center;margin-bottom:16px;">' +
    '<h1 style="color:#4f46e5;margin:0;font-size:22px;">📊 Reporte Semanal LabControl</h1>' +
    '<p style="color:#64748b;margin:4px 0 0;font-size:13px;">' + fecha + '</p></div>' +

    '<table width="100%" cellpadding="0" cellspacing="8"><tr>' +
    '<td style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#6366f1;">' + totalErrores + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Total</div></td>' +
    '<td style="background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#10b981;">' + totalAgregados + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Agregados</div></td>' +
    '<td style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#ef4444;">' + totalEliminados + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Eliminados</div></td>' +
    '</tr></table>' +

    '<div style="background:white;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-top:12px;">' +
    '<h3 style="margin:0 0 8px;font-size:14px;color:#1e293b;">👤 Top Usuarios con más Errores</h3>' +
    '<table width="100%" cellpadding="0" cellspacing="0">' + usersHtml + '</table></div>' +

    '<table width="100%" cellpadding="0" cellspacing="8" style="margin-top:4px;"><tr>' +
    '<td valign="top" width="50%" style="background:white;border:1px solid #e2e8f0;border-radius:10px;padding:14px;">' +
    '<h3 style="margin:0 0 8px;font-size:13px;color:#10b981;">✅ Top Exámenes Agregados</h3>' +
    '<table width="100%" cellpadding="0" cellspacing="0">' + exAgHtml + '</table></td>' +
    '<td valign="top" width="50%" style="background:white;border:1px solid #e2e8f0;border-radius:10px;padding:14px;">' +
    '<h3 style="margin:0 0 8px;font-size:13px;color:#ef4444;">❌ Top Exámenes Eliminados</h3>' +
    '<table width="100%" cellpadding="0" cellspacing="0">' + exElHtml + '</table></td>' +
    '</tr></table>' +

    '<div style="text-align:center;margin-top:16px;padding:12px;background:#eef2ff;border-radius:8px;">' +
    '<p style="margin:0;font-size:11px;color:#64748b;">📎 Se adjuntan 2 archivos CSV: Errores + Estadísticas</p></div>' +
    '</div>';

  // --- Construir saludo para el cuerpo del correo ---
  var saludoHtml = '<div style="font-family:Arial,sans-serif;max-width:650px;margin:0 auto 16px;padding:0 20px;">' +
    '<p style="color:#1e293b;font-size:14px;line-height:1.6;margin:0 0 8px;">Estimadas/os:</p>' +
    '<p style="color:#1e293b;font-size:14px;line-height:1.6;margin:0 0 16px;">Esperando que se encuentren bien, les comparto el reporte de errores semanales registrados por AIC durante la semana pasada.</p>' +
    '<p style="color:#1e293b;font-size:14px;line-height:1.6;margin:0 0 4px;">Saludos cordiales.</p>' +
    '<br>' +
    '<p style="color:#94a3b8;font-size:12px;font-style:italic;margin:0;">(Correo generado automáticamente)</p>' +
    '</div>';

  var htmlBodyCompleto = saludoHtml + htmlBody;

  // --- Enviar correo ---
  try {
    MailApp.sendEmail({
      to: "grivera@hospitaldetalca.cl;cdiazp@hospitaldetalca.cl;agarridom@hospitaldetalca.cl;jsmartin@hospitaldetalca.cl",
      cc: "ahormazabal@hospitaldetalca.cl;dcalderon@hospitaldetalca.cl",
      bcc: "cgonzalezmu@hospitaldetalca.cl",
      subject: "📊 Reporte Semanal LabControl - " + fecha + " (" + totalErrores + " errores)",
      body: "Estimadas/os:\n\nEsperando que se encuentren bien, les comparto el reporte de errores semanales registrados por AIC durante la semana pasada.\n\nSaludos cordiales.\n\n(Correo generado automáticamente)",
      htmlBody: htmlBodyCompleto,
      attachments: [blobErrores, blobStats]
    });
  } catch (e) {
    throw new Error("Error correo: " + e.toString());
  }

  // --- Backup en hoja dentro del mismo spreadsheet ---
  var backup = ss.insertSheet("Backup_" + fecha);
  backup.getRange(1, 1, data.length, data[0].length).setValues(data);



  // --- Limpiar SOLO Errores (NO Pizarra) ---
  if (sheetErrores.getMaxRows() > 1) {
    sheetErrores.getRange(2, 1, sheetErrores.getMaxRows() - 1, sheetErrores.getMaxColumns()).clearContent();
  }

  // --- Reiniciar Centros ---
  var centrosSheet = ss.getSheetByName("Centros");
  if (centrosSheet) {
    var cd = centrosSheet.getDataRange().getValues();
    for (var i = 1; i < cd.length; i++) {
      centrosSheet.getRange(i + 1, 2, 1, 5).setValues([["NO REVISADO :(", "", "", "", ""]]);
    }
  }

  return "Semana cerrada, correo enviado con XLSX y estadísticas.";
}








