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
      ["TOMA DE MUESTRA", "NO REVISADO :(", "", "", "", ""]
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

  // --- Generar CSV ---
  var csv = "\uFEFF";
  for (var i = 0; i < data.length; i++) {
    var row = data[i].map(function (item) {
      var str = typeof item === 'string' ? item.replace(/"/g, '""') : String(item);
      if (str.indexOf(',') > -1 || str.indexOf('\n') > -1 || str.indexOf('"') > -1) str = '"' + str + '"';
      return str;
    });
    csv += row.join(",") + "\n";
  }

  // --- Calcular Estadísticas ---
  var headers = data[0];
  var colAccion = headers.indexOf("Acción");
  var colExamen = headers.indexOf("Examen");
  var colUsuario = headers.indexOf("Usuario");
  var colDia = headers.indexOf("Día");

  var totalErrores = data.length - 1;
  var totalAgregados = 0, totalEliminados = 0;
  var userCount = {}, examCount = {};
  var diasSet = {};

  for (var i = 1; i < data.length; i++) {
    var accion = data[i][colAccion] ? data[i][colAccion].toString() : "";
    var examen = data[i][colExamen] ? data[i][colExamen].toString() : "";
    var usuario = data[i][colUsuario] ? data[i][colUsuario].toString() : "";
    var dia = data[i][colDia] ? data[i][colDia].toString() : "";

    if (accion === "AGREGADO") totalAgregados++;
    if (accion === "ELIMINADO") totalEliminados++;
    if (usuario) userCount[usuario] = (userCount[usuario] || 0) + 1;
    if (examen) examCount[examen] = (examCount[examen] || 0) + 1;
    if (dia) diasSet[dia] = true;
  }

  var diasConErrores = Object.keys(diasSet).length;

  // Top 8 usuarios
  var topUsers = [];
  for (var u in userCount) topUsers.push({ name: u, count: userCount[u] });
  topUsers.sort(function (a, b) { return b.count - a.count; });
  topUsers = topUsers.slice(0, 8);

  // Top 8 exámenes
  var topExams = [];
  for (var e in examCount) topExams.push({ name: e, count: examCount[e] });
  topExams.sort(function (a, b) { return b.count - a.count; });
  topExams = topExams.slice(0, 8);

  // --- Construir HTML del correo ---
  var maxU = topUsers.length > 0 ? topUsers[0].count : 1;
  var maxE = topExams.length > 0 ? topExams[0].count : 1;

  var usersHtml = topUsers.map(function (u) {
    var pct = Math.round(u.count / maxU * 100);
    return '<tr><td style="padding:4px 8px;font-weight:600;font-size:13px;">' + u.name + '</td>' +
      '<td style="padding:4px 8px;width:60%;"><div style="background:#fee2e2;border-radius:8px;overflow:hidden;height:18px;">' +
      '<div style="background:linear-gradient(90deg,#ef4444,#f97316);height:18px;width:' + pct + '%;border-radius:8px;"></div></div></td>' +
      '<td style="padding:4px 8px;text-align:right;font-weight:700;color:#ef4444;font-size:14px;">' + u.count + '</td></tr>';
  }).join("");

  var examsHtml = topExams.map(function (e) {
    var pct = Math.round(e.count / maxE * 100);
    return '<tr><td style="padding:4px 8px;font-weight:600;font-size:13px;">' + e.name + '</td>' +
      '<td style="padding:4px 8px;width:60%;"><div style="background:#e0e7ff;border-radius:8px;overflow:hidden;height:18px;">' +
      '<div style="background:linear-gradient(90deg,#6366f1,#8b5cf6);height:18px;width:' + pct + '%;border-radius:8px;"></div></div></td>' +
      '<td style="padding:4px 8px;text-align:right;font-weight:700;color:#6366f1;font-size:14px;">' + e.count + '</td></tr>';
  }).join("");

  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  var htmlBody = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8fafc;padding:20px;border-radius:12px;">' +
    '<div style="text-align:center;margin-bottom:16px;">' +
    '<h1 style="color:#4f46e5;margin:0;font-size:22px;">📊 Reporte Semanal LabControl</h1>' +
    '<p style="color:#64748b;margin:4px 0 0;font-size:13px;">' + fecha + '</p></div>' +

    '<div style="display:flex;gap:8px;margin-bottom:16px;text-align:center;">' +
    '<table width="100%" cellpadding="0" cellspacing="8"><tr>' +
    '<td style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#6366f1;">' + totalErrores + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Total</div></td>' +
    '<td style="background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#10b981;">' + totalAgregados + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Agregados</div></td>' +
    '<td style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#ef4444;">' + totalEliminados + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Eliminados</div></td>' +
    '<td style="background:#fefce8;border:1px solid #fde68a;border-radius:10px;padding:12px;text-align:center;"><div style="font-size:28px;font-weight:800;color:#f59e0b;">' + diasConErrores + '</div><div style="font-size:10px;color:#64748b;text-transform:uppercase;">Días</div></td>' +
    '</tr></table></div>' +

    '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-bottom:12px;">' +
    '<h3 style="margin:0 0 8px;font-size:14px;color:#ef4444;">🔝 Top Usuarios con más Errores</h3>' +
    (topUsers.length > 0 ? '<table width="100%" cellpadding="0" cellspacing="0">' + usersHtml + '</table>' : '<p style="color:#94a3b8;font-size:12px;">Sin datos</p>') + '</div>' +

    '<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-bottom:12px;">' +
    '<h3 style="margin:0 0 8px;font-size:14px;color:#6366f1;">🔬 Top Exámenes más Frecuentes</h3>' +
    (topExams.length > 0 ? '<table width="100%" cellpadding="0" cellspacing="0">' + examsHtml + '</table>' : '<p style="color:#94a3b8;font-size:12px;">Sin datos</p>') + '</div>' +

    '<p style="text-align:center;color:#94a3b8;font-size:11px;margin-top:16px;">Generado por LabControl · Adjunto CSV con datos completos</p>' +
    '</div>';

  var blob = Utilities.newBlob(csv, "text/csv", "Reporte_Errores_" + fecha + ".csv");

  try {
    MailApp.sendEmail({
      to: "grivera@hospitaldetalca.cl",
      subject: "📊 Reporte Semanal LabControl - " + fecha + " (" + totalErrores + " errores)",
      body: "Reporte semanal: " + totalErrores + " errores (" + totalAgregados + " agregados, " + totalEliminados + " eliminados). Top usuario: " + (topUsers.length > 0 ? topUsers[0].name + " (" + topUsers[0].count + ")" : "N/A") + ". Ver HTML para dashboard completo.",
      htmlBody: htmlBody,
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

  var pizarras = ["Pizarra_Curvas", "Pizarra_Urgentes", "Pizarra_Custom"];
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
      centrosSheet.getRange(i + 1, 2, 1, 5).setValues([["NO REVISADO :(", "", "", "", ""]]);
    }
  }

  return "Semana cerrada, correo enviado con estadísticas.";
}
