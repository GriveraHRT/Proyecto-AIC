/**
 * ========================================================
 * PROYECTO AIC - API DE GOOGLE SHEETS v2
 * ========================================================
 * INSTRUCCIONES:
 * 1. Reemplaza TODO el código anterior con este.
 * 2. Guarda (Ctrl+S).
 * 3. Ejecuta la función "setup" una vez para crear las hojas nuevas.
 * 4. Ve a Implementar > Administrar implementaciones > Editar (lápiz)
 *    > Versión: "Nueva versión" > Implementar.
 */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Crear hojas si no existen
  const sheets = {
    "Errores": [["Fecha", "Día", "Acción", "N° Petición", "Examen", "Usuario"]],
    "Pizarra_Muestras": [["ID", "Tipo", "N° Petición", "Examen"]],
    "Pizarra_Curvas": [["ID", "N° Petición", "Validada"]],
    "Pizarra_Urgentes": [["ID", "N° Petición", "Validada"]],
    "Pizarra_Recordatorios": [["ID", "Texto", "Usuario", "Fecha"]],
    "Pizarra_Custom": [["Cuadrante", "ID", "N° Petición", "Detalle"]],
    "Centros": [["Centro", "Estado", "Pdte Rev", "Pdte Val", "Responsable"]],
    "Maestro_Examenes": [["Examen"]]
  };

  for (const name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // Poner cabeceras si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, sheets[name][0].length).setValues(sheets[name]);
    }
  }

  // Poblar Centros si solo tiene cabeceras
  const centrosSheet = ss.getSheetByName("Centros");
  if (centrosSheet.getLastRow() <= 1) {
    const centros = [
      ["Cachureo","NO REVISADO :(","","",""],
      ["San Clemente","NO REVISADO :(","","",""],
      ["Maule","NO REVISADO :(","","",""],
      ["Externos","NO REVISADO :(","","",""],
      ["Hospitalizados","NO REVISADO :(","","",""],
      ["Ambulatorio","NO REVISADO :(","","",""],
      ["San Rafael","NO REVISADO :(","","",""],
      ["Rio Claro","NO REVISADO :(","","",""],
      ["Pencahue","NO REVISADO :(","","",""],
      ["Pelarco","NO REVISADO :(","","",""],
      ["TOMA DE MUESTRA","NO REVISADO :(","","",""]
    ];
    centrosSheet.getRange(2, 1, centros.length, 5).setValues(centros);
  }

  // Poblar exámenes maestros si solo tiene cabecera
  const maestroSheet = ss.getSheetByName("Maestro_Examenes");
  if (maestroSheet.getLastRow() <= 1) {
    const examenes = [
      "GLU","LDL","HDL","COL","TG","TSH","T3","T4","T4L",
      "HBA1C","CRE","URE","AC. URICO","BIL TOTAL","BIL DIRECTA",
      "GOT","GPT","GGT","FA","LDH","CK","CK-MB","TROPO",
      "NA","K","CL","CA","MG","P","FE","FERRITINA","VIT B12",
      "AC. FOLICO","VIT D","PTH","PCR","VHS","HEMOGRAMA",
      "PROTROMBINA","TTPA","FIBRINOGENO","INR","ORINA COMP",
      "PSA","CEA","AFP","CA 19-9","CA 125","INSULINA",
      "CORTISOL","PROLACTINA","FSH","LH","ESTRADIOL",
      "PROGESTERONA","TESTOSTERONA","DHEA-S","AMILASA","LIPASA",
      "PROTEINAS TOT","ALBUMINA","GLOBULINA","TP","ELECT. PROT"
    ];
    const rows = examenes.sort().map(e => [e]);
    maestroSheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
}

// ===================== GET =====================
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getSheetData = (name) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j] ? data[i][j].toString() : "";
      }
      rows.push(obj);
    }
    return rows;
  };

  // Lista simple de exámenes (solo columna A)
  const getMaestro = () => {
    const sheet = ss.getSheetByName("Maestro_Examenes");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) list.push(data[i][0].toString());
    }
    return list.sort();
  };

  const response = {
    errores: getSheetData("Errores"),
    muestras: getSheetData("Pizarra_Muestras"),
    curvas: getSheetData("Pizarra_Curvas"),
    urgentes: getSheetData("Pizarra_Urgentes"),
    recordatorios: getSheetData("Pizarra_Recordatorios"),
    custom: getSheetData("Pizarra_Custom"),
    centros: getSheetData("Centros"),
    maestro_examenes: getMaestro()
  };

  return ContentService.createTextOutput(JSON.stringify({ success: true, data: response }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== POST =====================
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    const result = handleAction(payload);
    output.setContent(JSON.stringify({ success: true, data: result }));
  } catch (error) {
    output.setContent(JSON.stringify({ success: false, error: error.toString() }));
  }
  return output;
}

function handleAction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = payload.action;

  if (action === "cerrar_semana") return procesarCierreSemana();

  // Acciones que requieren una hoja
  const sheet = ss.getSheetByName(payload.sheet);
  if (!sheet) throw new Error("No se encontró la hoja: " + payload.sheet);

  if (action === "insert") {
    sheet.appendRow(payload.row);
    return "OK";
  }

  if (action === "update") {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() == payload.id) {
        if (payload.row) {
          sheet.getRange(i + 1, 1, 1, payload.row.length).setValues([payload.row]);
        } else if (payload.updates) {
          for (const key in payload.updates) {
            const col = headers.indexOf(key);
            if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(payload.updates[key]);
          }
        }
        return "OK";
      }
    }
    throw new Error("ID no encontrado para actualizar");
  }

  if (action === "delete") {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() == payload.id) {
        sheet.deleteRow(i + 1);
        return "OK";
      }
    }
    throw new Error("ID no encontrado para eliminar");
  }

  if (action === "delete_by_col") {
    // Borrar todas las filas donde la columna indicada coincida con el valor
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = headers.indexOf(payload.column);
    if (col === -1) throw new Error("Columna no encontrada: " + payload.column);
    // Borrar de abajo hacia arriba para no perder índices
    let count = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][col].toString() == payload.value) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    return "Eliminadas " + count + " filas";
  }

  throw new Error("Acción desconocida: " + action);
}

function procesarCierreSemana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetErrores = ss.getSheetByName("Errores");
  const data = sheetErrores.getDataRange().getValues();

  if (data.length <= 1) return "No hay datos semanales para exportar.";

  // Crear CSV
  let csv = "\uFEFF";
  for (let i = 0; i < data.length; i++) {
    let row = data[i].map(item => {
      let str = typeof item === 'string' ? item.replace(/"/g, '""') : String(item);
      if (str.includes(',') || str.includes('\n') || str.includes('"')) str = '"' + str + '"';
      return str;
    });
    csv += row.join(",") + "\n";
  }

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const blob = Utilities.newBlob(csv, "text/csv", "Reporte_Errores_" + fecha + ".csv");

  try {
    MailApp.sendEmail({
      to: "grivera@hospitaldetalca.cl",
      subject: "Reporte de Errores - Cierre Semanal " + fecha,
      body: "Adjunto el reporte de errores de la semana. El sistema ha limpiado la tabla.",
      attachments: [blob]
    });
  } catch(e) {
    throw new Error("Error al enviar correo: " + e.toString());
  }

  // Backup
  const backup = ss.insertSheet("Backup_" + fecha);
  backup.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Limpiar errores
  if (sheetErrores.getMaxRows() > 1) {
    sheetErrores.getRange(2, 1, sheetErrores.getMaxRows() - 1, sheetErrores.getMaxColumns()).clearContent();
  }

  // Limpiar pizarras también
  ["Pizarra_Muestras","Pizarra_Curvas","Pizarra_Urgentes","Pizarra_Recordatorios","Pizarra_Custom"].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && s.getMaxRows() > 1) {
      s.getRange(2, 1, s.getMaxRows() - 1, s.getMaxColumns()).clearContent();
    }
  });

  // Resetear estados de centros
  const centrosSheet = ss.getSheetByName("Centros");
  if (centrosSheet) {
    const cd = centrosSheet.getDataRange().getValues();
    for (let i = 1; i < cd.length; i++) {
      centrosSheet.getRange(i + 1, 2, 1, 4).setValues([["NO REVISADO :(", "", "", ""]]);
    }
  }

  return "Semana cerrada, correo enviado y datos limpiados.";
}
