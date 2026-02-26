/**
 * PROYECTO AIC - API DE GOOGLE SHEETS
 * Copia y pega todo este código en el editor de Google Apps Script.
 * Luego, ejecuta la función "setup" una vez para crear las hojas necesarias.
 * Finalmente, publica esto como una "Aplicación Web" (Ejecutar como: Tú, Acceso: Cualquier persona).
 */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Errores", "Pizarra_Muestras", "Pizarra_Curvas", "Centros"];
  
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // Configurar las cabeceras
  ss.getSheetByName("Errores").getRange("A1:F1").setValues([["Fecha", "Día", "Acción", "N° Petición", "Examen", "Usuario"]]);
  ss.getSheetByName("Pizarra_Muestras").getRange("A1:D1").setValues([["ID", "Tipo", "N° Petición", "Examen"]]);
  ss.getSheetByName("Pizarra_Curvas").getRange("A1:C1").setValues([["ID", "N° Petición", "Validada"]]);
  ss.getSheetByName("Centros").getRange("A1:E1").setValues([["Centro", "Estado", "Pdte Rev", "Pdte Val", "Responsable"]]);
  
  // Poblar los Centros Iniciales
  const CentrosSheet = ss.getSheetByName("Centros");
  if(CentrosSheet.getLastRow() === 1 || CentrosSheet.getLastRow() === 0) {
     const centrosIniciales = [
       ["Cachureo", "NO REVISADO :(", "", "", ""],
       ["San Clemente", "NO REVISADO :(", "", "", ""],
       ["Maule", "NO REVISADO :(", "", "", ""],
       ["Externos", "NO REVISADO :(", "", "", ""],
       ["Hospitalizados", "NO REVISADO :(", "", "", ""],
       ["Ambulatorio", "NO REVISADO :(", "", "", ""]
     ];
     CentrosSheet.getRange(2, 1, centrosIniciales.length, 5).setValues(centrosIniciales);
  }
}

// Función GET: Se llama cuando la página lee la base de datos (Polling)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const getSheetData = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Solo cabeceras o vacío
    
    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        // Convertir strings para evitar problemas de fechas complejas de JS si no son necesarias
        obj[headers[j]] = data[i][j] ? data[i][j].toString() : ""; 
      }
      rows.push(obj);
    }
    return rows;
  };

  const response = {
    errores: getSheetData("Errores"),
    muestras: getSheetData("Pizarra_Muestras"),
    curvas: getSheetData("Pizarra_Curvas"),
    centros: getSheetData("Centros")
  };

  return ContentService.createTextOutput(JSON.stringify({ success: true, data: response }))
        .setMimeType(ContentService.MimeType.JSON);
}


// Función POST: Se llama cuando la página envía datos a la base de datos
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

  if (action === "cerrar_semana") {
    return procesarCierreSemana();
  }

  const sheet = ss.getSheetByName(payload.sheet);
  if (!sheet) throw new Error("Ups! No se encontró la hoja: " + payload.sheet);

  if (action === "insert") {
    sheet.appendRow(payload.row);
    return "Insertado correctamente";
  } 
  
  else if (action === "update") {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
        // Asumiendo que la Columna A (índice 0) siempre aloja el "ID" o Identificador único
        if (data[i][0] == payload.id) {
            if (payload.row) {
                 // Reemplazar toda la fila
                 sheet.getRange(i + 1, 1, 1, payload.row.length).setValues([payload.row]);
            } else if (payload.updates) {
                 // Reemplazar solo algunas celdas específicas
                 for (const key in payload.updates) {
                     const colIndex = headers.indexOf(key);
                     if (colIndex !== -1) {
                         sheet.getRange(i + 1, colIndex + 1).setValue(payload.updates[key]);
                     }
                 }
            }
            return "Actualizado correctamente";
        }
    }
    throw new Error("No se encontró el ID para actualizar");
  }
  
  else if (action === "delete") {
     const data = sheet.getDataRange().getValues();
     for (let i = 1; i < data.length; i++) {
        if (data[i][0] == payload.id) {
            sheet.deleteRow(i + 1);
            return "Borrado correctamente";
        }
     }
     throw new Error("No se encontró el ID para eliminar");
  }
  
  throw new Error("Acción desconocida");
}

function procesarCierreSemana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetErrores = ss.getSheetByName("Errores");
  const data = sheetErrores.getDataRange().getValues();
  
  if (data.length <= 1) {
      return "No hay datos semanales para exportar todavía.";
  }

  // Crear archivo CSV manualmente
  let csvContent = "";
  for (let i = 0; i < data.length; i++) {
    let row = data[i].map(item => {
      let str = typeof item === 'string' ? item.replace(/"/g, '""') : String(item);
      if (str.includes(',') || str.includes('\\n') || str.includes('"')) {
        str = `"${str}"`;
      }
      return str;
    });
    // Añadimos indicador BOM para que Excel respete los tildes al abrir el CSV
    if(i === 0) { csvContent += "\uFEFF"; }
    csvContent += row.join(",") + "\n";
  }

  const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const blob = Utilities.newBlob(csvContent, "text/csv", "Reporte_Errores_" + fechaActual + ".csv");

  try {
     MailApp.sendEmail({
      to: "grivera@hospitaldetalca.cl",
      subject: "Reporte de Errores - Cierre Semanal",
      body: "Adjunto el reporte de los errores de la semana en formato CSV. El sistema ha limpiado la tabla de la semana.",
      attachments: [blob]
    });
  } catch(e) {
    throw new Error("Imposible enviar correo: " + e.toString());
  }

  // Realizar copia de respaldo de la información en una nueva pestaña (por seguridad antes de borrar)
  const archiveName = "Backup_" + fechaActual;
  const wsBackup = ss.insertSheet(archiveName);
  wsBackup.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Limpiar hoja de errores, dejando solamente la fila 1 (encabezados)
  const maxRows = sheetErrores.getMaxRows();
  const maxCols = sheetErrores.getMaxColumns();
  if(maxRows > 1) {
      sheetErrores.getRange(2, 1, maxRows - 1, maxCols).clearContent();
  }

  return "Semana cerrada, correo enviado a grivera@hospitaldetalca.cl y datos limpios exitosamente.";
}
