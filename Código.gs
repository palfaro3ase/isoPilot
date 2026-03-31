// =============================================================
// ANCLA PRINCIPAL: ID DE LA HOJA DE CÁLCULO (BASE DE DATOS)
// =============================================================
var SS_ID = "1rs9NvJ512mouckfc78u9Fwwbi4jHcPv0dMZ3D2Tohp4";

// =============================================================
// FUNCIONES DE RENDERIZADO (WEB APP)
// =============================================================

function doGet(e) {
  var page = e.parameter.page || 'index';
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('ISO-Pilot: Gestión Inteligente')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getParam(paramName) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName("CONFIG");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == paramName) {
      return data[i][1];
    }
  }
  throw new Error("Parámetro no encontrado: " + paramName);
}

// =============================================================
// LÓGICA DE PROCESAMIENTO CON IA
// =============================================================

function procesarAuditoria(fileData) {
  try {
    console.log("🔵 [procesarAuditoria] Iniciando procesamiento de:", fileData.name);
    
    const FOLDER_ID = getParam("ID_CARPETA_ENTRADA");
    const API_KEY = getParam("GEMINI_API_KEY");
    
    if (!FOLDER_ID || !API_KEY) {
      throw new Error("Faltan parámetros de configuración. Verifica CONFIG sheet.");
    }
    
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.type, fileData.name);
    
    // 1. Guardar archivo original
    var file = folder.createFile(blob);
    var textoDocumento = "";

    // 2. EXTRAER TEXTO (Versión mejorada para PDF y DOCX)
    try {
      console.log("🔵 [OCR] Extrayendo texto de:", file.getName());
      console.log("🔵 [OCR] Tipo MIME:", file.getMimeType());
      console.log("🔵 [OCR] Tamaño:", file.getSize(), "bytes");
      
      var mimeType = file.getMimeType();
      var fileName = file.getName().toLowerCase();
      
      // ✅ MÉTODO PRINCIPAL: Conversión a Google Docs con OCR
      console.log("🔵 [OCR] Intentando conversión a Google Docs...");
      try {
        var resource = {
          title: file.getName() + "_temp",
          mimeType: MimeType.GOOGLE_DOCS
        };
        
        // Parámetros optimizados según tipo de archivo
        var convertParams = { 
          ocr: true, 
          ocrLanguage: 'es',
          convert: true
        };
        
        // Para DOCX/PDF, el OCR ayuda; para otros formatos, solo conversión
        if (mimeType !== MimeType.PDF && !fileName.endsWith('.pdf')) {
          // Para DOC/DOCX, intentar sin OCR primero (más rápido)
          try {
            var tempFile = Drive.Files.insert(resource, blob, { convert: true });
            var doc = DocumentApp.openById(tempFile.id);
            textoDocumento = doc.getBody().getText();
            Drive.Files.remove(tempFile.id);
            console.log("🟢 [OCR] Conversión directa exitosa, longitud:", textoDocumento.length);
          } catch (e) {
            // Fallback a OCR si falla conversión directa
            var tempFile = Drive.Files.insert(resource, blob, convertParams);
            var doc = DocumentApp.openById(tempFile.id);
            textoDocumento = doc.getBody().getText();
            Drive.Files.remove(tempFile.id);
            console.log("🟢 [OCR] Conversión con OCR exitosa, longitud:", textoDocumento.length);
          }
        } else {
          // Para PDF, usar OCR obligatorio
          var tempFile = Drive.Files.insert(resource, blob, convertParams);
          var doc = DocumentApp.openById(tempFile.id);
          textoDocumento = doc.getBody().getText();
          Drive.Files.remove(tempFile.id);
          console.log("🟢 [OCR] PDF procesado con OCR, longitud:", textoDocumento.length);
        }
        
      } catch (convertError) {
        console.warn("⚠️ [OCR] Conversión a Google Docs falló:", convertError.toString());
        
        // ✅ FALLBACK 1: Exportar como texto vía Drive API
        try {
          console.log("🔵 [OCR] Fallback: Exportando como texto...");
          var fileId = file.getId();
          var exportUrl = "https://docs.google.com/document/d/" + fileId + "/export?format=txt";
          var token = ScriptApp.getOAuthToken();
          
          var response = UrlFetchApp.fetch(exportUrl, {
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true,
            timeout: 30000
          });
          
          if (response.getResponseCode() === 200) {
            textoDocumento = response.getContentText();
            console.log("🟢 [OCR] Export como texto exitoso, longitud:", textoDocumento.length);
          }
        } catch (exportError) {
          console.warn("⚠️ [OCR] Export como texto falló:", exportError.toString());
        }
      }
      
      // ✅ Limpieza de texto extraído
      if (textoDocumento) {
        // Remover caracteres binarios/control que no son texto legible
        textoDocumento = textoDocumento.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, ' ');
        // Remover secuencias PK (indicador de archivo ZIP/DOCX sin procesar)
        textoDocumento = textoDocumento.replace(/PK[\x03\x04].*?$/s, '');
        // Normalizar espacios múltiples
        textoDocumento = textoDocumento.replace(/\s+/g, ' ').trim();
        
        console.log("🟢 [OCR] Texto limpio, longitud final:", textoDocumento.length);
      }
      
      // ✅ Validación final del texto extraído
      if (!textoDocumento || textoDocumento.trim().length < 100 || 
          textoDocumento.toLowerCase().includes('pk') || 
          textoDocumento.toLowerCase().includes('binary')) {
        console.warn("⚠️ [OCR] Texto insuficiente o inválido, usando metadatos...");
        textoDocumento = "Archivo: " + file.getName() + 
                        "\nTipo MIME: " + mimeType + 
                        "\nTamaño: " + file.getSize() + " bytes" +
                        "\n\n[Nota: El contenido completo no pudo extraerse automáticamente. " +
                        "El análisis se basará en el nombre del archivo y metadatos disponibles. " +
                        "Para mejor resultado, asegúrese que el archivo sea legible y no esté protegido.]";
      }
      
      // Limitar longitud para evitar timeout en Gemini
      if (textoDocumento.length > 45000) {
        textoDocumento = textoDocumento.substring(0, 45000) + "\n\n[... texto truncado por longitud ...]";
        console.log("⚠️ [OCR] Texto truncado a 45000 caracteres");
      }
      
    } catch (err) {
      console.error("❌ [OCR] Error crítico:", err.toString());
      textoDocumento = "Archivo: " + file.getName() + 
                      "\nTipo: " + file.getMimeType() + 
                      "\n\n[ERROR en extracción: " + err.message + "]";
    }

    // 3. Obtener contexto histórico
    var historial = getHistorialReciente();
    
    // ✅ 4. Construir prompt OPTIMIZADO (más corto = menos ancho de banda)
    var prompt = `Auditor ISO 9001:2015. Analiza:

${textoDocumento.substring(0, 30000)}

Historial: ${historial.substring(0, 5000)}

TAREA: Identifica no conformidades, reincidencias, riesgo (Alto/Medio/Bajo).

JSON: {"hallazgos":"texto","reincidencias":"texto","riesgo":"Alto|Medio|Bajo"}`;
    
    console.log("🔵 [Prompt] Enviando a Gemini...");
    
    // 5. Llamada a Gemini
    var resIA = callGemini(prompt, API_KEY);
    
    console.log("🔵 [Parseo] Longitud respuesta IA:", resIA.length);
    
    // 6. Parsear respuesta JSON (Resiliente a markdown y truncamiento)
    var dataIA;
    try {
      // Limpiar markdown
      var textoLimpio = resIA.replace(/```json/gi, "").replace(/```/gi, "").trim();
      
      // Extraer JSON entre llaves
      var jsonMatch = textoLimpio.match(/\{[\s\S]*\}/);
      var jsonStr = jsonMatch ? jsonMatch[0] : textoLimpio;
      
      // Intentar parsear
      try {
        dataIA = JSON.parse(jsonStr);
      } catch (jsonError) {
        // Fallback para JSON truncado
        console.warn("⚠️ [Parseo] JSON incompleto, recuperando...");
        dataIA = {
          hallazgos: extraerCampo(jsonStr, "hallazgos") || "Análisis en proceso",
          reincidencias: extraerCampo(jsonStr, "reincidencias") || "No determinado",
          riesgo: extraerCampo(jsonStr, "riesgo") || "Medio"
        };
      }
      
      // Validar y normalizar - ✅ AÑADIDO: Verificar que dataIA sea objeto
      if (!dataIA || typeof dataIA !== 'object') {
        console.warn("⚠️ [Parseo] dataIA no es objeto, creando fallback");
        dataIA = {};
      }
      if (!dataIA.hallazgos) dataIA.hallazgos = "Sin hallazgos específicos";
      if (!dataIA.reincidencias) dataIA.reincidencias = "Ninguna";
      var riesgosValidos = ["Alto", "Medio", "Bajo"];
      if (!dataIA.riesgo || !riesgosValidos.includes(dataIA.riesgo)) {
        dataIA.riesgo = "Medio";
      }
      
      console.log("🟢 [Parseo] Datos procesados");
      
    } catch (parseError) {
      console.error("❌ [Parseo] Error:", parseError.toString());
      dataIA = {
        hallazgos: "Error al procesar: " + parseError.message,
        reincidencias: "No determinado",
        riesgo: "Medio"
      };
    }
    
    // ✅ SAFETY CHECK FINAL: Asegurar que dataIA sea válido antes de continuar
    if (!dataIA || typeof dataIA !== 'object') {
      console.error("❌ [procesarAuditoria] dataIA inválido, usando fallback definitivo");
      dataIA = {
        hallazgos: "Error: datos no procesables",
        reincidencias: "No determinado",
        riesgo: "Medio"
      };
    }
    
    // 7. Registro en LOG_AUDITORIAS
    var ss = SpreadsheetApp.openById(SS_ID);
    var logSheet = ss.getSheetByName("LOG_AUDITORIAS");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("LOG_AUDITORIAS");
      logSheet.appendRow(["Fecha", "ID Archivo", "Nombre", "Hallazgos", "Reincidencias", "Usuario", "Riesgo"]);
    }
    
    logSheet.appendRow([
      new Date(), 
      file.getId(), 
      file.getName(), 
      dataIA.hallazgos, 
      dataIA.reincidencias, 
      Session.getActiveUser().getEmail(),
      dataIA.riesgo 
    ]);
    
    console.log("🟢 [procesarAuditoria] Completado");
    
    // ✅ FIX: Devolver 'data' (no 'dataIA') para coincidir con frontend
    return { success: true, data: dataIA };
    
  } catch (e) {
    console.error("❌ [procesarAuditoria] Error:", e.toString());
    return { success: false, error: "Error en el proceso: " + e.toString() };
  }
}

// ✅ FUNCIÓN AUXILIAR: Extraer campo de JSON truncado
function extraerCampo(jsonStr, campo) {
  try {
    var pattern = new RegExp('"' + campo + '"\\s*:\\s*"([^"]*)', 'i');
    var match = jsonStr.match(pattern);
    if (match && match[1]) {
      return match[1].trim();
    }
  } catch (e) {
    console.warn("⚠️ [extraerCampo] Error:", e.toString());
  }
  return null;
}

function callGemini(prompt, apiKey) {
  var model = getParam("MODEL_NAME") || "gemini-1.5-flash";
  var url = "https://generativelanguage.googleapis.com/v1/models/" + model + ":generateContent?key=" + apiKey;
  
  var payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { 
      "temperature": 0.2,
      "maxOutputTokens": 4096,
      "topP": 0.9,
      "topK": 40
    }
  };
  
  // ✅ MODIFICADO: Timeout aumentado a 120 segundos para evitar rate limit
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true,
    "headers": { "Content-Type": "application/json" },
    "timeout": 120000
  };
  
  try {
    console.log("🔵 [Gemini] Modelo:", model);
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode !== 200) {
      throw new Error("API Gemini error " + responseCode + ": " + responseText.substring(0, 500));
    }
    
    var json = JSON.parse(responseText);
    
    if (!json.candidates || json.candidates.length === 0) {
      if (json.promptFeedback && json.promptFeedback.blockReason) {
        throw new Error("Bloqueado: " + json.promptFeedback.blockReason);
      }
      throw new Error("Sin candidatos");
    }
    
    var candidate = json.candidates[0];
    
    if (candidate.finishReason === "SAFETY") {
      throw new Error("Contenido bloqueado por seguridad");
    }
    
    if (!candidate.content || !candidate.content.parts || candidate.content.parts.length === 0) {
      throw new Error("Respuesta sin contenido");
    }
    
    var resultText = candidate.content.parts[0].text;
    
    if (!resultText || resultText.trim() === "") {
      throw new Error("Respuesta vacía");
    }
    
    console.log("🟢 [Gemini] Éxito, longitud:", resultText.length);
    return resultText;
    
  } catch (e) {
    console.error("❌ [Gemini] Error:", e.toString());
    throw new Error("Error llamando a Gemini: " + e.toString());
  }
}

function getHistorialReciente() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName("LOG_AUDITORIAS");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "No hay historial previo.";
  
  var startRow = Math.max(2, lastRow - 15);
  var values = sheet.getRange(startRow, 4, (lastRow - startRow + 1), 2).getValues(); 
  return JSON.stringify(values);
}

function getMetricasDashboard() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    const contarFilas = (nombreHoja) => {
      const sheet = ss.getSheetByName(nombreHoja);
      if (!sheet) return 0;
      const lastRow = sheet.getLastRow();
      return lastRow > 1 ? (lastRow - 1) : 0;
    };
    
    // ✅ Contar riesgos "Alto" en LOG_AUDITORIAS
    const contarAlertas = () => {
      const sheet = ss.getSheetByName("LOG_AUDITORIAS");
      if (!sheet) return 0;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return 0;
      
      // Columna 7 (G) es donde está el Riesgo
      const riesgos = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
      let alertas = 0;
      
      riesgos.forEach(fila => {
        if (fila[0] === "Alto") alertas++;
      });
      
      return alertas;
    };
    
    return {
      docs: contarFilas("LISTADO_MAESTRO"),
      auditorias: contarFilas("LOG_AUDITORIAS"),
      nc: contarAlertas()  // ✅ Ahora cuenta riesgos "Alto"
    };
  } catch (e) {
    console.error("Error en getMetricasDashboard:", e);
    return { docs: 0, auditorias: 0, nc: 0 };
  }
}

function testConfig() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName("CONFIG");
  var data = sheet.getDataRange().getValues();
  console.log("CONFIG sheet ", data);
  data.forEach(function(row) {
    if (row[0] === "GEMINI_API_KEY") {
      console.log("API Key:", row[1] ? "Sí" : "No");
    }
  });
}

function listarModelosGemini() {
  var apiKey = getParam("GEMINI_API_KEY");
  var url = "https://generativelanguage.googleapis.com/v1/models?key=" + apiKey;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(response.getContentText());
  console.log("📋 Modelos disponibles:");
  if (json.models) {
    json.models.forEach(function(model) {
      console.log("  - " + model.name);
    });
  }
  return json;
}

function obtenerDocumentosMaestro() {
  try {
    // CAMBIO CLAVE: Usar el ID de la hoja externa
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('LISTADO_MAESTRO');
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 
    
    // Obtenemos los datos (ajusta el número de columnas si es necesario)
    // Según tu registro, tenemos: ID, Código, Nombre, Versión, Link, Usuario, Fecha
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    return data.map(fila => ({
      id: fila[0],
      codigo: fila[1],
      nombre: fila[2],
      version: fila[3],
      enlace: fila[4],
      usuario: fila[5],
      fecha: fila[6] instanceof Date ? Utilities.formatDate(fila[6], Session.getScriptTimeZone(), "dd/MM/yyyy") : fila[6]
    }));
  } catch (e) {
    console.error("Error en obtenerDocumentosMaestro: " + e.toString());
    return [];
  }
}

function obtenerLogsAuditorias() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    // ⚠️ ASEGÚRATE que el nombre sea exacto: "LOG_AUDITORIAS" o "LOGS_AUDITORIAS"
    var sheet = ss.getSheetByName("LOG_AUDITORIAS") || ss.getSheetByName("LOGS_AUDITORIAS");
    
    if (!sheet) {
      console.error("❌ Hoja de logs no encontrada");
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Obtenemos el rango (7 columnas)
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var logs = [];
    
    // Procesar de más reciente a más antiguo
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (!row[2]) continue; // Saltar filas vacías

      var fechaObj = row[0];
      var fechaFormateada = (fechaObj instanceof Date) 
        ? Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : String(fechaObj);
      
      logs.push({
        fecha: fechaFormateada,
        fileId: row[1] || "",
        nombre: row[2] || "Sin nombre",
        // Limpiamos saltos de línea para evitar errores en el JS del cliente
        hallazgos: String(row[3] || "").replace(/[\r\n]+/g, " "), 
        reincidencias: String(row[4] || "Ninguna").replace(/[\r\n]+/g, " "),
        usuario: row[5] || "",
        riesgo: row[6] || "Medio"
      });
    }
    return logs;
  } catch (e) {
    console.error("❌ Error en obtenerLogsAuditorias: " + e.toString());
    return [];
  }
}

// ✅ FUNCIÓN: Buscar logs con filtro (SERVER-SIDE)
function buscarLogsAuditorias(textoBusqueda) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var sheet = ss.getSheetByName("LOG_AUDITORIAS");
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var logs = [];
    var texto = textoBusqueda ? textoBusqueda.toString().toLowerCase().trim() : "";
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (!texto || texto.length < 2) {
        logs.push(crearLogDesdeFila(row));
        continue;
      }
      
      var fecha = String(row[0] || "").toLowerCase();
      var fileId = String(row[1] || "").toLowerCase();
      var nombre = String(row[2] || "").toLowerCase();
      var hallazgos = String(row[3] || "").toLowerCase();
      var reincidencias = String(row[4] || "").toLowerCase();
      var usuario = String(row[5] || "").toLowerCase();
      var riesgo = String(row[6] || "").toLowerCase();
      
      if (fecha.indexOf(texto) !== -1 || fileId.indexOf(texto) !== -1 ||
          nombre.indexOf(texto) !== -1 || hallazgos.indexOf(texto) !== -1 || 
          reincidencias.indexOf(texto) !== -1 || usuario.indexOf(texto) !== -1 ||
          riesgo.indexOf(texto) !== -1) {
        logs.push(crearLogDesdeFila(row));
      }
    }
    return logs;
  } catch (e) {
    console.error("❌ [BUSCAR] Error:", e.toString());
    return [];
  }
}

function crearLogDesdeFila(row) {
  var fechaObj = row[0];
  var fechaStr = fechaObj instanceof Date 
    ? Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    : String(fechaObj);
  
  return {
    fecha: fechaStr,
    fileId: row[1] || "",
    nombre: row[2] || "Sin nombre",
    hallazgos: String(row[3] || "").replace(/[\r\n]+/g, " "),
    reincidencias: String(row[4] || "Ninguna").replace(/[\r\n]+/g, " "),
    usuario: row[5] || "",
    riesgo: row[6] || "Medio"
  };
}

// =============================================================
// NUEVA FUNCIÓN DE REGISTRO PARA MAESTRO (DINÁMICA)
// =============================================================

function registrarNuevoDocumento(formData) {
  try {
    const FOLDER_ID = getParam("ID_CARPETA_VIGENTES");
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName("LISTADO_MAESTRO");
    
    if (!FOLDER_ID) throw new Error("No se encontró ID_CARPETA_VIGENTES en CONFIG");

    var folder = DriveApp.getFolderById(FOLDER_ID);
    var fileId = "";

    if (formData.fileBase64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(formData.fileBase64), formData.fileType, formData.fileName);
      var file = folder.createFile(blob);
      fileId = file.getId();
    }

    var lastRow = sheet.getLastRow();
    var nextId = lastRow < 2 ? 1 : parseInt(sheet.getRange(lastRow, 1).getValue()) + 1;

    sheet.appendRow([
      nextId,
      formData.codigo,
      formData.nombre,
      formData.version,
      "Vigente"
    ]);

    return { success: true, message: "Documento registrado", fileId: fileId };
  } catch (e) {
    console.error("❌ [Maestro] Error:", e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Registra un documento en el Listado Maestro y sube el archivo 
 * utilizando la base de datos externa definida en SS_ID.
 */
function registrarDocumentoMaestro(payload) {
  try {
    // 1. Abrir la hoja de cálculo externa usando el ID global
    var ss = SpreadsheetApp.openById(SS_ID);
    
    // 2. Obtener el ID de la carpeta dinámicamente usando tu función getParam
    var folderIdVigentes = getParam("ID_CARPETA_VIGENTES");
    
    if (!folderIdVigentes) {
      throw new Error("No se encontró el parámetro ID_CARPETA_VIGENTES en la hoja CONFIG.");
    }

    // 3. Acceder a la carpeta de Drive y crear el archivo
    var carpeta = DriveApp.getFolderById(folderIdVigentes);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(payload.archivo.base64), 
      payload.archivo.mimeType, 
      payload.archivo.nombre
    );
    var archivoSubido = carpeta.createFile(blob);
    
    // 4. Registrar los datos en la hoja LISTADO_MAESTRO de la base de datos externa
    var sheetMaestro = ss.getSheetByName("LISTADO_MAESTRO");
    if (!sheetMaestro) {
      throw new Error("No se encontró la pestaña 'LISTADO_MAESTRO' en la base de datos.");
    }
    
    // Obtenemos el último ID numérico para mantener la consistencia que tenías antes
    var lastRow = sheetMaestro.getLastRow();
    var nextId = lastRow < 2 ? 1 : (parseInt(sheetMaestro.getRange(lastRow, 1).getValue()) || lastRow) + 1;

    // Insertar fila: ID, Código, Nombre, Versión, Link, Usuario, Fecha
    sheetMaestro.appendRow([
      nextId,                             // A: ID Correlativo
      payload.codigo,                     // B: Código
      payload.nombre,                     // C: Nombre
      payload.version,                    // D: Versión
      archivoSubido.getUrl(),             // E: Enlace al archivo (Drive)
      Session.getActiveUser().getEmail(), // F: Usuario que registró
      new Date()                          // G: Fecha de registro
    ]);
    
    return { 
      success: true, 
      message: "Documento registrado con éxito", 
      fileId: archivoSubido.getId() 
    };
    
  } catch (e) {
    console.error("❌ [registrarDocumentoMaestro] Error:", e.toString());
    return { success: false, error: e.toString() };
  }
}