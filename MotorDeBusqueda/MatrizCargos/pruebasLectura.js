// ============================================================================
// SISTEMA DE AUTOMATIZACIÓN: REQUERIMIENTOS LEGALES - CÓDIGO COMPLETO
// ============================================================================
// Este es el código FINAL con KPIs integrados
// Copia TODO esto tal cual a Apps Script
// ============================================================================

// ============================================================================
// 1. SETUP: Crear el trigger automático
// ============================================================================

function crearTriggerAutomatico() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'verificaCorreosNuevos') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('verificaCorreosNuevos')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log("✅ Trigger creado: verificará correos cada 5 minutos");
}

// ============================================================================
// 2. VERIFICACIÓN DE CORREOS NUEVOS (TESTING)
// ============================================================================

function verificaCorreosNuevos() {
  try {
    const testingLabel = GmailApp.getUserLabelByName("Testing_Reqs");
    
    if (!testingLabel) {
      Logger.log("⚠️ Label 'Testing_Reqs' no encontrado");
      return;
    }
    
    const threads = testingLabel.getThreads(0, 50);
    Logger.log(`📧 Correos en Testing_Reqs: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log("   (Sin correos nuevos)");
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("REQ LEGALES SEC HACIENDA");
    
    threads.forEach((thread, idx) => {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      
      Logger.log(`\n--- CORREO ${idx + 1} ---`);
      Logger.log(`Asunto: ${lastMessage.getSubject()}`);
      
      procesarCorreo(lastMessage, sheet);
      
      const processedLabel = GmailApp.getUserLabelByName("Procesados") || 
                            GmailApp.createLabel("Procesados");
      thread.addLabel(processedLabel);
      thread.removeLabel(testingLabel);
      
      Logger.log("✅ Procesado");
    });
    
    Logger.log(`\n✅ Completado: ${threads.length} correo(s)`);
    
  } catch (error) {
    Logger.log(`❌ ERROR: ${error}`);
  }
}

// ============================================================================
// 3. PROCESAR CORREO INDIVIDUAL
// ============================================================================

function procesarCorreo(mensaje, sheet) {
  try {
    const asunto = mensaje.getSubject();
    const cuerpo = mensaje.getPlainBody();
    const emailTimestamp = mensaje.getDate();
    
    Logger.log(`🔍 Extrayendo datos...`);
    
    const datos = extraeYValidaTodoDelCorreo(asunto, cuerpo);
    
    Logger.log(`
      Municipio: ${datos.municipio}
      Tipo Oficio: ${datos.tipo_oficio}
      Vencimiento: ${datos.fecha_vencimiento}
    `);
    
    const id_req = generaID_REQ();
    
    const nuevaFila = [
      id_req,
      "por-completar",
      emailTimestamp,
      "=TODAY()-C" + (sheet.getLastRow() + 1),
      asunto,
      datos.fecha_recepcion_area,
      "",
      datos.municipio,
      datos.departamento,
      datos.tipo_oficio,
      datos.consulta_realizada_por,
      "",
      "",
      "",
      datos.fecha_vencimiento,
      "",
      "",
      "",
      "",
      new Date()
    ];
    
    sheet.appendRow(nuevaFila);
    Logger.log(`✅ Fila agregada - ID: ${id_req}`);
    
    actualizaMunicipioEnTabla(datos.municipio, datos.departamento);
    crearRecordatorio(id_req, emailTimestamp, 1);
    
    // ⭐ AQUÍ SE ACTUALIZA KPIs AUTOMÁTICAMENTE
    actualizaKPIs();
    
  } catch (error) {
    Logger.log(`❌ ERROR: ${error}`);
  }
}

// ============================================================================
// 4. EXTRACCIÓN DE DATOS (REGEX)
// ============================================================================

function extraeYValidaTodoDelCorreo(asunto, cuerpo) {
  const texto = asunto + "\n" + cuerpo;
  const textoMin = texto.toLowerCase();
  
  let datos = {
    municipio: "No identificado",
    departamento: "",
    tipo_oficio: "Otro",
    numero_oficio: "",
    fecha_recepcion_area: new Date(),
    fecha_vencimiento: null,
    consulta_realizada_por: "Sec. Hacienda",
    periodos_afectados: []
  };
  
  // EXTRAE MUNICIPIO
  Logger.log(`[REGEX] Buscando municipio...`);
  let match = texto.match(/Secretaría de Hacienda de\s+([A-Za-záéíóúñÁÉÍÓÚÑ\s]+?)(?:\n|,|$|\s-)/i);
  if (match) {
    datos.municipio = match[1].trim();
    Logger.log(`  ✅ Encontrado: "${datos.municipio}"`);
  }
  
  // EXTRAE TIPO DE OFICIO
  Logger.log(`[REGEX] Buscando tipo de oficio...`);
  if (textoMin.includes("pliego de cargos")) {
    datos.tipo_oficio = "Pliego de Cargos";
    match = texto.match(/PC\s+(\d{4,})/i);
    if (match) datos.numero_oficio = "PC " + match[1];
  } 
  else if (textoMin.includes("emplazamiento")) {
    datos.tipo_oficio = "Emplazamiento";
    match = texto.match(/Emplazamiento[^\n]*?No(?:\.|:)?\s*(\d+)/i);
    if (match) datos.numero_oficio = match[1];
  }
  else if (textoMin.includes("requerimiento")) {
    datos.tipo_oficio = "Requerimiento";
    match = texto.match(/Oficio\s*[\:\#]?\s*(\d+)/i);
    if (match) datos.numero_oficio = match[1];
  }
  else if (textoMin.includes("certificación")) {
    datos.tipo_oficio = "Certificación";
    match = texto.match(/solicitud[^\n]*?(\d+)/i);
    if (match) datos.numero_oficio = match[1];
  }
  
  // EXTRAE FECHA VENCIMIENTO
  Logger.log(`[REGEX] Buscando fecha...`);
  match = texto.match(/vencimiento[^\n]*?(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/i);
  if (match) {
    datos.fecha_vencimiento = convertirFecha(match[1], match[2], match[3]);
  } else {
    match = texto.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (match) {
      datos.fecha_vencimiento = new Date(match[3], match[2]-1, match[1]);
    }
  }
  
  // EXTRAE PERIODOS
  const periodos = texto.match(/\d{2}-\d{4}/g);
  if (periodos) {
    datos.periodos_afectados = [...new Set(periodos)];
  }
  
  return datos;
}

function convertirFecha(dia, mesTexto, año) {
  const meses = {
    enero: 0, febrero: 1, marzo: 2, abril: 3, mayo: 4, junio: 5,
    julio: 6, agosto: 7, septiembre: 8, octubre: 9, noviembre: 10, diciembre: 11
  };
  const mes = meses[mesTexto.toLowerCase()];
  if (mes === undefined) return null;
  return new Date(año, mes, dia);
}

function formatearFecha(date) {
  if (!date) return "N/A";
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

// ============================================================================
// 5. FUNCIONES DE UTILIDAD
// ============================================================================

function generaID_REQ() {
  const hoy = new Date();
  const año = hoy.getFullYear();
  const mes = (hoy.getMonth() + 1).toString().padStart(2, '0');
  const dia = hoy.getDate().toString().padStart(2, '0');
  const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return `REQ-${año}-${mes}-${dia}-${random}`;
}

function actualizaMunicipioEnTabla(municipio, departamento) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let refSheet = ss.getSheetByName("Municipios_Detectados");
    
    if (!refSheet) {
      refSheet = ss.insertSheet("Municipios_Detectados");
      refSheet.appendRow([
        "Municipio", "Departamento", "Primer_Avistamiento", 
        "Última_Actualización", "Validado"
      ]);
    }
    
    const datos = refSheet.getRange(2, 1, refSheet.getLastRow()-1, 5).getValues();
    
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0] === municipio) {
        refSheet.getRange(i + 2, 4).setValue(new Date());
        Logger.log(`📋 Municipio actualizado: ${municipio}`);
        return;
      }
    }
    
    refSheet.appendRow([
      municipio,
      departamento || "[vacío]",
      new Date(),
      new Date(),
      "❌ Revisar"
    ]);
    Logger.log(`📋 Municipio agregado: ${municipio}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error municipios: ${error}`);
  }
}

function crearRecordatorio(idReq, emailTimestamp, diasDespues) {
  try {
    const props = PropertiesService.getScriptProperties();
    const recordatorios = JSON.parse(props.getProperty("recordatorios") || "{}");
    
    recordatorios[idReq] = {
      emailTimestamp: emailTimestamp.toISOString(),
      diasDespues: diasDespues,
      creado: new Date().toISOString()
    };
    
    props.setProperty("recordatorios", JSON.stringify(recordatorios));
    Logger.log(`⏰ Recordatorio creado para ${idReq}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error recordatorio: ${error}`);
  }
}

// ============================================================================
// 6. ACTUALIZAR KPIs AUTOMÁTICAMENTE ⭐ NUEVO
// ============================================================================

function actualizaKPIs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetReqs = ss.getSheetByName("REQ LEGALES SEC HACIENDA");
    let sheetKPIs = ss.getSheetByName("KPIs");
    
    // Si no existe hoja KPIs, la crea
    if (!sheetKPIs) {
      sheetKPIs = crearHojaKPIs();
    }
    
    // Obtiene datos
    const ultimaFila = sheetReqs.getLastRow();
    if (ultimaFila <= 1) {
      Logger.log("⚠️ Sin requerimientos aún");
      return;
    }
    
    const datos = sheetReqs.getRange(2, 1, ultimaFila-1, 20).getValues();
    
    // Calcula KPIs
    let totalReqs = datos.length;
    let resueltos = 0;
    let porCompletados = 0;
    let enAtraso = 0;
    
    datos.forEach(row => {
      const status = row[1];
      const fechaVencimiento = new Date(row[14]);
      
      if (status === "Resuelto") {
        resueltos++;
      } else if (status === "por-completar") {
        porCompletados++;
      }
      
      if (fechaVencimiento < new Date() && status !== "Resuelto") {
        enAtraso++;
      }
    });
    
    // Calcula % cumplimiento
    const porcentajeCumplimiento = totalReqs > 0 ? 
      ((resueltos / totalReqs) * 100).toFixed(1) : 0;
    
    // Crea fila de KPIs
    const filaKPIs = [
      new Date(),
      totalReqs,
      resueltos,
      porCompletados,
      enAtraso,
      porcentajeCumplimiento + "%",
      "Actualizado"
    ];
    
    // Agrega fila a KPIs
    sheetKPIs.appendRow(filaKPIs);
    
    Logger.log(`✅ KPIs: Total=${totalReqs}, Resueltos=${resueltos}, Atraso=${enAtraso}, Cumpl=${porcentajeCumplimiento}%`);
    
  } catch (error) {
    Logger.log(`❌ Error KPIs: ${error}`);
  }
}

function crearHojaKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet("KPIs", 1);
  
  const headers = [
    "FECHA_ACTUALIZACIÓN",
    "TOTAL_REQS",
    "RESUELTOS",
    "POR_COMPLETAR",
    "EN_ATRASO",
    "% CUMPLIMIENTO",
    "STATUS"
  ];
  
  sheet.appendRow(headers);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  sheet.setColumnWidth(1, 22);
  sheet.setColumnWidth(2, 14);
  sheet.setColumnWidth(3, 12);
  sheet.setColumnWidth(4, 15);
  sheet.setColumnWidth(5, 12);
  sheet.setColumnWidth(6, 16);
  sheet.setColumnWidth(7, 15);
  
  sheet.freezeRows(1);
  
  Logger.log("✅ Hoja KPIs creada");
  
  return sheet;
}

// ============================================================================
// 7. TESTING: Función para pruebas manuales
// ============================================================================

function testeoExtraccion() {
  const ejemplos = [
    {
      nombre: "Template 1: Medellín - Pliego de Cargos",
      asunto: "REQ-TEST-001 | Oficio No. PC 0156 del 18 de septiembre de 2024",
      cuerpo: `Estimado Banco Davivienda,

Por este medio, la Secretaría de Hacienda de Medellín, le comunica 
un PLIEGO DE CARGOS.

Número: PC 0156
Fecha vencimiento: 08 de octubre de 2024

Medellín, Antioquia`
    }
  ];
  
  ejemplos.forEach((ejemplo) => {
    Logger.log(`\n${'='.repeat(60)}`);
    Logger.log(`📝 ${ejemplo.nombre}`);
    Logger.log(`${'='.repeat(60)}`);
    
    const datos = extraeYValidaTodoDelCorreo(ejemplo.asunto, ejemplo.cuerpo);
    
    Logger.log(`RESULTADOS:`);
    Logger.log(`  Municipio: ${datos.municipio}`);
    Logger.log(`  Tipo Oficio: ${datos.tipo_oficio}`);
    Logger.log(`  Vencimiento: ${formatearFecha(datos.fecha_vencimiento)}`);
  });
}

// ============================================================================
// ✅ FIN DEL CÓDIGO
// ============================================================================
// 
// PASOS DE INSTALACIÓN:
// 
// 1. DESCARGA EXCEL
//    ├─ Descarga: Requerimientos_Legales_Template.xlsx
//    └─ Sube a Google Drive
// 
// 2. CONVIERTE A SHEETS
//    ├─ Abre en Google Drive
//    └─ Clic derecho → Abrir con → Google Sheets
// 
// 3. COPIA TODO ESTE CÓDIGO
//    ├─ Extensions → Apps Script
//    ├─ Elimina código por defecto
//    └─ Pega TODO este código
// 
// 4. EJECUTA TRIGGER
//    ├─ Clic en crearTriggerAutomatico()
//    ├─ Ejecuta
//    └─ Autoriza permisos
// 
// 5. SETUP GMAIL
//    ├─ Abre Gmail
//    ├─ Crea label: "Testing_Reqs"
//    └─ Guarda
// 
// 6. PRUEBA
//    ├─ Envíate a ti misma un correo (de CORREOS-REBECA-A-REBECA.md)
//    ├─ Marca con label "Testing_Reqs"
//    ├─ Espera 5 minutos
//    └─ Chequea Sheets:
//        ├─ Aparece fila en "REQ LEGALES SEC HACIENDA" ✅
//        ├─ Aparece hoja "KPIs" ✅
//        └─ Se agrega fila de KPIs ✅
// 
// ✅ LISTO - Sistema funcionando
//