// ============================================================================
// SISTEMA DE AUTOMATIZACIÓN: REQUERIMIENTOS LEGALES
// Código AppScript VERSIÓN TESTING (busca en label específico)
// ============================================================================

// ============================================================================
// VERSIÓN: TESTING (busca en label "Testing_Reqs")
// Una vez valides todo, cambias a VERSIÓN PRODUCCIÓN
// ============================================================================

// ============================================================================
// 1. SETUP: Crear el trigger automático al recibir correos
// ============================================================================

/**
 * Ejecuta una sola vez para crear el trigger automático
 * Copia esta función en el Editor, corre una sola vez, listo.
 */
function crearTriggerAutomatico() {
  // Elimina triggers antiguos (para evitar duplicados)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'verificaCorreosNuevos') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crea trigger de tiempo: cada 5 min busca correos nuevos
  ScriptApp.newTrigger('verificaCorreosNuevos')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log("✅ Trigger creado: verificará correos cada 5 minutos");
}

// ============================================================================
// 2. VERIFICACIÓN DE CORREOS NUEVOS (VERSIÓN TESTING)
// ============================================================================

function verificaCorreosNuevos() {
  try {
    // ===== VERSIÓN TESTING: Busca en label "Testing_Reqs" =====
    const testingLabel = GmailApp.getUserLabelByName("Testing_Reqs");
    
    if (!testingLabel) {
      Logger.log("⚠️ Label 'Testing_Reqs' no encontrado. ¿Lo creaste en Gmail?");
      Logger.log("   1. Abre Gmail");
      Logger.log("   2. Crea un label llamado 'Testing_Reqs'");
      Logger.log("   3. Marca tus correos con ese label");
      return;
    }
    
    // Busca correos SIN procesar en label Testing_Reqs
    const threads = testingLabel.getThreads(0, 50);
    
    Logger.log(`📧 Correos en Testing_Reqs: ${threads.length}`);
    
    if (threads.length === 0) {
      Logger.log("   (Sin correos nuevos para procesar)");
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("REQ LEGALES SEC HACIENDA");
    
    threads.forEach((thread, idx) => {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      
      Logger.log(`\n--- PROCESANDO CORREO ${idx + 1} ---`);
      Logger.log(`Asunto: ${lastMessage.getSubject()}`);
      Logger.log(`De: ${lastMessage.getFrom()}`);
      Logger.log(`Fecha: ${lastMessage.getDate()}`);
      
      procesarCorreo(lastMessage, sheet);
      
      // Mueve a carpeta "Procesados" después de procesar
      const processedLabel = GmailApp.getUserLabelByName("Procesados") || 
                            GmailApp.createLabel("Procesados");
      thread.addLabel(processedLabel);
      thread.removeLabel(testingLabel);
      
      Logger.log("✅ Movido a 'Procesados'");
    });
    
    Logger.log(`\n✅ Procesamiento completado: ${threads.length} correo(s)`);
    
  } catch (error) {
    Logger.log(`❌ ERROR en verificaCorreosNuevos: ${error}`);
    Logger.log(error.stack);
  }
}

// ============================================================================
// 3. PROCESAR CORREO INDIVIDUAL
// ============================================================================

function procesarCorreo(mensaje, sheet) {
  try {
    const asunto = mensaje.getSubject();
    const cuerpo = mensaje.getPlainBody();
    const emailFrom = mensaje.getFrom();
    const emailTimestamp = mensaje.getDate();
    
    Logger.log(`🔍 Extrayendo datos...`);
    
    // Extrae todos los datos del correo
    const datos = extraeYValidaTodoDelCorreo(asunto, cuerpo);
    
    Logger.log(`
      Municipio: ${datos.municipio}
      Departamento: ${datos.departamento}
      Tipo Oficio: ${datos.tipo_oficio}
      Número: ${datos.numero_oficio}
      Vencimiento: ${datos.fecha_vencimiento}
    `);
    
    // Genera ID único
    const id_req = generaID_REQ();
    
    // Crea fila nueva
    const nuevaFila = [
      id_req,                          // 1. ID_REQ
      "por-completar",               // 2. STATUS
      emailTimestamp,                 // 3. EMAIL_TIMESTAMP
      "=TODAY()-C" + (sheet.getLastRow() + 1), // 4. DÍAS_ATRASO (fórmula)
      asunto,                         // 5. ASUNTO_CORREO
      datos.fecha_recepcion_area,     // 6. FECHA_RECEPCIÓN_ÁREA
      "",                             // 7. FECHA_RECEPCIÓN_BANCO (manual)
      datos.municipio,                // 8. CIUDAD_MUNICIPIO
      datos.departamento,             // 9. DEPARTAMENTO
      datos.tipo_oficio,              // 10. TIPO_OFICIO
      datos.consulta_realizada_por,   // 11. CONSULTA_REALIZADA_POR
      "",                             // 12. RESPONSABLE_ÁREA (manual)
      "",                             // 13. CASO (manual)
      "",                             // 14. ACCIÓN_EJECUTADA (manual)
      datos.fecha_vencimiento,        // 15. FECHA_VENCIMIENTO
      "",                             // 16. FECHA_RESPUESTA_ENVIADA (manual)
      "",                             // 17. MEDIO_RESPUESTA (manual)
      "",                             // 18. QUIÉN_TIENE_FÍSICO (manual)
      "",                             // 19. OBSERVACIONES (manual)
      new Date()                      // 20. ÚLTIMA_ACTUALIZACIÓN
    ];
    
    // Agrega fila a Sheets
    sheet.appendRow(nuevaFila);
    
    Logger.log(`✅ Fila agregada - ID: ${id_req}`);
    
    // Actualiza tabla de municipios detectados
    actualizaMunicipioEnTabla(datos.municipio, datos.departamento);
    
    // Crea recordatorio para 1 día después
    crearRecordatorio(id_req, emailTimestamp, 1);
    
  } catch (error) {
    Logger.log(`❌ ERROR procesando correo: ${error}\n${error.stack}`);
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
  
  // ===== EXTRAE MUNICIPIO =====
  Logger.log(`[REGEX] Buscando municipio...`);
  
  let match = texto.match(/Secretaría de Hacienda de\s+([A-Za-záéíóúñÁÉÍÓÚÑ\s]+?)(?:\n|,|$|\s-)/i);
  if (match) {
    datos.municipio = match[1].trim();
    Logger.log(`  ✅ Encontrado (patrón 1): "${datos.municipio}"`);
  } else {
    // Patrón 2: "Municipio-Departamento" o "Municipio, Departamento"
    match = texto.match(/([A-Za-záéíóúñÁÉÍÓÚÑ]+)(?:,|\s-)\s+([A-Za-záéíóúñÁÉÍÓÚÑ]+)$/m);
    if (match) {
      datos.municipio = match[1].trim();
      datos.departamento = match[2].trim();
      Logger.log(`  ✅ Encontrado (patrón 2): "${datos.municipio}" - "${datos.departamento}"`);
    } else {
      Logger.log(`  ❌ No se encontró municipio`);
    }
  }
  
  // ===== EXTRAE TIPO DE OFICIO =====
  Logger.log(`[REGEX] Buscando tipo de oficio...`);
  
  if (textoMin.includes("pliego de cargos")) {
    datos.tipo_oficio = "Pliego de Cargos";
    match = texto.match(/PC\s+(\d{4,})/i);
    if (match) datos.numero_oficio = "PC " + match[1];
    Logger.log(`  ✅ Tipo: Pliego de Cargos | Número: ${datos.numero_oficio}`);
  } 
  else if (textoMin.includes("emplazamiento")) {
    datos.tipo_oficio = "Emplazamiento";
    match = texto.match(/Emplazamiento[^\n]*?No(?:\.|:)?\s*(\d+)/i);
    if (match) datos.numero_oficio = match[1];
    Logger.log(`  ✅ Tipo: Emplazamiento | Número: ${datos.numero_oficio}`);
  }
  else if (textoMin.includes("requerimiento")) {
    datos.tipo_oficio = "Requerimiento";
    match = texto.match(/Oficio\s*[\:\#]?\s*(\d+)/i);
    if (match) datos.numero_oficio = match[1];
    Logger.log(`  ✅ Tipo: Requerimiento | Número: ${datos.numero_oficio}`);
  }
  else if (textoMin.includes("certificación")) {
    datos.tipo_oficio = "Certificación";
    match = texto.match(/solicitud[^\n]*?(\d+)/i);
    if (match) datos.numero_oficio = match[1];
    Logger.log(`  ✅ Tipo: Certificación | Número: ${datos.numero_oficio}`);
  }
  else {
    Logger.log(`  ⚠️ Tipo de oficio no identificado`);
  }
  
  // ===== EXTRAE FECHA VENCIMIENTO =====
  Logger.log(`[REGEX] Buscando fecha de vencimiento...`);
  
  // Prioridad 1: Busca palabra "vencimiento"
  match = texto.match(/vencimiento[^\n]*?(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/i);
  if (match) {
    datos.fecha_vencimiento = convertirFecha(match[1], match[2], match[3]);
    Logger.log(`  ✅ Encontrada (palabra clave): ${formatearFecha(datos.fecha_vencimiento)}`);
  } else {
    // Prioridad 2: Formato DD/MM/YYYY
    match = texto.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if (match) {
      datos.fecha_vencimiento = new Date(match[3], match[2]-1, match[1]);
      Logger.log(`  ✅ Encontrada (DD/MM/YYYY): ${formatearFecha(datos.fecha_vencimiento)}`);
    }
  }
  
  // ===== EXTRAE PERIODOS (si existen) =====
  const periodos = texto.match(/\d{2}-\d{4}/g);
  if (periodos) {
    datos.periodos_afectados = [...new Set(periodos)];
    Logger.log(`  ✅ Períodos: ${datos.periodos_afectados.join(", ")}`);
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
  // Genera: REQ-YYYY-MM-DD-NNN
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
    
    // Si no existe, la crea
    if (!refSheet) {
      refSheet = ss.insertSheet("Municipios_Detectados");
      refSheet.appendRow([
        "Municipio", "Departamento", "Primer_Avistamiento", 
        "Última_Actualización", "Validado"
      ]);
    }
    
    const datos = refSheet.getRange(2, 1, refSheet.getLastRow()-1, 5).getValues();
    
    // Verifica si municipio ya existe
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0] === municipio) {
        refSheet.getRange(i + 2, 4).setValue(new Date()); // Actualiza timestamp
        Logger.log(`📋 Municipio actualizado: ${municipio}`);
        return;
      }
    }
    
    // Si no existe, agrega nueva fila
    refSheet.appendRow([
      municipio,
      departamento || "[vacío - revisar]",
      new Date(),
      new Date(),
      "❌ Revisar"
    ]);
    Logger.log(`📋 Municipio agregado: ${municipio}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error actualizando municipios: ${error}`);
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
    Logger.log(`⏰ Recordatorio creado para ${idReq} en ${diasDespues} día(s)`);
    
  } catch (error) {
    Logger.log(`⚠️ Error creando recordatorio: ${error}`);
  }
}

// ============================================================================
// 6. TESTING: Función para pruebas manuales
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
    },
    {
      nombre: "Template 2: Bogotá - Emplazamiento",
      asunto: "REQ-TEST-002 | Emplazamiento No. 2024-5847",
      cuerpo: `La Secretaría de Hacienda de Bogotá-Cundinamarca requiere:

Emplazamiento No: 2024-5847
Periodos: 01-2024, 02-2024, 03-2024

Vencimiento: 28 de septiembre de 2024

Bogotá, Cundinamarca`
    }
  ];
  
  ejemplos.forEach((ejemplo) => {
    Logger.log(`\n${'='.repeat(60)}`);
    Logger.log(`📝 ${ejemplo.nombre}`);
    Logger.log(`${'='.repeat(60)}`);
    
    const datos = extraeYValidaTodoDelCorreo(ejemplo.asunto, ejemplo.cuerpo);
    
    Logger.log(`RESULTADOS:`);
    Logger.log(`  Municipio: ${datos.municipio}`);
    Logger.log(`  Departamento: ${datos.departamento}`);
    Logger.log(`  Tipo Oficio: ${datos.tipo_oficio}`);
    Logger.log(`  Número: ${datos.numero_oficio}`);
    Logger.log(`  Vencimiento: ${formatearFecha(datos.fecha_vencimiento)}`);
  });
}

// ============================================================================
// 7. INSTALACIÓN Y CAMBIO A PRODUCCIÓN
// ============================================================================

/**
 * PASOS PARA INSTALAR:
 * 
 * TESTING (AHORA):
 * 1. Copia TODO este código
 * 2. Apps Script → pega
 * 3. Ejecuta: crearTriggerAutomatico()
 * 4. Crea label en Gmail: "Testing_Reqs"
 * 5. Envíate correos a ti misma
 * 6. Marca con label "Testing_Reqs"
 * 7. Espera 5 min → chequea Sheets
 * 
 * PRODUCCIÓN (Después):
 * Cambias la función verificaCorreosNuevos() para:
 * 
 * const threads = GmailApp.search(
 *   'to:notificacionesjudiciales@davivienda.co -label:Procesado newer_than:1d'
 * );
 * 
 * Y listo. Ahora captura correos reales.
 */