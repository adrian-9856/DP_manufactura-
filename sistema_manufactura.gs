/**
 * ============================================================================
 * SISTEMA MANUFACTURA - COMPLETO
 * ============================================================================
 */

const NOMBRE_PROGRAMA = "NOMBRE DEL PROGRAMA";

// ── Hojas del sistema ─────────────────────────────────────────────────────────
const SHEET_RECEPCIONES    = "Recepciones";
const SHEET_CATALOGOS      = "Participantes Activos";
const SHEET_INACTIVOS      = "Participantes Inactivos";
const SHEET_RESUMEN_PART   = "Resumen Participantes";
const SHEET_PAGOS_PEND     = "Pagos Pendientes";
const SHEET_DASHBOARD      = "Dashboard";
const SHEET_REPORTES       = "Reportes PowerBI";
const SHEET_PERIODOS       = "PERIODOS";
const SHEET_HIST_QUINCENAS = "Historial Quincenas";
const SHEET_CHEQUES        = "Cheques";
const SHEET_TRANSFERENCIAS = "Transferencias";
const SHEET_PROGRAMAS      = "Programas_Participacion";
const SHEET_HISTORIAL_PAGOS= "Historial_Pagos";
const SHEET_ARCHIVO_REC    = "Archivo_Recepciones";
const SHEET_CATALOGO_PROD  = "Catálogo_Productos";

// ── Google Sheets externos ─────────────────────────────────────────────────────
const SS_ID_WOMEN_PAYMENT  = "1e3zlQQ827h_uXGE-0GryvpPs7r7kki0c6wcmc_jzIHk"; // Women Payment 26 (Transferencias)
const SS_ID_CHEQUES_EXT    = "14yNGv0ce8heGSeJ5L4EnUQFryJfF1n3NeTT3FUky02M"; // Cheques externo

const HEADER_ROW     = 4;
const DATA_START_ROW = 5;

// ── Mejora 1: Quincena insertada entre Fecha entrega y Participante ───────────
const RECEPCIONES_HEADERS = [
  "#", "Fecha entrega", "Quincena", "Participante", "Creamos ID", "Proyecto / Cliente", "Categoría", "Producto",
  "Unidades buenas", "Unidades rechazadas", "Precio unit. (Q)", "Total Q",
  "Impuesto PC (5%)", "Total a Pagar", "Estado pago"
];

const CAT_OPCIONES = ["Bisuteria","Servicios","Costura","Pulsera Project","Niñera","Brackish"];

// Columnas del catálogo de participantes (0-indexed)
const CAT_COL = {
  ID:             0,  // A – Creamos ID
  NOMBRE:         1,  // B – Nombre Completo
  ETAPA:          2,  // C – Estado
  NUM_CUENTA:     3,  // D – Nº Cuenta
  TIPO_CUENTA:    4,  // E – Tipo Cuenta
  BANCO:          5,  // F – Banco
  TITULAR:        6,  // G – Titular
  DPI:            7,  // H – DPI
  EDUCACION:      8,  // I – Educación (checkbox)
  INCLUSION_LAB:  9,  // J – Inclusión Laboral (checkbox)
  APOYO_EMOC:    10,  // K – Apoyo Emocional (checkbox)
  HIJOS:         11,  // L – N° Hijos
  CCI:           12,  // M – CCI (checkbox)
  CUNDE:         13,  // N – CUNDE (checkbox)
  CUENTA_PAGO:   14,  // O – Cuenta de Pago (Creamos / mi-eelo)
};

// ========================= MENU =========================
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    // Botón independiente — acceso directo a Nueva Entrega
    ui.createMenu("➕ Nueva Entrega")
      .addItem("➕ Abrir formulario de entrega", "agregarEntregaRapida")
      .addToUi();

    // Menú principal del sistema
    ui.createMenu("⚙️ Manufactura")
      // — Operaciones diarias —
      .addItem("✏️ Editar fila seleccionada",       "editarFilaSeleccionada")
      .addItem("✅ Marcar fila como pagada",         "marcarFilaPagada")
      .addItem("🔍 Historial de participante",       "buscarHistorialParticipante")
      .addSeparator()
      // — Quincenas —
      .addItem("🗓️ Crear quincena inicial",         "crearQuincenaInicial")
      .addItem("📅 Cerrar quincena actual",         "cerrarQuincenaActual")
      .addItem("🗃️ Cerrar mes manualmente",         "cerrarMes")
      .addSeparator()
      // — Mantenimiento —
      .addItem("🔢 Reparar montos y totales",       "repararRedondeoMontos")
      .addItem("🧹 Limpiar datos de prueba",        "limpiarDatosPrueba")
      .addItem("📋 Actualizar participación",       "actualizarParticipacionProgramas")
      .addItem("🔄 Actualizar todo",                "actualizarTodo")
      .addSeparator()
      // — Sistema —
      .addItem("⬆️ Aplicar actualizaciones",        "aplicarActualizaciones")
      .addItem("🚀 Instalar sistema",               "instalarSistema")
      .addItem("♻️ Reinstalar sistema",             "reinstalarSistema")
      .addItem("⛔ Desinstalar sistema",             "desinstalarSistema")
      .addSeparator()
      // — Triggers —
      .addItem("⏱️ Instalar triggers",              "crearTriggers")
      .addItem("👤 Trigger inactivos",              "instalarTriggerInactivos")
      .addItem("⏹️ Eliminar triggers",              "eliminarTriggers")
      .addToUi();
  } catch (e) {
    // Silenciosamente ignorar si getUi no está disponible
  }
}

// ========================= APLICAR ACTUALIZACIONES SIN BORRAR DATOS =========================

function aplicarActualizaciones() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const ui  = SpreadsheetApp.getUi();
    const log = [];

    // 1. Participantes Activos — eliminar mi_eelo (col I) si existe, agregar Cuenta de Pago
    const cat = ss.getSheetByName(SHEET_CATALOGOS);
    if (cat) {
      const headers = cat.getRange(1, 1, 1, Math.max(cat.getLastColumn(), 16)).getValues()[0];

      // Eliminar columna mi_eelo si todavía está en col I
      if (String(headers[8] || "").toLowerCase().replace(/[_\-]/g, "").includes("mieelo")) {
        cat.deleteColumn(9); // col I = 1-based index 9
        log.push("✅ Columna 'mi_eelo' eliminada de Participantes Activos");
        // Refrescar headers tras eliminar
        headers.splice(8, 1);
      } else {
        log.push("☑️ Columna 'mi_eelo' ya no existe");
      }

      // Agregar Cuenta de Pago en col O (15 = índice 14) si no existe
      const yaExiste = headers.some(h => String(h).toLowerCase().includes("cuenta de pago"));
      if (!yaExiste) {
        const colO = 15; // columna O
        cat.getRange(1, colO)
          .setValue("Cuenta de Pago")
          .setBackground("#e65100").setFontColor("#ffffff")
          .setFontWeight("bold").setHorizontalAlignment("center");
        cat.setColumnWidth(colO, 120);
        cat.getRange(1, colO).setValue("Cuenta de Pago")
          .setBackground("#e65100").setFontColor("#ffffff").setFontWeight("bold")
          .setHorizontalAlignment("center");
        cat.getRange(2, colO, 999)
          .setBackground("#fff3e0")
          .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID)
          .setDataValidation(
            SpreadsheetApp.newDataValidation()
              .requireValueInList(["Creamos", "mi-eelo"])
              .setAllowInvalid(false).build()
          );
        log.push("✅ Columna 'Cuenta de Pago' agregada en Participantes Activos (columna O)");
      } else {
        log.push("☑️ Columna 'Cuenta de Pago' ya existía");
      }
    }

    // 2. Recepciones — insertar columna Categoría si no existe, reaplica formato Q
    const shRec = ss.getSheetByName(SHEET_RECEPCIONES);
    if (shRec) {
      const mRec = _headerMap_(shRec);
      if (!mRec["categoría"]) {
        // Insertar después de "Proyecto / Cliente" (col F)
        const colProy = mRec["proyecto / cliente"] || 6;
        shRec.insertColumnAfter(colProy);
        shRec.getRange(HEADER_ROW, colProy + 1).setValue("Categoría")
          .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
          .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");
        shRec.setColumnWidth(colProy + 1, 130);
        shRec.getRange(DATA_START_ROW, colProy + 1, 996).setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(CAT_OPCIONES)
            .setAllowInvalid(true).build()
        );
        log.push("✅ Columna 'Categoría' insertada en Recepciones (col " + colProy + "+1)");
      } else {
        log.push("☑️ Columna 'Categoría' ya existe en Recepciones");
      }

      if (shRec.getLastRow() >= DATA_START_ROW) {
        const m2 = _headerMap_(shRec);
        const firstCol = m2["precio unit. (q)"] || m2["total q"];
        if (firstCol) {
          const nRows = shRec.getLastRow() - DATA_START_ROW + 1;
          shRec.getRange(DATA_START_ROW, firstCol, nRows, 4).setNumberFormat('"Q "#,##0.00');
        }
      }
      log.push("✅ Formato Q restaurado en Recepciones");
    }

    // 3. Cheques / Transferencias — reaplica formato Q en cols G-I
    [SHEET_CHEQUES, SHEET_TRANSFERENCIAS, SHEET_HISTORIAL_PAGOS].forEach(nombre => {
      const sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 3) return;
      sh.getRange(3, 7, sh.getLastRow() - 2, 3).setNumberFormat('"Q "#,##0.00');
      log.push("✅ Formato Q restaurado en " + nombre);
    });

    // 4. Recalcular todo
    actualizarTodo();
    log.push("✅ Cálculos y colores actualizados");

    ui.alert(
      "✅ Actualizaciones aplicadas",
      log.join("\n"),
      ui.ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
  }
}

// ========================= HISTORIAL POR PARTICIPANTE =========================

// Busca todas las entregas y pagos de un participante y genera una hoja de resumen
function buscarHistorialParticipante() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const ui  = SpreadsheetApp.getUi();

    // Pedir nombre del participante
    const resp = ui.prompt(
      "🔍 Historial de participante",
      "Escribe el nombre (o parte del nombre) del participante:",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const busqueda = resp.getResponseText().trim().toLowerCase();
    if (!busqueda) return;

    // Buscar coincidencias en Participantes Activos
    const cat     = ss.getSheetByName(SHEET_CATALOGOS);
    const dataCat = cat ? cat.getDataRange().getValues().slice(1) : [];
    const coincidencias = dataCat.filter(r =>
      String(r[CAT_COL.NOMBRE] || "").toLowerCase().includes(busqueda)
    );

    if (coincidencias.length === 0) {
      return ui.alert("Sin resultados", "No se encontró ningún participante con ese nombre.", ui.ButtonSet.OK);
    }

    // Si hay varias, preguntar cuál
    let infoParticipante;
    if (coincidencias.length === 1) {
      infoParticipante = coincidencias[0];
    } else {
      const lista = coincidencias.map((r, i) => (i + 1) + ". " + r[CAT_COL.NOMBRE]).join("\n");
      const sel = ui.prompt(
        "Varias coincidencias",
        "Se encontraron " + coincidencias.length + " participantes:\n\n" + lista + "\n\nEscribe el número:",
        ui.ButtonSet.OK_CANCEL
      );
      if (sel.getSelectedButton() !== ui.Button.OK) return;
      const idx = parseInt(sel.getResponseText().trim()) - 1;
      if (isNaN(idx) || idx < 0 || idx >= coincidencias.length) return ui.alert("❌ Número no válido.");
      infoParticipante = coincidencias[idx];
    }

    const nombreBuscado = String(infoParticipante[CAT_COL.NOMBRE] || "").trim();
    _generarHojaHistorial_(ss, nombreBuscado, infoParticipante);

  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
  }
}

function _generarHojaHistorial_(ss, nombre, infoRow) {
  const NOMBRE_HOJA = "Historial_Participante";
  const fmtQ = '"Q "#,##0.00';
  const tz   = Session.getScriptTimeZone();

  // Crear o reemplazar la hoja
  let sh = ss.getSheetByName(NOMBRE_HOJA);
  if (sh) {
    sh.clear();
    sh.clearFormats();
  } else {
    sh = ss.insertSheet(NOMBRE_HOJA);
  }
  ss.setActiveSheet(sh);

  // ── Encabezado principal ──
  sh.getRange("A1:H1").merge()
    .setValue("📋 HISTORIAL DE PAGOS — " + nombre.toUpperCase())
    .setFontWeight("bold").setFontSize(13)
    .setBackground("#263238").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 32);

  // ── Datos del participante ──
  const cuentaPago = String(infoRow[CAT_COL.CUENTA_PAGO] || "—").trim();
  const banco      = String(infoRow[CAT_COL.BANCO]       || "—").trim();
  const tipoCta    = String(infoRow[CAT_COL.TIPO_CUENTA]  || "—").trim();
  const numCta     = String(infoRow[CAT_COL.NUM_CUENTA]   || "—").trim();
  const cid        = String(infoRow[CAT_COL.ID]           || "—").trim();

  sh.getRange("A2:H2").setValues([[
    "Creamos ID: " + cid,
    "Banco: " + banco,
    "Tipo cta: " + tipoCta,
    "Nº Cuenta: " + numCta,
    "Cuenta de Pago: " + cuentaPago,
    "", "", ""
  ]]).setBackground("#eceff1").setFontColor("#37474f").setFontSize(10);
  sh.getRange("A2:H2").merge();
  sh.getRange("A2").setValue(
    "Creamos ID: " + cid + "   |   Banco: " + banco + "   |   Tipo: " + tipoCta +
    "   |   Nº Cuenta: " + numCta + "   |   Cuenta de Pago: " + cuentaPago
  );

  // ── Encabezados tabla de entregas ──
  const hdrsEntregas = ["Fecha", "Quincena", "Proyecto / Servicio", "Producto",
                        "Unidades buenas", "Total Q", "Impuesto PC", "Total a Pagar"];
  sh.getRange(3, 1, 1, hdrsEntregas.length).setValues([hdrsEntregas])
    .setFontWeight("bold").setBackground("#37474f").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sh.setRowHeight(3, 24);

  // ── Leer Recepciones y filtrar por participante ──
  const shRec = ss.getSheetByName(SHEET_RECEPCIONES);
  let filaActual = 4;
  let totalGeneral = 0;

  if (shRec && shRec.getLastRow() >= DATA_START_ROW) {
    const m       = _headerMap_(shRec);
    const dataRec = shRec.getRange(DATA_START_ROW, 1, shRec.getLastRow() - DATA_START_ROW + 1, shRec.getLastColumn()).getValues();
    const colFecha = m["fecha entrega"] ? m["fecha entrega"] - 1 : 1;
    const colQ     = m["quincena"]      ? m["quincena"] - 1      : 2;
    const colProy  = m["proyecto / cliente"] ? m["proyecto / cliente"] - 1 : 5;
    const colProd  = m["producto"]      ? m["producto"] - 1      : 6;
    const colUB    = m["unidades buenas"]? m["unidades buenas"] - 1 : 7;
    const colTQ    = m["total q"]       ? m["total q"] - 1       : 10;
    const colImp   = (m["impuesto pc (5%)"] || m["impuesto pc"]) ? (m["impuesto pc (5%)"] || m["impuesto pc"]) - 1 : 11;
    const colTAP   = m["total a pagar"] ? m["total a pagar"] - 1 : 12;

    const filas = dataRec.filter(r => String(r[m["participante"] - 1] || "").trim() === nombre);

    filas.forEach(r => {
      const fecha = r[colFecha] instanceof Date
        ? Utilities.formatDate(r[colFecha], tz, "dd/MM/yyyy")
        : String(r[colFecha] || "");
      const tap   = Math.round(Number(r[colTAP]) || 0);
      totalGeneral += tap;
      sh.getRange(filaActual, 1, 1, 8).setValues([[
        fecha,
        String(r[colQ]   || ""),
        String(r[colProy] || ""),
        String(r[colProd] || ""),
        Number(r[colUB]  || 0),
        Number(r[colTQ]  || 0),
        Number(r[colImp] || 0),
        tap
      ]]);
      sh.getRange(filaActual, 6, 1, 3).setNumberFormat(fmtQ);
      sh.getRange(filaActual, 1, 1, 8).setBackground(filaActual % 2 === 0 ? "#f5f5f5" : "#ffffff");
      filaActual++;
    });

    if (filas.length === 0) {
      sh.getRange(filaActual, 1).setValue("Sin entregas registradas para este participante.");
      filaActual++;
    }
  }

  // ── Subtotal entregas ──
  filaActual++;
  sh.getRange(filaActual, 6, 1, 2).setValues([["TOTAL ENTREGAS", totalGeneral]])
    .setFontWeight("bold").setBackground("#cfd8dc");
  sh.getRange(filaActual, 7).setNumberFormat(fmtQ);
  filaActual += 2;

  // ── Encabezados tabla de pagos realizados ──
  const hdrsPagos = ["Mes", "Quincena 1", "Quincena 2", "Total Mes", "Fuente"];
  sh.getRange(filaActual, 1, 1, hdrsPagos.length).setValues([hdrsPagos])
    .setFontWeight("bold").setBackground("#1b5e20").setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  filaActual++;

  // ── Leer Cheques + Transferencias + Historial_Pagos ──
  [SHEET_CHEQUES, SHEET_TRANSFERENCIAS, SHEET_HISTORIAL_PAGOS].forEach(sheetName => {
    const shP = ss.getSheetByName(sheetName);
    if (!shP || shP.getLastRow() < 3) return;
    const data = shP.getRange(3, 1, shP.getLastRow() - 2, 11).getValues();
    data.forEach(r => {
      if (String(r[0] || "").trim() !== nombre) return;
      sh.getRange(filaActual, 1, 1, 5).setValues([[
        String(r[9] || ""),     // Mes
        Number(r[6] || 0),      // Q1
        Number(r[7] || 0),      // Q2
        Number(r[8] || 0),      // Total
        sheetName               // Fuente
      ]]);
      sh.getRange(filaActual, 2, 1, 3).setNumberFormat(fmtQ);
      sh.getRange(filaActual, 1, 1, 5).setBackground(filaActual % 2 === 0 ? "#e8f5e9" : "#f1f8e9");
      filaActual++;
    });
  });

  // ── Formato columnas ──
  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5,  90); sh.setColumnWidth(6,  90);
  sh.setColumnWidth(7,  90); sh.setColumnWidth(8,  110);
  sh.setFrozenRows(3);

  ss.toast("✅ Historial de " + nombre + " generado en hoja '" + NOMBRE_HOJA + "'.", null, 5);
}

// ========================= MIGRACIÓN DE DATOS EXISTENTES =========================
// Aplica cambios de estructura sin borrar datos: encabezados, dropdowns, valores de etapa
function migrarDatosExistentes() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const ui  = SpreadsheetApp.getUi();
    const log = [];

    // 1. Participantes Activos: encabezado C1 y valores Estado → Etapa
    const shAct = ss.getSheetByName(SHEET_CATALOGOS);
    if (shAct) {
      // Encabezado
      const hdrs = shAct.getRange(1, 1, 1, 15).getValues()[0];
      if (hdrs[2] === "Estado") {
        shAct.getRange(1, 3).setValue("Etapa");
        log.push("Participantes Activos: encabezado C → Etapa");
      }
      // Quitar validación vieja antes de escribir nuevos valores
      shAct.getRange("C2:C1000").clearDataValidations();
      // Valores de la columna
      const lastRow = shAct.getLastRow();
      if (lastRow > 1) {
        const vals = shAct.getRange(2, 3, lastRow - 1, 1).getValues();
        const newVals = vals.map(([v]) => {
          if (v === "Activo")   return ["Inscritx"];
          if (v === "Inactivo") return ["Retiradx"];
          return [v];
        });
        shAct.getRange(2, 3, lastRow - 1, 1).setValues(newVals);
        log.push("Participantes Activos: Activo → Inscritx, Inactivo → Retiradx");
      }
      // Aplicar nueva validación
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Inscritx", "Retiradx", "Empleadx", "Ciclo de Vida Terminado"])
        .setAllowInvalid(false).build();
      shAct.getRange("C2:C1000").setDataValidation(rule);
      log.push("Participantes Activos: dropdown de Etapa actualizado");
    }

    // 2. Participantes Inactivos: encabezado C1
    const shIna = ss.getSheetByName(SHEET_INACTIVOS);
    if (shIna) {
      const hdrs = shIna.getRange(1, 1, 1, 8).getValues()[0];
      if (hdrs[2] === "Estado") {
        shIna.getRange(1, 3).setValue("Etapa");
        log.push("Participantes Inactivos: encabezado C → Etapa");
      }
      // Quitar validación vieja antes de escribir
      shIna.getRange("C2:C1000").clearDataValidations();
      if (shIna.getLastRow() > 1) {
        const vals = shIna.getRange(2, 3, shIna.getLastRow() - 1, 1).getValues();
        const newVals = vals.map(([v]) => {
          if (v === "Activo")   return ["Inscritx"];
          if (v === "Inactivo") return ["Retiradx"];
          return [v];
        });
        shIna.getRange(2, 3, shIna.getLastRow() - 1, 1).setValues(newVals);
        log.push("Participantes Inactivos: valores de Etapa actualizados");
      }
    }

    // 3. Cheques y Transferencias: nueva estructura (solo si están vacías)
    [SHEET_CHEQUES, SHEET_TRANSFERENCIAS, SHEET_HISTORIAL_PAGOS].forEach(nombre => {
      const sh = ss.getSheetByName(nombre);
      if (sh && sh.getLastRow() <= 2) {
        if (nombre === SHEET_CHEQUES)        _setupCheques_(sh);
        if (nombre === SHEET_TRANSFERENCIAS) _setupTransferencias_(sh);
        if (nombre === SHEET_HISTORIAL_PAGOS)_setupHistorialPagos_(sh);
        log.push(nombre + ": estructura actualizada");
      }
    });

    // 4. Recepciones: actualizar encabezados si faltan las columnas nuevas
    const shRecMig = ss.getSheetByName(SHEET_RECEPCIONES);
    if (shRecMig) {
      const hdrsRec   = shRecMig.getRange(HEADER_ROW, 1, 1, shRecMig.getLastColumn()).getValues()[0];
      const tieneImp  = hdrsRec.some(h => String(h).toLowerCase().includes("impuesto"));
      if (!tieneImp) {
        const nCols   = RECEPCIONES_HEADERS.length;
        const lastCol = _colLetter_(nCols);
        // Reescribir fila de encabezados
        shRecMig.getRange(HEADER_ROW, 1, 1, nCols)
          .setValues([RECEPCIONES_HEADERS])
          .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
          .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");
        // Actualizar rango del título en fila 1
        shRecMig.getRange(1, 1, 1, nCols).merge();
        // Anchos para columnas nuevas
        shRecMig.setColumnWidth(12, 90); // Impuesto PC
        shRecMig.setColumnWidth(13, 90); // Total a Pagar
        // Limpiar formato en datos
        if (shRecMig.getLastRow() >= DATA_START_ROW) {
          shRecMig.getRange(`A${DATA_START_ROW}:${lastCol}1000`)
            .setBackground("#ffffff").setFontColor("#212121");
        }
        // Corregir formatos de columnas monetarias y de fecha
        shRecMig.getRange(`B${DATA_START_ROW}:B1000`).setNumberFormat("dd/MM/yyyy");
        shRecMig.getRange(`O${DATA_START_ROW}:O1000`).setNumberFormat("dd/MM/yyyy");
        shRecMig.getRange(`J${DATA_START_ROW}:M1000`).setNumberFormat('"Q "#,##0.00');
        log.push("Recepciones: encabezados y formatos actualizados");
      }
    }

    // 5. Eliminar Archivo_Recepciones si existe
    const shArc = ss.getSheetByName(SHEET_ARCHIVO_REC);
    if (shArc) {
      ss.deleteSheet(shArc);
      log.push("Archivo_Recepciones: eliminada");
    }

    const resumen = log.length ? log.join("\n• ") : "Nada que migrar — todo estaba al día.";
    ui.alert("✅ Migración completada", "• " + resumen, ui.ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error en migración: " + e.message, null, 5);
  }
}

// ========================= INSTALAR / REINSTALAR / DESINSTALAR =========================
function instalarSistema() {
  try {
    _crearEstructura_(false);
    SpreadsheetApp.getActive().toast("✅ Sistema instalado", "Listo", 4);
  } catch (e) {
    // Error manejado silenciosamente
  }
}

function reinstalarSistema() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ok = ui.alert("Reinstalar sistema", "Esto borra hojas del sistema y las crea de nuevo. ¿Continuar?", ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
    _crearEstructura_(true);
    SpreadsheetApp.getActive().toast("✅ Sistema reinstalado", "Listo", 4);
  } catch (e) {
    console.log("Error al reinstalar: " + e.message);
  }
}

function desinstalarSistema() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ok = ui.alert("Desinstalar sistema", "Esto eliminará hojas y triggers del sistema. ¿Continuar?", ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;

    const ss = SpreadsheetApp.getActive();
    [
      SHEET_RECEPCIONES, SHEET_CATALOGOS, SHEET_INACTIVOS,
      SHEET_RESUMEN_PART, SHEET_HIST_QUINCENAS, SHEET_PAGOS_PEND,
      SHEET_DASHBOARD, SHEET_REPORTES,
      SHEET_PERIODOS, SHEET_CHEQUES, SHEET_TRANSFERENCIAS,
      SHEET_PROGRAMAS, SHEET_HISTORIAL_PAGOS,
      SHEET_ARCHIVO_REC, SHEET_CATALOGO_PROD,
      "Resumen Mensual"
    ].forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    });

    eliminarTriggers();
    SpreadsheetApp.getActive().toast("🗑️ Sistema desinstalado", "Listo", 4);
  } catch (e) {
    console.log("Error al desinstalar: " + e.message);
  }
}

// Borra solo datos de transacción (pruebas). Conserva Participantes y Catálogo.
function limpiarDatosPrueba() {
  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const ok = ui.alert(
      "🧹 Limpiar datos de prueba",
      "Esto borrará:\n\n" +
      "• Recepciones (todas las entregas)\n" +
      "• Historial Quincenas\n" +
      "• PERIODOS (quincenas creadas)\n" +
      "• Cheques / Transferencias\n" +
      "• Historial_Pagos\n" +
      "• Pagos Pendientes\n" +
      "• Resumen Participantes\n\n" +
      "NO toca: Participantes Activos, Participantes Inactivos, Catálogo_Productos.\n\n" +
      "¿Continuar?",
      ui.ButtonSet.YES_NO
    );
    if (ok !== ui.Button.YES) return;

    const log = [];

    // Hojas con Título(fila1) + Encabezados(fila2) + datos desde fila 3
    const hojasFila3 = [
      SHEET_HIST_QUINCENAS,
      SHEET_CHEQUES,
      SHEET_TRANSFERENCIAS,
      SHEET_HISTORIAL_PAGOS,
    ];
    hojasFila3.forEach(nombre => {
      const sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 3) return;
      const rng = sh.getRange(3, 1, sh.getLastRow() - 2, sh.getLastColumn());
      rng.clearContent();
      rng.setBackground(null).setFontColor(null).setFontWeight("normal");
      log.push(nombre);
    });

    // Hojas con Encabezados(fila1) + datos desde fila 2
    const hojasFila2 = [
      SHEET_PAGOS_PEND,
      SHEET_RESUMEN_PART,
    ];
    hojasFila2.forEach(nombre => {
      const sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 2) return;
      const rng = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
      rng.clearContent();
      rng.setBackground(null).setFontColor(null).setFontWeight("normal");
      log.push(nombre);
    });

    // PERIODOS: encabezado en fila 1, datos desde fila 2
    const shPer = ss.getSheetByName(SHEET_PERIODOS);
    if (shPer && shPer.getLastRow() > 1) {
      const rng = shPer.getRange(2, 1, shPer.getLastRow() - 1, shPer.getLastColumn());
      rng.clearContent();
      rng.setBackground(null).setFontColor(null).setFontWeight("normal");
      log.push(SHEET_PERIODOS);
    }

    // Recepciones: encabezado en fila 4, datos desde fila 5
    const shRec = ss.getSheetByName(SHEET_RECEPCIONES);
    if (shRec) {
      const totalRows = shRec.getMaxRows();
      if (totalRows >= DATA_START_ROW) {
        const nCols  = RECEPCIONES_HEADERS.length;
        const rngDat = shRec.getRange(DATA_START_ROW, 1, totalRows - DATA_START_ROW + 1, nCols);
        rngDat.clearContent();
        rngDat.setBackground("#ffffff").setFontColor("#212121").setFontWeight("normal");
        rngDat.setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
      }
      log.push(SHEET_RECEPCIONES);
    }

    ui.alert(
      "✅ Datos de prueba eliminados",
      "Se limpiaron:\n• " + log.join("\n• ") + "\n\nParticipantes y catálogo intactos.",
      ui.ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
  }
}

function limpiarHojasObsoletas() {
  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const obsoletas = [SHEET_PROGRAMAS, "Resumen Mensual", SHEET_PAGOS_PEND, SHEET_ARCHIVO_REC];
    const existentes = obsoletas.filter(n => !!ss.getSheetByName(n));

    if (existentes.length === 0) {
      return ss.toast("✅ No hay hojas obsoletas.", null, 3);
    }

    const ok = ui.alert(
      "🗑️ Eliminar hojas obsoletas",
      "Se eliminarán permanentemente:\n\n• " + existentes.join("\n• ") + "\n\n¿Continuar?",
      ui.ButtonSet.YES_NO
    );
    if (ok !== ui.Button.YES) return;

    existentes.forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    });

    ss.toast("✅ " + existentes.length + " hoja(s) eliminada(s).", null, 4);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

// Repara Nº Cuenta (texto) y Creamos ID en filas existentes de Cheques/Transferencias
function repararDatosPago() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const cat = ss.getSheetByName(SHEET_CATALOGOS);
    const log = [];

    // Mapa nombre → { cid, numCuenta }
    const mapaParticipantes = {};
    if (cat) {
      cat.getDataRange().getValues().slice(1).forEach(r => {
        const nombre = String(r[CAT_COL.NOMBRE] || "").trim();
        if (!nombre) return;
        mapaParticipantes[nombre] = {
          cid:       String(r[CAT_COL.ID]        || "").trim(),
          numCuenta: String(r[CAT_COL.NUM_CUENTA] || "").trim(),
        };
      });
    }

    [SHEET_CHEQUES, SHEET_TRANSFERENCIAS, SHEET_HISTORIAL_PAGOS].forEach(nombre => {
      const sh = ss.getSheetByName(nombre);
      if (!sh || sh.getLastRow() < 3) return;
      const nRows = sh.getLastRow() - 2;
      const data  = sh.getRange(3, 1, nRows, 11).getValues();

      for (let i = 0; i < data.length; i++) {
        const row      = data[i];
        const nomPart  = String(row[0] || "").trim();
        if (!nomPart) continue;
        const fila     = i + 3;
        const info     = mapaParticipantes[nomPart] || {};

        // Rellenar Creamos ID si está vacío
        if (!String(row[1] || "").trim() && info.cid) {
          sh.getRange(fila, 2).setValue(info.cid);
        }
        // Forzar Nº Cuenta como texto exacto
        const cuentaActual = String(row[4] || "").trim();
        const cuentaCorr   = info.numCuenta || cuentaActual;
        sh.getRange(fila, 5).setValue(cuentaCorr).setNumberFormat("@");
      }
      log.push(nombre + " (" + nRows + " filas)");
    });

    SpreadsheetApp.getUi().alert(
      "✅ Reparación completada",
      "Se repararon:\n• " + log.join("\n• "),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
  }
}

// Redondea y formatea montos en TODAS las hojas del sistema
function repararRedondeoMontos() {
  try {
    const ss   = SpreadsheetApp.getActive();
    const ui   = SpreadsheetApp.getUi();
    const log  = [];
    const fmtQ = '"Q "#,##0.00';

    // --- helper: redondea valores numéricos en un rango y aplica formato Q ---
    function _limpiarRango_(sh, startRow, col, numCols, label) {
      if (!sh || sh.getLastRow() < startRow) return;
      const nRows = sh.getLastRow() - startRow + 1;
      const rng   = sh.getRange(startRow, col, nRows, numCols);
      const vals  = rng.getValues();
      let fixes = 0;
      for (let i = 0; i < vals.length; i++) {
        for (let j = 0; j < vals[i].length; j++) {
          const v = Number(vals[i][j]);
          if (v && !isNaN(v) && v !== Math.round(v)) { vals[i][j] = Math.round(v); fixes++; }
        }
      }
      rng.setValues(vals).setNumberFormat(fmtQ);
      log.push(label + ": " + fixes + " valor(es) ajustado(s)");
    }

    // 1. Recepciones — recalcula Impuesto y Total a Pagar desde Total Q con redondeo
    const shRec = ss.getSheetByName(SHEET_RECEPCIONES);
    if (shRec && shRec.getLastRow() >= DATA_START_ROW) {
      const m      = _headerMap_(shRec);
      const colPU  = m["precio unit. (q)"] || m["precio unit.(q)"];
      const colTQ  = m["total q"];
      const colImp = m["impuesto pc (5%)"] || m["impuesto pc"];
      const colTAP = m["total a pagar"];
      const nRows  = shRec.getLastRow() - DATA_START_ROW + 1;
      const ncols  = shRec.getLastColumn();
      const data   = shRec.getRange(DATA_START_ROW, 1, nRows, ncols).getValues();
      let fixes = 0;
      for (let i = 0; i < data.length; i++) {
        const tq = Number(data[i][(colTQ || 1) - 1]) || 0;
        if (tq <= 0) continue;
        const newImp = Math.round(tq * 0.05);
        const newTap = Math.round(tq + newImp);
        if (colImp) { data[i][colImp - 1] = newImp; fixes++; }
        if (colTAP) { data[i][colTAP - 1] = newTap; fixes++; }
      }
      shRec.getRange(DATA_START_ROW, 1, nRows, ncols).setValues(data);
      const firstCol = colPU || colTQ;
      if (firstCol) shRec.getRange(DATA_START_ROW, firstCol, nRows, 4).setNumberFormat(fmtQ);
      log.push("Recepciones: " + fixes + " valor(es) recalculado(s), formato Q aplicado");
    }

    // 2. Cheques, Transferencias, Historial_Pagos — re-sumar desde Recepciones (fuente única de verdad)
    const logPagos = repararTotalesPagos(true); // true = modo silencioso
    log.push(logPagos);

    // 3. Resumen Participantes — cols F,G,H (6,7,8) = Total generado, Total pagado, Pendiente
    _limpiarRango_(ss.getSheetByName(SHEET_RESUMEN_PART), 2, 6, 3, "Resumen Participantes");

    // 4. Historial Quincenas — col G (7) = Total_Q
    _limpiarRango_(ss.getSheetByName(SHEET_HIST_QUINCENAS), 3, 7, 1, "Historial Quincenas");

    // 5. Reportes PowerBI — col J (10) = Total a Pagar
    _limpiarRango_(ss.getSheetByName(SHEET_REPORTES), 3, 10, 1, "Reportes PowerBI");

    // 6. Pagos Pendientes — col F (6) = Total Q
    _limpiarRango_(ss.getSheetByName(SHEET_PAGOS_PEND), 2, 6, 1, "Pagos Pendientes");

    // 7. Refrescar todos los cálculos y colores
    actualizarTodo();

    ui.alert(
      "✅ Redondeo y formato aplicados en todo el sistema",
      "• " + log.join("\n• "),
      ui.ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
  }
}

// Re-suma Quincena_1, Quincena_2 y Total_Mes directamente desde Recepciones para cada participante
// Pasar silencioso=true para llamarlo desde repararRedondeoMontos sin segundo alert
function repararTotalesPagos(silencioso) {
  try {
    const ss   = SpreadsheetApp.getActive();
    const tz   = Session.getScriptTimeZone();
    const fmtQ = '"Q "#,##0.00';

    // Mapa quincena_nombre → {ordenQ:"Q1"/"Q2", mesKey:"June 2026"}
    const shP = ss.getSheetByName(SHEET_PERIODOS);
    if (!shP) return "PERIODOS no encontrado";
    const mapPeriodo = {};
    shP.getDataRange().getValues().slice(1).forEach(r => {
      const nombre = String(r[1] || "").trim();
      const fin    = r[3] instanceof Date ? r[3] : new Date(r[3]);
      const orden  = String(r[5] || "").trim();
      if (!nombre) return;
      mapPeriodo[nombre] = { orden, mesKey: Utilities.formatDate(fin, tz, "MMMM yyyy") };
    });

    // Suma por participante+quincena desde Recepciones (valores ya redondeados individualmente)
    const shRec = ss.getSheetByName(SHEET_RECEPCIONES);
    if (!shRec || shRec.getLastRow() < DATA_START_ROW) return "Recepciones vacía";
    const mRec   = _headerMap_(shRec);
    const colPar = mRec["participante"];
    const colQ   = mRec["quincena"];
    const colTAP = mRec["total a pagar"] || mRec["total q"];
    const dataRec = shRec.getRange(DATA_START_ROW, 1, shRec.getLastRow() - DATA_START_ROW + 1, shRec.getLastColumn()).getValues();

    // recSumas["nombre|||quincenaNombre"] = suma de TAP redondeada individualmente
    const recSumas = {};
    dataRec.forEach(r => {
      const nombre  = String(r[colPar - 1] || "").trim();
      const qNombre = String(r[colQ   - 1] || "").trim();
      if (!nombre || !qNombre) return;
      const tap = Math.round(Number(r[colTAP - 1]) || 0);
      const key = nombre + "|||" + qNombre;
      recSumas[key] = (recSumas[key] || 0) + tap;
    });

    let totalFilas = 0;
    [SHEET_CHEQUES, SHEET_TRANSFERENCIAS, SHEET_HISTORIAL_PAGOS].forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 3) return;
      const nRows = sh.getLastRow() - 2;
      const data  = sh.getRange(3, 1, nRows, 11).getValues();

      for (let i = 0; i < data.length; i++) {
        const nombre = String(data[i][0] || "").trim();
        const mesKey = String(data[i][9] || "").trim(); // col J = Mes
        if (!nombre || !mesKey) continue;

        // Encontrar nombres de quincena Q1 y Q2 para este mes
        let q1Nombre = "", q2Nombre = "";
        Object.entries(mapPeriodo).forEach(([nQ, info]) => {
          if (info.mesKey === mesKey) {
            if (info.orden === "Q1") q1Nombre = nQ;
            else if (info.orden === "Q2") q2Nombre = nQ;
          }
        });

        const q1 = q1Nombre ? (recSumas[nombre + "|||" + q1Nombre] || 0) : (Number(data[i][6]) || 0);
        const q2 = q2Nombre ? (recSumas[nombre + "|||" + q2Nombre] || 0) : (Number(data[i][7]) || 0);
        const total = q1 + q2;

        sh.getRange(i + 3, 7, 1, 3).setValues([[q1, q2, total]]).setNumberFormat(fmtQ);
        totalFilas++;
      }
    });

    const msg = "Cheques/Transferencias/Historial: " + totalFilas + " fila(s) resincronizadas desde Recepciones";
    if (!silencioso) {
      SpreadsheetApp.getUi().alert("✅ Totales resincronizados", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return msg;
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 4);
    return "Error: " + e.message;
  }
}

/ Aplica la nueva estructura a Cheques y Transferencias (seguro: solo si están vacías o con OK del usuario)
function actualizarHojasNuevas() {
  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const hojas = [
      { nombre: SHEET_CHEQUES,        fn: _setupCheques_,        minRow: 3 },
      { nombre: SHEET_TRANSFERENCIAS, fn: _setupTransferencias_, minRow: 3 },
      { nombre: SHEET_HISTORIAL_PAGOS,fn: _setupHistorialPagos_, minRow: 3 }
    ];

    const conDatos = hojas.filter(h => {
      const sh = ss.getSheetByName(h.nombre);
      return sh && sh.getLastRow() >= h.minRow;
    });

    if (conDatos.length > 0) {
      const aviso = ui.alert(
        "⚠️ Hojas con datos",
        "Estas hojas tienen datos que se borrarán al actualizar la estructura:\n\n• " +
        conDatos.map(h => h.nombre).join("\n• ") +
        "\n\n¿Continuar?",
        ui.ButtonSet.YES_NO
      );
      if (aviso !== ui.Button.YES) return;
    }

    hojas.forEach(h => {
      const sh = ss.getSheetByName(h.nombre) || ss.insertSheet(h.nombre);
      h.fn(sh);
    });

    ss.toast("✅ Cheques, Transferencias e Historial_Pagos actualizados.", null, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

function _crearEstructura_(recrear) {
  const ss = SpreadsheetApp.getActive();

  // Programas_Participacion ya no se crea — sus columnas están en Participantes Activos
  const hojasSistema = [
    SHEET_RECEPCIONES, SHEET_CATALOGOS, SHEET_INACTIVOS,
    SHEET_RESUMEN_PART, SHEET_HIST_QUINCENAS,
    SHEET_DASHBOARD, SHEET_REPORTES,
    SHEET_PERIODOS, SHEET_CHEQUES, SHEET_TRANSFERENCIAS,
    SHEET_HISTORIAL_PAGOS, SHEET_CATALOGO_PROD
  ];

  if (recrear) {
    // Reinstalar completo: borra solo si el usuario lo pidió explícitamente
    hojasSistema.forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    });
  }

  // Solo llama setup en hojas recién creadas — las existentes conservan sus datos
  function setup(name, fn) {
    const existia = !!ss.getSheetByName(name);
    const sh = existia ? ss.getSheetByName(name) : ss.insertSheet(name);
    if (!existia || recrear) fn(sh);
    return sh;
  }

  setup(SHEET_RECEPCIONES,     _setupRecepciones_);
  setup(SHEET_CATALOGOS,       _setupCatalogos_);
  setup(SHEET_INACTIVOS,       _setupInactivos_);
  setup(SHEET_RESUMEN_PART,    sh => _setupResumen_(sh, "RESUMEN POR PARTICIPANTE"));
  setup(SHEET_HIST_QUINCENAS,  _setupHistorialQuincenas_);
  setup(SHEET_DASHBOARD,       _setupDashboard_);
  setup(SHEET_REPORTES,        _setupReportes_);
  setup(SHEET_PERIODOS,        _setupPeriodos_);
  setup(SHEET_CHEQUES,         _setupCheques_);
  setup(SHEET_TRANSFERENCIAS,  _setupTransferencias_);
  setup(SHEET_HISTORIAL_PAGOS, _setupHistorialPagos_);
  setup(SHEET_CATALOGO_PROD,   _setupCatalogo_);

  crearValidacionesDatos();
  crearQuincenaInicial();
  actualizarTodo();
}


// ========================= SETUP DE HOJAS =========================

function _setupRecepciones_(sh) {
  sh.clear();
  sh.clearFormats();

  const nCols    = RECEPCIONES_HEADERS.length;
  const lastCol  = _colLetter_(nCols);

  sh.getRange(1, 1, 1, nCols).merge().setValue("📦 REGISTRO DE RECEPCIONES")
    .setFontWeight("bold").setFontSize(14).setBackground("#263238").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 28);

  sh.getRange(HEADER_ROW, 1, 1, nCols)
    .setValues([RECEPCIONES_HEADERS])
    .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");

  sh.setFrozenRows(HEADER_ROW);
  sh.setRowHeight(HEADER_ROW, 30);

  // #|Fecha|Quincena|Participante|CreamosID|Proyecto|Categoría|Producto|UBuenas|URechaz|PrecioU|TotalQ|ImpPC|TotalPagar|Estado
  const widths = [45, 90, 110, 150, 100, 160, 130, 120, 100, 120, 90, 90, 95, 95, 90];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange(`A${HEADER_ROW + 1}:${lastCol}1000`).setBackground("#ffffff").setFontColor("#212121");
  sh.getRange(`A${HEADER_ROW + 1}:${lastCol}1000`).setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(`B${HEADER_ROW + 1}:B1000`).setNumberFormat("dd/MM/yyyy");       // Fecha entrega
  sh.getRange(`K${HEADER_ROW + 1}:N1000`).setNumberFormat('"Q "#,##0.00');     // Precio, Total Q, Impuesto, Total a Pagar

  // Dropdown Categoría (col G = 7)
  sh.getRange(DATA_START_ROW, 7, 996).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CAT_OPCIONES)
      .setAllowInvalid(true).build()
  );
}

function _colLetter_(col) {
  let r = '';
  while (col > 0) { const m = (col - 1) % 26; r = String.fromCharCode(65 + m) + r; col = Math.floor((col - 1) / 26); }
  return r;
}

function _setupCatalogos_(sh) {
  sh.clear();
  sh.clearFormats();

  // A–H: datos personales/bancarios
  sh.getRange("A1:H1").setValues([["Creamos ID", "Nombre Completo", "Etapa", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI"]])
    .setBackground("#263238").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // I–K: programas sociales
  sh.getRange("I1:K1").setValues([["Educación", "Inclusión Laboral", "Apoyo Emocional"]])
    .setBackground("#2e7d32").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // L–N: datos sociales
  sh.getRange("L1").setValue("N° Hijos")
    .setBackground("#4a148c").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("M1:N1").setValues([["CCI", "CUNDE"]])
    .setBackground("#4a148c").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // O: fuente de pago
  sh.getRange("O1").setValue("Cuenta de Pago")
    .setBackground("#e65100").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 28);

  sh.setColumnWidth(1,  100); sh.setColumnWidth(2,  180); sh.setColumnWidth(3,  100);
  sh.setColumnWidth(4,  130); sh.setColumnWidth(5,  130); sh.setColumnWidth(6,  120);
  sh.setColumnWidth(7,  150); sh.setColumnWidth(8,  110);
  sh.setColumnWidth(9,   90); sh.setColumnWidth(10, 140); sh.setColumnWidth(11, 140);
  sh.setColumnWidth(12,  80); sh.setColumnWidth(13,  60); sh.setColumnWidth(14,  70);
  sh.setColumnWidth(15, 120);

  sh.getRange("A2:H1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:H1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("I2:K1000").insertCheckboxes();
  sh.getRange("I2:K1000").setBackground("#f1f8e9");
  sh.getRange("I2:K1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("L2:L1000").setBackground("#fce4ec").setFontColor("#212121");
  sh.getRange("L2:L1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("M2:N1000").insertCheckboxes();
  sh.getRange("M2:N1000").setBackground("#fce4ec");
  sh.getRange("M2:N1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("C2:C1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Inscritx", "Retiradx", "Empleadx", "Ciclo de Vida Terminado"])
      .setAllowInvalid(false).build()
  );
  sh.getRange("E2:E1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Monetaria", "Ahorro", "Corriente"])
      .setAllowInvalid(true).build()
  );

  // O: Cuenta de Pago dropdown
  sh.getRange("O2:O1000")
    .setBackground("#fff3e0").setFontColor("#212121")
    .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Creamos", "mi-eelo"])
        .setAllowInvalid(false).build()
    );
}

function _setupInactivos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:H1").setValues([["Creamos ID", "Nombre Completo", "Etapa", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI"]])
    .setBackground("#5d4037").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 130); sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150); sh.setColumnWidth(8, 110);

  sh.getRange("A2:H1000").setBackground("#fafafa").setFontColor("#212121");
  sh.getRange("A2:H1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupResumen_(sh, titulo) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1").setValue(titulo).setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");
}

function _setupPagosPendientes_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:K1").setValues([[
    "Participante", "Creamos ID", "Fecha entrega", "Producto",
    "Unidades buenas", "Total Q", "Días pendiente",
    "Método Pago", "Nº Cuenta", "Tipo Cuenta", "Ejecución Pago"
  ]])
    .setBackground("#c62828").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 100);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 130);
  sh.setColumnWidth(10, 120);
  sh.setColumnWidth(11, 140);

  sh.getRange("A2:K1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:K1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupDashboard_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:D1").merge().setValue("📊 DASHBOARD")
    .setFontWeight("bold").setFontSize(14).setBackground("#1565c0").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 28);

  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 150);
  sh.setColumnWidth(4, 120);
}

function _setupReportes_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:L1").merge().setValue("REPORTES PARA POWER BI").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white").setHorizontalAlignment("center");

  sh.getRange("A2:L2").setValues([[
    "Fecha", "Mes", "Creamos ID", "Participante", "Producto", "Proyecto",
    "Unidades Buenas", "Unidades Rechazadas", "Tasa Rechazo %",
    "Total a Pagar", "Estado Pago", "Días Pendiente"
  ]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(2);
  sh.setRowHeight(2, 25);

  sh.setColumnWidth(1, 100);  sh.setColumnWidth(2,  90);  sh.setColumnWidth(3,  110);
  sh.setColumnWidth(4, 160);  sh.setColumnWidth(5, 120);  sh.setColumnWidth(6,  130);
  sh.setColumnWidth(7, 110);  sh.setColumnWidth(8, 130);  sh.setColumnWidth(9,  110);
  sh.setColumnWidth(10, 110); sh.setColumnWidth(11, 150); sh.setColumnWidth(12, 110);
}

// ── Mejora 1: PERIODOS ────────────────────────────────────────────────────────
function _setupPeriodos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:F1").setValues([["#", "Nombre_Quincena", "Fecha_Inicio", "Fecha_Fin", "Estado", "Notas"]])
    .setBackground("#1565c0").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 50);  sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 100); sh.setColumnWidth(6, 220);

  sh.getRange("C2:D1000").setNumberFormat("yyyy-mm-dd");
  sh.getRange("E2:E1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Activo", "Cerrado"])
      .setAllowInvalid(false)
      .build()
  );
  sh.getRange("A2:F1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:F1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

// ── Mejora 1: Historial Quincenas (reemplaza Resumen Mensual) ─────────────────
function _setupHistorialQuincenas_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1").setValue("HISTORIAL DE QUINCENAS").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");

  sh.getRange("A2:J2").setValues([[
    "Quincena", "Fecha_Inicio", "Fecha_Fin", "Entregas", "Unidades_Buenas",
    "Unidades_Rechazadas", "Total_Q", "Participantes_Con_Entrega",
    "Participantes_Sin_Entrega", "Notas"
  ]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(2);
  sh.setRowHeight(2, 25);

  sh.setColumnWidth(1, 210); sh.setColumnWidth(2, 110);  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 90);  sh.setColumnWidth(5, 130);  sh.setColumnWidth(6, 150);
  sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 180);  sh.setColumnWidth(9, 180);
  sh.setColumnWidth(10, 220);
}

// ── Mejora 2: CHEQUES ─────────────────────────────────────────────────────────
function _setupPagosHoja_(sh, titulo, color) {
  sh.clear();
  sh.clearFormats();

  const COLS = 11;
  sh.getRange(1, 1, 1, COLS).merge()
    .setValue(titulo).setFontWeight("bold").setFontSize(13)
    .setBackground(color).setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 32);

  sh.getRange(2, 1, 1, COLS).setValues([[
    "Nombre", "Creamos ID", "Banco", "Tipo Cuenta", "Nº Cuenta", "Titular",
    "Quincena_1", "Quincena_2", "Total_Mes", "Mes", "Año"
  ]])
    .setFontWeight("bold").setBackground(color).setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 28);
  sh.setFrozenRows(2);

  sh.getRange("G3:I1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("E3:E1000").setNumberFormat("@");   // Nº Cuenta como texto (evita notación científica)
  sh.getRange("A3:K1000").setBackground("#fafafa").setFontColor("#212121");
  sh.getRange("A3:K1000").setBorder(true, true, true, true, false, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  const widths = [160, 100, 130, 110, 130, 150, 100, 100, 100, 110, 70];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function _setupCheques_(sh) {
  _setupPagosHoja_(sh, "PAGOS POR CHEQUE", "#4a148c");
}

function _setupTransferencias_(sh) {
  _setupPagosHoja_(sh, "PAGOS POR TRANSFERENCIA", "#006064");
}

// ── Mejora 3: PROGRAMAS_PARTICIPACION ────────────────────────────────────────
function _setupProgramas_(sh) {
  sh.clear();
  sh.clearFormats();

  // Programas sociales: Manufactura es el sistema completo, NO un programa aquí
  sh.getRange("A1:G1").setValues([[
    "Creamos_ID", "Participante", "mi_eelo", "Educación", "Inclusión_Laboral", "Apoyo_Emocional", "Notas"
  ]])
    .setBackground("#2e7d32").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 190); sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 100); sh.setColumnWidth(5, 140); sh.setColumnWidth(6, 140);
  sh.setColumnWidth(7, 240);

  sh.getRange("C2:F1000").insertCheckboxes();
  sh.getRange("A2:G1000").setBackground("#f1f8e9").setFontColor("#212121");
  sh.getRange("A2:G1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupHistorialPagos_(sh) {
  sh.clear();
  sh.clearFormats();

  const COLS = 12;
  sh.getRange(1, 1, 1, COLS).merge()
    .setValue("HISTORIAL DE PAGOS").setFontWeight("bold").setFontSize(13)
    .setBackground("#1b5e20").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 32);

  sh.getRange(2, 1, 1, COLS).setValues([[
    "Nombre", "Creamos ID", "Banco", "Tipo Cuenta", "Nº Cuenta", "Titular",
    "Quincena_1", "Quincena_2", "Total_Mes", "Mes", "Año", "Método"
  ]])
    .setFontWeight("bold").setBackground("#2e7d32").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 28);
  sh.setFrozenRows(2);

  sh.getRange("G3:I1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("E3:E1000").setNumberFormat("@"); // Nº Cuenta como texto
  sh.getRange("A3:L1000").setBackground("#f1f8e9").setFontColor("#212121");
  sh.getRange("A3:L1000").setBorder(true, true, true, true, false, true, "#c8e6c9", SpreadsheetApp.BorderStyle.SOLID);

  const widths = [160, 100, 130, 110, 130, 150, 100, 100, 100, 110, 60, 110];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}


// ── Catálogo de Productos ─────────────────────────────────────────────────────
function _setupCatalogo_(sh) {
  sh.clear();
  sh.clearFormats();

  // Col A = Categoría, B = Diseño/Producto, C = Precio, D = Notas
  sh.getRange("A1:D1").setValues([["Categoría", "Diseño / Producto", "Precio (Q)", "Notas"]])
    .setBackground("#1565c0").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 260);

  sh.getRange("C2:C1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("A2:D1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:D1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

// ── Archivo de Recepciones (quincenas cerradas) ───────────────────────────────
function _setupArchivo_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange(1, 1, 1, RECEPCIONES_HEADERS.length)
    .setValues([RECEPCIONES_HEADERS])
    .setBackground("#546e7a").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  const widths = [45, 90, 110, 150, 100, 160, 120, 100, 120, 90, 80, 80, 80, 100, 120, 150];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange("A2:P2000").setBackground("#fafafa").setFontColor("#212121");
  sh.getRange("A2:P2000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}


// ========================= CAPTURA RAPIDA =========================
function agregarEntregaRapida() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const cat = ss.getSheetByName(SHEET_CATALOGOS);

    // Solo participantes con Etapa = Inscritx
    const participantesData = [];
    if (cat) {
      const dataCat = cat.getDataRange().getValues();
      for (let i = 1; i < dataCat.length; i++) {
        const nombre = String(dataCat[i][CAT_COL.NOMBRE] || "").trim();
        const etapa  = String(dataCat[i][CAT_COL.ETAPA]  || "").trim();
        const id     = String(dataCat[i][CAT_COL.ID]      || "").trim();
        if (nombre && etapa === "Inscritx") participantesData.push({ nombre, id });
      }
    }
    const proyectos     = _obtenerHistorico_("proyectos");
    const metodos       = _obtenerHistorico_("metodos");

    // Leer productos, precios y categorías del Catálogo_Productos
    // Col A = Categoría, Col B = Diseño/Producto, Col C = Precio
    const productosConPrecio = [];
    const catProd = ss.getSheetByName(SHEET_CATALOGO_PROD);
    if (catProd && catProd.getLastRow() > 1) {
      catProd.getRange(2, 1, catProd.getLastRow() - 1, 3).getValues().forEach(r => {
        const categoria = String(r[0] || "").trim();
        const prod      = String(r[1] || "").trim();
        if (!prod) return;
        const raw    = r[2];
        const precio = typeof raw === 'number'
          ? raw
          : parseFloat(String(raw).replace(/[^0-9.]/g, '')) || 0;
        productosConPrecio.push({ nombre: prod, precio, categoria });
      });
    }

    const productosHistorico = _obtenerHistorico_("productos");

    // Obtener quincena activa para mostrar en el form
    const shPer = ss.getSheetByName(SHEET_PERIODOS);
    let quincenaActiva = "Sin quincena activa";
    if (shPer && shPer.getLastRow() > 1) {
      const dPer = shPer.getDataRange().getValues();
      for (let i = 1; i < dPer.length; i++) {
        if (String(dPer[i][4] || "").trim() === "Activo") {
          quincenaActiva = String(dPer[i][1] || "").trim();
          break;
        }
      }
    }

    const tz       = Session.getScriptTimeZone();
    const fechaHoy = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
          font-family: 'Segoe UI', sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          padding: 16px; min-height: 100vh;
        }
        .container {
          max-width: 460px; margin: 0 auto; background: white;
          border-radius: 10px; box-shadow: 0 8px 30px rgba(0,0,0,0.2); padding: 22px;
        }
        .header { margin-bottom: 16px; }
        .header h1 { color: #333; font-size: 20px; margin-bottom: 4px; }
        .badge {
          display: inline-block; background: #ede7f6; color: #5e35b1;
          font-size: 11px; font-weight: 600; padding: 3px 10px;
          border-radius: 20px; margin-bottom: 4px;
        }
        .subtitle { color: #999; font-size: 12px; }
        .form-group { margin-bottom: 12px; }
        label {
          display: block; color: #555; font-weight: 600; margin-bottom: 4px;
          font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;
        }
        select, input {
          width: 100%; padding: 8px 10px; border: 1.5px solid #e0e0e0;
          border-radius: 6px; font-size: 14px; transition: border-color 0.2s; font-family: inherit;
        }
        select:focus, input:focus { outline: none; border-color: #667eea; }
        select { cursor: pointer; background: white; }
        .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        .resumen {
          background: #f3f0ff; border-left: 3px solid #667eea;
          padding: 10px 14px; border-radius: 6px; margin: 14px 0 0 0; font-size: 13px;
        }
        .resumen-row { display: flex; justify-content: space-between; padding: 2px 0; }
        .resumen-row.total {
          border-top: 1px solid #d1c4e9; margin-top: 6px; padding-top: 6px;
          font-weight: 700; color: #5e35b1; font-size: 14px;
        }
        .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 16px; }
        button {
          padding: 10px; border: none; border-radius: 6px; font-size: 13px;
          font-weight: 600; cursor: pointer; transition: all 0.2s;
        }
        .btn-guardar { background: linear-gradient(135deg, #667eea, #764ba2); color: white; }
        .btn-guardar:hover { opacity: 0.9; transform: translateY(-1px); }
        .btn-cancelar { background: #f0f0f0; color: #666; }
        .btn-cancelar:hover { background: #e0e0e0; }
        .error { color: #e53935; font-size: 11px; margin-top: 2px; display: none; }
        .loading { display: none; text-align: center; color: #667eea; font-size: 13px; margin-top: 10px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📦 Nueva Entrega</h1>
          <span class="badge">📅 ${quincenaActiva}</span><br>
          <span class="subtitle">${productosConPrecio.length} producto(s) en catálogo</span>
        </div>

        <div class="form-group">
          <label>Fecha de entrega</label>
          <input type="date" id="fecha" value="${fechaHoy}">
        </div>

        <div class="form-group">
          <label>Participante *</label>
          <select id="participante" required>
            <option value="">— Seleccionar —</option>
            ${participantesData.map(p => '<option value="'+p.nombre+'">'+(p.id ? p.nombre+' – '+p.id : p.nombre)+'</option>').join('')}
          </select>
          <div class="error" id="err-part">Requerido</div>
        </div>

        <div class="form-group">
          <label>Producto *</label>
          <input type="text" id="producto" list="productos-list" placeholder="Escribe o selecciona..." required>
          <datalist id="productos-list">
            ${productosConPrecio.map(p => '<option value="'+p.nombre+'">').join('')}
            ${productosHistorico.map(p => '<option value="'+p+'">').join('')}
          </datalist>
          <div class="error" id="err-prod">Requerido</div>
        </div>

        <div class="form-group">
          <label>Categoría *</label>
          <select id="categoria">
            <option value="">— Se llena al elegir producto —</option>
            ${CAT_OPCIONES.map(c => '<option value="'+c+'">'+c+'</option>').join('')}
          </select>
        </div>

        <div class="form-group">
          <label>Proyecto / Cliente *</label>
          <input type="text" id="proyecto" list="proyectos-list" placeholder="Creamos, Cliente X..." required>
          <datalist id="proyectos-list">
            ${proyectos.map(p => '<option value="'+p+'">').join('')}
          </datalist>
          <div class="error" id="err-proy">Requerido</div>
        </div>

        <div class="row">
          <div class="form-group">
            <label>Unidades Buenas *</label>
            <input type="number" id="buenas" min="0" value="0" required>
            <div class="error" id="err-buenas">Debe ser &gt; 0</div>
          </div>
          <div class="form-group">
            <label>Unidades Rechazadas</label>
            <input type="number" id="rechazadas" min="0" value="0">
          </div>
        </div>

        <div class="form-group">
          <label>Precio Unitario (Q) *</label>
          <input type="number" id="precio" min="0" step="0.01" value="0" required>
          <div class="error" id="err-precio">Debe ser &gt; 0</div>
        </div>

        <div class="resumen" id="resumen" style="display:none;">
          <div class="resumen-row"><span>Unidades buenas:</span><span id="res-buenas">0</span></div>
          <div class="resumen-row"><span>Precio unitario:</span><span>Q <span id="res-precio">0.00</span></span></div>
          <div class="resumen-row"><span>Total bruto:</span><span>Q <span id="res-bruto">0.00</span></span></div>
          <div class="resumen-row" style="color:#1565c0;"><span>+ Imp. Pequeño Contribuyente (5%):</span><span>+ Q <span id="res-impuesto">0.00</span></span></div>
          <div class="resumen-row total"><span>Total a pagar:</span><span>Q <span id="res-total">0.00</span></span></div>
        </div>

        <div id="toast" style="display:none; background:#e8f5e9; border-left:3px solid #43a047;
          padding:10px 14px; border-radius:6px; margin-top:12px; font-size:13px; color:#2e7d32;
          font-weight:600;"></div>

        <div class="buttons">
          <button class="btn-guardar" id="btn-guardar" onclick="guardarEntrega()">✅ Guardar</button>
          <button class="btn-cancelar" onclick="google.script.host.close()">✕ Cerrar</button>
        </div>
        <div class="loading" id="loading">⏳ Guardando...</div>
      </div>

      <script>
        const productosConPrecio = ${JSON.stringify(productosConPrecio)};

        ['buenas','precio'].forEach(id =>
          document.getElementById(id).addEventListener('input', actualizarResumen)
        );
        document.getElementById('producto').addEventListener('change', actualizarPrecio);
        document.getElementById('producto').addEventListener('input',  actualizarPrecio);

        function actualizarPrecio() {
          setTimeout(function() {
            const val   = document.getElementById('producto').value.trim().toLowerCase();
            const campo = document.getElementById('precio');
            if (!val) return;
            const match = productosConPrecio.find(p => p.nombre.trim().toLowerCase() === val);
            if (match) {
              if (match.precio > 0) { campo.value = match.precio; actualizarResumen(); }
              if (match.categoria) {
                const sel = document.getElementById('categoria');
                sel.value = match.categoria;
                if (sel.value !== match.categoria) {
                  // valor no está en la lista, agregarlo temporalmente
                  const opt = document.createElement('option');
                  opt.value = match.categoria; opt.textContent = match.categoria;
                  sel.appendChild(opt); sel.value = match.categoria;
                }
              }
            }
          }, 80);
        }

        function actualizarResumen() {
          const b       = Number(document.getElementById('buenas').value) || 0;
          const p       = Number(document.getElementById('precio').value) || 0;
          const bruto   = b * p;
          const imp     = Math.round(bruto * 0.05);
          const total   = Math.round(bruto + imp);  // programa paga bruto + impuesto
          document.getElementById('res-buenas').textContent   = b;
          document.getElementById('res-precio').textContent   = p.toFixed(2);
          document.getElementById('res-bruto').textContent    = bruto.toFixed(2);
          document.getElementById('res-impuesto').textContent = imp.toFixed(2);
          document.getElementById('res-total').textContent    = total.toFixed(2);
          document.getElementById('resumen').style.display    = (b > 0 || p > 0) ? 'block' : 'none';
        }

        function validar() {
          const part = document.getElementById('participante').value.trim();
          const prod = document.getElementById('producto').value.trim();
          const proy = document.getElementById('proyecto').value.trim();
          const b    = Number(document.getElementById('buenas').value) || 0;
          const p    = Number(document.getElementById('precio').value) || 0;
          document.getElementById('err-part').style.display   = !part     ? 'block' : 'none';
          document.getElementById('err-prod').style.display   = !prod     ? 'block' : 'none';
          document.getElementById('err-proy').style.display   = !proy     ? 'block' : 'none';
          document.getElementById('err-buenas').style.display = (b <= 0)  ? 'block' : 'none';
          document.getElementById('err-precio').style.display = (p <= 0)  ? 'block' : 'none';
          return part && prod && proy && b > 0 && p > 0;
        }

        let contadorRegistros = 0;

        function mostrarToast(msg) {
          const t = document.getElementById('toast');
          t.textContent = msg;
          t.style.display = 'block';
          setTimeout(function() { t.style.display = 'none'; }, 3000);
        }

        function limpiarParaSiguiente(buenas, precio, total) {
          contadorRegistros++;
          mostrarToast('✅ Registro #' + contadorRegistros + ' guardado — Q ' + total.toFixed(2));
          // Conserva participante y proyecto; limpia producto, unidades y precio
          document.getElementById('producto').value    = '';
          document.getElementById('buenas').value      = '0';
          document.getElementById('rechazadas').value  = '0';
          document.getElementById('precio').value      = '0';
          document.getElementById('resumen').style.display = 'none';
          document.getElementById('producto').focus();
        }

        function guardarEntrega() {
          if (!validar()) return;
          const btn   = document.getElementById('btn-guardar');
          const b     = Number(document.getElementById('buenas').value) || 0;
          const p     = Number(document.getElementById('precio').value) || 0;
          const total = b * p * 1.05;           // bruto + impuesto 5%
          btn.disabled = true; btn.style.opacity = '0.6';
          document.getElementById('loading').style.display = 'block';
          google.script.run
            .withSuccessHandler(function() {
              btn.disabled = false; btn.style.opacity = '1';
              document.getElementById('loading').style.display = 'none';
              limpiarParaSiguiente(b, p, total);
            })
            .withFailureHandler(function(err) {
              alert('Error: ' + err);
              btn.disabled = false; btn.style.opacity = '1';
              document.getElementById('loading').style.display = 'none';
            })
            .guardarEntregaServer(
              document.getElementById('participante').value.trim(),
              document.getElementById('producto').value.trim(),
              document.getElementById('proyecto').value.trim(),
              Number(document.getElementById('buenas').value),
              Number(document.getElementById('rechazadas').value),
              Number(document.getElementById('precio').value),
              "",
              document.getElementById('fecha').value,
              document.getElementById('categoria').value.trim()
            );
        }
      </script>
    </body>
    </html>
  `)
    .setWidth(500)
    .setHeight(700);

    SpreadsheetApp.getUi().showModalDialog(html, "Nueva Entrega Rápida");
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

function guardarEntregaServer(participante, producto, proyecto, buenas, rechazadas, precio, metodo, fechaStr, categoria) {
  try {
    const ss  = SpreadsheetApp.getActive();
    const sh  = ss.getSheetByName(SHEET_RECEPCIONES);
    const cat = ss.getSheetByName(SHEET_CATALOGOS);
    const m   = _headerMap_(sh);
    const row = Math.max(sh.getLastRow() + 1, DATA_START_ROW);

    const total      = buenas * precio;         // subtotal bruto
    const impuesto   = Math.round(total * 0.05);          // 5% pequeño contribuyente
    const totalPagar = Math.round(total + impuesto);      // lo que el programa paga
    const quincena = _obtenerQuincenaActiva_();

    // Fecha desde el formulario (YYYY-MM-DD) o hoy si no viene
    const tz    = Session.getScriptTimeZone();
    const fecha = fechaStr
      ? new Date(fechaStr + "T12:00:00")
      : new Date();

    let creamosID = "";
    if (cat) {
      const dataCat = cat.getDataRange().getValues();
      for (let i = 1; i < dataCat.length; i++) {
        if (String(dataCat[i][CAT_COL.NOMBRE] || "").trim() === participante) {
          creamosID = String(dataCat[i][CAT_COL.ID] || "").trim();
          break;
        }
      }
    }

    sh.getRange(row, m["#"]).setValue(row - DATA_START_ROW + 1);
    const celdaFecha = sh.getRange(row, m["fecha entrega"]);
    celdaFecha.setValue(fecha).setNumberFormat("dd/MM/yyyy");
    sh.getRange(row, m["quincena"]).setValue(quincena);
    sh.getRange(row, m["participante"]).setValue(participante || "");
    sh.getRange(row, m["creamos id"]).setValue(creamosID);
    sh.getRange(row, m["proyecto / cliente"]).setValue(proyecto || "");
    if (m["categoría"]) sh.getRange(row, m["categoría"]).setValue(categoria || "");
    sh.getRange(row, m["producto"]).setValue(producto || "");
    sh.getRange(row, m["unidades buenas"]).setValue(buenas || 0);
    sh.getRange(row, m["unidades rechazadas"]).setValue(rechazadas || 0);
    sh.getRange(row, m["precio unit. (q)"]).setValue(precio || 0).setNumberFormat('"Q "#,##0.00');
    sh.getRange(row, m["total q"]).setValue(total || 0).setNumberFormat('"Q "#,##0.00');
    if (m["impuesto pc (5%)"]) sh.getRange(row, m["impuesto pc (5%)"]).setValue(impuesto).setNumberFormat('"Q "#,##0.00');
    if (m["total a pagar"])    sh.getRange(row, m["total a pagar"]).setValue(totalPagar).setNumberFormat('"Q "#,##0.00');
    sh.getRange(row, m["estado pago"]).setValue("Espera cierre quincena");

    _guardarHistorico_("productos", producto);
    _guardarHistorico_("proyectos", proyecto);

    SpreadsheetApp.getActive().toast("✅ Total Q " + totalPagar.toFixed(2), null, 2);
    actualizarTodo();
    return true;
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
    throw e;
  }
}

// ========================= EDICION =========================
function editarFilaSeleccionada() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== SHEET_RECEPCIONES) return SpreadsheetApp.getUi().alert("Selecciona una fila en Recepciones.");
    const row = sh.getActiveRange().getRow();
    if (row < DATA_START_ROW) return SpreadsheetApp.getUi().alert("Selecciona una fila de datos (desde fila 5).");

    const m  = _headerMap_(sh);
    const ui = SpreadsheetApp.getUi();

    const ubActual    = sh.getRange(row, m["unidades buenas"]).getValue();
    const urActual    = sh.getRange(row, m["unidades rechazadas"]).getValue();
    const puActual    = sh.getRange(row, m["precio unit. (q)"]).getValue();
    const notasActual = sh.getRange(row, m["notas / calidad"]).getValue();

    const datos = ui.prompt(
      "Editar entrega",
      `Unidades buenas, rechazadas, precio (ej: ${ubActual}, ${urActual}, ${puActual})`,
      ui.ButtonSet.OK_CANCEL
    );
    if (datos.getSelectedButton() !== ui.Button.OK) return;

    const partes = datos.getResponseText().split(",").map(x => Number(x.trim()) || 0);
    const buenas = partes[0];
    const rechaz = partes[1];
    const precio = partes[2];

    const notas = ui.prompt("Editar entrega", `Notas / calidad (actual: ${notasActual})`, ui.ButtonSet.OK_CANCEL);
    if (notas.getSelectedButton() !== ui.Button.OK) return;

    sh.getRange(row, m["unidades buenas"]).setValue(buenas);
    sh.getRange(row, m["unidades rechazadas"]).setValue(rechaz);
    sh.getRange(row, m["precio unit. (q)"]).setValue(precio);
    sh.getRange(row, m["total q"]).setValue(buenas * precio);
    sh.getRange(row, m["notas / calidad"]).setValue(notas.getResponseText());

    SpreadsheetApp.getActive().toast("✅ Entrega actualizada", "Listo", 3);
    actualizarTodo();
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

// ========================= PAGOS =========================
function marcarFilaPagada() {
  try {
    const sh = SpreadsheetApp.getActiveSheet();
    if (sh.getName() !== SHEET_RECEPCIONES) return SpreadsheetApp.getUi().alert("Selecciona una fila en Recepciones.");
    const row = sh.getActiveRange().getRow();
    if (row < DATA_START_ROW) return SpreadsheetApp.getUi().alert("Selecciona una fila de datos (desde fila 5).");

    const m   = _headerMap_(sh);
    const hoy = new Date();

    sh.getRange(row, m["estado pago"]).setValue("Pagado");

    // Registrar en Historial_Pagos
    const rowData    = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const metodo     = "Transferencia";
    const quincena   = String(rowData[m["quincena"]         - 1] || "").trim();
    const participante = String(rowData[m["participante"]   - 1] || "").trim();
    const creamosID  = String(rowData[m["creamos id"]       - 1] || "").trim();
    const producto   = String(rowData[m["producto"]         - 1] || "").trim();
    const ub         = Number(rowData[m["unidades buenas"]  - 1]) || 0;
    const tq         = Number(rowData[m["total q"]          - 1]) || 0;

    _appendHistorialPago_(participante, { quincena, creamosID, producto, ub, tq, metodo }, hoy);

    actualizarTodo();
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

// ========================= ACTUALIZACION GENERAL =========================
function actualizarTodo() {
  calcularTotalesColumnas();
  actualizarEstadoPagosAutomatico();
  procesarParticipantesInactivos();
  actualizarResumenParticipantes();
  actualizarHistorialQuincenas();
  aplicarColoresAutomaticos();
  actualizarDashboard();
  actualizarReportes();
}

function calcularTotalesColumnas() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RECEPCIONES);
  if (!sh || sh.getLastRow() < DATA_START_ROW) return;
  const m      = _headerMap_(sh);
  const colPart = m["participante"];
  const colUB   = m["unidades buenas"];
  const colPU   = m["precio unit. (q)"];
  const colTQ   = m["total q"];
  const colImp  = m["impuesto pc (5%)"] || m["impuesto pc"];
  const colTAP  = m["total a pagar"];
  const nRows   = sh.getLastRow() - DATA_START_ROW + 1;
  const ncols   = sh.getLastColumn();
  const data    = sh.getRange(DATA_START_ROW, 1, nRows, ncols).getValues();

  for (let i = 0; i < data.length; i++) {
    if (!data[i][colPart - 1]) continue;
    const b   = Number(data[i][colUB  - 1]) || 0;
    const p   = Number(data[i][colPU  - 1]) || 0;
    const tq  = b * p;
    const imp = Math.round(tq * 0.05);
    const tap = Math.round(tq + imp);
    data[i][colTQ  - 1] = tq;
    if (colImp) data[i][colImp - 1] = imp;
    if (colTAP) data[i][colTAP - 1] = tap;
  }
  sh.getRange(DATA_START_ROW, 1, nRows, ncols).setValues(data);

  // Reaplica formato Q a columnas monetarias
  const firstCol = m["precio unit. (q)"] || colTQ;
  if (firstCol) sh.getRange(DATA_START_ROW, firstCol, nRows, 4).setNumberFormat('"Q "#,##0.00');
}

function actualizarEstadoPagosAutomatico() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RECEPCIONES);
  if (!sh || sh.getLastRow() < DATA_START_ROW) return;
  const m    = _headerMap_(sh);
  const colP = m["participante"];
  const colE = m["estado pago"];
  const colT = m["total a pagar"] || m["total q"];
  if (!colP || !colE || !colT) return;

  for (let r = DATA_START_ROW; r <= sh.getLastRow(); r++) {
    const part = sh.getRange(r, colP).getValue();
    if (!part) continue;
    const estado = String(sh.getRange(r, colE).getValue() || "").trim();
    // Solo asigna estado si la celda está vacía o tiene un valor obsoleto
    if (!estado || estado === "Sin monto") {
      const total = Number(sh.getRange(r, colT).getValue()) || 0;
      sh.getRange(r, colE).setValue(total > 0 ? "Espera cierre quincena" : "");
    }
  }
}

// ========================= RESUMENES =========================
function actualizarResumenParticipantes() {
  const ss  = SpreadsheetApp.getActive();
  const dst = ss.getSheetByName(SHEET_RESUMEN_PART);
  if (!dst) return;

  // Acumular desde Archivo_Recepciones (quincenas cerradas) + Recepciones (quincena actual)
  const acc = {}; // { nombre: { cid, b, re, q, pagado, pendiente, quincenas: Set, primera, ultima } }

  function _acumularHoja_(sh, startRow) {
    if (!sh || sh.getLastRow() < startRow) return;
    const m    = _headerMap_(sh);
    const data = sh.getDataRange().getValues();
    for (let i = startRow - 1; i < data.length; i++) {
      const r = data[i];
      const p = String(r[m["participante"] - 1] || "").trim();
      if (!p) continue;
      if (!acc[p]) acc[p] = { cid: "", b: 0, re: 0, q: 0, pagado: 0, pendiente: 0, quincenas: new Set(), primera: null, ultima: null };
      const a = acc[p];
      if (!a.cid) a.cid = String(r[m["creamos id"] - 1] || "").trim();
      a.b  += Number(r[m["unidades buenas"]    - 1]) || 0;
      a.re += Number(r[m["unidades rechazadas"] - 1]) || 0;
      const tap = m["total a pagar"] ? Number(r[m["total a pagar"] - 1]) || 0 : 0;
      const tqb = Number(r[m["total q"] - 1]) || 0;
      const tq  = tap > 0 ? tap : tqb * 1.05;
      a.q  += tq;
      const estado = String(r[m["estado pago"] - 1] || "").trim().toLowerCase();
      if (estado === "quincena cerrada")          a.pagado    += tq;
      if (estado === "espera cierre quincena")    a.pendiente += tq;
      // compatibilidad con valores anteriores
      if (estado === "pagado")    a.pagado    += tq;
      if (estado === "pendiente") a.pendiente += tq;
      const q = String(r[m["quincena"] - 1] || "").trim();
      if (q) a.quincenas.add(q);
      const fe = r[m["fecha entrega"] - 1];
      const fd = fe ? new Date(fe) : null;
      if (fd && !isNaN(fd)) {
        if (!a.primera || fd < a.primera) a.primera = fd;
        if (!a.ultima  || fd > a.ultima)  a.ultima  = fd;
      }
    }
  }

  _acumularHoja_(ss.getSheetByName(SHEET_RECEPCIONES), DATA_START_ROW);

  // Construir tabla
  dst.clear();
  dst.clearFormats();

  const fmt = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy") : "";

  const titulo = "RESUMEN HISTÓRICO POR PARTICIPANTE";
  dst.getRange("A1:J1").merge()
    .setValue(titulo).setFontWeight("bold").setFontSize(13)
    .setBackground("#1a237e").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dst.setRowHeight(1, 32);

  const headers = [
    "Participante", "Creamos ID",
    "Unidades buenas", "Unidades rechazadas", "Tasa rechazo %",
    "Total Q generado", "Total pagado", "Total pendiente",
    "Quincenas activas", "Primera entrega", "Última entrega"
  ];
  dst.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#283593").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dst.setRowHeight(2, 28);
  dst.setFrozenRows(2);

  const rows = Object.keys(acc).sort().map(k => {
    const a = acc[k];
    const tot = a.b + a.re;
    return [
      k, a.cid,
      a.b, a.re,
      tot ? a.re / tot : 0,
      a.q, a.pagado, a.pendiente,
      a.quincenas.size,
      fmt(a.primera), fmt(a.ultima)
    ];
  });

  if (rows.length) {
    const rng = dst.getRange(3, 1, rows.length, headers.length);
    rng.setValues(rows)
      .setBackground("#ffffff").setFontColor("#212121")
      .setBorder(true, true, true, true, false, true, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    // Formatos numéricos
    dst.getRange(3, 5, rows.length, 1).setNumberFormat("0.00%");
    dst.getRange(3, 6, rows.length, 3).setNumberFormat('"Q " #,##0.00');

    // Filas alternadas
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1)
        dst.getRange(3 + i, 1, 1, headers.length).setBackground("#e8eaf6");
    }

    // Resaltar pendientes > 0 en rojo suave
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][7] > 0)
        dst.getRange(3 + i, 8, 1, 1).setBackground("#ffcdd2").setFontColor("#b71c1c");
    }
  }

  const widths = [160, 100, 110, 120, 100, 110, 110, 110, 110, 110, 110];
  widths.forEach((w, i) => dst.setColumnWidth(i + 1, w));

  dst.getRange("A1").offset(rows.length + 2, 0)
    .setValue("Actualizado: " + fmt(new Date()))
    .setFontColor("#9e9e9e").setFontStyle("italic").setFontSize(9);
}

function actualizarPagosPendientes() {
  const ss      = SpreadsheetApp.getActive();
  const src     = ss.getSheetByName(SHEET_RECEPCIONES);
  const catAct  = ss.getSheetByName(SHEET_CATALOGOS);
  const dst     = ss.getSheetByName(SHEET_PAGOS_PEND);
  const m       = _headerMap_(src);
  const data    = src.getDataRange().getValues();
  const dataCat = catAct.getDataRange().getValues();
  const hoy     = new Date();

  const pendientes = [];
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r      = data[i];
    const p      = r[m["participante"] - 1];
    const cid    = r[m["creamos id"]   - 1] || "";
    const estado = r[m["estado pago"]  - 1];
    if (!p || estado !== "Pendiente") continue;

    const fe     = new Date(r[m["fecha entrega"] - 1]);
    const dias   = isNaN(fe) ? 0 : Math.floor((hoy - fe) / (1000 * 60 * 60 * 24));
    const prod   = r[m["producto"]          - 1] || "";
    const ub     = Number(r[m["unidades buenas"] - 1]) || 0;
    const tq     = Number(r[m["total q"]         - 1]) || 0;
    const metodo = r[m["método pago"] - 1] || "";

    let numCuenta  = "";
    let tipoCuenta = "";
    for (let j = 1; j < dataCat.length; j++) {
      if (dataCat[j][CAT_COL.NOMBRE] === p) {
        numCuenta  = dataCat[j][CAT_COL.NUM_CUENTA]  || "";
        tipoCuenta = dataCat[j][CAT_COL.TIPO_CUENTA] || "";
        break;
      }
    }

    pendientes.push([p, cid, fe, prod, ub, tq, dias, metodo, numCuenta, tipoCuenta]);
  }

  pendientes.sort((a, b) => b[6] - a[6]);

  dst.clear();
  dst.getRange("A1:K1").setValues([[
    "Participante", "Creamos ID", "Fecha entrega", "Producto",
    "Unidades buenas", "Total Q", "Días pendiente",
    "Método Pago", "Nº Cuenta", "Tipo Cuenta", "Ejecución Pago"
  ]])
    .setBackground("#c62828").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dst.setFrozenRows(1);
  dst.setRowHeight(1, 25);

  if (pendientes.length) {
    const datosConAccion = pendientes.map(p => [...p, ""]);
    dst.getRange(2, 1, pendientes.length, 11).setValues(datosConAccion);
    dst.getRange(2, 3, pendientes.length, 1).setNumberFormat("yyyy-mm-dd");
    dst.getRange(2, 6, pendientes.length, 1).setNumberFormat('"Q " #,##0.00');
    dst.getRange(2, 1, pendientes.length, 11)
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

    dst.getRange(2, 11, pendientes.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["", "Pago", "Cheque entregado"])
        .setAllowInvalid(true)
        .build()
    );

    for (let i = 0; i < pendientes.length; i++) {
      if (pendientes[i][6] > 30) {
        dst.getRange(2 + i, 1, 1, 11).setBackground("#ffebee").setFontColor("#c62828").setFontWeight("bold");
      } else {
        dst.getRange(2 + i, 1, 1, 11).setBackground("#ffffff").setFontColor("#212121");
      }
    }
  }

  dst.setColumnWidth(1, 150); dst.setColumnWidth(2, 100); dst.setColumnWidth(3, 100);
  dst.setColumnWidth(4, 120); dst.setColumnWidth(5, 120); dst.setColumnWidth(6, 90);
  dst.setColumnWidth(7, 100); dst.setColumnWidth(8, 110); dst.setColumnWidth(9, 130);
  dst.setColumnWidth(10, 120); dst.setColumnWidth(11, 140);
}

function procesarParticipantesInactivos() {
  const ss    = SpreadsheetApp.getActive();
  const shAct = ss.getSheetByName(SHEET_CATALOGOS);
  const shIna = ss.getSheetByName(SHEET_INACTIVOS);
  if (!shAct || !shIna) return;

  const data          = shAct.getDataRange().getValues();
  const rowsAEliminar = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][CAT_COL.ETAPA] === "Retiradx") {
      shIna.appendRow(data[i].slice(0, 8));
      rowsAEliminar.push(i + 1);
    }
  }

  rowsAEliminar.forEach(row => shAct.deleteRow(row));
}

function aplicarColoresAutomaticos() {
  const sh   = SpreadsheetApp.getActive().getSheetByName(SHEET_RECEPCIONES);
  const m    = _headerMap_(sh);
  const data = sh.getDataRange().getValues();
  const hoy  = new Date();

  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r      = data[i];
    const p      = r[m["participante"] - 1];
    if (!p) continue;

    const estado = r[m["estado pago"]  - 1];
    const fe     = new Date(r[m["fecha entrega"] - 1]);
    const dias   = isNaN(fe) ? 0 : Math.floor((hoy - fe) / (1000 * 60 * 60 * 24));
    const rowNum = i + 1;
    const ncols  = RECEPCIONES_HEADERS.length;

    if (estado === "Quincena cerrada" || estado === "Pagado") {
      // Verde: quincena procesada / pago registrado
      sh.getRange(rowNum, 1, 1, ncols).setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("normal");
    } else if (estado === "Espera cierre quincena" || estado === "Pendiente") {
      // Amarillo suave: esperando cierre
      sh.getRange(rowNum, 1, 1, ncols).setBackground("#fff9c4").setFontColor("#f57f17").setFontWeight("normal");
    } else {
      sh.getRange(rowNum, 1, 1, ncols).setBackground("#ffffff").setFontColor("#212121").setFontWeight("normal");
    }

    sh.getRange(rowNum, 1, 1, ncols)
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function actualizarDashboard() {
  const ss   = SpreadsheetApp.getActive();
  const src  = ss.getSheetByName(SHEET_RECEPCIONES);
  const dst  = ss.getSheetByName(SHEET_DASHBOARD);
  const m    = _headerMap_(src);
  const data = src.getDataRange().getValues();
  const hoy  = new Date();
  const mesActual = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "MMMM yyyy");

  const stats = {
    participantes:  new Set(),
    clientes:       new Set(),
    pedidos:        0,
    totalQ:         0,
    ubTotal:        0,
    urTotal:        0,
    pagados:        0,
    montoPagado:    0,
    pendientes:     0,
    montoPendiente: 0
  };

  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r = data[i];
    const p = r[m["participante"] - 1];
    if (!p) continue;

    stats.participantes.add(p);
    stats.clientes.add(r[m["proyecto / cliente"] - 1] || "");
    stats.pedidos++;

    const ub     = Number(r[m["unidades buenas"]    - 1]) || 0;
    const ur     = Number(r[m["unidades rechazadas"] - 1]) || 0;
    const tapD   = m["total a pagar"] ? Number(r[m["total a pagar"] - 1]) || 0 : 0;
    const tqbD   = Number(r[m["total q"] - 1]) || 0;
    const tq     = tapD > 0 ? tapD : tqbD * 1.05;
    const estado = String(r[m["estado pago"] - 1] || "").trim();

    stats.totalQ  += tq;
    stats.ubTotal += ub;
    stats.urTotal += ur;

    const esCerrado  = estado === "Quincena cerrada" || estado === "Pagado";
    const esPendiente = estado === "Espera cierre quincena" || estado === "Pendiente";
    if (esCerrado)  { stats.pagados++;    stats.montoPagado    += tq; }
    if (esPendiente){ stats.pendientes++; stats.montoPendiente += tq; }
  }

  const promIngresos = stats.participantes.size > 0 ? stats.totalQ / stats.participantes.size : 0;
  const tasaRechazo  = (stats.ubTotal + stats.urTotal) > 0 ? stats.urTotal / (stats.ubTotal + stats.urTotal) : 0;

  dst.clear();
  dst.getRange("A1:D1").merge().setValue("📊 DASHBOARD DE INDICADORES")
    .setFontWeight("bold").setFontSize(16).setBackground("#1976d2").setFontColor("white")
    .setHorizontalAlignment("center");
  dst.getRange("A2:D2").merge().setValue(`Período: ${mesActual}`)
    .setHorizontalAlignment("center").setFontColor("#666").setFontSize(11);

  let row = 4;

  const indicadores = [
    ["ME.P.01", "Participantes Activos",   stats.participantes.size,               "personas", "#1976d2"],
    ["ME.P.07", "Clientes Únicos",          stats.clientes.size,                    "clientes", "#388e3c"],
    ["ME.P.08", "Pedidos Completados",      stats.pedidos,                          "pedidos",  "#f57c00"],
    ["ME.P.10", "Ingresos Totales",         "Q " + stats.totalQ.toFixed(2),         "GTQ",      "#d32f2f"],
    ["ME.R.01", "Promedio/Participante",    "Q " + promIngresos.toFixed(2),         "GTQ",      "#7b1fa2"],
    ["CALIDAD", "Tasa de Rechazo",          (tasaRechazo * 100).toFixed(1) + "%",   "%",        "#e64a19"],
    ["PAGOS",   "Pendientes",               "Q " + stats.montoPendiente.toFixed(2), "GTQ",      "#c62828"]
  ];

  for (const ind of indicadores) {
    dst.getRange(row, 1).setValue(ind[0]).setFontWeight("bold").setFontSize(9).setBackground("#eceff1").setFontColor("#37474f");
    dst.getRange(row, 2).setValue(ind[1]).setFontWeight("bold").setFontSize(10).setBackground("#ffffff");
    dst.getRange(row, 3).setValue(ind[2]).setFontWeight("bold").setFontSize(13).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.getRange(row, 4).setValue(ind[3]).setFontSize(9).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.setRowHeight(row, 28);
    row++;
  }

  // ── Mejora 2: sección métodos de pago ──────────────────────────────────────
  row++;
  dst.getRange(row, 1, 1, 4).merge().setValue("💳 MÉTODOS DE PAGO")
    .setFontWeight("bold").setBackground("#37474f").setFontColor("white")
    .setHorizontalAlignment("center");
  dst.setRowHeight(row, 22);
  row++;

  const shCheques = ss.getSheetByName(SHEET_CHEQUES);
  const shTrans   = ss.getSheetByName(SHEET_TRANSFERENCIAS);

  let chequesPend = 0, montoCheqPend = 0, chequesOk = 0, montoCheqOk = 0;
  let transPend   = 0, montoTransPend = 0, transOk   = 0, montoTransOk = 0;

  if (shCheques && shCheques.getLastRow() > 1) {
    shCheques.getRange(2, 1, shCheques.getLastRow() - 1, 9).getValues().forEach(r => {
      const estado = String(r[7] || "").trim();
      const total  = Number(r[5]) || 0;
      if (estado === "Pendiente") { chequesPend++; montoCheqPend += total; }
      if (estado === "Entregado") { chequesOk++;   montoCheqOk   += total; }
    });
  }

  if (shTrans && shTrans.getLastRow() > 1) {
    shTrans.getRange(2, 1, shTrans.getLastRow() - 1, 12).getValues().forEach(r => {
      const estado = String(r[10] || "").trim();
      const total  = Number(r[7])  || 0;
      if (estado === "Pendiente") { transPend++; montoTransPend += total; }
      if (estado === "Ejecutado") { transOk++;   montoTransOk   += total; }
    });
  }

  const indicadoresPago = [
    ["CHQ-PEND", "Cheques pendientes",       chequesPend > 0 ? "Q " + montoCheqPend.toFixed(2)  : "—", "GTQ", "#7b1fa2"],
    ["CHQ-OK",   "Cheques entregados",        "Q " + montoCheqOk.toFixed(2),                            "GTQ", "#4a148c"],
    ["TRF-PEND", "Transferencias pendientes", transPend > 0 ? "Q " + montoTransPend.toFixed(2)   : "—", "GTQ", "#00838f"],
    ["TRF-OK",   "Transferencias ejecutadas", "Q " + montoTransOk.toFixed(2),                           "GTQ", "#006064"]
  ];

  for (const ind of indicadoresPago) {
    dst.getRange(row, 1).setValue(ind[0]).setFontWeight("bold").setFontSize(9).setBackground("#eceff1").setFontColor("#37474f");
    dst.getRange(row, 2).setValue(ind[1]).setFontWeight("bold").setFontSize(10).setBackground("#ffffff");
    dst.getRange(row, 3).setValue(ind[2]).setFontWeight("bold").setFontSize(13).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.getRange(row, 4).setValue(ind[3]).setFontSize(9).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.setRowHeight(row, 28);
    row++;
  }

  dst.setColumnWidth(1, 100);
  dst.setColumnWidth(2, 220);
  dst.setColumnWidth(3, 160);
  dst.setColumnWidth(4, 100);
}

function actualizarReportes() {
  const ss     = SpreadsheetApp.getActive();
  const src    = ss.getSheetByName(SHEET_RECEPCIONES);
  const catAct = ss.getSheetByName(SHEET_CATALOGOS);
  const catIna = ss.getSheetByName(SHEET_INACTIVOS);
  const dst    = ss.getSheetByName(SHEET_REPORTES);
  const m      = _headerMap_(src);
  const data   = src.getDataRange().getValues();
  const hoy    = new Date();

  const mapaCreamosID = {};

  if (catAct) {
    catAct.getDataRange().getValues().slice(1).forEach(r => {
      const nombre = String(r[CAT_COL.NOMBRE] || "").trim();
      const id     = String(r[CAT_COL.ID]     || "").trim();
      if (nombre && id) mapaCreamosID[nombre] = id;
    });
  }

  if (catIna) {
    catIna.getDataRange().getValues().slice(1).forEach(r => {
      const nombre = String(r[CAT_COL.NOMBRE] || "").trim();
      const id     = String(r[CAT_COL.ID]     || "").trim();
      if (nombre && id) mapaCreamosID[nombre] = id;
    });
  }

  const reportes = [];
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r = data[i];
    const p = String(r[m["participante"] - 1] || "").trim();
    if (!p) continue;

    const fe          = new Date(r[m["fecha entrega"] - 1]);
    const mesAño      = isNaN(fe) ? "" : Utilities.formatDate(fe, Session.getScriptTimeZone(), "yyyy-MM");
    const ub          = Number(r[m["unidades buenas"]    - 1]) || 0;
    const ur          = Number(r[m["unidades rechazadas"] - 1]) || 0;
    const tasaRechazo = (ub + ur) > 0 ? ur / (ub + ur) : 0;
    const tapR    = m["total a pagar"] ? Number(r[m["total a pagar"] - 1]) || 0 : 0;
    const tqR     = Number(r[m["total q"] - 1]) || 0;
    const total   = tapR > 0 ? tapR : tqR * 1.05;
    const estado  = String(r[m["estado pago"] - 1] || "").trim();
    const dias    = isNaN(fe) ? 0 : Math.floor((hoy - fe) / (1000 * 60 * 60 * 24));
    const esPend  = estado === "Espera cierre quincena" || estado === "Pendiente";
    const creamosID = mapaCreamosID[p] || "";

    reportes.push([
      fe, mesAño, creamosID, p,
      String(r[m["producto"]           - 1] || "").trim(),
      String(r[m["proyecto / cliente"] - 1] || "").trim(),
      ub, ur, tasaRechazo, total, estado,
      esPend ? dias : 0
    ]);
  }

  dst.getRange("A3:L1000").clearContent();

  if (reportes.length) {
    dst.getRange(3, 1, reportes.length, 12).setValues(reportes);
    dst.getRange(3, 1, reportes.length, 12)
      .setBackground("#ffffff").setFontColor("#212121")
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    dst.getRange(3, 1,  reportes.length, 1).setNumberFormat("yyyy-mm-dd");
    dst.getRange(3, 9,  reportes.length, 1).setNumberFormat("0.00%");
    dst.getRange(3, 10, reportes.length, 1).setNumberFormat('"Q " #,##0.00');
  }
}


// ========================= VALIDACIONES =========================
function crearValidacionesDatos() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName(SHEET_RECEPCIONES);
  const cat = ss.getSheetByName(SHEET_CATALOGOS);
  if (!sh || !cat) return;

  const m = _headerMap_(sh);
  sh.getRange(DATA_START_ROW, m["participante"], 2000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(cat.getRange("B2:B1000"), true)
      .setAllowInvalid(true)
      .build()
  );
}

// ========================= TRIGGERS =========================
function crearTriggers() {
  eliminarTriggers();
  ScriptApp.newTrigger("actualizarTodo").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("alEditarHoja").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

function eliminarTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

function alEditarHoja(e) {
  if (!e || !e.range) return;
  const sh     = e.range.getSheet();
  const shName = sh.getName();

  if (shName === SHEET_RECEPCIONES) {
    if (e.range.getRow() < DATA_START_ROW) return;
    const m = _headerMap_(sh);
    const c = e.range.getColumn();
    if (c === m["unidades buenas"] || c === m["precio unit. (q)"] || c === m["fecha pago"]) {
      actualizarTodo();
    }
  } else if (shName === SHEET_PAGOS_PEND) {
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2 || col !== 11) return;

    const valor = e.value;
    if (!valor || (valor !== "Pago" && valor !== "Cheque entregado")) return;

    _ejecutarPago_(row, valor);
  }
}

function _ejecutarPago_(rowPagoPend, tipo) {
  try {
    const ss     = SpreadsheetApp.getActive();
    const shPago = ss.getSheetByName(SHEET_PAGOS_PEND);
    const shRec  = ss.getSheetByName(SHEET_RECEPCIONES);

    const participante = String(shPago.getRange(rowPagoPend, 1).getValue()).trim();
    const fechaEntrega = shPago.getRange(rowPagoPend, 3).getValue();

    // Normalizar a medianoche para comparar solo por día
    const diaEntrega = new Date(fechaEntrega);
    diaEntrega.setHours(0, 0, 0, 0);

    const mRec    = _headerMap_(shRec);
    const dataRec = shRec.getDataRange().getValues();
    const hoy     = new Date();

    let encontrado    = false;
    let datosFilaPago = null;

    for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
      const r  = dataRec[i];
      const p  = String(r[mRec["participante"] - 1] || "").trim();
      const fe = new Date(r[mRec["fecha entrega"] - 1]);

      if (p === participante && !isNaN(fe)) {
        const diaFe = new Date(fe);
        diaFe.setHours(0, 0, 0, 0);

        if (diaFe.getTime() === diaEntrega.getTime()) {
          const rowRec      = i + 1;
          const metodoValor = tipo === "Cheque entregado" ? "Cheque" : "Transferencia";

          shRec.getRange(rowRec, mRec["estado pago"]).setValue("Pagado");
          shRec.getRange(rowRec, mRec["fecha pago"]).setValue(hoy);
          shRec.getRange(rowRec, mRec["método pago"]).setValue(metodoValor);

          // Capturar datos de la fila para Mejoras 2 y 4
          datosFilaPago = {
            quincena:  String(r[mRec["quincena"]          - 1] || "").trim(),
            creamosID: String(r[mRec["creamos id"]        - 1] || "").trim(),
            producto:  String(r[mRec["producto"]          - 1] || "").trim(),
            ub:        Number(r[mRec["unidades buenas"]   - 1]) || 0,
            tq:        Number(r[mRec["total q"]           - 1]) || 0,
            metodo:    metodoValor
          };

          encontrado = true;
          break;
        }
      }
    }

    shPago.getRange(rowPagoPend, 11).clearContent();

    if (encontrado && datosFilaPago) {
      // Mejoras 2 y 4: registrar en Cheques/Transferencias e Historial_Pagos
      _registrarPagoEnHojas_(participante, datosFilaPago, hoy, tipo);
    }

    if (encontrado) {
      SpreadsheetApp.getActive().toast("✅ Pago marcado como " + tipo, null, 2);
    } else {
      SpreadsheetApp.getActive().toast("⚠️ No se encontró la entrega", null, 3);
    }

    actualizarTodo();
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}


// ========================= MEJORA 1 – QUINCENAS =========================

function _obtenerQuincenaActiva_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PERIODOS);
  if (!sh || sh.getLastRow() < 2) return "";
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4] || "").trim() === "Activo") return String(data[i][1] || "").trim();
  }
  return "";
}

function _formatNombreQuincena_(inicio, fin) {
  const mes = ["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"];
  return `Q_${inicio.getDate()}_${mes[inicio.getMonth()]}_${fin.getDate()}_${mes[fin.getMonth()]}_${fin.getFullYear()}`;
}

// Convierte texto DD/MM/AAAA o AAAA-MM-DD a Date. Retorna null si no es válido.
function _parsearFecha_(texto) {
  if (!texto) return null;
  texto = texto.trim();
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(texto)) {
    const p = texto.split(/[\/\-]/);
    const f = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    return isNaN(f.getTime()) ? null : f;
  }
  if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(texto)) {
    const p = texto.split(/[\/\-]/);
    const f = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
    return isNaN(f.getTime()) ? null : f;
  }
  return null;
}

// Muestra un prompt con fecha propuesta. El usuario escribe otra o presiona OK para aceptar.
function _pedirFecha_(ui, titulo, propuesta) {
  const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const resp = ui.prompt(
    titulo,
    "Propuesta: " + fmt(propuesta) + "\n\nPresiona OK para aceptar, o escribe otra fecha (DD/MM/AAAA):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  const texto = resp.getResponseText().trim();
  return texto ? (_parsearFecha_(texto) || propuesta) : propuesta;
}

function crearQuincenaInicial() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const shP = ss.getSheetByName(SHEET_PERIODOS);
    const ui  = SpreadsheetApp.getUi();
    if (!shP) return ui.alert("Instala el sistema primero.");

    // Bloquear si ya existe quincena activa
    if (shP.getLastRow() > 1) {
      const data = shP.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][4] || "").trim() === "Activo") {
          return ui.alert(
            "Ya existe una quincena activa",
            "Quincena activa: " + data[i][1] + "\n\nUsa 'Cerrar quincena actual' para avanzar al siguiente período.",
            ui.ButtonSet.OK
          );
        }
      }
    }

    // Pedir fecha de INICIO al usuario
    const rInicio = ui.prompt(
      "🗓️ Primera quincena — Fecha de INICIO",
      "Escribe la fecha de inicio de la primera quincena.\nFormato: DD/MM/AAAA\n\nEjemplo: 11/06/2026",
      ui.ButtonSet.OK_CANCEL
    );
    if (rInicio.getSelectedButton() !== ui.Button.OK) return;
    const inicio = _parsearFecha_(rInicio.getResponseText().trim());
    if (!inicio) return ui.alert("❌ Fecha no válida. Usa el formato DD/MM/AAAA (ejemplo: 11/06/2026).");

    // Proponer fin (+14 días) y dejar que el usuario lo confirme o cambie
    const finPropuesto = new Date(inicio);
    finPropuesto.setDate(finPropuesto.getDate() + 14);
    const fin = _pedirFecha_(ui, "🗓️ Primera quincena — Fecha de FIN", finPropuesto);
    if (!fin) return;

    const nombre  = _formatNombreQuincena_(inicio, fin);
    shP.appendRow([shP.getLastRow(), nombre, inicio, fin, "Activo", "Q1"]);
    shP.getRange(shP.getLastRow(), 3, 1, 2).setNumberFormat("yyyy-mm-dd");

    actualizarHistorialQuincenas();
    ss.toast("✅ Primera quincena creada: " + nombre, null, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

// Cambia el estado de todas las filas de una quincena a "Quincena cerrada"
function _marcarQuincenaCerrada_(ss, nombreQ) {
  const sh = ss.getSheetByName(SHEET_RECEPCIONES);
  if (!sh) return;
  const m    = _headerMap_(sh);
  const data = sh.getDataRange().getValues();
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const q = String(data[i][m["quincena"] - 1] || "").trim();
    if (q !== nombreQ) continue;
    const colEstado = m["estado pago"];
    if (colEstado) sh.getRange(i + 1, colEstado).setValue("Quincena cerrada");
  }
}

// Calcula montos por participante en Recepciones (filtrado por quincena) y los escribe en Cheques o Transferencias
function _generarPagosQuincena_(ss, ordenQ, fechaFin, nombreQ) {
  const rec = ss.getSheetByName(SHEET_RECEPCIONES);
  const cat = ss.getSheetByName(SHEET_CATALOGOS);
  const shT = ss.getSheetByName(SHEET_TRANSFERENCIAS);
  const shC = ss.getSheetByName(SHEET_CHEQUES);
  if (!rec || !cat || !shT || !shC) return;

  const m       = _headerMap_(rec);
  const dataRec = rec.getDataRange().getValues();
  const dataCat = cat.getDataRange().getValues();
  const colPagar     = m["total a pagar"] ? m["total a pagar"] - 1 : m["total q"] - 1;
  const colCategoria = m["categoría"] ? m["categoría"] - 1 : (m["proyecto / cliente"] ? m["proyecto / cliente"] - 1 : -1);

  // totales[participante] = monto total (interno)
  // desglose[participante][servicio] = monto (para hojas externas)
  const totales  = {};
  const desglose = {};
  for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
    const r = dataRec[i];
    const p = String(r[m["participante"] - 1] || "").trim();
    const q = String(r[m["quincena"]     - 1] || "").trim();
    if (!p) continue;
    if (q && q !== nombreQ) continue;
    const tap      = Math.round(Number(r[colPagar]) || 0);
    const servicio = colCategoria >= 0 ? String(r[colCategoria] || "Sin categoría").trim() : "Sin categoría";
    totales[p] = (totales[p] || 0) + tap;
    if (!desglose[p]) desglose[p] = {};
    desglose[p][servicio] = (desglose[p][servicio] || 0) + tap;
  }

  // Info bancaria de Participantes Activos (incluye Cuenta de Pago)
  const infoPart = {};
  for (let i = 1; i < dataCat.length; i++) {
    const row    = dataCat[i];
    const nombre = String(row[CAT_COL.NOMBRE] || "").trim();
    if (!nombre) continue;
    infoPart[nombre] = {
      cid:         String(row[CAT_COL.ID]          || "").trim(),
      banco:       String(row[CAT_COL.BANCO]        || "").trim(),
      tipoCuenta:  String(row[CAT_COL.TIPO_CUENTA]  || "").trim(),
      numCuenta:   String(row[CAT_COL.NUM_CUENTA]   || "").trim(),
      titular:     String(row[CAT_COL.TITULAR]      || "").trim(),
      cuentaPago:  String(row[CAT_COL.CUENTA_PAGO]  || "Creamos").trim(),
    };
  }

  const tz   = Session.getScriptTimeZone();
  const mes  = Utilities.formatDate(fechaFin, tz, "MMMM yyyy");
  const anio = fechaFin.getFullYear();
  const esQ1 = ordenQ === "Q1";

  // — Hojas internas (Cheques / Transferencias) — agrupadas por participante
  for (const [nombre, monto] of Object.entries(totales)) {
    if (monto <= 0) continue;
    const info  = infoPart[nombre] || {};
    const esChq = (info.tipoCuenta || "").toLowerCase() === "cheque";
    const sh    = esChq ? shC : shT;

    if (esQ1) {
      const newRow = sh.getLastRow() + 1;
      sh.getRange(newRow, 1, 1, 11).setValues([[
        nombre, info.cid, info.banco, info.tipoCuenta, String(info.numCuenta), info.titular,
        monto, "", monto, mes, anio
      ]]);
      sh.getRange(newRow, 5).setNumberFormat("@");
    } else {
      const nRows = sh.getLastRow() - 2;
      const data  = nRows > 0 ? sh.getRange(3, 1, nRows, 11).getValues() : [];
      let found   = false;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === nombre && String(data[i][9]).trim() === mes) {
          const q1 = Number(data[i][6]) || 0;
          sh.getRange(i + 3, 8).setValue(monto);
          sh.getRange(i + 3, 9).setValue(q1 + monto);
          found = true;
          break;
        }
      }
      if (!found) {
        const newRow2 = sh.getLastRow() + 1;
        sh.getRange(newRow2, 1, 1, 11).setValues([[
          nombre, info.cid, info.banco, info.tipoCuenta, String(info.numCuenta), info.titular,
          0, monto, monto, mes, anio
        ]]);
        sh.getRange(newRow2, 5).setNumberFormat("@");
      }
    }
  }

  // — Hojas externas (Women Payment 26 + Cheques externo) —
  try {
    _enviarPagosExterno_(desglose, infoPart, ordenQ, mes);
  } catch (e) {
    SpreadsheetApp.getActive().toast("⚠️ Hojas internas OK. Error al escribir hojas externas: " + e.message, null, 6);
  }
}

// Devuelve el tab del mes en una hoja externa; lo crea con encabezados si no existe
function _obtenerTabMes_(ssExt, mes, esCheque) {
  const HDR = ["Nombre","Tipo de Pago","Servicio","Banco","Tipo de Cuenta",
               "Numero de Cta","Cuenta de Pago","Quincena 1","Quincena 2","Total Mes"];
  const COLOR_HDR = esCheque ? "#1a237e" : "#263238";
  let sh = ssExt.getSheetByName(mes);
  if (!sh) {
    sh = ssExt.insertSheet(mes);
    sh.getRange(1, 1, 1, HDR.length).setValues([HDR])
      .setFontWeight("bold").setBackground(COLOR_HDR).setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 140);
    sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 140);
    sh.setColumnWidth(7, 100); sh.setColumnWidth(8, 90);  sh.setColumnWidth(9, 90);
    sh.setColumnWidth(10, 100);
    sh.getRange("F2:F1000").setNumberFormat("@");
    sh.getRange("H2:J1000").setNumberFormat('"Q "#,##0.00');
  }
  return sh;
}

// Escribe pagos desglosados por Servicio en Women Payment 26 y Cheques externo (tab por mes)
function _enviarPagosExterno_(desglose, infoPart, ordenQ, mes) {
  const esQ1 = ordenQ === "Q1";
  const colQ = esQ1 ? 8 : 9;   // col H = Q1, col I = Q2

  // Separar participantes: cheques vs transferencias
  const filasTrans = {};
  const filasChq   = {};

  for (const [nombre, servicios] of Object.entries(desglose)) {
    const info  = infoPart[nombre] || {};
    const esChq = (info.tipoCuenta || "").toLowerCase() === "cheque";
    const dest  = esChq ? filasChq : filasTrans;
    for (const [servicio, monto] of Object.entries(servicios)) {
      if (!dest[servicio]) dest[servicio] = [];
      dest[servicio].push({ nombre, info, monto });
    }
  }

  // ── Women Payment 26 — tab por mes ──
  try {
    const ssExt = SpreadsheetApp.openById(SS_ID_WOMEN_PAYMENT);
    const sh    = _obtenerTabMes_(ssExt, mes, false);
    const lastRow = sh.getLastRow();
    const data    = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 10).getValues() : [];

    for (const [servicio, filas] of Object.entries(filasTrans)) {
      for (const { nombre, info, monto } of filas) {
        if (monto <= 0) continue;
        let found = false;
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][0]).trim() === nombre && String(data[i][2]).trim() === servicio) {
            const fila = i + 2;
            sh.getRange(fila, colQ).setValue(monto);
            const q1 = Number(data[i][7]) || 0;
            const q2 = Number(data[i][8]) || 0;
            sh.getRange(fila, 10).setValue(esQ1 ? monto + q2 : q1 + monto);
            found = true;
            break;
          }
        }
        if (!found) {
          const nr = sh.getLastRow() + 1;
          sh.getRange(nr, 1, 1, 10).setValues([[
            nombre, "Pago Cuenta", servicio, info.banco, info.tipoCuenta,
            String(info.numCuenta), info.cuentaPago,
            esQ1 ? monto : 0, esQ1 ? 0 : monto, monto
          ]]);
          sh.getRange(nr, 6).setNumberFormat("@");
          sh.getRange(nr, 8, 1, 3).setNumberFormat('"Q "#,##0.00');
        }
      }
    }
  } catch (e) {
    throw new Error("Women Payment 26: " + e.message);
  }

  // ── Cheques externo — tab por mes ──
  try {
    const ssChq = SpreadsheetApp.openById(SS_ID_CHEQUES_EXT);
    const shChq = _obtenerTabMes_(ssChq, mes, true);
    const lastRow = shChq.getLastRow();
    const data    = lastRow > 1 ? shChq.getRange(2, 1, lastRow - 1, 10).getValues() : [];

    for (const [servicio, filas] of Object.entries(filasChq)) {
      for (const { nombre, info, monto } of filas) {
        if (monto <= 0) continue;
        let found = false;
        for (let i = 0; i < data.length; i++) {
          if (String(data[i][0]).trim() === nombre && String(data[i][2]).trim() === servicio) {
            shChq.getRange(i + 2, colQ).setValue(monto);
            found = true;
            break;
          }
        }
        if (!found) {
          const nr = shChq.getLastRow() + 1;
          shChq.getRange(nr, 1, 1, 10).setValues([[
            nombre, "Cheque", servicio, info.banco, info.tipoCuenta,
            String(info.numCuenta), info.cuentaPago,
            esQ1 ? monto : 0, esQ1 ? 0 : monto, monto
          ]]);
          shChq.getRange(nr, 6).setNumberFormat("@");
          shChq.getRange(nr, 8, 1, 3).setNumberFormat('"Q "#,##0.00');
        }
      }
    }
  } catch (e) {
    throw new Error("Cheques externo: " + e.message);
  }
}

// Al cierre de mes: copia registros a Historial_Pagos y limpia Cheques/Transferencias
function _archivarPagosFinMes_(ss) {
  const shT  = ss.getSheetByName(SHEET_TRANSFERENCIAS);
  const shC  = ss.getSheetByName(SHEET_CHEQUES);
  const hist = ss.getSheetByName(SHEET_HISTORIAL_PAGOS);
  if (!hist) return;

  [[shT, "Transferencia"], [shC, "Cheque"]].forEach(([sh, tipo]) => {
    if (!sh || sh.getLastRow() < 3) return;
    const data = sh.getRange(3, 1, sh.getLastRow() - 2, 11).getValues();
    data.forEach(row => { if (row[0]) hist.appendRow([...row, tipo]); });
    sh.deleteRows(3, sh.getLastRow() - 2);
  });
}

// ── Trigger instalable: cuando participante cambia a Inactivo ─────────────────
function instalarTriggerInactivos() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onEditParticipantes")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onEditParticipantes")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getActive().toast("✅ Trigger de participantes instalado.", null, 4);
}

function onEditParticipantes(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_CATALOGOS) return;
    if (e.range.getColumn() !== CAT_COL.ETAPA + 1) return;
    if (String(e.value || "").trim() !== "Retiradx") return;

    const row    = e.range.getRow();
    const nombre = sheet.getRange(row, CAT_COL.NOMBRE + 1).getValue();
    mostrarFormularioInactivo(nombre, row);
  } catch (err) {
    console.log("onEditParticipantes error: " + err.message);
  }
}

function mostrarFormularioInactivo(nombre, rowNum) {
  const motivos = [
    "Otras prioridades",
    "Horario laboral",
    "Retos/problemas familiares",
    "Violencia de parte de la pareja/violencia de género",
    "Migración (por motivos económicos/por violencia)",
    "Embarazo",
    "Retos/problemas de salud física",
    "Retos/problemas de salud mental",
    "Retos/Problemas legales/Privación de libertad",
    "Falta de apoyo",
    "Compromisos religiosos",
    "Problemas financieros",
    "Violencia comunitaria",
    "No querer continuar en el proceso",
    "Descontento con la organización",
    "Asesinato/Fallecimiento",
    "Cuidado de terceras personas",
    "Falta de adaptabilidad",
    "Pérdida de contacto / Inaccesibilidad",
    "Expectativas no alineadas con el programa",
    "Cambio de prioridad personal",
    "Sobrecarga personal / Dificultad para sostener el proceso"
  ];

  const opcionesHtml = motivos.map(m => '<option value="' + m + '">' + m + '</option>').join('');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    '* { margin:0; padding:0; box-sizing:border-box; }' +
    'body { font-family:"Segoe UI",sans-serif; background:#f5f5f5; padding:20px; }' +
    '.container { background:white; border-radius:10px; padding:22px; box-shadow:0 4px 20px rgba(0,0,0,0.1); }' +
    'h2 { color:#333; font-size:17px; margin-bottom:4px; }' +
    '.nombre { color:#5e35b1; font-weight:700; font-size:15px; margin-bottom:16px; }' +
    'label { display:block; font-size:11px; font-weight:600; color:#666; text-transform:uppercase; letter-spacing:0.4px; margin-bottom:6px; }' +
    'select, textarea { width:100%; padding:9px 12px; border:1.5px solid #e0e0e0; border-radius:6px; font-size:13px; font-family:inherit; background:white; }' +
    'select { cursor:pointer; } select:focus, textarea:focus { outline:none; border-color:#7e57c2; }' +
    'textarea { resize:vertical; margin-top:12px; }' +
    '.buttons { display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-top:16px; }' +
    'button { padding:10px; border:none; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; }' +
    '.btn-ok { background:linear-gradient(135deg,#7e57c2,#5e35b1); color:white; }' +
    '.btn-no { background:#f0f0f0; color:#666; }' +
    '</style></head><body><div class="container">' +
    '<h2>Participante inactiva</h2>' +
    '<p class="nombre">👤 ' + nombre + '</p>' +
    '<label>Motivo de inactividad *</label>' +
    '<select id="motivo"><option value="">— Seleccionar motivo —</option>' + opcionesHtml + '</select>' +
    '<textarea id="notas" rows="2" placeholder="Notas adicionales (opcional)..."></textarea>' +
    '<div class="buttons">' +
    '<button class="btn-ok" onclick="guardar()">✅ Guardar</button>' +
    '<button class="btn-no" onclick="google.script.host.close()">✕ Cancelar</button>' +
    '</div></div>' +
    '<script>' +
    'function guardar() {' +
    '  var m = document.getElementById("motivo").value;' +
    '  if (!m) { alert("Selecciona un motivo"); return; }' +
    '  var n = document.getElementById("notas").value.trim();' +
    '  var t = n ? m + " — " + n : m;' +
    '  google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).guardarMotivoInactivo(t, ' + rowNum + ');' +
    '}' +
    '<\/script></body></html>'
  ).setWidth(440).setHeight(330);

  SpreadsheetApp.getUi().showModalDialog(html, "Motivo de inactividad");
}

function guardarMotivoInactivo(motivo, rowNum) {
  const ss     = SpreadsheetApp.getActive();
  const shAct  = ss.getSheetByName(SHEET_CATALOGOS);
  const shInac = ss.getSheetByName(SHEET_INACTIVOS);
  if (!shAct || !shInac) return;
  const rowData = shAct.getRange(rowNum, 1, 1, 8).getValues()[0];
  shInac.appendRow([...rowData, motivo, new Date()]);
  shInac.getRange(shInac.getLastRow(), 10).setNumberFormat("yyyy-mm-dd");
}

function cerrarQuincenaActual() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const shP = ss.getSheetByName(SHEET_PERIODOS);
    const rec = ss.getSheetByName(SHEET_RECEPCIONES);
    const ui  = SpreadsheetApp.getUi();
    if (!shP) return ui.alert("Hoja PERIODOS no encontrada. Instala el sistema primero.");
    if (!rec)  return ui.alert("Hoja Recepciones no encontrada.");

    const dataPer = shP.getDataRange().getValues();
    let filaActiva = -1, nombreQ = "", fechaIni = null, fechaFin = null, ordenQ = "";

    for (let i = 1; i < dataPer.length; i++) {
      if (String(dataPer[i][4] || "").trim() === "Activo") {
        filaActiva = i + 1;
        nombreQ    = String(dataPer[i][1] || "").trim();
        fechaIni   = new Date(dataPer[i][2]);
        fechaFin   = new Date(dataPer[i][3]);
        ordenQ     = String(dataPer[i][5] || "").trim().toUpperCase();
        break;
      }
    }

    if (filaActiva === -1) {
      return ui.alert("No hay quincena activa.\n\nUsa '🗓️ Crear quincena inicial' primero.");
    }

    if (ordenQ === "Q2") {
      const hayQ1Cerrada = dataPer.slice(1).some(
        r => String(r[4] || "").trim() === "Cerrado" && String(r[5] || "").trim().toUpperCase() === "Q1"
      );
      if (!hayQ1Cerrada) {
        return ui.alert("⚠️ No puedes cerrar Q2 sin haber cerrado primero Q1.");
      }
    }

    const fmt  = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const mRec = _headerMap_(rec);
    const dataRec = rec.getDataRange().getValues();

    // Identificar filas pendientes de esta quincena
    const filasPendient = [];
    for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
      const r = dataRec[i];
      const p = String(r[mRec["participante"] - 1] || "").trim();
      const q = String(r[mRec["quincena"]     - 1] || "").trim();
      if (!p) continue;
      if (q && q !== nombreQ) continue;
      if (String(r[mRec["estado pago"] - 1] || "").trim() === "Pendiente") {
        filasPendient.push({ rowIndex: i, data: r });
      }
    }

    const esQ2 = (ordenQ === "Q2");
    let msgConfirm = "Quincena a cerrar:\n" + nombreQ + "\n(" + fmt(fechaIni) + " al " + fmt(fechaFin) + ")";
    if (filasPendient.length > 0)
      msgConfirm += "\n\n⚠️ " + filasPendient.length + " pago(s) pendiente(s) → quedarán en Recepciones asignados a la siguiente quincena.";
    if (esQ2)
      msgConfirm += "\n\n🗓️ Es la segunda quincena → el mes se cerrará y se guardará copia en Drive.";
    msgConfirm += "\n\n¿Continuar?";

    if (ui.alert("Cerrar quincena", msgConfirm, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    // 1. Generar pagos en Cheques / Transferencias
    _generarPagosQuincena_(ss, ordenQ, fechaFin, nombreQ);

    // 2. Marcar filas de esta quincena como "Quincena cerrada" en Recepciones
    _marcarQuincenaCerrada_(ss, nombreQ);

    // 3. Marcar quincena como Cerrada en PERIODOS
    shP.getRange(filaActiva, 5).setValue("Cerrado");

    // 3. Si es Q2 → archivar pagos a Historial y guardar Drive
    if (esQ2) _ejecutarCierreMes_(ss, false);

    // 4. Pedir fechas para la siguiente quincena
    const siguienteOrden = esQ2 ? "Q1" : "Q2";
    const propInicio     = new Date(fechaFin); propInicio.setDate(propInicio.getDate() + 1);
    const propFin        = new Date(propInicio); propFin.setDate(propFin.getDate() + 14);

    const nuevaInicio = _pedirFecha_(ui, "🗓️ " + siguienteOrden + " — Fecha de INICIO", propInicio);
    if (!nuevaInicio) {
      actualizarHistorialQuincenas();
      return ss.toast("✅ " + nombreQ + " cerrada. Crea la siguiente quincena cuando quieras.", null, 6);
    }

    const propNuevaFin = new Date(nuevaInicio); propNuevaFin.setDate(propNuevaFin.getDate() + 14);
    const nuevaFin     = _pedirFecha_(ui, "🗓️ " + siguienteOrden + " — Fecha de FIN", propNuevaFin);
    if (!nuevaFin) {
      actualizarHistorialQuincenas();
      return ss.toast("✅ " + nombreQ + " cerrada. Crea la siguiente quincena cuando quieras.", null, 6);
    }

    const nombreNueva = _formatNombreQuincena_(nuevaInicio, nuevaFin);
    shP.appendRow([shP.getLastRow(), nombreNueva, nuevaInicio, nuevaFin, "Activo", siguienteOrden]);
    shP.getRange(shP.getLastRow(), 3, 1, 2).setNumberFormat("yyyy-mm-dd");

    // 5. Reasignar pendientes a la nueva quincena (actualización en el lugar, sin borrar filas)
    if (filasPendient.length > 0) {
      for (const entry of filasPendient) {
        const shRow = entry.rowIndex + 1;
        rec.getRange(shRow, mRec["quincena"]).setValue(nombreNueva);
        const nota = String(entry.data[mRec["notas / calidad"] - 1] || "").trim();
        rec.getRange(shRow, mRec["notas / calidad"]).setValue(nota ? nota + " (arrastrado)" : "(arrastrado)");
      }
    }

    actualizarHistorialQuincenas();

    const pMes  = esQ2 ? " 🗓️ Mes cerrado y guardado en Drive." : "";
    const pPend = filasPendient.length > 0 ? ` ${filasPendient.length} pendiente(s) reasignado(s) a ${nombreNueva}.` : "";
    ss.toast(`✅ ${nombreQ} cerrada.${pPend}${pMes} Nueva: ${nombreNueva}`, null, 8);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

function actualizarHistorialQuincenas() {
  const ss   = SpreadsheetApp.getActive();
  const src  = ss.getSheetByName(SHEET_RECEPCIONES);
  const per  = ss.getSheetByName(SHEET_PERIODOS);
  const dst  = ss.getSheetByName(SHEET_HIST_QUINCENAS);
  const catS = ss.getSheetByName(SHEET_CATALOGOS);
  if (!src || !per || !dst) return;

  const mRec    = _headerMap_(src);
  const dataRec = src.getDataRange().getValues();
  const dataPer = per.getDataRange().getValues();

  // Todos los participantes activos para calcular "sin entrega"
  const todosParticipantes = new Set();
  if (catS && catS.getLastRow() > 1) {
    catS.getRange("B2:B" + catS.getLastRow()).getValues()
      .flat().filter(x => x).forEach(n => todosParticipantes.add(String(n).trim()));
  }

  // Agrupar recepciones por nombre de quincena
  const byQ = {};
  for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
    const r = dataRec[i];
    const p = String(r[mRec["participante"] - 1] || "").trim();
    if (!p) continue;
    const q = String(r[mRec["quincena"] - 1] || "").trim() || "Sin quincena";
    if (!byQ[q]) byQ[q] = { entregas: 0, ub: 0, ur: 0, tq: 0, parts: new Set() };
    byQ[q].entregas++;
    byQ[q].ub  += Number(r[mRec["unidades buenas"]    - 1]) || 0;
    byQ[q].ur  += Number(r[mRec["unidades rechazadas"] - 1]) || 0;
    const tapH = mRec["total a pagar"] ? Number(r[mRec["total a pagar"] - 1]) || 0 : 0;
    const tqbH = Number(r[mRec["total q"] - 1]) || 0;
    byQ[q].tq += tapH > 0 ? tapH : tqbH * 1.05;
    byQ[q].parts.add(p);
  }

  dst.clearContents();
  dst.clearFormats();
  dst.getRange("A1").setValue("HISTORIAL DE QUINCENAS").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");

  dst.getRange("A2:J2").setValues([[
    "Quincena", "Fecha_Inicio", "Fecha_Fin", "Entregas", "Unidades_Buenas",
    "Unidades_Rechazadas", "Total_Q", "Participantes_Con_Entrega",
    "Participantes_Sin_Entrega", "Notas"
  ]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dst.setFrozenRows(2);
  dst.setRowHeight(2, 25);

  const rows    = [];
  const estados = [];

  for (let i = 1; i < dataPer.length; i++) {
    const nombreQ  = String(dataPer[i][1] || "").trim();
    const fechaIni = dataPer[i][2];
    const fechaFin = dataPer[i][3];
    const estado   = String(dataPer[i][4] || "").trim();
    const notas    = dataPer[i][5] || "";
    const d        = byQ[nombreQ] || { entregas: 0, ub: 0, ur: 0, tq: 0, parts: new Set() };
    const conEntrega = d.parts.size;
    const sinEntrega = Math.max(0, todosParticipantes.size - conEntrega);
    rows.push([nombreQ, fechaIni, fechaFin, d.entregas, d.ub, d.ur, d.tq, conEntrega, sinEntrega, notas]);
    estados.push(estado);
  }

  if (rows.length) {
    dst.getRange(3, 1, rows.length, 10).setValues(rows);
    dst.getRange(3, 2, rows.length, 2).setNumberFormat("yyyy-mm-dd");
    dst.getRange(3, 7, rows.length, 1).setNumberFormat('"Q " #,##0.00');
    dst.getRange(3, 1, rows.length, 10)
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

    for (let i = 0; i < rows.length; i++) {
      if (estados[i] === "Activo") {
        // Resaltar la quincena en curso
        dst.getRange(3 + i, 1, 1, 10).setBackground("#e8f5e9").setFontColor("#1b5e20").setFontWeight("bold");
      } else {
        dst.getRange(3 + i, 1, 1, 10).setBackground("#ffffff").setFontColor("#212121");
      }
    }
  }

  dst.setColumnWidth(1, 210); dst.setColumnWidth(2, 110);  dst.setColumnWidth(3, 110);
  dst.setColumnWidth(4, 90);  dst.setColumnWidth(5, 130);  dst.setColumnWidth(6, 150);
  dst.setColumnWidth(7, 110); dst.setColumnWidth(8, 180);  dst.setColumnWidth(9, 180);
  dst.setColumnWidth(10, 220);
}


// ========================= CIERRE DE MES =========================

// Llamado manual desde el menú (pide confirmación)
function cerrarMes() {
  try {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();
    const ok = ui.alert(
      "🗓️ Cierre de mes manual",
      "Guardará copia en Drive y limpiará Recepciones y Archivo_Recepciones.\n⚠️ Historial_Pagos NUNCA se borra.\n\n¿Continuar?",
      ui.ButtonSet.YES_NO
    );
    if (ok !== ui.Button.YES) return;
    _ejecutarCierreMes_(ss, true);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

// Lógica real del cierre — llamada automáticamente desde cerrarQuincenaActual o manual desde cerrarMes
function _ejecutarCierreMes_(ss, mostrarToast) {
  const hoy   = new Date();
  const tz    = Session.getScriptTimeZone();
  const mesLabel = Utilities.formatDate(hoy, tz, "MMMM_yyyy");
  const nombrePDF = "Manufactura_" + NOMBRE_PROGRAMA.replace(/ /g, "_") + "_" + mesLabel;

  // Generar PDF con Recepciones + Transferencias + Cheques (en lugar de copia del documento)
  try {
    _generarPDFCierreMes_(ss, nombrePDF);
  } catch (e) {
    ss.toast("⚠️ PDF no generado: " + e.message + ". Continuando cierre...", null, 5);
  }

  // Archivar pagos a Historial_Pagos y limpiar Cheques/Transferencias
  _archivarPagosFinMes_(ss);

  // Recepciones NO se limpia — acumula el historial completo de entregas

  if (mostrarToast) {
    ss.toast("✅ Mes cerrado. PDF guardado en Drive: '" + nombrePDF + ".pdf'.", null, 8);
  }
}

// Exporta Recepciones + Cheques + Transferencias como PDF y lo guarda en Drive
function _generarPDFCierreMes_(ss, nombreArchivo) {
  const ssId  = ss.getId();
  const token = ScriptApp.getOAuthToken();

  // IDs de las hojas a incluir en el PDF
  const sheetIds = [
    SHEET_RECEPCIONES, SHEET_CHEQUES, SHEET_TRANSFERENCIAS
  ].map(nombre => {
    const sh = ss.getSheetByName(nombre);
    return sh ? sh.getSheetId() : null;
  }).filter(id => id !== null);

  // Exportar cada hoja como PDF y combinar en un blob
  const blobs = sheetIds.map(gid => {
    const url = "https://docs.google.com/spreadsheets/d/" + ssId +
      "/export?exportFormat=pdf&format=pdf" +
      "&size=A4&portrait=false&fitw=true" +
      "&sheetnames=true&printtitle=false&pagenumbers=true" +
      "&gridlines=false&fzr=false&gid=" + gid;
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });
    return resp.getBlob();
  });

  // Guardar cada hoja como PDF independiente (Drive no soporta merge nativo)
  const nombres = [SHEET_RECEPCIONES, SHEET_CHEQUES, SHEET_TRANSFERENCIAS];
  blobs.forEach((blob, i) => {
    blob.setName(nombreArchivo + "_" + nombres[i] + ".pdf");
    DriveApp.createFile(blob);
  });
}


// ========================= MEJORA 2 – REGISTRAR EN CHEQUES / TRANSFERENCIAS =========================

// Agrega fila a Historial_Pagos (nunca se borra)
function _appendHistorialPago_(participante, datos, fechaPago) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_HISTORIAL_PAGOS);
  if (!sh) return;
  sh.appendRow([
    fechaPago,
    datos.quincena   || "",
    participante     || "",
    datos.creamosID  || "",
    datos.producto   || "",
    datos.ub         || 0,
    datos.tq         || 0,
    datos.metodo     || "",
    datos.referencia || "",
    datos.notas      || ""
  ]);
}

// Agrega en Cheques o Transferencias según tipo, y siempre en Historial_Pagos
function _registrarPagoEnHojas_(participante, datos, fechaPago, tipo) {
  const ss     = SpreadsheetApp.getActive();
  const catAct = ss.getSheetByName(SHEET_CATALOGOS);

  let numCuenta = "", tipoCuenta = "", banco = "", titular = "";
  if (catAct) {
    const dataCat = catAct.getDataRange().getValues();
    for (let j = 1; j < dataCat.length; j++) {
      if (String(dataCat[j][CAT_COL.NOMBRE] || "").trim() === participante) {
        numCuenta  = dataCat[j][CAT_COL.NUM_CUENTA]  || "";
        tipoCuenta = dataCat[j][CAT_COL.TIPO_CUENTA] || "";
        banco      = dataCat[j][CAT_COL.BANCO]        || "";
        titular    = dataCat[j][CAT_COL.TITULAR]      || "";
        break;
      }
    }
  }

  if (tipo === "Cheque entregado") {
    const shCheques = ss.getSheetByName(SHEET_CHEQUES);
    if (shCheques) {
      // Quincena | Participante | Creamos_ID | Nº_Cheque | Banco | Total_Q | Fecha_Entrega | Estado | Notas
      shCheques.appendRow([datos.quincena, participante, datos.creamosID, "", banco, datos.tq, fechaPago, "Entregado", ""]);
    }
  } else {
    const shTrans = ss.getSheetByName(SHEET_TRANSFERENCIAS);
    if (shTrans) {
      // Quincena | Participante | Creamos_ID | Nº_Cuenta | Tipo_Cuenta | Banco | Titular | Total_Q | Fecha_Trans | Comprobante | Estado | Notas
      shTrans.appendRow([datos.quincena, participante, datos.creamosID, numCuenta, tipoCuenta, banco, titular, datos.tq, fechaPago, "", "Ejecutado", ""]);
    }
  }

  _appendHistorialPago_(participante, datos, fechaPago);
}


// ========================= MEJORA 3 – PARTICIPACIÓN EN PROGRAMAS =========================

function actualizarParticipacionProgramas() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const cat = ss.getSheetByName(SHEET_CATALOGOS);
    if (!cat) return SpreadsheetApp.getUi().alert("Instala el sistema primero.");

    // Los programas (I–L) ya están en Participantes Activos.
    // Esta función solo asegura que las filas nuevas tengan checkboxes en I:O.
    const lastRow = cat.getLastRow();
    if (lastRow < 2) return ss.toast("No hay participantes.", null, 3);

    // Insertar checkboxes en columnas I–K (programas) y M–N (CCI, CUNDE)
    cat.getRange(2, 9,  lastRow - 1, 3).insertCheckboxes(); // I–K: Educación, Inclusión, Apoyo
    cat.getRange(2, 13, lastRow - 1, 2).insertCheckboxes(); // M–N: CCI, CUNDE

    ss.toast("✅ Participación al día — programas en columnas I–K de Participantes Activos.", null, 4);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}


// ========================= UTILIDADES =========================

function _getOrCreate_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh   = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _headerMap_(sheet) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map     = {};
  headers.forEach((v, i) => {
    const key = String(v || "").replace(/\n/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
    if (key) map[key] = i + 1;
  });
  return map;
}

function _obtenerHistorico_(tipo) {
  const props = PropertiesService.getScriptProperties();
  const json  = props.getProperty("historico_" + tipo) || "[]";
  return JSON.parse(json);
}

function _guardarHistorico_(tipo, valor) {
  if (!valor || valor.trim() === "") return;
  const props = PropertiesService.getScriptProperties();
  const key   = "historico_" + tipo;
  let lista   = _obtenerHistorico_(tipo);

  if (!lista.includes(valor)) {
    lista.unshift(valor);
    lista = lista.slice(0, 20);
    props.setProperty(key, JSON.stringify(lista));
  }
}
