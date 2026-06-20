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

const HEADER_ROW     = 4;
const DATA_START_ROW = 5;

// ── Mejora 1: Quincena insertada entre Fecha entrega y Participante ───────────
const RECEPCIONES_HEADERS = [
  "#", "Fecha entrega", "Quincena", "Participante", "Creamos ID", "Proyecto / Cliente", "Producto",
  "Unidades buenas", "Unidades rechazadas", "Precio unit. (Q)", "Total Q",
  "Estado pago", "Fecha pago", "Método pago", "Comprobante", "Notas / calidad"
];

// Columnas del catálogo de participantes (0-indexed)
const CAT_COL = {
  ID:             0,  // A – Creamos ID
  NOMBRE:         1,  // B – Nombre Completo
  ESTADO:         2,  // C – Estado
  NUM_CUENTA:     3,  // D – Nº Cuenta
  TIPO_CUENTA:    4,  // E – Tipo Cuenta
  BANCO:          5,  // F – Banco
  TITULAR:        6,  // G – Titular
  DPI:            7,  // H – DPI
  MI_EELO:        8,  // I – mi_eelo (checkbox)
  EDUCACION:      9,  // J – Educación (checkbox)
  INCLUSION_LAB: 10,  // K – Inclusión Laboral (checkbox)
  APOYO_EMOC:    11,  // L – Apoyo Emocional (checkbox)
  HIJOS:         12,  // M – N° Hijos
  CCI:           13,  // N – CCI (checkbox)
  CUNDE:         14,  // O – CUNDE (checkbox)
};

// ========================= MENU =========================
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("⚙️ Automatización")
      .addItem("🚀 Instalar sistema",     "instalarSistema")
      .addItem("♻️ Reinstalar sistema",   "reinstalarSistema")
      .addItem("🗑️ Desinstalar sistema",  "desinstalarSistema")
      .addSeparator()
      .addItem("➕ Nueva entrega (rápida)",    "agregarEntregaRapida")
      .addItem("✏️ Editar fila seleccionada",  "editarFilaSeleccionada")
      .addItem("✅ Marcar fila como pagada",   "marcarFilaPagada")
      .addSeparator()
      .addItem("🗓️ Crear quincena inicial",        "crearQuincenaInicial")
      .addItem("📅 Cerrar quincena actual",         "cerrarQuincenaActual")
      .addItem("🗃️ Cerrar mes manualmente",         "cerrarMes")
      .addSeparator()
      .addItem("📋 Actualizar participación",      "actualizarParticipacionProgramas")
      .addItem("🔄 Actualizar todo",               "actualizarTodo")
      .addSeparator()
      .addItem("⏱️ Crear triggers",    "crearTriggers")
      .addItem("⏹️ Eliminar triggers", "eliminarTriggers")
      .addToUi();
  } catch (e) {
    // Silenciosamente ignorar si getUi no está disponible
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

function _crearEstructura_(recrear) {
  const ss = SpreadsheetApp.getActive();
  const hojasSistema = [
    SHEET_RECEPCIONES, SHEET_CATALOGOS, SHEET_INACTIVOS,
    SHEET_RESUMEN_PART, SHEET_HIST_QUINCENAS, SHEET_PAGOS_PEND,
    SHEET_DASHBOARD, SHEET_REPORTES,
    SHEET_PERIODOS, SHEET_CHEQUES, SHEET_TRANSFERENCIAS,
    SHEET_PROGRAMAS, SHEET_HISTORIAL_PAGOS,
    SHEET_ARCHIVO_REC, SHEET_CATALOGO_PROD
  ];

  if (recrear) {
    hojasSistema.forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    });
  }

  const rec  = _getOrCreate_(SHEET_RECEPCIONES);
  const cat  = _getOrCreate_(SHEET_CATALOGOS);
  const ina  = _getOrCreate_(SHEET_INACTIVOS);
  const rp   = _getOrCreate_(SHEET_RESUMEN_PART);
  const hq   = _getOrCreate_(SHEET_HIST_QUINCENAS);
  const pp   = _getOrCreate_(SHEET_PAGOS_PEND);
  const db   = _getOrCreate_(SHEET_DASHBOARD);
  const rep  = _getOrCreate_(SHEET_REPORTES);
  const per  = _getOrCreate_(SHEET_PERIODOS);
  const chq  = _getOrCreate_(SHEET_CHEQUES);
  const tra  = _getOrCreate_(SHEET_TRANSFERENCIAS);
  const prog = _getOrCreate_(SHEET_PROGRAMAS);

  const hist = _getOrCreate_(SHEET_HISTORIAL_PAGOS);
  const arc  = _getOrCreate_(SHEET_ARCHIVO_REC);
  const cat2 = _getOrCreate_(SHEET_CATALOGO_PROD);

  _setupRecepciones_(rec);
  _setupCatalogos_(cat);
  _setupInactivos_(ina);
  _setupResumen_(rp, "RESUMEN POR PARTICIPANTE");
  _setupHistorialQuincenas_(hq);
  _setupPagosPendientes_(pp);
  _setupDashboard_(db);
  _setupReportes_(rep);
  _setupPeriodos_(per);
  _setupCheques_(chq);
  _setupTransferencias_(tra);
  _setupProgramas_(prog);
  _setupHistorialPagos_(hist);
  _setupArchivo_(arc);
  _setupCatalogo_(cat2);

  crearValidacionesDatos();
  crearQuincenaInicial();  // crea la quincena del período actual si no existe
  actualizarTodo();
}


// ========================= SETUP DE HOJAS =========================

function _setupRecepciones_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:P1").merge().setValue("📦 REGISTRO DE RECEPCIONES")
    .setFontWeight("bold").setFontSize(14).setBackground("#263238").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 28);

  sh.getRange(HEADER_ROW, 1, 1, RECEPCIONES_HEADERS.length)
    .setValues([RECEPCIONES_HEADERS])
    .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");

  sh.setFrozenRows(HEADER_ROW);
  sh.setRowHeight(HEADER_ROW, 30);

  // 16 columnas: # | Fecha | Quincena | Participante | CreamosID | Proyecto | Producto |
  //              UBuenas | URechaz | PrecioU | TotalQ | Estado | FechaPago | Metodo | Comprobante | Notas
  const widths = [45, 90, 110, 150, 100, 160, 120, 100, 120, 90, 80, 80, 80, 100, 120, 150];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange(`A${HEADER_ROW + 1}:P1000`).setBackground("#ffffff").setFontColor("#212121");
  sh.getRange(`A${HEADER_ROW + 1}:P1000`).setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupCatalogos_(sh) {
  sh.clear();
  sh.clearFormats();

  // A–H: datos personales/bancarios
  sh.getRange("A1:H1").setValues([["Creamos ID", "Nombre Completo", "Estado", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI"]])
    .setBackground("#263238").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // I–L: programas sociales (antes hoja separada Programas_Participacion)
  sh.getRange("I1:L1").setValues([["mi_eelo", "Educación", "Inclusión Laboral", "Apoyo Emocional"]])
    .setBackground("#2e7d32").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // M–O: datos sociales
  sh.getRange("M1:O1").setValues([["N° Hijos", "CCI", "CUNDE"]])
    .setBackground("#4a148c").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 28);

  sh.setColumnWidth(1,  100); sh.setColumnWidth(2,  180); sh.setColumnWidth(3,  100);
  sh.setColumnWidth(4,  130); sh.setColumnWidth(5,  130); sh.setColumnWidth(6,  120);
  sh.setColumnWidth(7,  150); sh.setColumnWidth(8,  110);
  sh.setColumnWidth(9,   90); sh.setColumnWidth(10,  90); sh.setColumnWidth(11, 140);
  sh.setColumnWidth(12, 140); sh.setColumnWidth(13,  80); sh.setColumnWidth(14,  60);
  sh.setColumnWidth(15,  70);

  sh.getRange("A2:H1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:H1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("I2:L1000").insertCheckboxes();
  sh.getRange("I2:L1000").setBackground("#f1f8e9");
  sh.getRange("I2:L1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("M2:M1000").setBackground("#fce4ec").setFontColor("#212121");
  sh.getRange("M2:M1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("N2:O1000").insertCheckboxes();
  sh.getRange("N2:O1000").setBackground("#fce4ec");
  sh.getRange("N2:O1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("C2:C1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Activo", "Inactivo"])
      .setAllowInvalid(false).build()
  );
  sh.getRange("E2:E1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Monetaria", "Ahorro", "Corriente"])
      .setAllowInvalid(true).build()
  );
}

function _setupInactivos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:H1").setValues([["Creamos ID", "Nombre Completo", "Estado", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI"]])
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

  sh.getRange("A1").setValue("REPORTES PARA POWER BI").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");

  sh.getRange("A2:M2").setValues([[
    "Fecha", "Mes", "Creamos ID", "Participante", "Producto", "Proyecto",
    "Unidades Buenas", "Unidades Rechazadas", "Tasa Rechazo %",
    "Total Q", "Estado Pago", "Método Pago", "Días Pendiente"
  ]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(2);
  sh.setRowHeight(2, 25);

  sh.setColumnWidth(1, 100);  sh.setColumnWidth(2, 100);  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 150);  sh.setColumnWidth(5, 120);  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 120);  sh.setColumnWidth(8, 130);  sh.setColumnWidth(9, 120);
  sh.setColumnWidth(10, 100); sh.setColumnWidth(11, 120); sh.setColumnWidth(12, 120);
  sh.setColumnWidth(13, 120);
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
function _setupCheques_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:I1").setValues([[
    "Quincena", "Participante", "Creamos_ID", "Nº_Cheque",
    "Banco", "Total_Q", "Fecha_Entrega", "Estado", "Notas"
  ]])
    .setBackground("#4a148c").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 140); sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 130); sh.setColumnWidth(8, 110); sh.setColumnWidth(9, 200);

  sh.getRange("F2:F1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("G2:G1000").setNumberFormat("yyyy-mm-dd");
  sh.getRange("A2:I1000").setBackground("#fff8ff").setFontColor("#212121");
  sh.getRange("A2:I1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

// ── Mejora 2: TRANSFERENCIAS ──────────────────────────────────────────────────
function _setupTransferencias_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:L1").setValues([[
    "Quincena", "Participante", "Creamos_ID", "Nº_Cuenta",
    "Tipo_Cuenta", "Banco", "Titular", "Total_Q",
    "Fecha_Transferencia", "Comprobante", "Estado", "Notas"
  ]])
    .setBackground("#006064").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 200);  sh.setColumnWidth(2, 170);  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 140);  sh.setColumnWidth(5, 140);  sh.setColumnWidth(6, 140);
  sh.setColumnWidth(7, 160);  sh.setColumnWidth(8, 110);  sh.setColumnWidth(9, 150);
  sh.setColumnWidth(10, 130); sh.setColumnWidth(11, 120); sh.setColumnWidth(12, 200);

  sh.getRange("H2:H1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("I2:I1000").setNumberFormat("yyyy-mm-dd");
  sh.getRange("A2:L1000").setBackground("#f0fdfd").setFontColor("#212121");
  sh.getRange("A2:L1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
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

// ── Mejora 4: HISTORIAL_PAGOS ─────────────────────────────────────────────────
function _setupHistorialPagos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:J1").setValues([[
    "Fecha_Pago", "Quincena", "Participante", "Creamos_ID",
    "Producto", "Unidades", "Total_Q", "Método_Pago", "Referencia", "Notas"
  ]])
    .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 140); sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 130); sh.setColumnWidth(9, 140);
  sh.setColumnWidth(10, 200);

  sh.getRange("A2:A1000").setNumberFormat("yyyy-mm-dd");
  sh.getRange("G2:G1000").setNumberFormat('"Q " #,##0.00');
  sh.getRange("A2:J1000").setBackground("#fafafa").setFontColor("#212121");
  sh.getRange("A2:J1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
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

    const participantes = cat.getRange("B2:B1000").getValues().flat().filter(x => x);
    const proyectos     = _obtenerHistorico_("proyectos");
    const metodos       = _obtenerHistorico_("metodos");

    // Leer productos y precios del Catálogo_Productos (col B = nombre, col C = precio, col D = activo)
    const productosConPrecio = [];
    const catProd = ss.getSheetByName(SHEET_CATALOGO_PROD);
    if (catProd && catProd.getLastRow() > 1) {
      catProd.getRange(2, 1, catProd.getLastRow() - 1, 3).getValues().forEach(r => {
        const prod   = String(r[1] || "").trim();
        const precio = Number(r[2]) || 0;
        if (prod) productosConPrecio.push({ nombre: prod, precio });
      });
    }

    const productosHistorico = _obtenerHistorico_("productos");

    const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          padding: 20px; min-height: 100vh;
        }
        .container {
          max-width: 500px; margin: 0 auto; background: white;
          border-radius: 12px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); padding: 30px;
        }
        h1 { color: #333; margin-bottom: 8px; font-size: 24px; }
        .subtitle { color: #888; font-size: 13px; margin-bottom: 25px; }
        .form-group { margin-bottom: 18px; }
        label {
          display: block; color: #555; font-weight: 600; margin-bottom: 6px;
          font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px;
        }
        select, input {
          width: 100%; padding: 10px 12px; border: 2px solid #e0e0e0;
          border-radius: 6px; font-size: 14px; transition: all 0.3s; font-family: inherit;
        }
        select:focus, input:focus {
          outline: none; border-color: #667eea;
          box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        select { cursor: pointer; }
        .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .resumen {
          background: #f5f5f5; border-left: 4px solid #667eea;
          padding: 15px; border-radius: 6px; margin: 20px 0;
          font-size: 13px; line-height: 1.6;
        }
        .resumen-row { display: flex; justify-content: space-between; margin-bottom: 6px; }
        .resumen-row:last-child {
          margin-bottom: 0; border-top: 1px solid #ddd;
          padding-top: 8px; margin-top: 8px; font-weight: 600; color: #667eea;
        }
        .total-q { color: #d32f2f; font-weight: bold; }
        .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 25px; }
        button {
          padding: 12px; border: none; border-radius: 6px; font-size: 14px;
          font-weight: 600; cursor: pointer; transition: all 0.3s;
          text-transform: uppercase; letter-spacing: 0.5px;
        }
        .btn-guardar { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
        .btn-guardar:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(102, 126, 234, 0.3); }
        .btn-guardar:active { transform: translateY(0); }
        .btn-cancelar { background: #e0e0e0; color: #666; }
        .btn-cancelar:hover { background: #d0d0d0; }
        .error { color: #d32f2f; font-size: 12px; margin-top: 4px; display: none; }
        .loading { display: none; text-align: center; color: #667eea; margin-top: 15px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>📦 Nueva Entrega</h1>
        <p class="subtitle">Completa los datos de la entrega</p>

        <div class="form-group">
          <label for="participante">Participante *</label>
          <select id="participante" required>
            <option value="">-- Seleccionar participante --</option>
            ${participantes.map(p => `<option value="${p}">${p}</option>`).join('')}
          </select>
          <div class="error" id="err-part">Campo requerido</div>
        </div>

        <div class="form-group">
          <label for="producto">Producto * (escribe o selecciona)</label>
          <input type="text" id="producto" list="productos-list" placeholder="Ej: Pulsera, Mantel..." required onchange="actualizarPrecio()" oninput="actualizarPrecio()">
          <datalist id="productos-list">
            ${productosConPrecio.map(p => `<option value="${p.nombre}" data-precio="${p.precio}">`).join('')}
            ${productosHistorico.map(p => `<option value="${p}">`).join('')}
          </datalist>
          <div class="error" id="err-prod">Campo requerido</div>
        </div>

        <div class="form-group">
          <label for="proyecto">Proyecto / Cliente * (escribe o selecciona)</label>
          <input type="text" id="proyecto" list="proyectos-list" placeholder="Ej: Creamos, Cliente X..." required>
          <datalist id="proyectos-list">
            ${proyectos.map(p => `<option value="${p}">`).join('')}
          </datalist>
          <div class="error" id="err-proy">Campo requerido</div>
        </div>

        <div class="row">
          <div class="form-group">
            <label for="buenas">Unidades Buenas *</label>
            <input type="number" id="buenas" min="0" value="0" required>
            <div class="error" id="err-buenas">Campo requerido</div>
          </div>
          <div class="form-group">
            <label for="rechazadas">Unidades Rechazadas *</label>
            <input type="number" id="rechazadas" min="0" value="0" required>
          </div>
        </div>

        <div class="form-group">
          <label for="precio">Precio Unitario (Q) *</label>
          <input type="number" id="precio" min="0" step="0.01" value="0" required>
          <div class="error" id="err-precio">Campo requerido</div>
        </div>

        <div class="form-group">
          <label for="metodo">Método de Pago (opcional)</label>
          <input type="text" id="metodo" list="metodos-list" placeholder="Ej: Efectivo, Transferencia...">
          <datalist id="metodos-list">
            ${metodos.map(m => `<option value="${m}">`).join('')}
          </datalist>
        </div>

        <div class="resumen" id="resumen" style="display:none;">
          <div class="resumen-row"><span>Unidades Buenas:</span><span id="res-buenas">0</span></div>
          <div class="resumen-row"><span>Precio Unitario:</span><span>Q <span id="res-precio">0.00</span></span></div>
          <div class="resumen-row"><span>Total a Pagar:</span><span class="total-q">Q <span id="res-total">0.00</span></span></div>
        </div>

        <div class="buttons">
          <button class="btn-guardar" onclick="guardarEntrega()">✅ Guardar</button>
          <button class="btn-cancelar" onclick="google.script.host.close()">❌ Cancelar</button>
        </div>
        <div class="loading" id="loading">Guardando entrega...</div>
      </div>

      <script>
        const productosConPrecio = ${JSON.stringify(productosConPrecio)};

        ['buenas', 'precio'].forEach(id => {
          document.getElementById(id).addEventListener('input', actualizarResumen);
        });

        function actualizarPrecio() {
          const productoInput = document.getElementById('producto').value.trim().toLowerCase();
          const precioInput   = document.getElementById('precio');
          if (!productoInput) return;
          const producto = productosConPrecio.find(p => p.nombre.toLowerCase() === productoInput);
          if (producto && producto.precio > 0) {
            precioInput.value = producto.precio;
            actualizarResumen();
          }
        }

        function actualizarResumen() {
          const buenas = Number(document.getElementById('buenas').value) || 0;
          const precio = Number(document.getElementById('precio').value) || 0;
          const total  = buenas * precio;
          document.getElementById('res-buenas').textContent = buenas;
          document.getElementById('res-precio').textContent = precio.toFixed(2);
          document.getElementById('res-total').textContent  = total.toFixed(2);
          if (buenas > 0 || precio > 0) document.getElementById('resumen').style.display = 'block';
        }

        function validar() {
          const participante = document.getElementById('participante').value.trim();
          const producto     = document.getElementById('producto').value.trim();
          const proyecto     = document.getElementById('proyecto').value.trim();
          const buenas       = Number(document.getElementById('buenas').value) || 0;
          const precio       = Number(document.getElementById('precio').value) || 0;

          document.getElementById('err-part').style.display   = !participante ? 'block' : 'none';
          document.getElementById('err-prod').style.display   = !producto     ? 'block' : 'none';
          document.getElementById('err-proy').style.display   = !proyecto     ? 'block' : 'none';
          document.getElementById('err-buenas').style.display = (buenas <= 0) ? 'block' : 'none';
          document.getElementById('err-precio').style.display = (precio <= 0) ? 'block' : 'none';

          return participante && producto && proyecto && buenas > 0 && precio > 0;
        }

        function guardarEntrega() {
          if (!validar()) { alert('⚠️ Por favor completa todos los campos requeridos'); return; }
          const btn = document.querySelector('.btn-guardar');
          btn.disabled = true;
          btn.style.opacity = '0.5';
          document.getElementById('loading').style.display = 'block';

          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(err) {
              alert('Error al guardar: ' + err);
              btn.disabled = false;
              btn.style.opacity = '1';
              document.getElementById('loading').style.display = 'none';
            })
            .guardarEntregaServer(
              document.getElementById('participante').value.trim(),
              document.getElementById('producto').value.trim(),
              document.getElementById('proyecto').value.trim(),
              Number(document.getElementById('buenas').value),
              Number(document.getElementById('rechazadas').value),
              Number(document.getElementById('precio').value),
              document.getElementById('metodo').value.trim()
            );
        }
      </script>
    </body>
    </html>
  `)
    .setWidth(540)
    .setHeight(720);

    SpreadsheetApp.getUi().showModalDialog(html, "Nueva Entrega Rápida");
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

function guardarEntregaServer(participante, producto, proyecto, buenas, rechazadas, precio, metodo) {
  try {
    const ss  = SpreadsheetApp.getActive();
    const sh  = ss.getSheetByName(SHEET_RECEPCIONES);
    const cat = ss.getSheetByName(SHEET_CATALOGOS);
    const m   = _headerMap_(sh);
    const row = Math.max(sh.getLastRow() + 1, DATA_START_ROW);
    const total    = buenas * precio;
    const quincena = _obtenerQuincenaActiva_(); // Mejora 1

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
    sh.getRange(row, m["fecha entrega"]).setValue(new Date());
    sh.getRange(row, m["quincena"]).setValue(quincena);           // Mejora 1
    sh.getRange(row, m["participante"]).setValue(participante || "");
    sh.getRange(row, m["creamos id"]).setValue(creamosID);
    sh.getRange(row, m["proyecto / cliente"]).setValue(proyecto || "");
    sh.getRange(row, m["producto"]).setValue(producto || "");
    sh.getRange(row, m["unidades buenas"]).setValue(buenas || 0);
    sh.getRange(row, m["unidades rechazadas"]).setValue(rechazadas || 0);
    sh.getRange(row, m["precio unit. (q)"]).setValue(precio || 0);
    sh.getRange(row, m["total q"]).setValue(total || 0);
    sh.getRange(row, m["estado pago"]).setValue("Pendiente");
    if (metodo && metodo.length > 0) sh.getRange(row, m["método pago"]).setValue(metodo);

    _guardarHistorico_("productos", producto);
    _guardarHistorico_("proyectos", proyecto);
    if (metodo && metodo.length > 0) _guardarHistorico_("metodos", metodo);

    SpreadsheetApp.getActive().toast("✅ Q " + total.toFixed(2), null, 2);
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
    sh.getRange(row, m["fecha pago"]).setValue(hoy);
    if (!sh.getRange(row, m["método pago"]).getValue()) sh.getRange(row, m["método pago"]).setValue("Efectivo");

    // Mejora 4: registrar en Historial_Pagos
    const rowData    = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const metodo     = String(rowData[m["método pago"]      - 1] || "Efectivo").trim();
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
  actualizarHistorialQuincenas();      // Mejora 1 (reemplaza actualizarResumenMensual en el ciclo)
  actualizarPagosPendientes();
  aplicarColoresAutomaticos();
  actualizarDashboard();
  actualizarReportes();
}

function calcularTotalesColumnas() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RECEPCIONES);
  const m  = _headerMap_(sh);

  for (let r = DATA_START_ROW; r <= sh.getLastRow(); r++) {
    const part = sh.getRange(r, m["participante"]).getValue();
    if (!part) continue;
    const b = Number(sh.getRange(r, m["unidades buenas"]).getValue()) || 0;
    const p = Number(sh.getRange(r, m["precio unit. (q)"]).getValue()) || 0;
    sh.getRange(r, m["total q"]).setValue(b * p);
  }
}

function actualizarEstadoPagosAutomatico() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RECEPCIONES);
  const m  = _headerMap_(sh);

  for (let r = DATA_START_ROW; r <= sh.getLastRow(); r++) {
    const part = sh.getRange(r, m["participante"]).getValue();
    if (!part) continue;
    const total = Number(sh.getRange(r, m["total q"]).getValue()) || 0;
    const fp    = sh.getRange(r, m["fecha pago"]).getValue();
    sh.getRange(r, m["estado pago"]).setValue(fp ? "Pagado" : (total > 0 ? "Pendiente" : "Sin monto"));
  }
}

// ========================= RESUMENES =========================
function actualizarResumenParticipantes() {
  const ss   = SpreadsheetApp.getActive();
  const src  = ss.getSheetByName(SHEET_RECEPCIONES);
  const dst  = ss.getSheetByName(SHEET_RESUMEN_PART);
  const m    = _headerMap_(src);
  const data = src.getDataRange().getValues();

  const acc = {};
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r = data[i];
    const p = r[m["participante"] - 1];
    if (!p) continue;
    if (!acc[p]) acc[p] = { b: 0, re: 0, q: 0, cid: "" };
    acc[p].b  += Number(r[m["unidades buenas"]    - 1]) || 0;
    acc[p].re += Number(r[m["unidades rechazadas"] - 1]) || 0;
    acc[p].q  += Number(r[m["total q"]            - 1]) || 0;
    if (!acc[p].cid) acc[p].cid = r[m["creamos id"] - 1] || "";
  }

  dst.clear();
  dst.getRange("A1").setValue("RESUMEN POR PARTICIPANTE").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");

  dst.getRange("A2:G2").setValues([["Participante","Creamos ID","Unidades buenas","Unidades rechazadas","Tasa rechazo %","Total Q","Estado"]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dst.setRowHeight(2, 25);
  dst.setFrozenRows(2);

  const rows = Object.keys(acc).sort().map(k => {
    const t = acc[k].b + acc[k].re;
    return [k, acc[k].cid, acc[k].b, acc[k].re, t ? acc[k].re / t : 0, acc[k].q, acc[k].b > 0 ? "✓ Activo" : "⚠ Sin movimiento"];
  });

  if (rows.length) {
    dst.getRange(3, 1, rows.length, 7).setValues(rows);
    dst.getRange(3, 1, rows.length, 7)
      .setBackground("#ffffff").setFontColor("#212121")
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    dst.getRange(3, 5, rows.length, 1).setNumberFormat("0.00%");
    dst.getRange(3, 6, rows.length, 1).setNumberFormat('"Q " #,##0.00');
  }

  dst.setColumnWidth(1, 150); dst.setColumnWidth(2, 100); dst.setColumnWidth(3, 120);
  dst.setColumnWidth(4, 130); dst.setColumnWidth(5, 120); dst.setColumnWidth(6, 100);
  dst.setColumnWidth(7, 130);
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
    if (data[i][CAT_COL.ESTADO] === "Inactivo") {
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

    if (estado === "Pagado") {
      sh.getRange(rowNum, 1, 1, ncols).setBackground("#c8e6c9").setFontColor("#1b5e20").setFontWeight("bold");
    } else if (estado === "Pendiente") {
      if (dias > 14) {
        sh.getRange(rowNum, 1, 1, ncols).setBackground("#ffebee").setFontColor("#c62828").setFontWeight("bold");
      } else {
        sh.getRange(rowNum, 1, 1, ncols).setBackground("#fff9c4").setFontColor("#f57f17");
      }
    } else {
      sh.getRange(rowNum, 1, 1, ncols).setBackground("#ffffff").setFontColor("#212121");
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
    const tq     = Number(r[m["total q"]            - 1]) || 0;
    const estado = r[m["estado pago"] - 1];

    stats.totalQ  += tq;
    stats.ubTotal += ub;
    stats.urTotal += ur;

    if (estado === "Pagado")   { stats.pagados++;    stats.montoPagado    += tq; }
    if (estado === "Pendiente"){ stats.pendientes++; stats.montoPendiente += tq; }
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
    const tq          = Number(r[m["total q"]  - 1]) || 0;
    const estado      = String(r[m["estado pago"]  - 1] || "").trim();
    const metodo      = String(r[m["método pago"]  - 1] || "").trim();
    const dias        = isNaN(fe) ? 0 : Math.floor((hoy - fe) / (1000 * 60 * 60 * 24));
    const creamosID   = mapaCreamosID[p] || "";

    reportes.push([
      fe, mesAño, creamosID, p,
      String(r[m["producto"]           - 1] || "").trim(),
      String(r[m["proyecto / cliente"] - 1] || "").trim(),
      ub, ur, tasaRechazo, tq, estado, metodo,
      estado === "Pendiente" ? dias : 0
    ]);
  }

  dst.getRange("A3:M1000").clearContent();

  if (reportes.length) {
    dst.getRange(3, 1, reportes.length, 13).setValues(reportes);
    dst.getRange(3, 1, reportes.length, 13)
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

function crearQuincenaInicial() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const shP = ss.getSheetByName(SHEET_PERIODOS);
    const ui  = SpreadsheetApp.getUi();
    if (!shP) return ui.alert("Instala el sistema primero.");

    // Verificar si ya hay quincena activa
    if (shP.getLastRow() > 1) {
      const data = shP.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][4] || "").trim() === "Activo") {
          return ui.alert("Ya existe una quincena activa: " + data[i][1] + "\n\nUsa 'Cerrar quincena actual' para avanzar al siguiente período.");
        }
      }
    }

    const hoy   = new Date();
    const dia   = hoy.getDate();
    const mes   = hoy.getMonth();
    const anio  = hoy.getFullYear();

    // Primera quincena = días 1-15 | Segunda quincena = 16 al último día del mes
    let inicioD, finD;
    if (dia <= 15) {
      inicioD = new Date(anio, mes, 1);
      finD    = new Date(anio, mes, 15);
    } else {
      inicioD = new Date(anio, mes, 16);
      finD    = new Date(anio, mes + 1, 0); // último día del mes
    }

    const nombre  = _formatNombreQuincena_(inicioD, finD);
    const nextNum = Math.max(shP.getLastRow(), 1);
    shP.appendRow([nextNum, nombre, inicioD, finD, "Activo", "Creada automáticamente"]);
    shP.getRange(shP.getLastRow(), 3, 1, 2).setNumberFormat("yyyy-mm-dd");

    actualizarHistorialQuincenas();
    ss.toast("✅ Quincena creada: " + nombre, null, 5);
  } catch (e) {
    SpreadsheetApp.getActive().toast("❌ Error: " + e.message, null, 3);
  }
}

function cerrarQuincenaActual() {
  try {
    const ss  = SpreadsheetApp.getActive();
    const shP = ss.getSheetByName(SHEET_PERIODOS);
    const rec = ss.getSheetByName(SHEET_RECEPCIONES);
    if (!shP) return SpreadsheetApp.getUi().alert("Hoja PERIODOS no encontrada. Instala el sistema primero.");
    if (!rec)  return SpreadsheetApp.getUi().alert("Hoja Recepciones no encontrada.");

    const arc = _getOrCreate_(SHEET_ARCHIVO_REC);
    if (arc.getLastRow() === 0) _setupArchivo_(arc);

    const dataPer = shP.getDataRange().getValues();
    let filaActiva = -1, nombreQ = "", fechaFinActiva = null;

    for (let i = 1; i < dataPer.length; i++) {
      if (String(dataPer[i][4] || "").trim() === "Activo") {
        filaActiva     = i + 1;
        nombreQ        = String(dataPer[i][1] || "").trim();
        fechaFinActiva = new Date(dataPer[i][3]);
        break;
      }
    }

    if (filaActiva === -1) return SpreadsheetApp.getUi().alert("No hay ninguna quincena activa.");

    const mRec    = _headerMap_(rec);
    const dataRec = rec.getDataRange().getValues();

    // Identificar filas de la quincena actual (y las sin quincena asignada)
    const filasDeEstaQ  = [];
    const filasPendient = [];

    for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
      const r = dataRec[i];
      const p = String(r[mRec["participante"] - 1] || "").trim();
      if (!p) continue;
      const q = String(r[mRec["quincena"] - 1] || "").trim();
      if (q === nombreQ || q === "") {
        filasDeEstaQ.push({ rowIndex: i, data: r });
        if (String(r[mRec["estado pago"] - 1] || "").trim() === "Pendiente") {
          filasPendient.push([...r]);
        }
      }
    }

    const ui = SpreadsheetApp.getUi();

    if (filasPendient.length > 0) {
      const aviso = ui.alert(
        "⚠️ Hay " + filasPendient.length + " pago(s) pendiente(s)",
        "Se arrastrarán a la nueva quincena con la nota '(arrastrado)'.\n¿Continuar con el cierre?",
        ui.ButtonSet.YES_NO
      );
      if (aviso !== ui.Button.YES) return;
    } else {
      const ok = ui.alert(
        "Cerrar quincena",
        "¿Cerrar '" + nombreQ + "' y crear la siguiente (+15 días)?",
        ui.ButtonSet.YES_NO
      );
      if (ok !== ui.Button.YES) return;
    }

    // Archivar todas las filas de esta quincena
    for (const entry of filasDeEstaQ) arc.appendRow(entry.data);

    // Borrar de Recepciones de abajo hacia arriba
    const numBorrar = filasDeEstaQ.map(e => e.rowIndex + 1).sort((a, b) => b - a);
    for (const rn of numBorrar) rec.deleteRow(rn);

    // Marcar quincena como Cerrada
    shP.getRange(filaActiva, 5).setValue("Cerrado");

    // Calcular siguiente quincena
    const nuevaInicio = new Date(fechaFinActiva);
    nuevaInicio.setDate(nuevaInicio.getDate() + 1);
    const nuevaFin = new Date(nuevaInicio);
    nuevaFin.setDate(nuevaFin.getDate() + 14);
    const nombreNueva = _formatNombreQuincena_(nuevaInicio, nuevaFin);

    // ── Detectar fin de mes ────────────────────────────────────────────────────
    // Si la nueva quincena empieza en un mes diferente al que terminó → fin de mes
    const esCambioMes = nuevaInicio.getMonth() !== fechaFinActiva.getMonth()
                     || nuevaInicio.getFullYear() !== fechaFinActiva.getFullYear();

    if (esCambioMes) {
      // Cierre de mes automático: guardar Drive + limpiar Archivo_Recepciones
      _ejecutarCierreMes_(ss, false); // false = no pedir confirmación, ya se confirmó arriba
    }

    // Abrir la nueva quincena
    const nextNum = shP.getLastRow();
    shP.appendRow([nextNum, nombreNueva, nuevaInicio, nuevaFin, "Activo", ""]);
    shP.getRange(shP.getLastRow(), 3, 1, 2).setNumberFormat("yyyy-mm-dd");

    // Arrastrar pendientes a la nueva quincena
    if (filasPendient.length > 0) {
      for (const r of filasPendient) {
        const newRow = [...r];
        newRow[mRec["quincena"] - 1] = nombreNueva;
        const notaActual = String(newRow[mRec["notas / calidad"] - 1] || "").trim();
        newRow[mRec["notas / calidad"] - 1] = notaActual ? notaActual + " (arrastrado)" : "(arrastrado)";
        rec.appendRow(newRow);
      }
      for (let r = DATA_START_ROW; r <= rec.getLastRow(); r++) {
        if (rec.getRange(r, mRec["participante"]).getValue()) {
          rec.getRange(r, mRec["#"]).setValue(r - DATA_START_ROW + 1);
        }
      }
    }

    actualizarHistorialQuincenas();

    const partesMes = esCambioMes ? " 🗓️ Mes cerrado y guardado en Drive." : "";
    const partesPend = filasPendient.length > 0 ? ` ${filasPendient.length} pendiente(s) arrastrado(s).` : "";
    ss.toast(`✅ Quincena cerrada.${partesPend}${partesMes} Nueva: ${nombreNueva}`, null, 8);
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
    byQ[q].tq  += Number(r[mRec["total q"]            - 1]) || 0;
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
  const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
  const sufijo = meses[hoy.getMonth()] + hoy.getFullYear();
  const nombre = "Manufactura_" + NOMBRE_PROGRAMA.replace(/ /g, "_") + "_" + sufijo;

  // Guardar copia en Drive
  DriveApp.getFileById(ss.getId()).makeCopy(nombre);

  // Limpiar Recepciones (las filas ya fueron borradas al cerrar quincena, pero por si acaso)
  const rec = ss.getSheetByName(SHEET_RECEPCIONES);
  if (rec && rec.getLastRow() >= DATA_START_ROW) {
    rec.deleteRows(DATA_START_ROW, rec.getLastRow() - DATA_START_ROW + 1);
  }

  // Limpiar Archivo_Recepciones (ya tiene todo guardado en Drive)
  const arc = ss.getSheetByName(SHEET_ARCHIVO_REC);
  if (arc && arc.getLastRow() > 1) {
    arc.deleteRows(2, arc.getLastRow() - 1);
  }

  // Historial_Pagos, Cheques y Transferencias NUNCA se limpian — son registros permanentes

  if (mostrarToast) {
    ss.toast("✅ Mes cerrado. Copia en Drive: '" + nombre + "'. Hojas limpiadas.", null, 8);
  }
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

    // Insertar checkboxes en columnas I–L y N–O para filas que los necesiten
    cat.getRange(2, 9,  lastRow - 1, 4).insertCheckboxes(); // I–L: programas
    cat.getRange(2, 14, lastRow - 1, 2).insertCheckboxes(); // N–O: CCI, CUNDE

    ss.toast("✅ Participación al día — programas en columnas I–L de Participantes Activos.", null, 4);
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
