// ============================================================
//  SCRIPT PARA: Cheques externo
//  Pegar en: Extensiones → Apps Script de la hoja de Cheques
//  Columnas: #, Nombre, Nº Cheque, Q1, Q2, Total Mes, Mes, Programa, Cta de Cheques, Status
// ============================================================

const HDRS_CHEQUES = [
  "#","Nombre","Nº Cheque","Quincena 1","Quincena 2",
  "Total Mes","Mes","Programa","Cta de Cheques","Status"
];

const STATUS_OPCIONES = ["Emitido | No liberado","Liberado","Cobrado","Anulado"];

// ─── Menú ────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏦 Cheques")
    .addItem("📅 Crear tab del mes actual",  "crearTabMesActual")
    .addItem("🧹 Limpiar tab seleccionado",  "limpiarTabActual")
    .addItem("📊 Ver resumen del mes",       "verResumenMes")
    .addToUi();
}

// ─── Crear tab para el mes actual si no existe ───────────────
function crearTabMesActual() {
  const ss  = SpreadsheetApp.getActive();
  const tz  = Session.getScriptTimeZone();
  const mes = Utilities.formatDate(new Date(), tz, "MMMM yyyy");
  _crearOObtenerTab_(ss, mes);
  ss.toast("✅ Tab '" + mes + "' listo.", null, 4);
}

// ─── Crear o reutilizar tab con nombre = mes ─────────────────
function _crearOObtenerTab_(ss, mes) {
  let sh = ss.getSheetByName(mes);
  if (sh) return sh;

  sh = ss.insertSheet(mes);
  sh.getRange(1, 1, 1, HDRS_CHEQUES.length)
    .setValues([HDRS_CHEQUES])
    .setFontWeight("bold")
    .setBackground("#1a237e")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sh.setFrozenRows(1);

  [40,180,110,90,90,100,80,140,120,130].forEach((w, i) => sh.setColumnWidth(i+1, w));
  sh.getRange("D2:F1000").setNumberFormat('"Q "#,##0.00');
  sh.setRowHeight(1, 28);

  // Dropdown Status (col J)
  sh.getRange("J2:J1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_OPCIONES)
      .setAllowInvalid(false).build()
  );

  return sh;
}

// ─── Limpiar datos del tab activo (mantiene encabezado) ──────
function limpiarTabActual() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const ok = ui.alert(
    "¿Limpiar tab?",
    "Se borrarán todos los datos de '" + sh.getName() + "'. ¿Continuar?",
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }
  ss.toast("✅ Tab limpiado.", null, 3);
}

// ─── Resumen del mes activo: total Q1, Q2 por programa ───────
function verResumenMes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sh.getLastRow() < 2) return ui.alert("Sin datos en este tab.");

  // Col H (índice 7) = Programa, Col D (índice 3) = Q1, Col E (índice 4) = Q2
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  const grupos = {};
  data.forEach(r => {
    const nombre = String(r[1] || "").trim();
    if (!nombre || nombre.toUpperCase() === "TOTAL") return;
    const prog = String(r[7] || "Sin programa").trim();
    if (!grupos[prog]) grupos[prog] = { q1: 0, q2: 0 };
    grupos[prog].q1 += Number(r[3]) || 0;
    grupos[prog].q2 += Number(r[4]) || 0;
  });

  let msg = "Resumen cheques — " + sh.getName() + "\n\n";
  let totalQ1 = 0, totalQ2 = 0;
  Object.entries(grupos).forEach(([prog, t]) => {
    msg += prog + ":  Q1=Q" + t.q1.toFixed(0) + "  Q2=Q" + t.q2.toFixed(0) + "\n";
    totalQ1 += t.q1;
    totalQ2 += t.q2;
  });
  msg += "\nTOTAL MES: Q" + (totalQ1 + totalQ2).toFixed(0);

  ui.alert("📊 Resumen cheques", msg, ui.ButtonSet.OK);
}
