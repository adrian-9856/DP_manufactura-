// ============================================================
//  SCRIPT PARA: Cheques externo
//  Pegar en: Extensiones → Apps Script de la hoja de Cheques
//  Función: gestiona tabs mensuales para pagos por cheque
// ============================================================

const HDRS_CHEQUES = [
  "Nombre","Tipo de Pago","Servicio","Banco","Tipo de Cuenta",
  "Numero de Cta","Cuenta de Pago","Quincena 1","Quincena 2","Total Mes"
];

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

  [180,110,140,110,110,140,100,90,90,100].forEach((w, i) => sh.setColumnWidth(i+1, w));
  sh.getRange("F2:F1000").setNumberFormat("@");
  sh.getRange("H2:J1000").setNumberFormat('"Q "#,##0.00');
  sh.setRowHeight(1, 28);
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

// ─── Resumen del mes activo: total Q1, Q2 y mes por servicio ─
function verResumenMes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sh.getLastRow() < 2) return ui.alert("Sin datos en este tab.");

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  const grupos = {};
  data.forEach(r => {
    const serv = String(r[2] || "").trim();
    if (!serv || serv === "TOTAL") return;
    if (!grupos[serv]) grupos[serv] = { q1: 0, q2: 0 };
    grupos[serv].q1 += Number(r[7]) || 0;
    grupos[serv].q2 += Number(r[8]) || 0;
  });

  let msg = "Resumen — " + sh.getName() + "\n\n";
  let totalQ1 = 0, totalQ2 = 0;
  Object.entries(grupos).forEach(([serv, t]) => {
    msg += serv + ":  Q1=" + t.q1.toFixed(0) + "  Q2=" + t.q2.toFixed(0) + "\n";
    totalQ1 += t.q1;
    totalQ2 += t.q2;
  });
  msg += "\nTOTAL MES: Q" + (totalQ1 + totalQ2).toFixed(0);

  ui.alert("📊 Resumen cheques", msg, ui.ButtonSet.OK);
}
