// ============================================================
//  SCRIPT PARA: Women Payment 26
//  Pegar en: Extensiones → Apps Script de la hoja Women Payment 26
//  Función: gestiona tabs mensuales y estructura de pagos por transferencia
// ============================================================

const HDRS_TRANSFER = [
  "Nombre","Tipo de Pago","Servicio","Banco","Tipo de Cuenta",
  "Numero de Cta","Cuenta de Pago","Quincena 1","Quincena 2","Total Mes"
];

// ─── Menú ────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💳 Women Payment")
    .addItem("📅 Crear tab del mes actual",  "crearTabMesActual")
    .addItem("🧹 Limpiar tab seleccionado",  "limpiarTabActual")
    .addItem("📊 Agregar fila TOTAL por servicio", "agregarTotalesPorServicio")
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
  sh.getRange(1, 1, 1, HDRS_TRANSFER.length)
    .setValues([HDRS_TRANSFER])
    .setFontWeight("bold")
    .setBackground("#263238")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sh.setFrozenRows(1);

  // Anchos de columna
  [180,110,140,110,110,140,100,90,90,100].forEach((w, i) => sh.setColumnWidth(i+1, w));

  // Formato texto en Numero de Cta; formato Q en montos
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

// ─── Insertar filas TOTAL por cada grupo de Servicio ─────────
function agregarTotalesPorServicio() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  if (sh.getLastRow() < 2) return ss.toast("Sin datos.", null, 3);

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();

  // Agrupar por Servicio (col C = índice 2)
  const grupos = {};
  data.forEach((r, i) => {
    const servicio = String(r[2] || "").trim();
    if (!servicio || servicio === "TOTAL") return;
    if (!grupos[servicio]) grupos[servicio] = [];
    grupos[servicio].push({ row: i + 2, q1: Number(r[7])||0, q2: Number(r[8])||0, total: Number(r[9])||0 });
  });

  // Limpiar filas TOTAL existentes y reescribir al final de cada grupo
  sh.clearContents();
  sh.getRange(1, 1, 1, HDRS_TRANSFER.length).setValues([HDRS_TRANSFER])
    .setFontWeight("bold").setBackground("#263238").setFontColor("#ffffff");

  let fila = 2;
  Object.entries(grupos).forEach(([servicio, filas]) => {
    let sumQ1 = 0, sumQ2 = 0, sumTotal = 0;
    filas.forEach(f => {
      const orig = data[f.row - 2];
      sh.getRange(fila, 1, 1, 10).setValues([orig]);
      sumQ1    += f.q1;
      sumQ2    += f.q2;
      sumTotal += f.total;
      fila++;
    });
    // Fila TOTAL del servicio
    sh.getRange(fila, 1, 1, 10).setValues([[
      "", "", servicio, "", "", "", "TOTAL", sumQ1, sumQ2, sumTotal
    ]]).setFontWeight("bold").setBackground("#cfd8dc");
    sh.getRange(fila, 8, 1, 3).setNumberFormat('"Q "#,##0.00');
    fila++;
  });

  sh.getRange("F2:F" + fila).setNumberFormat("@");
  sh.getRange("H2:J" + fila).setNumberFormat('"Q "#,##0.00');
  ss.toast("✅ Totales por servicio agregados.", null, 4);
}
