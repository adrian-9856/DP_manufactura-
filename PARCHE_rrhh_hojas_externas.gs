// ============================================================
//  PARCHE — Conexión de sistema_rrhh_mieelo.gs con hojas externas
//  (Women Payment 26 + Cheques externo)
//
//  Instrucciones: copia cada bloque en el lugar indicado de tu
//  archivo sistema_rrhh_mieelo.gs real. NO reemplaces el archivo
//  completo — solo agrega/edita las partes señaladas.
// ============================================================


// ────────────────────────────────────────────────────────────
// BLOQUE 1 — Agregar estas constantes justo DESPUÉS del cierre
// del objeto CFG (después de la línea "};" que cierra CFG).
// ────────────────────────────────────────────────────────────

const SS_ID_WOMEN_PAYMENT = "1e3zlQQ827h_uXGE-0GryvpPs7r7kki0c6wcmc_jzIHk"; // Women Payment 26 (finanzas)
const SS_ID_CHEQUES_EXT   = "14yNGv0ce8heGSeJ5L4EnUQFryJfF1n3NeTT3FUky02M"; // Cheques externo (administración)
const SERVICIO_EXTERNO    = "Costura"; // Servicio que se reporta en Women Payment 26 para mi eelo


// ────────────────────────────────────────────────────────────
// BLOQUE 2 — Dentro de la función registrarPagosQuincena(),
// BUSCA esta línea (dentro del bloque que construye mapaPagoById):
//
//     var servicio = "Textil"; // siempre Textil para el taller
//
// REEMPLÁZALA por:
// ────────────────────────────────────────────────────────────

    var servicio = SERVICIO_EXTERNO; // "Costura" — lo que debe llegar a Women Payment 26


// ────────────────────────────────────────────────────────────
// BLOQUE 3 — Dentro de registrarPagosQuincena(), BUSCA el final
// del bloque "// ── Paso 7: Enrutar pagos ─────" — específicamente
// el cierre de:
//
//     Object.keys(pagosReporte).forEach(function(nombre) {
//        ... (todo el enrutamiento a Cheques/Transferencias) ...
//     });
//
// Justo DESPUÉS de ese "});" (y ANTES del comentario
// "// ── Paso 8: Resumen ───"), AGREGA este bloque nuevo:
// ────────────────────────────────────────────────────────────

  // ── Paso 7b: Enviar a hojas externas (Women Payment 26 + Cheques externo) ──
  try {
    var desgloseExt = {}; // { nombre: { servicio: monto } }
    var infoPartExt = {}; // { nombre: { banco, tipoCuenta, numCuenta, cuentaPago } }

    Object.keys(pagosReporte).forEach(function(nombre) {
      var entradaExt = pagosReporte[nombre];
      var montoExt   = entradaExt.neto;
      if (montoExt <= 0) return;
      var infoExt = _buscarInfoPago(nombre, entradaExt.id);
      if (!infoExt) return;
      var nombreOf = infoExt.nombreOficial || nombre;

      if (!desgloseExt[nombreOf]) desgloseExt[nombreOf] = {};
      desgloseExt[nombreOf][SERVICIO_EXTERNO] = (desgloseExt[nombreOf][SERVICIO_EXTERNO] || 0) + montoExt;

      infoPartExt[nombreOf] = {
        banco:      infoExt.banco,
        tipoCuenta: infoExt.tipoCuenta,
        numCuenta:  infoExt.numCuenta,
        cuentaPago: "mi-eelo"
      };
    });

    if (Object.keys(desgloseExt).length > 0) {
      _enviarPagosExterno_(desgloseExt, infoPartExt, quincenaLabel, mesNombre + " " + anioNum);
    }
  } catch (eExt) {
    SpreadsheetApp.getActive().toast("⚠️ Pagos internos OK. Error hojas externas: " + eExt.message, null, 6);
  }


// ────────────────────────────────────────────────────────────
// BLOQUE 4 — Agrega estas 3 funciones COMPLETAS al final del
// archivo sistema_rrhh_mieelo.gs (después de la última función).
// ────────────────────────────────────────────────────────────

// Crea o reutiliza el tab del mes en el Sheet externo (Women Payment 26 / Cheques externo)
function _obtenerTabMes_(ssExt, mes, esCheque) {
  var HDR_TRANS = ["Nombre","Tipo de Pago","Servicio","Banco","Tipo de Cuenta",
                    "Numero de Cta","Cuenta de Pago","Quincena 1","Quincena 2","Total Mes"];
  var HDR_CHQ   = ["#","Nombre","Nº Cheque","Quincena 1","Quincena 2",
                    "Total Mes","Mes","Programa","Cta de Cheques","Status"];
  var HDR       = esCheque ? HDR_CHQ : HDR_TRANS;
  var COLOR_HDR = esCheque ? "#1a237e" : "#263238";

  var sh = ssExt.getSheetByName(mes);
  if (!sh) {
    sh = ssExt.insertSheet(mes);
    sh.getRange(1, 1, 1, HDR.length).setValues([HDR])
      .setFontWeight("bold").setBackground(COLOR_HDR).setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    sh.setFrozenRows(1);

    if (esCheque) {
      [40,180,110,90,90,100,80,140,120,130].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
      sh.getRange("D2:F1000").setNumberFormat('"Q "#,##0.00');
      sh.getRange("J2:J1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(["Emitido | No liberado","Liberado","Cobrado","Anulado"])
          .setAllowInvalid(false).build()
      );
    } else {
      [180,110,140,110,110,140,100,90,90,100].forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
      sh.getRange("F2:F1000").setNumberFormat("@");
      sh.getRange("H2:J1000").setNumberFormat('"Q "#,##0.00');
    }
  }
  return sh;
}

// Escribe pagos desglosados por Servicio en Women Payment 26 y Cheques externo (tab por mes)
// ordenQ recibe quincenaLabel ("Q1" o "Q2", tal como lo usa registrarPagosQuincena)
function _enviarPagosExterno_(desglose, infoPart, ordenQ, mes) {
  var esQ1 = String(ordenQ) === "Q1";

  var filasTrans = {};   // { servicio: [{nombre, info, monto}] } — TODOS
  var filasChq   = {};   // { servicio: [{nombre, info, monto}] } — solo Cheque

  Object.keys(desglose).forEach(function(nombre) {
    var servicios = desglose[nombre];
    var info  = infoPart[nombre] || {};
    var esChq = String(info.tipoCuenta || "").toLowerCase().trim() === "cheque";
    Object.keys(servicios).forEach(function(servicio) {
      var monto = servicios[servicio];
      if (!filasTrans[servicio]) filasTrans[servicio] = [];
      filasTrans[servicio].push({ nombre: nombre, info: info, monto: monto });
      if (esChq) {
        if (!filasChq[servicio]) filasChq[servicio] = [];
        filasChq[servicio].push({ nombre: nombre, info: info, monto: monto });
      }
    });
  });

  function _bloqueServicio_(sh, servicio, filas, esChequeSheet) {
    var nCols = 10;
    var lastR = sh.getLastRow();
    var datos = lastR > 1 ? sh.getRange(2, 1, lastR - 1, nCols).getValues() : [];

    var COL_NOMBRE   = esChequeSheet ? 1 : 0;
    var COL_SERVICIO = esChequeSheet ? 7 : 2;
    var COL_MARCADOR = esChequeSheet ? 9 : 6;
    var COL_Q1       = esChequeSheet ? 3 : 7;
    var COL_Q2       = esChequeSheet ? 4 : 8;
    var COL_TOT      = esChequeSheet ? 5 : 9;

    filas.forEach(function(f) {
      var nombre = f.nombre, info = f.info, monto = f.monto;
      if (monto <= 0) return;
      var found = false;
      for (var i = 0; i < datos.length; i++) {
        var esTotalRow = String(datos[i][COL_MARCADOR] || "").trim().toUpperCase() === "TOTAL";
        if (esTotalRow) continue;
        if (String(datos[i][COL_NOMBRE]).trim()  === nombre &&
            String(datos[i][COL_SERVICIO]).trim() === servicio) {
          var fila = i + 2;
          var q1 = esQ1 ? monto : (Number(datos[i][COL_Q1]) || 0);
          var q2 = esQ1 ? (Number(datos[i][COL_Q2]) || 0) : monto;
          sh.getRange(fila, COL_Q1 + 1).setValue(q1);
          sh.getRange(fila, COL_Q2 + 1).setValue(q2);
          sh.getRange(fila, COL_TOT + 1).setValue(q1 + q2);
          found = true; break;
        }
      }
      if (!found) {
        var nr = sh.getLastRow() + 1;
        if (!esChequeSheet) {
          sh.getRange(nr, 1, 1, 10).setValues([[
            nombre, "Pago Cuenta", servicio,
            info.banco || "", info.tipoCuenta || "", String(info.numCuenta || ""),
            info.cuentaPago || "mi-eelo",
            esQ1 ? monto : 0, esQ1 ? 0 : monto, monto
          ]]);
          sh.getRange(nr, 6).setNumberFormat("@");
          sh.getRange(nr, 8, 1, 3).setNumberFormat('"Q "#,##0.00');
        } else {
          var contadorChq = datos.filter(function(r) {
            return String(r[COL_SERVICIO]).trim() === servicio &&
                   String(r[COL_MARCADOR]).trim().toUpperCase() !== "TOTAL" && r[1];
          }).length + 1;
          sh.getRange(nr, 1, 1, 10).setValues([[
            contadorChq, nombre, "",
            esQ1 ? monto : 0, esQ1 ? 0 : monto, monto,
            mes, servicio, info.cuentaPago || "mi-eelo", "Emitido | No liberado"
          ]]);
          sh.getRange(nr, 4, 1, 3).setNumberFormat('"Q "#,##0.00');
        }
      }
    });

    var lastR2 = sh.getLastRow();
    var datos2 = lastR2 > 1 ? sh.getRange(2, 1, lastR2 - 1, nCols).getValues() : [];
    var totalQ1 = 0, totalQ2 = 0, totalRowIdx = -1;

    for (var i2 = 0; i2 < datos2.length; i2++) {
      var esTotalRow2 = String(datos2[i2][COL_MARCADOR] || "").trim().toUpperCase() === "TOTAL";
      var sameServ    = String(datos2[i2][COL_SERVICIO]).trim() === servicio;
      if (!sameServ) continue;
      if (esTotalRow2 && !datos2[i2][COL_NOMBRE]) {
        totalRowIdx = i2 + 2;
      } else if (datos2[i2][COL_NOMBRE]) {
        totalQ1 += Number(datos2[i2][COL_Q1]) || 0;
        totalQ2 += Number(datos2[i2][COL_Q2]) || 0;
      }
    }

    if (totalRowIdx > 0) {
      sh.getRange(totalRowIdx, COL_Q1 + 1, 1, 3).setValues([[totalQ1, totalQ2, totalQ1 + totalQ2]]);
    } else {
      var nrT = sh.getLastRow() + 1;
      if (!esChequeSheet) {
        sh.getRange(nrT, 1, 1, 10).setValues([[
          "", "", servicio, "", "", "", "TOTAL",
          totalQ1, totalQ2, totalQ1 + totalQ2
        ]]).setFontWeight("bold").setBackground("#cfd8dc");
        sh.getRange(nrT, 8, 1, 3).setNumberFormat('"Q "#,##0.00');
      } else {
        sh.getRange(nrT, 1, 1, 10).setValues([[
          "", "TOTAL", "", totalQ1, totalQ2, totalQ1 + totalQ2,
          "", servicio, "", ""
        ]]).setFontWeight("bold").setBackground("#c5cae9");
        sh.getRange(nrT, 4, 1, 3).setNumberFormat('"Q "#,##0.00');
      }
    }
  }

  try {
    var ssExt = SpreadsheetApp.openById(SS_ID_WOMEN_PAYMENT);
    var shExt = _obtenerTabMes_(ssExt, mes, false);
    Object.keys(filasTrans).forEach(function(servicio) {
      _bloqueServicio_(shExt, servicio, filasTrans[servicio], false);
    });
  } catch (e) {
    throw new Error("Women Payment 26: " + e.message);
  }

  try {
    var ssChq = SpreadsheetApp.openById(SS_ID_CHEQUES_EXT);
    var shChq = _obtenerTabMes_(ssChq, mes, true);
    Object.keys(filasChq).forEach(function(servicio) {
      _bloqueServicio_(shChq, servicio, filasChq[servicio], true);
    });
  } catch (e) {
    throw new Error("Cheques externo: " + e.message);
  }
}
