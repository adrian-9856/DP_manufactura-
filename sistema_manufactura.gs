/**
 * ============================================================================
 * SISTEMA MANUFACTURA - COMPLETO
 * ============================================================================
 */

const NOMBRE_PROGRAMA = "NOMBRE DEL PROGRAMA";

const SHEET_RECEPCIONES  = "Recepciones";
const SHEET_CATALOGOS    = "Participantes Activos";
const SHEET_INACTIVOS    = "Participantes Inactivos";
const SHEET_RESUMEN_PART = "Resumen Participantes";
const SHEET_RESUMEN_MENS = "Resumen Mensual";
const SHEET_PAGOS_PEND   = "Pagos Pendientes";
const SHEET_DASHBOARD    = "Dashboard";
const SHEET_REPORTES     = "Reportes PowerBI";

const HEADER_ROW   = 4;
const DATA_START_ROW = 5;

const RECEPCIONES_HEADERS = [
  "#", "Fecha entrega", "Participante", "Creamos ID", "Proyecto / Cliente", "Producto",
  "Unidades buenas", "Unidades rechazadas", "Precio unit. (Q)", "Total Q",
  "Estado pago", "Fecha pago", "Método pago", "Comprobante", "Notas / calidad"
];

// Columnas del catálogo de participantes (0-indexed)
const CAT_COL = {
  ID:          0,  // A – Creamos ID
  NOMBRE:      1,  // B – Nombre Completo
  ESTADO:      2,  // C – Estado
  NUM_CUENTA:  3,  // D – Nº Cuenta
  TIPO_CUENTA: 4,  // E – Tipo Cuenta
  BANCO:       5,  // F – Banco
  TITULAR:     6,  // G – Titular
  DPI:         7,  // H – DPI
  EMAIL:       8,  // I – Email
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
      .addItem("🔄 Actualizar todo", "actualizarTodo")
      .addSeparator()
      .addItem("⏱️ Crear triggers",   "crearTriggers")
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
      SHEET_RESUMEN_PART, SHEET_RESUMEN_MENS, SHEET_PAGOS_PEND, SHEET_DASHBOARD, SHEET_REPORTES
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
    SHEET_RESUMEN_PART, SHEET_RESUMEN_MENS, SHEET_PAGOS_PEND, SHEET_DASHBOARD, SHEET_REPORTES
  ];

  if (recrear) {
    hojasSistema.forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) ss.deleteSheet(sh);
    });
  }

  const rec = _getOrCreate_(SHEET_RECEPCIONES);
  const cat = _getOrCreate_(SHEET_CATALOGOS);
  const ina = _getOrCreate_(SHEET_INACTIVOS);
  const rp  = _getOrCreate_(SHEET_RESUMEN_PART);
  const rm  = _getOrCreate_(SHEET_RESUMEN_MENS);
  const pp  = _getOrCreate_(SHEET_PAGOS_PEND);
  const db  = _getOrCreate_(SHEET_DASHBOARD);
  const rep = _getOrCreate_(SHEET_REPORTES);

  _setupRecepciones_(rec);
  _setupCatalogos_(cat);
  _setupInactivos_(ina);
  _setupResumen_(rp, "RESUMEN POR PARTICIPANTE");
  _setupResumen_(rm, "RESUMEN MENSUAL");
  _setupPagosPendientes_(pp);
  _setupDashboard_(db);
  _setupReportes_(rep);

  crearValidacionesDatos();
  actualizarTodo();
}


function _setupRecepciones_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:O1").merge().setValue("📦 REGISTRO DE RECEPCIONES")
    .setFontWeight("bold").setFontSize(14).setBackground("#263238").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 28);

  sh.getRange(HEADER_ROW, 1, 1, RECEPCIONES_HEADERS.length)
    .setValues([RECEPCIONES_HEADERS])
    .setBackground("#37474f").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setWrap(true).setVerticalAlignment("middle");

  sh.setFrozenRows(HEADER_ROW);
  sh.setRowHeight(HEADER_ROW, 30);

  const widths = [45, 90, 150, 100, 160, 120, 100, 120, 90, 80, 80, 80, 100, 120, 150];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  sh.getRange(`A${HEADER_ROW + 1}:O1000`).setBackground("#ffffff").setFontColor("#212121");
  sh.getRange(`A${HEADER_ROW + 1}:O1000`).setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupCatalogos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:I1").setValues([["Creamos ID", "Nombre Completo", "Estado", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI", "Email"]])
    .setBackground("#263238").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("K1:L1").setValues([["Producto", "Precio (Q)"]])
    .setBackground("#1565c0").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 150);
  sh.setColumnWidth(11, 150);
  sh.setColumnWidth(12, 100);

  sh.getRange("A2:I1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("A2:I1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange("K2:L1000").setBackground("#ffffff").setFontColor("#212121");
  sh.getRange("K2:L1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange("L2:L1000").setNumberFormat('"Q " #,##0.00');

  sh.getRange("C2:C1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Activo", "Inactivo"])
      .setAllowInvalid(false)
      .build()
  );

  sh.getRange("E2:E1000").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Monetaria", "Ahorro", "Corriente"])
      .setAllowInvalid(true)
      .build()
  );
}

// FIX: _setupInactivos_ ahora conserva todas las columnas del catálogo activo (A:I)
function _setupInactivos_(sh) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1:I1").setValues([["Creamos ID", "Nombre Completo", "Estado", "Nº Cuenta", "Tipo Cuenta", "Banco", "Titular", "DPI", "Email"]])
    .setBackground("#5d4037").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 25);

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(9, 150);

  sh.getRange("A2:I1000").setBackground("#fafafa").setFontColor("#212121");
  sh.getRange("A2:I1000").setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
}

function _setupResumen_(sh, titulo) {
  sh.clear();
  sh.clearFormats();

  sh.getRange("A1").setValue(titulo).setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");
}

// FIX: encabezado ahora coincide con las 11 columnas que escribe actualizarPagosPendientes
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
  sh.setColumnWidth(2, 180);
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

  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 150);
  sh.setColumnWidth(5, 120);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 120);
  sh.setColumnWidth(8, 130);
  sh.setColumnWidth(9, 120);
  sh.setColumnWidth(10, 100);
  sh.setColumnWidth(11, 120);
  sh.setColumnWidth(12, 120);
  sh.setColumnWidth(13, 120);
}


// ========================= CAPTURA RAPIDA =========================
function agregarEntregaRapida() {
  try {
    const ss = SpreadsheetApp.getActive();
    const cat = ss.getSheetByName(SHEET_CATALOGOS);

    const participantes = cat.getRange("B2:B1000").getValues().flat().filter(x => x);
    const proyectos = _obtenerHistorico_("proyectos");
    const metodos  = _obtenerHistorico_("metodos");

    const dataCat = cat.getDataRange().getValues();
    const productosConPrecio = [];
    for (let i = 1; i < dataCat.length; i++) {
      const prod   = String(dataCat[i][10] || "").trim();
      const precio = Number(dataCat[i][11]) || 0;
      if (prod) productosConPrecio.push({ nombre: prod, precio: precio });
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
          padding: 20px;
          min-height: 100vh;
        }
        .container {
          max-width: 500px;
          margin: 0 auto;
          background: white;
          border-radius: 12px;
          box-shadow: 0 10px 40px rgba(0,0,0,0.2);
          padding: 30px;
        }
        h1 { color: #333; margin-bottom: 8px; font-size: 24px; }
        .subtitle { color: #888; font-size: 13px; margin-bottom: 25px; }
        .form-group { margin-bottom: 18px; }
        label {
          display: block; color: #555; font-weight: 600;
          margin-bottom: 6px; font-size: 13px;
          text-transform: uppercase; letter-spacing: 0.5px;
        }
        select, input {
          width: 100%; padding: 10px 12px;
          border: 2px solid #e0e0e0; border-radius: 6px;
          font-size: 14px; transition: all 0.3s; font-family: inherit;
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
          padding-top: 8px; margin-top: 8px;
          font-weight: 600; color: #667eea;
        }
        .total-q { color: #d32f2f; font-weight: bold; }
        .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 25px; }
        button {
          padding: 12px; border: none; border-radius: 6px;
          font-size: 14px; font-weight: 600; cursor: pointer;
          transition: all 0.3s; text-transform: uppercase; letter-spacing: 0.5px;
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
          <input type="text" id="producto" list="productos-list" placeholder="Ej: Pulsera, Mantel..." required onchange="actualizarPrecio()">
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
          <div class="resumen-row">
            <span>Unidades Buenas:</span>
            <span id="res-buenas">0</span>
          </div>
          <div class="resumen-row">
            <span>Precio Unitario:</span>
            <span>Q <span id="res-precio">0.00</span></span>
          </div>
          <div class="resumen-row">
            <span>Total a Pagar:</span>
            <span class="total-q">Q <span id="res-total">0.00</span></span>
          </div>
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
          const productoInput = document.getElementById('producto').value.trim();
          const precioInput   = document.getElementById('precio');
          if (!productoInput) return;
          const producto = productosConPrecio.find(p => p.nombre === productoInput);
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

          if (buenas > 0 || precio > 0) {
            document.getElementById('resumen').style.display = 'block';
          }
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
          if (!validar()) {
            alert('⚠️ Por favor completa todos los campos requeridos');
            return;
          }
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
    const total = buenas * precio;

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
    sh.getRange(row, m["participante"]).setValue(participante || "");
    sh.getRange(row, m["creamos id"]).setValue(creamosID);
    sh.getRange(row, m["proyecto / cliente"]).setValue(proyecto || "");
    sh.getRange(row, m["producto"]).setValue(producto || "");
    sh.getRange(row, m["unidades buenas"]).setValue(buenas || 0);
    sh.getRange(row, m["unidades rechazadas"]).setValue(rechazadas || 0);
    sh.getRange(row, m["precio unit. (q)"]).setValue(precio || 0);
    sh.getRange(row, m["total q"]).setValue(total || 0);
    sh.getRange(row, m["estado pago"]).setValue("Pendiente");
    if (metodo && metodo.length > 0) {
      sh.getRange(row, m["método pago"]).setValue(metodo);
    }

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

    const m = _headerMap_(sh);
    sh.getRange(row, m["estado pago"]).setValue("Pagado");
    sh.getRange(row, m["fecha pago"]).setValue(new Date());
    if (!sh.getRange(row, m["método pago"]).getValue()) sh.getRange(row, m["método pago"]).setValue("Efectivo");

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
  actualizarResumenMensual();
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
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SHEET_RECEPCIONES);
  const dst = ss.getSheetByName(SHEET_RESUMEN_PART);
  const m   = _headerMap_(src);
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

  dst.setColumnWidth(1, 150);
  dst.setColumnWidth(2, 100);
  dst.setColumnWidth(3, 120);
  dst.setColumnWidth(4, 130);
  dst.setColumnWidth(5, 120);
  dst.setColumnWidth(6, 100);
  dst.setColumnWidth(7, 130);
}

function actualizarResumenMensual() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SHEET_RECEPCIONES);
  const dst = ss.getSheetByName(SHEET_RESUMEN_MENS);
  const m   = _headerMap_(src);
  const data = src.getDataRange().getValues();

  const byMonth = {};
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r = data[i];
    const p = r[m["participante"] - 1];
    if (!p) continue;
    const f = new Date(r[m["fecha entrega"] - 1]);
    if (isNaN(f)) continue;
    const key = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM");

    if (!byMonth[key]) byMonth[key] = { reg: 0, b: 0, re: 0, q: 0, p: {} };

    const ub = Number(r[m["unidades buenas"]    - 1]) || 0;
    const ur = Number(r[m["unidades rechazadas"] - 1]) || 0;
    const tq = Number(r[m["total q"]            - 1]) || 0;

    byMonth[key].reg++;
    byMonth[key].b  += ub;
    byMonth[key].re += ur;
    byMonth[key].q  += tq;
    byMonth[key].p[p] = (byMonth[key].p[p] || 0) + ub;
  }

  dst.clear();
  dst.getRange("A1").setValue("RESUMEN MENSUAL").setFontWeight("bold").setFontSize(12)
    .setBackground("#37474f").setFontColor("white");

  dst.getRange(3, 1, 1, 7).setValues([["Mes","Registros","Unidades buenas","Unidades rechazadas","Tasa rechazo %","Total Q","Top 3 participantes"]])
    .setFontWeight("bold").setBackground("#455a64").setFontColor("white").setHorizontalAlignment("center");
  dst.setRowHeight(3, 25);

  const keys = Object.keys(byMonth).sort();
  const out  = keys.map(k => {
    const d   = byMonth[k];
    const top = Object.entries(d.p).sort((a, b) => b[1] - a[1]).slice(0, 3).map(x => `${x[0]} (${x[1]})`).join(", ");
    return [k, d.reg, d.b, d.re, (d.b + d.re) ? d.re / (d.b + d.re) : 0, d.q, top];
  });

  if (out.length) {
    dst.getRange(4, 1, out.length, 7).setValues(out);
    dst.getRange(4, 1, out.length, 7)
      .setBackground("#ffffff").setFontColor("#212121")
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    dst.getRange(4, 5, out.length, 1).setNumberFormat("0.00%");
    dst.getRange(4, 6, out.length, 1).setNumberFormat('"Q " #,##0.00');
  }

  dst.setColumnWidth(1, 100);
  dst.setColumnWidth(2, 100);
  dst.setColumnWidth(3, 120);
  dst.setColumnWidth(4, 130);
  dst.setColumnWidth(5, 120);
  dst.setColumnWidth(6, 100);
  dst.setColumnWidth(7, 180);
}

function actualizarPagosPendientes() {
  const ss     = SpreadsheetApp.getActive();
  const src    = ss.getSheetByName(SHEET_RECEPCIONES);
  const catAct = ss.getSheetByName(SHEET_CATALOGOS);
  const dst    = ss.getSheetByName(SHEET_PAGOS_PEND);
  const m      = _headerMap_(src);
  const data   = src.getDataRange().getValues();
  const dataCat = catAct.getDataRange().getValues();
  const hoy    = new Date();

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

  dst.setColumnWidth(1, 150);
  dst.setColumnWidth(2, 100);
  dst.setColumnWidth(3, 100);
  dst.setColumnWidth(4, 120);
  dst.setColumnWidth(5, 120);
  dst.setColumnWidth(6, 90);
  dst.setColumnWidth(7, 100);
  dst.setColumnWidth(8, 110);
  dst.setColumnWidth(9, 130);
  dst.setColumnWidth(10, 120);
  dst.setColumnWidth(11, 140);
}

// FIX: ahora copia todas las 9 columnas (incluyendo datos bancarios) al mover a inactivos
function procesarParticipantesInactivos() {
  const ss    = SpreadsheetApp.getActive();
  const shAct = ss.getSheetByName(SHEET_CATALOGOS);
  const shIna = ss.getSheetByName(SHEET_INACTIVOS);
  if (!shAct || !shIna) return;

  const data         = shAct.getDataRange().getValues();
  const rowsAEliminar = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][CAT_COL.ESTADO] === "Inactivo") {
      // Copia las 9 columnas completas: ID, Nombre, Estado, Cuenta, Tipo, Banco, Titular, DPI, Email
      shIna.appendRow(data[i].slice(0, 9));
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

    if (estado === "Pagado") {
      sh.getRange(rowNum, 1, 1, 15).setBackground("#c8e6c9").setFontColor("#1b5e20").setFontWeight("bold");
    } else if (estado === "Pendiente") {
      if (dias > 14) {
        sh.getRange(rowNum, 1, 1, 15).setBackground("#ffebee").setFontColor("#c62828").setFontWeight("bold");
      } else {
        sh.getRange(rowNum, 1, 1, 15).setBackground("#fff9c4").setFontColor("#f57f17");
      }
    } else {
      sh.getRange(rowNum, 1, 1, 15).setBackground("#ffffff").setFontColor("#212121");
    }

    sh.getRange(rowNum, 1, 1, 15)
      .setBorder(true, true, true, true, false, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function actualizarDashboard() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SHEET_RECEPCIONES);
  const dst = ss.getSheetByName(SHEET_DASHBOARD);
  const m   = _headerMap_(src);
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

    if (estado === "Pagado") {
      stats.pagados++;
      stats.montoPagado += tq;
    } else if (estado === "Pendiente") {
      stats.pendientes++;
      stats.montoPendiente += tq;
    }
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
    ["ME.P.01", "Participantes Activos",    stats.participantes.size,                     "personas", "#1976d2"],
    ["ME.P.07", "Clientes Únicos",           stats.clientes.size,                          "clientes", "#388e3c"],
    ["ME.P.08", "Pedidos Completados",       stats.pedidos,                                "pedidos",  "#f57c00"],
    ["ME.P.10", "Ingresos Totales",          "Q " + stats.totalQ.toFixed(2),               "GTQ",      "#d32f2f"],
    ["ME.R.01", "Promedio/Participante",     "Q " + promIngresos.toFixed(2),               "GTQ",      "#7b1fa2"],
    ["CALIDAD", "Tasa de Rechazo",           (tasaRechazo * 100).toFixed(1) + "%",         "%",        "#e64a19"],
    ["PAGOS",   "Pendientes",                "Q " + stats.montoPendiente.toFixed(2),       "GTQ",      "#c62828"]
  ];

  for (const ind of indicadores) {
    dst.getRange(row, 1).setValue(ind[0]).setFontWeight("bold").setFontSize(9).setBackground("#eceff1").setFontColor("#37474f");
    dst.getRange(row, 2).setValue(ind[1]).setFontWeight("bold").setFontSize(10).setBackground("#ffffff");
    dst.getRange(row, 3).setValue(ind[2]).setFontWeight("bold").setFontSize(13).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.getRange(row, 4).setValue(ind[3]).setFontSize(9).setBackground(ind[4]).setFontColor("white").setHorizontalAlignment("center");
    dst.setRowHeight(row, 28);
    row++;
  }

  dst.setColumnWidth(1, 100);
  dst.setColumnWidth(2, 200);
  dst.setColumnWidth(3, 150);
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
    const dataCatAct = catAct.getDataRange().getValues();
    for (let j = 1; j < dataCatAct.length; j++) {
      const nombre = String(dataCatAct[j][CAT_COL.NOMBRE] || "").trim();
      const id     = String(dataCatAct[j][CAT_COL.ID]     || "").trim();
      if (nombre && id) mapaCreamosID[nombre] = id;
    }
  }

  if (catIna) {
    const dataCatIna = catIna.getDataRange().getValues();
    for (let j = 1; j < dataCatIna.length; j++) {
      const nombre = String(dataCatIna[j][CAT_COL.NOMBRE] || "").trim();
      const id     = String(dataCatIna[j][CAT_COL.ID]     || "").trim();
      if (nombre && id) mapaCreamosID[nombre] = id;
    }
  }

  const reportes = [];
  for (let i = DATA_START_ROW - 1; i < data.length; i++) {
    const r = data[i];
    const p = String(r[m["participante"] - 1] || "").trim();
    if (!p) continue;

    const fe         = new Date(r[m["fecha entrega"] - 1]);
    const mesAño     = isNaN(fe) ? "" : Utilities.formatDate(fe, Session.getScriptTimeZone(), "yyyy-MM");
    const ub         = Number(r[m["unidades buenas"]    - 1]) || 0;
    const ur         = Number(r[m["unidades rechazadas"] - 1]) || 0;
    const tasaRechazo = (ub + ur) > 0 ? ur / (ub + ur) : 0;
    const tq         = Number(r[m["total q"]  - 1]) || 0;
    const estado     = String(r[m["estado pago"]  - 1] || "").trim();
    const metodo     = String(r[m["método pago"]  - 1] || "").trim();
    const dias       = isNaN(fe) ? 0 : Math.floor((hoy - fe) / (1000 * 60 * 60 * 24));
    const creamosID  = mapaCreamosID[p] || "";

    reportes.push([
      fe, mesAño, creamosID, p,
      String(r[m["producto"]          - 1] || "").trim(),
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

    const participante  = String(shPago.getRange(rowPagoPend, 1).getValue()).trim();
    const fechaEntrega  = shPago.getRange(rowPagoPend, 3).getValue();

    // Normalizar a medianoche para comparar solo por día (evita fallos por diferencia de horas)
    const diaEntrega = new Date(fechaEntrega);
    diaEntrega.setHours(0, 0, 0, 0);

    const mRec   = _headerMap_(shRec);
    const dataRec = shRec.getDataRange().getValues();
    const hoy    = new Date();

    let encontrado = false;
    for (let i = DATA_START_ROW - 1; i < dataRec.length; i++) {
      const r = dataRec[i];
      const p = String(r[mRec["participante"] - 1] || "").trim();
      const fe = new Date(r[mRec["fecha entrega"] - 1]);

      if (p === participante && !isNaN(fe)) {
        const diaFe = new Date(fe);
        diaFe.setHours(0, 0, 0, 0);

        if (diaFe.getTime() === diaEntrega.getTime()) {
          const rowRec = i + 1;
          shRec.getRange(rowRec, mRec["estado pago"]).setValue("Pagado");
          shRec.getRange(rowRec, mRec["fecha pago"]).setValue(hoy);
          shRec.getRange(rowRec, mRec["método pago"]).setValue(tipo === "Cheque entregado" ? "Cheque" : "Efectivo");
          encontrado = true;
          break;
        }
      }
    }

    shPago.getRange(rowPagoPend, 11).clearContent();

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
