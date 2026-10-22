/**
 * Wings Zone Anapoima: conexión entre la página y la hoja de inventario.
 *
 * Se pega en la hoja de Google (Extensiones > Apps Script) y se publica como
 * aplicación web. Los pasos están en CONECTAR-INVENTARIO.txt.
 *
 *   GET  ?action=stock  -> unidades disponibles, estado, precio y visibilidad por SKU
 *                          (lo usa la página para mostrar "Agotado", "Quedan N", etc.)
 *   GET  ?action=reviews -> reseñas marcadas "Sí" en la hoja Reseñas (la crea si no existe)
 *   POST action=order   -> registra el pedido como "Pendiente" en Pedidos y Detalle
 *                          (sus unidades quedan reservadas hasta que lo confirmes)
 *
 * Solo escribe en las columnas que se llenan a mano (las azules); las fórmulas no se tocan.
 */

const HOJA_INVENTARIO = "Inventario";
const HOJA_PEDIDOS = "Pedidos";
const HOJA_DETALLE = "Detalle";
const HOJA_RESENAS = "Reseñas";
const MAX_RESENAS = 12;    // cuántas reseñas muestra la página como máximo
const PRIMERA_FILA = 5;   // los datos empiezan en la fila 5 en todas las hojas
const ULTIMA_FILA = 2000; // hasta dónde llegan las fórmulas (repararFormulas las extiende hasta aquí)

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  if (action === "stock") return json_(leerStock_());
  if (action === "reviews") return json_(leerResenas_());
  return json_({ ok: true, mensaje: "Inventario Wings Zone conectado." });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "JSON inválido" });
  }
  if (data.action !== "order") return json_({ ok: false, error: "Acción desconocida" });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);   // evita que dos pedidos al tiempo caigan en la misma fila
  try {
    return json_(registrarPedido_(data));
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Inventario ---------- */

function leerStock_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_INVENTARIO);
  const n = sh.getLastRow() - PRIMERA_FILA + 1;
  const items = {};
  if (n <= 0) return { ok: true, items };
  // A SKU · D Precio · J Disponible · L Estado · M Visible en página
  const filas = sh.getRange(PRIMERA_FILA, 1, n, 13).getValues();
  filas.forEach(f => {
    const sku = String(f[0]).trim();
    if (!sku) return;
    items[sku] = {
      precio: Number(f[3]) || 0,
      disponible: Math.max(0, Math.floor(Number(f[9]) || 0)),
      estado: String(f[11] || ""),
      visible: String(f[12]).trim().toLowerCase() !== "no"
    };
  });
  return { ok: true, items };
}

/* ---------- Reseñas ---------- */

// Columnas: A Nombre · B Estrellas (1 a 5) · C Comentario · D Fecha · E Publicar (Sí/No)
function leerResenas_() {
  const sh = hojaResenas_();
  const n = sh.getLastRow() - PRIMERA_FILA + 1;
  if (n <= 0) return { ok: true, reviews: [] };
  const reviews = sh.getRange(PRIMERA_FILA, 1, n, 5).getValues()
    .filter(f => String(f[4]).trim().toLowerCase().replace("í", "i") === "si" && String(f[2]).trim())
    .map(f => ({
      nombre: String(f[0]).trim(),
      estrellas: Math.min(5, Math.max(1, Math.round(Number(f[1]) || 5))),
      comentario: String(f[2]).trim().slice(0, 400),
      fecha: f[3] instanceof Date ? Utilities.formatDate(f[3], "America/Bogota", "yyyy-MM-dd") : String(f[3] || "")
    }))
    .sort((a, b) => (b.fecha > a.fecha ? 1 : b.fecha < a.fecha ? -1 : 0))
    .slice(0, MAX_RESENAS);
  return { ok: true, reviews };
}

// Si la hoja Reseñas no existe, la crea con sus títulos y la lista Sí/No.
function hojaResenas_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA_RESENAS);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA_RESENAS);
  sh.getRange(1, 1).setValue("Reseñas de clientes").setFontWeight("bold").setFontSize(14);
  sh.getRange(2, 1).setValue("Un renglón por opinión. Solo se muestran en la página las que tengan Sí en Publicar. Usa nombre e inicial del apellido (ej: Laura M.).");
  sh.getRange(4, 1, 1, 5).setValues([["Nombre", "Estrellas (1 a 5)", "Comentario", "Fecha", "Publicar"]])
    .setFontWeight("bold").setBackground("#1A1A1A").setFontColor("#FFFFFF");
  const reglas = SpreadsheetApp.newDataValidation();
  sh.getRange("B5:B500").setDataValidation(reglas.requireValueInList(["5", "4", "3", "2", "1"], true).build());
  sh.getRange("E5:E500").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Sí", "No"], true).build());
  sh.getRange("D5:D500").setNumberFormat("yyyy-mm-dd");
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 480); sh.setColumnWidth(4, 110); sh.setColumnWidth(5, 90);
  sh.setFrozenRows(4);
  return sh;
}

/* ---------- Pedidos ---------- */

function registrarPedido_(d) {
  const ss = SpreadsheetApp.getActive();
  const pedidos = ss.getSheetByName(HOJA_PEDIDOS);
  const detalle = ss.getSheetByName(HOJA_DETALLE);
  const id = String(d.id || "").trim();
  if (!id) return { ok: false, error: "Pedido sin ID" };

  // Si el mismo pedido llega dos veces (reintento del navegador), no se duplica.
  const fp = filaLibre_(pedidos);
  if (fp > PRIMERA_FILA) {
    const ids = pedidos.getRange(PRIMERA_FILA, 1, fp - PRIMERA_FILA, 1).getValues().map(r => String(r[0]));
    if (ids.indexOf(id) >= 0) return { ok: true, duplicado: true };
  }

  const skus = skusValidos_(ss);
  const items = (d.items || [])
    .map(i => ({ sku: String(i.sku || "").trim(), cantidad: Math.floor(Number(i.cantidad) || 0) }))
    .filter(i => i.cantidad > 0 && skus[i.sku]);
  if (!items.length) return { ok: false, error: "Pedido sin productos válidos" };

  const fecha = d.fecha ? new Date(d.fecha) : new Date();
  const notas = [d.cupon ? "Cupón " + d.cupon : "", d.notas || ""].filter(String).join(" · ");

  // Pedidos: A ID · B Fecha · C Cliente · D Teléfono · E Dirección · F Correo · G Descripción
  //          H Cantidad · I Subtotal (fórmulas) · J Domicilio (lo llenas tú) · K Total (fórmula) · L Estado · M Notas
  pedidos.getRange(fp, 1, 1, 7).setValues([[
    id, fecha, texto_(d.cliente, 150), texto_(d.telefono, 30), texto_(d.direccion, 200),
    texto_(d.correo, 150), texto_(d.descripcion, 1000)
  ]]);
  pedidos.getRange(fp, 8, 1, 2).setFormulas([[fPedidoCant_(fp), fPedidoSubtotal_(fp)]]);
  pedidos.getRange(fp, 11).setFormula(fPedidoTotal_(fp));
  pedidos.getRange(fp, 12, 1, 2).setValues([["Pendiente", texto_(notas, 300)]]);

  // Detalle: A ID · B SKU · C Producto (fórmula) · D Cantidad · E Precio (se guarda el del momento
  //          del pedido, para que un cambio de precio después no altere ventas pasadas) · F Subtotal · G Estado
  const precios = preciosInventario_(ss);
  const fd = filaLibre_(detalle);
  items.forEach((it, k) => {
    const r = fd + k;
    detalle.getRange(r, 1, 1, 2).setValues([[id, it.sku]]);
    detalle.getRange(r, 3).setFormula(fDetalleProducto_(r));
    detalle.getRange(r, 4, 1, 2).setValues([[it.cantidad, precios[it.sku] || 0]]);
    detalle.getRange(r, 6, 1, 2).setFormulas([[fDetalleSubtotal_(r), fDetalleEstado_(r)]]);
  });

  return { ok: true, id };
}

/* ---------- Fórmulas ---------- */
// Se escriben desde aquí para no depender de que la hoja las tenga: si alguien borra una fila
// completa (por ejemplo el pedido de ejemplo), los pedidos nuevos siguen calculando solos.

// Separador de argumentos según la configuración regional de la hoja: con "Español (Colombia)"
// Google Sheets puede esperar punto y coma (;) en vez de coma (,). En lugar de suponerlo, se prueba
// una vez con una hoja temporal y se guarda el resultado en las propiedades del script.
function separador_(forzar) {
  const props = PropertiesService.getScriptProperties();
  let sep = forzar ? null : props.getProperty("SEP_FORMULAS");
  if (sep) return sep;
  const ss = SpreadsheetApp.getActive();
  const tmp = ss.insertSheet("tmp_formulas_" + Date.now());
  try {
    const c = tmp.getRange(1, 1);
    c.setFormula("=IF(TRUE,7,1)");
    SpreadsheetApp.flush();
    sep = String(c.getValue()) === "7" ? "," : ";";
  } finally {
    ss.deleteSheet(tmp);
  }
  props.setProperty("SEP_FORMULAS", sep);
  return sep;
}
// Ninguna fórmula de este archivo lleva comas dentro de textos, así que reemplazarlas es seguro.
const F_ = f => (separador_() === "," ? f : f.replace(/,/g, ";"));

const R_ = (hoja, col) => hoja + "!$" + col + "$" + PRIMERA_FILA + ":$" + col + "$" + ULTIMA_FILA;
const fPedidoCant_     = r => F_('=IF(A' + r + '="","",SUMIFS(' + R_("Detalle", "D") + "," + R_("Detalle", "A") + ",A" + r + "))");
const fPedidoSubtotal_ = r => F_('=IF(A' + r + '="","",SUMIFS(' + R_("Detalle", "F") + "," + R_("Detalle", "A") + ",A" + r + "))");
const fPedidoTotal_    = r => F_('=IF(A' + r + '="","",I' + r + "+N(J" + r + "))");
const fDetalleProducto_ = r => F_('=IF(B' + r + '="","",IFERROR(INDEX(' + R_("Inventario", "B") + ",MATCH(B" + r + "," + R_("Inventario", "A") + ',0)),"SKU no existe"))');
const fDetallePrecio_   = r => F_('=IF(B' + r + '="","",IFERROR(INDEX(' + R_("Inventario", "D") + ",MATCH(B" + r + "," + R_("Inventario", "A") + ",0)),0))");
const fDetalleSubtotal_ = r => F_('=IF(OR(B' + r + '="",D' + r + '=""),"",D' + r + "*E" + r + ")");
const fDetalleEstado_   = r => F_('=IF(A' + r + '="","",IFERROR(INDEX(' + R_("Pedidos", "L") + ",MATCH(A" + r + "," + R_("Pedidos", "A") + ',0)),"Pedido no existe"))');

function preciosInventario_(ss) {
  const sh = ss.getSheetByName(HOJA_INVENTARIO);
  const n = sh.getLastRow() - PRIMERA_FILA + 1;
  const out = {};
  if (n > 0) sh.getRange(PRIMERA_FILA, 1, n, 4).getValues().forEach(f => {
    const s = String(f[0]).trim(); if (s) out[s] = Number(f[3]) || 0;
  });
  return out;
}

/**
 * REPARAR FÓRMULAS: ejecútala UNA vez desde el editor (elige "repararFormulas" arriba y toca Ejecutar).
 * - Vuelve a poner las fórmulas de Pedidos (Cantidad, Subtotal, Total) y Detalle (Producto, Subtotal,
 *   Estado) hasta la fila 2000, incluidos los pedidos que llegaron sin ellas.
 * - En Detalle, si la columna Precio está vacía, pone la fórmula del precio actual; si ya tiene un
 *   precio, lo respeta.
 * - Actualiza Inventario y Resumen para que cuenten hasta la fila 2000 (antes llegaban a la 300).
 * No borra ni cambia los datos que escribiste a mano (IDs, clientes, cantidades, estados, domicilios).
 */
function repararFormulas() {
  const ss = SpreadsheetApp.getActive();
  separador_(true);   // vuelve a detectar el separador por si cambió la configuración regional
  const filas = ULTIMA_FILA - PRIMERA_FILA + 1;
  const rango = f => Array.from({ length: filas }, (_, i) => f(PRIMERA_FILA + i));

  const ped = ss.getSheetByName(HOJA_PEDIDOS);
  ped.getRange(PRIMERA_FILA, 8, filas, 2).setFormulas(rango(r => [fPedidoCant_(r), fPedidoSubtotal_(r)]));
  ped.getRange(PRIMERA_FILA, 11, filas, 1).setFormulas(rango(r => [fPedidoTotal_(r)]));

  const det = ss.getSheetByName(HOJA_DETALLE);
  det.getRange(PRIMERA_FILA, 3, filas, 1).setFormulas(rango(r => [fDetalleProducto_(r)]));
  det.getRange(PRIMERA_FILA, 6, filas, 2).setFormulas(rango(r => [fDetalleSubtotal_(r), fDetalleEstado_(r)]));
  // Precio: se respeta el número guardado al momento del pedido; las celdas vacías o con fórmula
  // (incluidas las que hayan quedado en #ERROR!) reciben la fórmula del precio actual.
  const precio = det.getRange(PRIMERA_FILA, 5, filas, 1);
  const pv = precio.getValues(), pf = precio.getFormulas();
  precio.setFormulas(pf.map((f, i) => [(!f[0] && pv[i][0] !== "") ? "" : fDetallePrecio_(PRIMERA_FILA + i)]));
  pf.forEach((f, i) => { if (!f[0] && pv[i][0] !== "") precio.getCell(i + 1, 1).setValue(pv[i][0]); });

  const inv = ss.getSheetByName(HOJA_INVENTARIO);
  const ni = inv.getLastRow() - PRIMERA_FILA + 1;
  if (ni > 0) {
    const skus = inv.getRange(PRIMERA_FILA, 1, ni, 1).getValues();
    skus.forEach((s, i) => {
      if (!String(s[0]).trim()) return;
      const r = PRIMERA_FILA + i;
      inv.getRange(r, 6, 1, 5).setFormulas([[
        "=SUMIFS(" + R_("Movimientos", "D") + "," + R_("Movimientos", "B") + ",A" + r + ")",
        "=SUMIFS(" + R_("Detalle", "D") + "," + R_("Detalle", "B") + ",A" + r + "," + R_("Detalle", "G") + ',"Confirmado")',
        "=SUMIFS(" + R_("Detalle", "D") + "," + R_("Detalle", "B") + ",A" + r + "," + R_("Detalle", "G") + ',"Pendiente")',
        "=E" + r + "+F" + r + "-G" + r,
        "=I" + r + "-H" + r
      ].map(F_)]);
      inv.getRange(r, 12).setFormula(F_('=IF(J' + r + '<=0,"Agotado",IF(J' + r + '<=K' + r + ',"Pocas unidades","Disponible"))'));
    });
  }

  const res = ss.getSheetByName("Resumen");
  if (res) res.getRange(3, 2, 6, 1).setFormulas([
    ['=COUNTIF(' + R_("Pedidos", "L") + ',"Pendiente")'],
    ['=COUNTIF(' + R_("Pedidos", "L") + ',"Confirmado")'],
    ['=COUNTIF(' + R_("Pedidos", "L") + ',"Cancelado")'],
    ['=SUMIFS(' + R_("Pedidos", "K") + "," + R_("Pedidos", "L") + ',"Confirmado")'],
    ['=COUNTIF(' + R_("Inventario", "L") + ',"Agotado")'],
    ['=COUNTIF(' + R_("Inventario", "L") + ',"Pocas unidades")']
  ].map(f => [F_(f[0])]));

  // Lista Estado (Pendiente/Confirmado/Cancelado) en toda la columna de Pedidos.
  ped.getRange(PRIMERA_FILA, 12, filas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Pendiente", "Confirmado", "Cancelado"], true).build());
  SpreadsheetApp.flush();
}

/* ---------- Utilidades ---------- */

// Primera fila con la columna A vacía (las fórmulas de otras columnas no cuentan).
function filaLibre_(sh) {
  const ultima = Math.max(sh.getLastRow(), PRIMERA_FILA);
  const col = sh.getRange(PRIMERA_FILA, 1, ultima - PRIMERA_FILA + 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === "") return PRIMERA_FILA + i;
  return PRIMERA_FILA + col.length;
}

function skusValidos_(ss) {
  const sh = ss.getSheetByName(HOJA_INVENTARIO);
  const n = sh.getLastRow() - PRIMERA_FILA + 1;
  const out = {};
  if (n > 0) sh.getRange(PRIMERA_FILA, 1, n, 1).getValues().forEach(r => { const s = String(r[0]).trim(); if (s) out[s] = true; });
  return out;
}

function texto_(v, max) {
  let s = String(v == null ? "" : v).trim().slice(0, max);
  // Evita que un texto que empieza con =, +, - o @ se tome como fórmula.
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
