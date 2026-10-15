/**
 * Wings Zone Anapoima: conexión entre la página y la hoja de inventario.
 *
 * Se pega en la hoja de Google (Extensiones > Apps Script) y se publica como
 * aplicación web. Los pasos están en CONECTAR-INVENTARIO.txt.
 *
 *   GET  ?action=stock  -> unidades disponibles, estado, precio y visibilidad por SKU
 *                          (lo usa la página para mostrar "Agotado", "Quedan N", etc.)
 *   POST action=order   -> registra el pedido como "Pendiente" en Pedidos y Detalle
 *                          (sus unidades quedan reservadas hasta que lo confirmes)
 *
 * Solo escribe en las columnas que se llenan a mano (las azules); las fórmulas no se tocan.
 */

const HOJA_INVENTARIO = "Inventario";
const HOJA_PEDIDOS = "Pedidos";
const HOJA_DETALLE = "Detalle";
const PRIMERA_FILA = 5;   // los datos empiezan en la fila 5 en todas las hojas

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  if (action === "stock") return json_(leerStock_());
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
  //          (H e I son fórmulas) · J Domicilio (lo llenas tú) · (K fórmula) · L Estado · M Notas
  pedidos.getRange(fp, 1, 1, 7).setValues([[
    id, fecha, texto_(d.cliente, 150), texto_(d.telefono, 30), texto_(d.direccion, 200),
    texto_(d.correo, 150), texto_(d.descripcion, 1000)
  ]]);
  pedidos.getRange(fp, 12, 1, 2).setValues([["Pendiente", texto_(notas, 300)]]);

  // Detalle: A ID · B SKU · (C fórmula) · D Cantidad · (E, F, G fórmulas)
  const fd = filaLibre_(detalle);
  detalle.getRange(fd, 1, items.length, 2).setValues(items.map(i => [id, i.sku]));
  detalle.getRange(fd, 4, items.length, 1).setValues(items.map(i => [i.cantidad]));
  copiarFormulas_(detalle, fd, items.length, [3, 5, 6, 7]);
  copiarFormulas_(pedidos, fp, 1, [8, 9, 11]);

  return { ok: true, id };
}

/* ---------- Utilidades ---------- */

// Primera fila con la columna A vacía (las fórmulas de otras columnas no cuentan).
function filaLibre_(sh) {
  const ultima = Math.max(sh.getLastRow(), PRIMERA_FILA);
  const col = sh.getRange(PRIMERA_FILA, 1, ultima - PRIMERA_FILA + 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (String(col[i][0]).trim() === "") return PRIMERA_FILA + i;
  return PRIMERA_FILA + col.length;
}

// Las hojas traen fórmulas listas hasta la fila 300. Si un pedido cae más abajo,
// copia las fórmulas de la fila de arriba para que siga calculando solo.
function copiarFormulas_(sh, fila, filas, columnas) {
  if (fila <= PRIMERA_FILA) return;
  columnas.forEach(c => {
    for (let r = fila; r < fila + filas; r++) {
      const celda = sh.getRange(r, c);
      if (celda.getFormula()) continue;
      sh.getRange(r - 1, c).copyTo(celda);
    }
  });
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
