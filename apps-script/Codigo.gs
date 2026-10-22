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
 * Las columnas calculadas (totales, reservado, salidas, stock, estado) las escribe este script como
 * números, sin fórmulas; se actualizan solas al llegar un pedido y al editar la hoja.
 */

const HOJA_INVENTARIO = "Inventario";
const HOJA_PEDIDOS = "Pedidos";
const HOJA_DETALLE = "Detalle";
const HOJA_RESENAS = "Reseñas";
const MAX_RESENAS = 12;    // cuántas reseñas muestra la página como máximo
const PRIMERA_FILA = 5;   // los datos empiezan en la fila 5 en todas las hojas
const ULTIMA_FILA = 2000; // filas con la lista Pendiente/Confirmado/Cancelado en Pedidos

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

// Se calcula aquí mismo (no se lee de las columnas de la hoja), así la página recibe números
// correctos aunque alguna celda de la hoja tenga un error.
function leerStock_() {
  const c = calcular_(SpreadsheetApp.getActive());
  const items = {};
  c.inv.forEach(p => {
    items[p.sku] = { precio: p.precio, disponible: Math.max(0, Math.floor(p.disponible)), estado: p.estado, visible: p.visible };
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

  const precios = {};
  leerInventario_(ss).forEach(p => { precios[p.sku] = p.precio; });
  const items = (d.items || [])
    .map(i => ({ sku: String(i.sku || "").trim(), cantidad: Math.floor(Number(i.cantidad) || 0) }))
    .filter(i => i.cantidad > 0 && (i.sku in precios));
  if (!items.length) return { ok: false, error: "Pedido sin productos válidos" };

  const fecha = d.fecha ? new Date(d.fecha) : new Date();
  const notas = [d.cupon ? "Cupón " + d.cupon : "", d.notas || ""].filter(String).join(" · ");

  // Pedidos: A ID · B Fecha · C Cliente · D Teléfono · E Dirección · F Correo · G Descripción
  //          H Cantidad · I Subtotal (los calcula el script) · J Domicilio (lo llenas tú) · K Total · L Estado · M Notas
  pedidos.getRange(fp, 1, 1, 7).setValues([[
    id, fecha, texto_(d.cliente, 150), texto_(d.telefono, 30), texto_(d.direccion, 200),
    texto_(d.correo, 150), texto_(d.descripcion, 1000)
  ]]);
  pedidos.getRange(fp, 12, 1, 2).setValues([["Pendiente", texto_(notas, 300)]]);

  // Detalle: A ID · B SKU · D Cantidad · E Precio (el del momento del pedido, así un cambio de
  //          precio después no altera ventas pasadas). C, F y G los calcula el script.
  const fd = filaLibre_(detalle);
  detalle.getRange(fd, 1, items.length, 5).setValues(items.map(i => [id, i.sku, "", i.cantidad, precios[i.sku] || 0]));

  recalcular_(ss);
  return { ok: true, id };
}

/* ---------- Cálculos (sin fórmulas) ----------
   Las cuentas (totales de pedidos, reservado, salidas, stock y estado) las hace este script y las
   escribe como números. Así no dependen de fórmulas que se puedan borrar o que la configuración
   regional de la hoja no entienda. Se recalculan solas cuando llega un pedido y cuando editas
   Pedidos, Detalle, Movimientos o Inventario (por ejemplo, al pasar un pedido a Confirmado). */

const num_ = v => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
const filasDatos_ = (sh, cols) => {
  const n = sh.getLastRow() - PRIMERA_FILA + 1;
  return n > 0 ? sh.getRange(PRIMERA_FILA, 1, n, cols).getValues() : [];
};

// Inventario: A SKU · B Producto · C Categoría · D Precio · E Stock inicial · K Stock mínimo · M Visible
function leerInventario_(ss) {
  return filasDatos_(ss.getSheetByName(HOJA_INVENTARIO), 13)
    .map((f, i) => ({ fila: i, sku: String(f[0]).trim(), nombre: String(f[1] || ""), precio: num_(f[3]),
                      inicial: num_(f[4]), minimo: num_(f[10]), visible: String(f[12]).trim().toLowerCase() !== "no" }));
}

function calcular_(ss) {
  const inv = leerInventario_(ss);
  const porSku = {};
  inv.forEach(p => { if (p.sku) { p.entradas = 0; p.salidas = 0; p.reservado = 0; porSku[p.sku] = p; } });

  const mov = filasDatos_(ss.getSheetByName("Movimientos"), 4);       // B SKU · D Cantidad
  mov.forEach(f => { const p = porSku[String(f[1]).trim()]; if (p) p.entradas += num_(f[3]); });

  const ped = filasDatos_(ss.getSheetByName(HOJA_PEDIDOS), 12);        // A ID · J Domicilio · L Estado
  const estado = {};
  ped.forEach(f => { const id = String(f[0]).trim(); if (id) estado[id] = String(f[11]).trim(); });

  const det = filasDatos_(ss.getSheetByName(HOJA_DETALLE), 5);         // A ID · B SKU · D Cantidad · E Precio
  const porPedido = {};
  const detalle = det.map(f => {
    const id = String(f[0]).trim(), sku = String(f[1]).trim();
    if (!id && !sku) return ["", "", "", ""];
    const p = porSku[sku], cant = num_(f[3]);
    const precio = (f[4] !== "" && Number.isFinite(Number(f[4]))) ? Number(f[4]) : (p ? p.precio : 0);
    const est = id ? (id in estado ? estado[id] : "Pedido no existe") : "";
    if (p && est === "Confirmado") p.salidas += cant;
    if (p && est === "Pendiente") p.reservado += cant;
    if (id) { const t = porPedido[id] || (porPedido[id] = { cant: 0, sub: 0 }); t.cant += cant; t.sub += cant * precio; }
    return [sku ? (p ? p.nombre : "SKU no existe") : "", precio, cant * precio, est];
  });

  inv.forEach(p => {
    if (!p.sku) return;
    p.actual = p.inicial + p.entradas - p.salidas;
    p.disponible = p.actual - p.reservado;
    p.estado = p.disponible <= 0 ? "Agotado" : p.disponible <= p.minimo ? "Pocas unidades" : "Disponible";
  });

  const pedidos = ped.map(f => {
    const id = String(f[0]).trim();
    if (!id) return ["", "", ""];
    const t = porPedido[id] || { cant: 0, sub: 0 };
    return [t.cant, t.sub, t.sub + num_(f[9])];
  });
  return { inv, detalle, pedidos, ped, movNombres: mov.map(f => { const s = String(f[1]).trim(); return [s ? (porSku[s] ? porSku[s].nombre : "SKU no existe") : ""]; }) };
}

function recalcular_(ss) {
  const c = calcular_(ss);
  const det = ss.getSheetByName(HOJA_DETALLE), ped = ss.getSheetByName(HOJA_PEDIDOS);
  const inv = ss.getSheetByName(HOJA_INVENTARIO), mov = ss.getSheetByName("Movimientos");

  if (c.detalle.length) {
    det.getRange(PRIMERA_FILA, 3, c.detalle.length, 1).setValues(c.detalle.map(r => [r[0]]));
    det.getRange(PRIMERA_FILA, 5, c.detalle.length, 3).setValues(c.detalle.map(r => [r[1] === "" ? "" : r[1], r[2] === "" ? "" : r[2], r[3]]));
  }
  if (c.pedidos.length) {
    ped.getRange(PRIMERA_FILA, 8, c.pedidos.length, 2).setValues(c.pedidos.map(r => [r[0], r[1]]));
    ped.getRange(PRIMERA_FILA, 11, c.pedidos.length, 1).setValues(c.pedidos.map(r => [r[2]]));
  }
  if (c.inv.length) {
    inv.getRange(PRIMERA_FILA, 6, c.inv.length, 5).setValues(c.inv.map(p => p.sku
      ? [p.entradas, p.salidas, p.reservado, p.actual, p.disponible] : ["", "", "", "", ""]));
    inv.getRange(PRIMERA_FILA, 12, c.inv.length, 1).setValues(c.inv.map(p => [p.sku ? p.estado : ""]));
  }
  if (mov && c.movNombres.length) mov.getRange(PRIMERA_FILA, 3, c.movNombres.length, 1).setValues(c.movNombres);

  const res = ss.getSheetByName("Resumen");
  if (res) {
    const cuenta = e => c.ped.filter(f => String(f[0]).trim() && String(f[11]).trim() === e).length;
    const ventas = c.ped.reduce((t, f, i) => t + (String(f[0]).trim() && String(f[11]).trim() === "Confirmado" ? c.pedidos[i][2] : 0), 0);
    res.getRange(3, 2, 6, 1).setValues([[cuenta("Pendiente")], [cuenta("Confirmado")], [cuenta("Cancelado")], [ventas],
      [c.inv.filter(p => p.sku && p.estado === "Agotado").length], [c.inv.filter(p => p.sku && p.estado === "Pocas unidades").length]]);
  }
}

// Se ejecuta solo cada vez que editas la hoja (activador simple de Google).
function onEdit(e) {
  try {
    const hoja = e && e.range && e.range.getSheet().getName();
    if ([HOJA_PEDIDOS, HOJA_DETALLE, HOJA_INVENTARIO, "Movimientos"].indexOf(hoja) < 0) return;
    if (e.range.getRow() < PRIMERA_FILA && e.range.getLastRow() < PRIMERA_FILA) return;
    recalcular_(SpreadsheetApp.getActive());
  } catch (err) { console.error(err); }
}

/**
 * REPARAR LA HOJA: ejecútala UNA vez desde el editor (elige "repararFormulas" arriba y toca Ejecutar).
 * Quita las fórmulas de las columnas calculadas (las que mostraban #ERROR! o quedaron vacías) y escribe
 * los valores correctos. Desde ahí el script mantiene esas columnas al día solo.
 * No cambia los datos que escribiste a mano (IDs, clientes, cantidades, precios guardados, estados,
 * domicilios, stock inicial, stock mínimo).
 */
function repararFormulas() {
  const ss = SpreadsheetApp.getActive();
  const limpiar = (hoja, cols) => {
    const sh = ss.getSheetByName(hoja); if (!sh) return;
    const n = Math.max(sh.getLastRow(), PRIMERA_FILA) - PRIMERA_FILA + 1;
    cols.forEach(c => sh.getRange(PRIMERA_FILA, c, n, 1).clearContent());
  };
  // Detalle: el precio (E) se conserva si es un número guardado; si era fórmula se recalcula.
  const det = ss.getSheetByName(HOJA_DETALLE);
  const nd = Math.max(det.getLastRow(), PRIMERA_FILA) - PRIMERA_FILA + 1;
  const pr = det.getRange(PRIMERA_FILA, 5, nd, 1), pf = pr.getFormulas(), pv = pr.getValues();
  const guardados = pf.map((f, i) => [f[0] ? "" : pv[i][0]]);
  limpiar(HOJA_DETALLE, [3, 5, 6, 7]);
  pr.setValues(guardados);
  limpiar(HOJA_PEDIDOS, [8, 9, 11]);
  limpiar(HOJA_INVENTARIO, [6, 7, 8, 9, 10, 12]);
  limpiar("Movimientos", [3]);
  const res = ss.getSheetByName("Resumen"); if (res) res.getRange(3, 2, 6, 1).clearContent();

  recalcular_(ss);
  // Lista Estado (Pendiente/Confirmado/Cancelado) en la columna de Pedidos.
  ss.getSheetByName(HOJA_PEDIDOS).getRange(PRIMERA_FILA, 12, ULTIMA_FILA - PRIMERA_FILA + 1, 1).setDataValidation(
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

function texto_(v, max) {
  let s = String(v == null ? "" : v).trim().slice(0, max);
  // Evita que un texto que empieza con =, +, - o @ se tome como fórmula.
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
