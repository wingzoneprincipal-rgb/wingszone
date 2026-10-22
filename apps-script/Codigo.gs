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
const HOJA_INFORME = "Informe";
const CORREO_INFORMES = "wingzoneprincipal@gmail.com";   // a dónde llegan los informes automáticos
const TZ = "America/Bogota";
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
    if (hoja === HOJA_INFORME) {   // cambiar Periodo, Desde o Hasta actualiza el informe
      if (e.range.getColumn() <= 2 && e.range.getLastColumn() >= 2 && e.range.getRow() <= 4 && e.range.getLastRow() >= 2) generarInforme();
      return;
    }
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

/* ---------- Informe de ventas ----------
   Hoja "Informe": eliges el periodo arriba (B2, y Desde/Hasta en B3:B4 si es Personalizado) y el
   informe se actualiza solo. También desde el menú "Wing Zone" de la hoja. Solo cuenta como venta
   lo que está en Confirmado. El menú permite activar el envío automático por correo (semanal los
   lunes y/o mensual el día 1) a CORREO_INFORMES. */

const PERIODOS = ["Hoy", "Ayer", "Últimos 7 días", "Esta semana", "Últimos 30 días", "Este mes",
                  "Mes anterior", "Este año", "Personalizado"];
const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Wing Zone")
    .addItem("Actualizar informe", "generarInforme")
    .addItem("Enviar informe por correo ahora", "enviarInformeAhora")
    .addSeparator()
    .addItem("Activar informe semanal por correo (lunes)", "activarInformeSemanal")
    .addItem("Activar informe mensual por correo (día 1)", "activarInformeMensual")
    .addItem("Desactivar informes por correo", "desactivarInformes")
    .addSeparator()
    .addItem("Reparar hoja (cálculos)", "repararFormulas")
    .addToUi();
}

// --- Fechas como texto "aaaa-mm-dd" en hora de Colombia (evita líos de zona horaria) ---
const clave_ = d => Utilities.formatDate(d, TZ, "yyyy-MM-dd");
const sumarDias_ = (k, n) => { const [y, m, d] = k.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
const diaSemana_ = k => { const [y, m, d] = k.split("-").map(Number); return (new Date(Date.UTC(y, m - 1, d)).getUTCDay() + 6) % 7; }; // 0 = lunes
const diasEntre_ = (a, b) => Math.round((Date.parse(b + "T00:00:00Z") - Date.parse(a + "T00:00:00Z")) / 86400000);
const fechaDe_ = v => { if (v instanceof Date) return v; const d = new Date(v); return isNaN(d) ? null : d; };
const fmtFecha_ = k => k.split("-").reverse().join("/");

function rangoPeriodo_(sel, desde, hasta) {
  const hoy = clave_(new Date());
  const [y, m] = hoy.split("-");
  const primeroMes = y + "-" + m + "-01";
  switch (sel) {
    case "Hoy": return [hoy, hoy];
    case "Ayer": return [sumarDias_(hoy, -1), sumarDias_(hoy, -1)];
    case "Últimos 7 días": return [sumarDias_(hoy, -6), hoy];
    case "Esta semana": return [sumarDias_(hoy, -diaSemana_(hoy)), hoy];
    case "Últimos 30 días": return [sumarDias_(hoy, -29), hoy];
    case "Mes anterior": { const fin = sumarDias_(primeroMes, -1); return [fin.slice(0, 8) + "01", fin]; }
    case "Este año": return [y + "-01-01", hoy];
    case "Personalizado": {
      const a = fechaDe_(desde), b = fechaDe_(hasta);
      if (a && b) { const ka = clave_(a), kb = clave_(b); return ka <= kb ? [ka, kb] : [kb, ka]; }
      return [primeroMes, hoy];
    }
    default: return [primeroMes, hoy];   // "Este mes"
  }
}

// Calcula todo el informe para un rango [desde, hasta] (claves aaaa-mm-dd, inclusivo).
function datosInforme_(ss, desde, hasta) {
  const inv = {};
  leerInventario_(ss).forEach(p => { if (p.sku) inv[p.sku] = p; });
  const categoria = {};
  filasDatos_(ss.getSheetByName(HOJA_INVENTARIO), 3).forEach(f => { categoria[String(f[0]).trim()] = String(f[2] || "Otros"); });

  const ped = filasDatos_(ss.getSheetByName(HOJA_PEDIDOS), 13);
  const det = filasDatos_(ss.getSheetByName(HOJA_DETALLE), 5);
  const pedidos = {};        // id -> info del pedido
  const historico = {};      // cliente -> nº de pedidos confirmados en toda la historia
  ped.forEach(f => {
    const id = String(f[0]).trim(); if (!id) return;
    const fecha = fechaDe_(f[1]);
    const cliente = String(f[3]).replace(/\D/g, "") || String(f[5]).trim().toLowerCase() || String(f[2]).trim().toLowerCase();
    const o = { id, fecha, clave: fecha ? clave_(fecha) : "", hora: fecha ? Number(Utilities.formatDate(fecha, TZ, "H")) : null,
      nombre: String(f[2]).trim(), cliente, desc: String(f[6] || ""), productos: num_(f[8]), domicilio: num_(f[9]),
      total: num_(f[10]), estado: String(f[11]).trim(), notas: String(f[12] || "") };
    pedidos[id] = o;
    if (o.estado === "Confirmado" && cliente) historico[cliente] = (historico[cliente] || 0) + 1;
  });

  const enRango = (o, a, b) => o.clave && o.clave >= a && o.clave <= b;
  const resumen = (a, b) => {
    const conf = Object.values(pedidos).filter(o => enRango(o, a, b) && o.estado === "Confirmado");
    const canc = Object.values(pedidos).filter(o => enRango(o, a, b) && o.estado === "Cancelado");
    const ventas = conf.reduce((t, o) => t + o.total, 0);
    return { pedidos: conf.length, ventas, productos: conf.reduce((t, o) => t + o.productos, 0),
             domicilios: conf.reduce((t, o) => t + o.domicilio, 0), ticket: conf.length ? ventas / conf.length : 0,
             cancelados: canc.length, valorCancelado: canc.reduce((t, o) => t + o.total, 0), lista: conf };
  };
  const dias = diasEntre_(desde, hasta) + 1;
  const actual = resumen(desde, hasta);
  const anterior = resumen(sumarDias_(desde, -dias), sumarDias_(desde, -1));
  const conf = actual.lista;
  const ids = {}; conf.forEach(o => { ids[o.id] = true; });

  // Ventas en el tiempo: por día (hasta 62 días) o por mes.
  const porMes = dias > 62;
  const tiempo = {};
  if (!porMes) for (let k = desde; k <= hasta; k = sumarDias_(k, 1)) tiempo[k] = { pedidos: 0, ventas: 0 };
  conf.forEach(o => { const k = porMes ? o.clave.slice(0, 7) : o.clave; const t = tiempo[k] || (tiempo[k] = { pedidos: 0, ventas: 0 }); t.pedidos++; t.ventas += o.total; });
  const serieTiempo = Object.keys(tiempo).sort().map(k => [porMes ? k : fmtFecha_(k), tiempo[k].pedidos, tiempo[k].ventas]);

  const semana = DIAS_SEMANA.map(d => [d, 0, 0]);
  const horas = {};
  conf.forEach(o => {
    const w = semana[diaSemana_(o.clave)]; w[1]++; w[2] += o.total;
    if (o.hora !== null) { const h = horas[o.hora] || (horas[o.hora] = [0, 0]); h[0]++; h[1] += o.total; }
  });
  const serieHoras = Object.keys(horas).map(Number).sort((a, b) => a - b).map(h => [h + ":00", horas[h][0], horas[h][1]]);

  // Productos y categorías (desde Detalle: cantidad × precio guardado).
  const prod = {}, cat = {};
  det.forEach(f => {
    const id = String(f[0]).trim(), sku = String(f[1]).trim();
    if (!ids[id] || !sku) return;
    const cant = num_(f[3]), monto = cant * num_(f[4]);
    const p = prod[sku] || (prod[sku] = { nombre: inv[sku] ? inv[sku].nombre : sku, cant: 0, monto: 0 });
    p.cant += cant; p.monto += monto;
    const c = categoria[sku] || "Otros", x = cat[c] || (cat[c] = { cant: 0, monto: 0 }); x.cant += cant; x.monto += monto;
  });
  const totalProd = Object.values(prod).reduce((t, p) => t + p.monto, 0) || 1;
  const ranking = Object.keys(prod).map(k => prod[k]).sort((a, b) => b.cant - a.cant || b.monto - a.monto)
    .map((p, i) => [i + 1, p.nombre, p.cant, p.monto, p.monto / totalProd]);
  const sinVentas = Object.keys(inv).filter(k => !prod[k] && inv[k].visible).map(k => [inv[k].nombre]);
  const categorias = Object.keys(cat).sort((a, b) => cat[b].monto - cat[a].monto).map(c => [c, cat[c].cant, cat[c].monto, cat[c].monto / totalProd]);

  // Salsas y acompañamientos: salen de la descripción, ej. "1 x Combo x12 alitas [BBQ clásica · Papas Cronch + Plátano chips]".
  const salsas = {}, acomp = {};
  conf.forEach(o => {
    const re = /(\d+)\s*x\s*[^;\[]*\[([^\]]+)\]/g; let m;
    while ((m = re.exec(o.desc))) {
      const n = Number(m[1]) || 1, partes = m[2].split("·").map(t => t.trim());
      if (partes[0]) partes[0].split("+").forEach(t => { t = t.trim(); if (t) salsas[t] = (salsas[t] || 0) + n; });
      if (partes[1]) partes[1].split("+").forEach(t => { t = t.trim(); if (t) acomp[t] = (acomp[t] || 0) + n; });
    }
  });
  const tablaConteo = obj => { const tot = Object.values(obj).reduce((a, b) => a + b, 0) || 1;
    return Object.keys(obj).sort((a, b) => obj[b] - obj[a]).map(k => [k, obj[k], obj[k] / tot]); };

  // Medios de pago, cupones, clientes y promociones (desde la columna Notas).
  const pagos = {}, cupones = {};
  let aceptanPromos = 0;
  conf.forEach(o => {
    const mp = /Pago:\s*([^·(]+)/.exec(o.notas); const pago = mp ? mp[1].trim() : "Sin dato";
    const p = pagos[pago] || (pagos[pago] = [0, 0]); p[0]++; p[1] += o.total;
    const mc = /Cup[oó]n\s+([A-Z0-9]+)/i.exec(o.notas);
    if (mc) { const c = cupones[mc[1].toUpperCase()] || (cupones[mc[1].toUpperCase()] = [0, 0]); c[0]++; c[1] += o.total; }
    if (/Promociones:\s*S[ií]/.test(o.notas)) aceptanPromos++;
  });
  const clientes = {};
  conf.forEach(o => { if (!o.cliente) return; const c = clientes[o.cliente] || (clientes[o.cliente] = { nombre: o.nombre, pedidos: 0, total: 0 }); c.pedidos++; c.total += o.total; });
  const unicos = Object.keys(clientes).length;
  const recurrentes = Object.keys(clientes).filter(k => (historico[k] || 0) > 1).length;
  const topClientes = Object.keys(clientes).map(k => clientes[k]).sort((a, b) => b.total - a.total).slice(0, 10)
    .map(c => [c.nombre || "(sin nombre)", c.pedidos, c.total]);

  return { desde, hasta, dias, porMes, actual, anterior, serieTiempo, semana, serieHoras, ranking, sinVentas, categorias,
           salsas: tablaConteo(salsas), acomp: tablaConteo(acomp),
           pagos: Object.keys(pagos).sort((a, b) => pagos[b][1] - pagos[a][1]).map(k => [k, pagos[k][0], pagos[k][1]]),
           cupones: Object.keys(cupones).map(k => [k, cupones[k][0], cupones[k][1]]),
           unicos, recurrentes, topClientes, aceptanPromos };
}

const variacion_ = (a, b) => (b ? (a - b) / b : "—");   // sin datos del periodo anterior no hay comparación

function hojaInforme_(ss) {
  let sh = ss.getSheetByName(HOJA_INFORME);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA_INFORME);
  sh.getRange("A1").setValue("Informe de ventas · Wing Zone").setFontSize(16).setFontWeight("bold");
  sh.getRange("A2:A4").setValues([["Periodo"], ["Desde"], ["Hasta"]]).setFontWeight("bold");
  sh.getRange("B2").setValue("Este mes").setBackground("#FFF2CC").setFontColor("#0000FF")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PERIODOS, true).build());
  sh.getRange("B3:B4").setBackground("#FFF2CC").setFontColor("#0000FF").setNumberFormat("dd/mm/yyyy")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  sh.getRange("C3").setValue("← solo si eliges Personalizado").setFontColor("#5E5A56");
  sh.setColumnWidth(1, 210); sh.setColumnWidths(2, 4, 120); sh.setFrozenRows(5);
  return sh;
}

function generarInforme() {
  const ss = SpreadsheetApp.getActive();
  const sh = hojaInforme_(ss);
  const sel = String(sh.getRange("B2").getValue() || "Este mes");
  const [desde, hasta] = rangoPeriodo_(sel, sh.getRange("B3").getValue(), sh.getRange("B4").getValue());
  const d = datosInforme_(ss, desde, hasta);

  sh.getCharts().forEach(c => sh.removeChart(c));
  const ultima = Math.max(sh.getLastRow(), 6);
  sh.getRange(6, 1, ultima - 5, Math.max(sh.getLastColumn(), 6)).clear();
  sh.getRange("A5").setValue("Periodo: " + fmtFecha_(desde) + " a " + fmtFecha_(hasta) + " · Actualizado " +
    Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm")).setFontColor("#5E5A56");

  const MONEDA = "$#,##0", PCT = "0%", ENT = "0";
  let fila = 7;
  const bloque = (titulo, encabezados, filas, formatos, vacio) => {
    sh.getRange(fila, 1).setValue(titulo).setFontWeight("bold").setFontSize(12).setFontColor("#D62828");
    fila++;
    if (!filas.length) { sh.getRange(fila, 1).setValue(vacio || "Sin datos en este periodo.").setFontColor("#5E5A56"); fila += 2; return null; }
    sh.getRange(fila, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight("bold").setBackground("#1A1A1A").setFontColor("#FFFFFF");
    const r = sh.getRange(fila + 1, 1, filas.length, encabezados.length).setValues(filas);
    (formatos || []).forEach((f, j) => { if (f) sh.getRange(fila + 1, j + 1, filas.length, 1).setNumberFormat(f); });
    const rango = sh.getRange(fila, 1, filas.length + 1, encabezados.length);
    fila += filas.length + 2;
    return rango;
  };
  let libreGrafico = 7;   // los gráficos van a la derecha (columna G), uno debajo de otro sin encimarse
  const grafico = (tipo, rangos, titulo, filaAncla, opciones) => {
    if (!rangos) return;
    const ancla = Math.max(filaAncla, libreGrafico);
    libreGrafico = ancla + 16;
    let b = sh.newChart().setChartType(tipo).setPosition(ancla, 7, 0, 0).setOption("title", titulo)
      .setOption("legend", { position: "none" }).setOption("width", 520).setOption("height", 300);
    rangos.forEach(r => { b = b.addRange(r); });
    Object.keys(opciones || {}).forEach(k => { b = b.setOption(k, opciones[k]); });
    sh.insertChart(b.build());
  };

  const a = d.actual, p = d.anterior;
  bloque("Resumen del periodo", ["Indicador", "Este periodo", "Periodo anterior", "Variación"], [
    ["Ventas totales (con domicilio)", a.ventas, p.ventas, variacion_(a.ventas, p.ventas)],
    ["Ventas de productos", a.productos, p.productos, variacion_(a.productos, p.productos)],
    ["Domicilios cobrados", a.domicilios, p.domicilios, variacion_(a.domicilios, p.domicilios)],
    ["Pedidos confirmados", a.pedidos, p.pedidos, variacion_(a.pedidos, p.pedidos)],
    ["Ticket promedio", a.ticket, p.ticket, variacion_(a.ticket, p.ticket)],
    ["Pedidos cancelados", a.cancelados, p.cancelados, variacion_(a.cancelados, p.cancelados)],
    ["Valor de pedidos cancelados", a.valorCancelado, p.valorCancelado, variacion_(a.valorCancelado, p.valorCancelado)]
  ], [null, MONEDA, MONEDA, "+0%;-0%;0%"]);
  [3, 5].forEach(k => sh.getRange(fila - 8 + k, 2, 1, 2).setNumberFormat(ENT)); // filas de pedidos y cancelados: cantidades, no pesos

  let f0 = fila;
  const rt = bloque(d.porMes ? "Ventas por mes" : "Ventas por día", [d.porMes ? "Mes" : "Día", "Pedidos", "Ventas"], d.serieTiempo, [null, ENT, MONEDA]);
  if (rt) grafico(Charts.ChartType.COLUMN, [rt.offset(0, 0, rt.getNumRows(), 1), rt.offset(0, 2, rt.getNumRows(), 1)], d.porMes ? "Ventas por mes" : "Ventas por día", f0, { vAxis: { format: "$#,##0" } });

  f0 = fila;
  const rr = bloque("Productos más vendidos", ["#", "Producto", "Unidades", "Ventas", "% de ventas"], d.ranking, [ENT, null, ENT, MONEDA, PCT]);
  if (rr) grafico(Charts.ChartType.BAR, [rr.offset(0, 1, Math.min(rr.getNumRows(), 11), 1), rr.offset(0, 2, Math.min(rr.getNumRows(), 11), 1)], "Top productos (unidades)", f0);
  bloque("Ventas por categoría", ["Categoría", "Unidades", "Ventas", "% de ventas"], d.categorias, [null, ENT, MONEDA, PCT]);
  bloque("Salsas más pedidas en combos", ["Salsa", "Combos", "%"], d.salsas, [null, ENT, PCT]);
  bloque("Acompañamientos más pedidos", ["Acompañamiento", "Veces", "%"], d.acomp, [null, ENT, PCT]);
  bloque("Productos sin ventas en el periodo", ["Producto"], d.sinVentas, null, "Todos los productos visibles tuvieron ventas.");

  bloque("Ventas por día de la semana", ["Día", "Pedidos", "Ventas"], d.semana, [null, ENT, MONEDA]);
  f0 = fila;
  const rh = bloque("Pedidos por hora", ["Hora", "Pedidos", "Ventas"], d.serieHoras, [null, ENT, MONEDA]);
  if (rh) grafico(Charts.ChartType.COLUMN, [rh.offset(0, 0, rh.getNumRows(), 2)], "¿A qué hora piden?", f0);

  f0 = fila;
  const rp = bloque("Medios de pago", ["Medio", "Pedidos", "Ventas"], d.pagos, [null, ENT, MONEDA]);
  if (rp) grafico(Charts.ChartType.PIE, [rp.offset(0, 0, rp.getNumRows(), 1), rp.offset(0, 2, rp.getNumRows(), 1)], "Ventas por medio de pago", f0, { legend: { position: "right" } });
  bloque("Cupones usados", ["Cupón", "Pedidos", "Ventas con el cupón"], d.cupones, [null, ENT, MONEDA], "No se usaron cupones en este periodo.");
  bloque("Clientes", ["Indicador", "Valor"], [
    ["Clientes que compraron", d.unicos],
    ["De ellos, ya habían comprado antes o repitieron", d.recurrentes],
    ["Pedidos que aceptaron recibir promociones", d.aceptanPromos]], [null, ENT]);
  bloque("Mejores clientes del periodo", ["Cliente", "Pedidos", "Total comprado"], d.topClientes, [null, ENT, MONEDA]);
  SpreadsheetApp.flush();
  return d;
}

/* ---------- Informe por correo ---------- */

function enviarInformeAhora() {
  const sh = hojaInforme_(SpreadsheetApp.getActive());
  const sel = String(sh.getRange("B2").getValue() || "Este mes");
  const [desde, hasta] = rangoPeriodo_(sel, sh.getRange("B3").getValue(), sh.getRange("B4").getValue());
  enviarInforme_(desde, hasta, "Informe " + sel.toLowerCase());
  try { SpreadsheetApp.getUi().alert("Informe enviado a " + CORREO_INFORMES); } catch (e) {}
}

// Lunes: semana anterior (lunes a domingo).
function informeSemanalAutomatico() {
  const hoy = clave_(new Date());
  const lunes = sumarDias_(hoy, -diaSemana_(hoy) - 7);
  enviarInforme_(lunes, sumarDias_(lunes, 6), "Informe semanal");
}

// Día 1: mes anterior completo.
function informeMensualAutomatico() {
  const [desde, hasta] = rangoPeriodo_("Mes anterior");
  enviarInforme_(desde, hasta, "Informe mensual");
}

function enviarInforme_(desde, hasta, titulo) {
  const d = datosInforme_(SpreadsheetApp.getActive(), desde, hasta);
  const $ = n => "$" + Math.round(n).toLocaleString("es-CO");
  const pct = n => typeof n !== "number" ? "—" : (n >= 0 ? "+" : "") + Math.round(n * 100) + "%";
  const esc = t => String(t).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const tabla = (enc, filas) => !filas.length ? "<p style='color:#5E5A56'>Sin datos.</p>" :
    "<table cellpadding='6' style='border-collapse:collapse;font-size:14px'><tr>" + enc.map(h => "<th style='background:#1A1A1A;color:#fff;text-align:left'>" + h + "</th>").join("") + "</tr>" +
    filas.map(r => "<tr>" + r.map(c => "<td style='border-bottom:1px solid #E1DFDB'>" + esc(c) + "</td>").join("") + "</tr>").join("") + "</table>";
  const a = d.actual, p = d.anterior;
  const html =
    "<div style='font-family:Arial,sans-serif;color:#1A1A1A'>" +
    "<h2 style='color:#D62828;margin:0'>🔥 " + esc(titulo) + " · Wing Zone</h2>" +
    "<p style='color:#5E5A56'>" + fmtFecha_(desde) + " a " + fmtFecha_(hasta) + "</p>" +
    tabla(["Indicador", "Este periodo", "Anterior", "Variación"], [
      ["Ventas totales", $(a.ventas), $(p.ventas), pct(variacion_(a.ventas, p.ventas))],
      ["Pedidos confirmados", a.pedidos, p.pedidos, pct(variacion_(a.pedidos, p.pedidos))],
      ["Ticket promedio", $(a.ticket), $(p.ticket), pct(variacion_(a.ticket, p.ticket))],
      ["Domicilios cobrados", $(a.domicilios), $(p.domicilios), pct(variacion_(a.domicilios, p.domicilios))],
      ["Pedidos cancelados", a.cancelados, p.cancelados, ""]]) +
    "<h3>🍗 Top 5 productos</h3>" + tabla(["Producto", "Unidades", "Ventas"], d.ranking.slice(0, 5).map(r => [r[1], r[2], $(r[3])])) +
    "<h3>🌶️ Salsas más pedidas</h3>" + tabla(["Salsa", "Combos"], d.salsas.slice(0, 5).map(r => [r[0], r[1]])) +
    "<h3>💳 Medios de pago</h3>" + tabla(["Medio", "Pedidos", "Ventas"], d.pagos.map(r => [r[0], r[1], $(r[2])])) +
    "<h3>📅 Mejor día de la semana</h3>" + tabla(["Día", "Pedidos", "Ventas"], d.semana.filter(r => r[1]).sort((x, y) => y[2] - x[2]).slice(0, 3).map(r => [r[0], r[1], $(r[2])])) +
    "<p style='color:#5E5A56;font-size:13px'>Informe completo con gráficos: hoja <b>Informe</b> de tu Google Sheets (menú Wing Zone → Actualizar informe).</p></div>";
  MailApp.sendEmail({ to: CORREO_INFORMES, subject: "🔥 " + titulo + " Wing Zone · " + fmtFecha_(desde) + " a " + fmtFecha_(hasta) + " · " + $(a.ventas), htmlBody: html });
}

function programar_(funcion, crear) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === funcion).forEach(t => ScriptApp.deleteTrigger(t));
  if (crear) crear(ScriptApp.newTrigger(funcion).timeBased()).atHour(7).inTimezone(TZ).create();
}
function activarInformeSemanal() {
  programar_("informeSemanalAutomatico", t => t.onWeekDay(ScriptApp.WeekDay.MONDAY));
  try { SpreadsheetApp.getUi().alert("Listo: cada lunes a las 7 a. m. llegará el informe de la semana anterior a " + CORREO_INFORMES); } catch (e) {}
}
function activarInformeMensual() {
  programar_("informeMensualAutomatico", t => t.onMonthDay(1));
  try { SpreadsheetApp.getUi().alert("Listo: el día 1 de cada mes a las 7 a. m. llegará el informe del mes anterior a " + CORREO_INFORMES); } catch (e) {}
}
function desactivarInformes() {
  programar_("informeSemanalAutomatico"); programar_("informeMensualAutomatico");
  try { SpreadsheetApp.getUi().alert("Informes automáticos por correo desactivados."); } catch (e) {}
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
