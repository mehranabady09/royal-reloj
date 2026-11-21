/**
 * Royal Reloj - reseñas de clientes (backend en Google Sheets)
 *
 * Recibe las reseñas del formulario de la web, las guarda en esta hoja y las
 * devuelve a la web para mostrarlas en el banner "Clientes reales".
 *
 * Para ocultar una reseña: pon NO en la columna "Visible" (o borra la fila).
 * Para marcar una compra comprobada: pon SI en la columna "Verificada".
 */
var SHEET_NAME = 'Reseñas';
var AUTO_PUBLISH = true;   // true: la reseña aparece al instante. false: tú escribes SI en "Visible" para publicarla.
var CLOUDINARY_PREFIX = 'https://res.cloudinary.com/wpgawuak/';
var HEADERS = ['Fecha', 'Nombre', 'Ciudad', 'Reloj', 'Estrellas', 'Opinión', 'Foto', 'WhatsApp del pedido', 'Visible', 'Verificada'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function yes_(v) {
  return String(v == null ? '' : v).trim().toUpperCase().replace('Í', 'I') === 'SI';
}

function clean_(s, max) {
  return String(s == null ? '' : s)
    .replace(/[\u0000-\u001F\u007F<>]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, max);
}

function hasLink_(s) {
  s = String(s);
  return /https?:|www\./i.test(s) ||
         /\b[a-z0-9-]+\.(com|es|net|org|io|ru|cn|xyz|top|shop|site|online|info)\b/.test(s) ||
         /@\w/.test(s);
}

function validate_(d) {
  if (!d || typeof d !== 'object') return { error: 'bad' };
  if (d.hp) return { error: 'bot' };
  if (!d.consent) return { error: 'consent' };
  if (Number(d.elapsed || 0) < 6000) return { error: 'fast' };
  var name = clean_(d.name, 40), city = clean_(d.city, 40), watch = clean_(d.watch, 60), text = clean_(d.text, 600);
  if (name.length < 2 || text.length < 15) return { error: 'short' };
  if (hasLink_(name + ' ' + city + ' ' + watch + ' ' + text)) return { error: 'link' };
  var stars = Math.max(0, Math.min(5, Math.round(Number(d.stars) || 0)));
  var photo = String(d.photo || '');
  if (photo && (photo.indexOf(CLOUDINARY_PREFIX) !== 0 || photo.length > 300 || /\s/.test(photo))) photo = '';
  return { name: name, city: city, watch: watch, text: text, stars: stars, photo: photo, phone: clean_(d.phone, 20) };
}

function rateLimited_(cache) {
  var now = Date.now();
  var km = 'rl_m_' + Math.floor(now / 60000), kh = 'rl_h_' + Math.floor(now / 3600000);
  var m = Number(cache.get(km) || 0), h = Number(cache.get(kh) || 0);
  if (m >= 3 || h >= 20) return true;
  cache.put(km, String(m + 1), 120);
  cache.put(kh, String(h + 1), 7200);
  return false;
}


/* ---------- STOCK (en stock / agotado en la web) ----------
 * El dashboard envía aquí las unidades AAA+ en stock de cada reloj; la web las lee y muestra
 * "En stock" / "Solo quedan N" / "Agotado". La hoja "Stock" es solo un espejo: edita el stock en el dashboard.
 * Antes de publicar, define la contraseña: Configuración del proyecto > Propiedades del script > STOCK_SECRET.
 */
var STOCK_SHEET = 'Stock';
var STOCK_HEADERS = ['Reloj (slug)', 'AAA+ en stock', 'Actualizado'];

function stockSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STOCK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STOCK_SHEET);
    sh.appendRow(STOCK_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, STOCK_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function stockGet_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('stock');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var sh = stockSheet_();
  var last = sh.getLastRow();
  var stock = {}, updated = '';
  if (last > 1) {
    var rows = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      var slug = String(rows[i][0]).trim();
      if (!/^[a-z0-9-]{2,60}$/.test(slug)) continue;
      var n = Number(rows[i][1]);
      if (isNaN(n) || n < 0) continue;
      stock[slug] = Math.min(99, Math.floor(n));
      updated = String(rows[i][2]);
    }
  }
  var body = JSON.stringify({ stock: stock, updated: updated });
  cache.put('stock', body, 30);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function stockPost_(d) {
  var secret = PropertiesService.getScriptProperties().getProperty('STOCK_SECRET');
  if (!secret || String(d.secret || '') !== secret) return out_({ ok: false, error: 'auth' });
  var items = d.stock;
  if (!items || typeof items !== 'object') return out_({ ok: false, error: 'bad' });
  var rows = [], now = new Date().toISOString();
  for (var slug in items) {
    if (!/^[a-z0-9-]{2,60}$/.test(slug)) continue;
    var n = Number(items[slug]);
    if (isNaN(n) || n < 0) continue;
    rows.push([slug, Math.min(99, Math.floor(n)), now]);
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = stockSheet_();
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
    CacheService.getScriptCache().remove('stock');
    return out_({ ok: true, count: rows.length, updated: now });
  } catch (err) {
    return out_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var d = JSON.parse(e.postData.contents);
    if (d && d.action === 'stock') { try { lock.releaseLock(); } catch (e0) {} return stockPost_(d); }
    var v = validate_(d);
    if (v.error) return out_({ ok: false, error: v.error });
    var cache = CacheService.getScriptCache();
    if (rateLimited_(cache)) return out_({ ok: false, error: 'rate' });

    var sh = sheet_();
    var last = sh.getLastRow();
    if (last > 1) {   // same opinion already stored? (double click / copy-paste spam)
      var from = Math.max(2, last - 200 + 1);
      var prev = sh.getRange(from, 6, last - from + 1, 1).getValues();
      for (var i = 0; i < prev.length; i++) if (String(prev[i][0]) === v.text) return out_({ ok: false, error: 'dup' });
    }
    var row = last + 1;
    var date = new Date().toISOString();
    var range = sh.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormat('@');   // plain text: nothing typed by a visitor can run as a formula
    range.setValues([[date, v.name, v.city, v.watch, String(v.stars), v.text, v.photo, v.phone, AUTO_PUBLISH ? 'SI' : 'NO', 'NO']]);
    cache.remove('items');
    return out_({ ok: true, visible: AUTO_PUBLISH, item: { date: date, name: v.name, city: v.city, watch: v.watch, stars: v.stars, text: v.text, photo: v.photo, verified: false } });
  } catch (err) {
    return out_({ ok: false, error: 'server' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'stock') return stockGet_();
  var cache = CacheService.getScriptCache();
  var hit = cache.get('items');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var sh = sheet_();
  var last = sh.getLastRow();
  var items = [];
  if (last > 1) {
    var rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {   // newest first; the visitor's WhatsApp is never sent out
      var r = rows[i];
      if (!yes_(r[8])) continue;
      items.push({
        date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        name: String(r[1]), city: String(r[2]), watch: String(r[3]),
        stars: Number(r[4]) || 0, text: String(r[5]), photo: String(r[6]),
        verified: yes_(r[9])
      });
      if (items.length >= 30) break;
    }
  }
  var body = JSON.stringify({ items: items });
  cache.put('items', body, 60);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
