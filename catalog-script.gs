/**
 * Spares Pro — public catalog & order gateway
 * Paste into: your Spares Pro Google Sheet → Extensions → Apps Script
 * Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone
 *
 * Exposes ONLY: part name, SKU, category, selling price (incl. GST) and an availability word.
 * Never exposes: buying price, supplier, exact quantity, HSN, location, customers, invoices.
 */
var SHOP_NAME = 'AutoParts Spares Pro';
var PARTS_SHEET = 'Parts';
var ORDERS_SHEET = 'OnlineOrders';
var ORDER_HEADER = ['id', 'createdAt', 'name', 'phone', 'note', 'items_json', 'total', 'status'];
var MAX_LINES = 30;            // max different parts per order
var MAX_QTY = 99;              // max quantity per part
var MAX_PER_PHONE_HOUR = 5;    // orders from one phone number per hour
var MAX_ORDERS_HOUR = 60;      // orders from everyone per hour (spam brake)
// Only needed if you created this script at script.google.com instead of from the Sheet's
// Extensions → Apps Script menu. The Spares Pro app fills this in for you when you press "Copy script".
var SHEET_ID = '';

function doGet(e) {
  // Never let an error escape: an uncaught error makes Google send an HTML page, which browsers
  // report as a confusing "CORS" failure. Always answer with JSON that says what is wrong.
  try {
    var action = (e && e.parameter && e.parameter.action) || 'catalog';
    if (action !== 'catalog') return out_({ ok: false, error: 'unknown action' });
    return out_({ ok: true, shop: SHOP_NAME, updated: new Date().toISOString(), items: readCatalog_() });
  } catch (err) {
    return out_({ ok: false, error: 'Script problem: ' + (err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Honeypot: real people never fill this hidden field. Pretend success so bots learn nothing.
    if (body.website) return out_({ ok: true, id: 'W0' });

    var name = clean_(body.name, 60);
    var phone = String(body.phone || '').replace(/[^0-9+]/g, '').slice(0, 16);
    var note = clean_(body.note, 300);
    if (name.length < 2) return out_({ ok: false, error: 'Please enter your name.' });
    if (phone.replace(/\D/g, '').length < 7) return out_({ ok: false, error: 'Please enter a valid phone number.' });

    var lines = (body.items || []).slice(0, MAX_LINES);
    if (!lines.length) return out_({ ok: false, error: 'Your cart is empty.' });

    // Spam brakes (per phone and overall, per hour)
    var cache = CacheService.getScriptCache();
    var bucket = Math.floor(Date.now() / 3600000);
    var kPhone = 'p' + bucket + '_' + phone, kAll = 'a' + bucket;
    var nPhone = Number(cache.get(kPhone) || 0), nAll = Number(cache.get(kAll) || 0);
    if (nPhone >= MAX_PER_PHONE_HOUR || nAll >= MAX_ORDERS_HOUR)
      return out_({ ok: false, error: 'Too many orders just now. Please call or WhatsApp the shop.' });

    // Prices ALWAYS come from the sheet — anything the browser sends about price is ignored.
    var catalog = {};
    readCatalog_().forEach(function (it) { catalog[it.id] = it; });
    var items = [], total = 0;
    lines.forEach(function (l) {
      var it = catalog[String(l.id)];
      var qty = Math.max(1, Math.min(MAX_QTY, Math.floor(Number(l.qty) || 0)));
      if (!it || !qty) return;
      items.push({ id: it.id, sku: it.sku, name: it.name, qty: qty, price: it.price });
      total += it.price * qty;
    });
    if (!items.length) return out_({ ok: false, error: 'Those parts are no longer listed. Please refresh the page.' });

    var sh = ordersSheet_();
    var id = 'W' + Date.now().toString(36).toUpperCase().slice(-6);
    var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    sh.appendRow([id, when, name, phone, note, JSON.stringify(items), Math.round(total * 100) / 100, 'new']);
    cache.put(kPhone, String(nPhone + 1), 3700);
    cache.put(kAll, String(nAll + 1), 3700);
    return out_({ ok: true, id: id, total: Math.round(total * 100) / 100 });
  } catch (err) {
    return out_({ ok: false, error: 'Could not place the order. Please try again or WhatsApp the shop.' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function ss_() {
  var active = SpreadsheetApp.getActive();           // works when the script was opened from the Sheet
  if (active) return active;
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  throw new Error('This script is not attached to your Spares Pro sheet. Open the Sheet → Extensions → Apps Script and paste it there.');
}

function readCatalog_() {
  var sh = ss_().getSheetByName(PARTS_SHEET);
  if (!sh) throw new Error('No "' + PARTS_SHEET + '" tab found in the sheet this script is attached to.');
  if (sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues();
  var H = {};
  v[0].forEach(function (h, i) { H[String(h)] = i; });
  var items = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (!row[H.id]) continue;
    var sell = Number(row[H.sellPrice]) || 0;
    if (sell <= 0) continue;                                   // unpriced parts are not listed
    var gst = (Number(row[H.cgst]) || 0) + (Number(row[H.sgst]) || 0);
    var qty = Number(row[H.qty]) || 0, min = Number(row[H.minQty]) || 0;
    items.push({
      id: String(row[H.id]),
      sku: String(row[H.sku] || ''),
      name: String(row[H.name] || ''),
      cat: String(row[H.category] || 'Other'),
      price: Math.round(sell * (1 + gst / 100) * 100) / 100,   // price the customer actually pays
      st: qty <= 0 ? 'out' : (qty <= min ? 'low' : 'in')
    });
  }
  return items;
}

function ordersSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(ORDERS_SHEET);
  if (!sh) { sh = ss.insertSheet(ORDERS_SHEET); sh.appendRow(ORDER_HEADER); }
  return sh;
}

// Strip characters that could be abused (HTML, spreadsheet formulas) from public text
function clean_(s, max) {
  return String(s || '')
    .replace(/[<>&"'`\\]/g, '')
    .replace(/[\u0000-\u001f]/g, ' ')
    .replace(/^[=+\-@\s]+/, '')
    .trim()
    .slice(0, max);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
