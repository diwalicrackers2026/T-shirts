/**
 * Farside Paris backend: Google Apps Script + Google Sheets.
 * Serves the storefront (index.html): products, reviews, OTP sign-in,
 * orders (COD + Razorpay), order history.
 * Setup: see the steps in the chat reply. Run setup() once.
 */

const SHEETS = {
  Products: ['id', 'name', 'category', 'gender', 'price', 'mrp', 'gsm', 'fit', 'colors', 'sizes', 'description', 'images', 'status', 'stock', 'chart'],
  Orders: ['orderId', 'date', 'identity', 'email', 'name', 'phone', 'address', 'pincode', 'items', 'summary', 'subtotal', 'shipping', 'total', 'payment', 'paymentId', 'razorpayOrderId', 'status', 'discount', 'coupon', 'tracking'],
  Reviews: ['id', 'productId', 'identity', 'name', 'rating', 'text', 'verified', 'date'],
  Sessions: ['hash', 'identity', 'expires'],
  Coupons: ['code', 'type', 'value', 'minOrder', 'status'],
  SizeCharts: ['chart', 'table'],
  Delivery: ['pincode', 'days', 'cod', 'status']
};

/* ---------- one-time setup ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid() + Utilities.getUuid());
  Object.keys(SHEETS).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[n]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, SHEETS[n].length).setFontWeight('bold');
    }
  });
  const p = ss.getSheetByName('Products');
  if (p.getLastRow() === 1) {
    [
      ['t1', 'Heavy loose tee', 'Loose tees', 'men', 799, 1199, 240, 'Loose fit, dropped shoulder', 'Black', 'M,L,XL,XXL', 'Thick, soft cotton that holds its shape wash after wash.', '', 'active'],
      ['t2', 'Cropped black tee', 'Crop tops', 'women', 599, 899, 220, 'Regular crop', 'Black', 'XS,S,M,L', 'A clean cropped cut in dense black cotton.', '', 'active']
    ].forEach(r => p.appendRow(r));
  }
  Object.keys(SHEETS).forEach(n => {   // older sheet: add any new columns at the end
    const sh = ss.getSheetByName(n), have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    SHEETS[n].forEach(k => { if (have.indexOf(k) < 0) sh.getRange(1, sh.getLastColumn() + 1).setValue(k).setFontWeight('bold'); });
  });
  const cp = ss.getSheetByName('Coupons');
  if (cp.getLastRow() === 1) cp.appendRow(['WELCOME10', 'percent', 10, 499, 'active']);
  const dl = ss.getSheetByName('Delivery'); dl.getRange('A:D').setNumberFormat('@');
  if (dl.getLastRow() === 1) dl.appendRow(['*', '4-8', 'yes', 'active']);   // * = every other pincode
  const sc = ss.getSheetByName('SizeCharts');
  if (sc.getLastRow() === 1) {   // sample numbers: replace with your real measurements
    sc.appendRow(['Loose tees', 'Size, Chest (in), Length (in), Shoulder (in)\nM, 42, 28, 20\nL, 44, 29, 21\nXL, 46, 30, 22\nXXL, 48, 31, 23']);
    sc.appendRow(['Crop tops', 'Size, Chest (in), Length (in)\nXS, 32, 17\nS, 34, 18\nM, 36, 19\nL, 38, 20']);
  }
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  ss.getSheetByName('Sessions').hideSheet();
  ['Orders', 'Reviews', 'Sessions'].forEach(n => ss.getSheetByName(n).getRange('A:Z').setNumberFormat('@'));
  ss.getSheetByName('Orders').getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  ss.getSheetByName('Reviews').getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
  Logger.log('Setup done. Now add Script Properties (OWNER_EMAIL, ADMIN_KEY etc.) and deploy as a Web app.');
}

/* ---------- plumbing ---------- */
class ApiError extends Error { constructor(m, code) { super(m); this.code = code; } }
const prop = (k, d) => PropertiesService.getScriptProperties().getProperty(k) || d;
const hex = b => b.map(x => (x < 0 ? x + 256 : x).toString(16).padStart(2, '0')).join('');
const sha = s => hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
const salt = () => prop('SALT', 'change-me');
const split = (v, nl) => String(v == null ? '' : v).split(nl ? /[\n,]+/ : /,/).map(x => x.trim()).filter(Boolean);
const safe = v => (typeof v === 'string' && /^[=+\-@\t\r]/.test(v)) ? "'" + v : v;   // blocks sheet formula injection
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function fail(err) {
  if (err instanceof ApiError) return out({ ok: false, error: err.message, code: err.code });
  console.error(err && err.stack || err);
  return out({ ok: false, error: 'Something went wrong. Try again.' });
}
const sheet = n => SpreadsheetApp.openById(prop('SS_ID')).getSheetByName(n);
function rows(n) {
  const v = sheet(n).getDataRange().getValues(); const h = v.shift();
  return v.map(r => { const o = {}; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add(n, o) { sheet(n).appendRow(SHEETS[n].map(k => safe(o[k] === undefined ? '' : o[k]))); }

function doGet(e) {
  try {
    const q = e.parameter || {};
    if (q.action === 'products') return out({ ok: true, products: getProducts(), charts: getCharts(), config: publicConfig() });
    if (q.action === 'pincode') return out(Object.assign({ ok: true }, pinInfo(String(q.pin || ''))));
    if (q.action === 'reviews') return out({ ok: true, reviews: getReviews(String(q.productId || '')) });
    return out({ ok: true, service: 'Farside Paris API' });
  } catch (err) { return fail(err); }
}
function doPost(e) {
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    const handlers = { sendOtp, verifyOtp, myOrders, addReview, createPayment, placeOrder, checkCoupon, adminList, adminSaveProduct, adminUpload, adminSetStatus, adminSaveChart };
    if (!Object.prototype.hasOwnProperty.call(handlers, d.action)) throw new ApiError('Unknown request.');
    return out(Object.assign({ ok: true }, handlers[d.action](d)));
  } catch (err) { return fail(err); }
}

/* ---------- catalogue ---------- */
const publicConfig = () => ({ shipping: +prop('SHIPPING', 60), freeAbove: +prop('FREE_ABOVE', 999), razorpay: prop('RAZORPAY_KEY_ID', '') });
const isHidden = p => String(p.status || 'active').toLowerCase() === 'hidden';

function getProducts() {
  const cache = CacheService.getScriptCache(), hit = cache.get('products');
  if (hit) return JSON.parse(hit);
  const stats = {};
  rows('Reviews').forEach(r => { const s = stats[r.productId] || (stats[r.productId] = { n: 0, t: 0 }); s.n++; s.t += +r.rating; });
  const items = rows('Products').filter(p => p.id !== '' && !isHidden(p) && +p.price > 0).map(p => {
    const s = stats[p.id];
    return {
      id: String(p.id), name: p.name, category: p.category, gender: p.gender, price: +p.price, mrp: +p.mrp || +p.price,
      gsm: p.gsm, fit: p.fit, colors: split(p.colors), sizes: split(p.sizes), description: p.description,
      images: split(p.images, true), stock: p.stock === '' || p.stock == null ? null : Math.max(0, +p.stock), chart: String(p.chart || ''), rating: s ? s.t / s.n : 0, reviewCount: s ? s.n : 0
    };
  });
  try { cache.put('products', JSON.stringify(items), 60); } catch (e) {}
  return items;
}
function getReviews(pid) {
  return rows('Reviews').filter(r => String(r.productId) === pid)
    .sort((a, b) => new Date(b.date) - new Date(a.date))
    .map(r => ({ name: r.name, rating: +r.rating, text: r.text, verified: String(r.verified) === 'yes', date: new Date(r.date).toISOString() }));
}

/* ---------- sign-in (OTP by email or WhatsApp) ---------- */
function normId(d) {
  const raw = String(d.email || d.phone || '').trim().toLowerCase().replace(/[\s-]/g, '');
  const ph = raw.replace(/^(\+?91|0)(?=[6-9]\d{9}$)/, '');
  if (/^[6-9]\d{9}$/.test(ph)) return { id: ph, phone: true };
  if (raw.length <= 120 && /^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(raw)) return { id: raw, phone: false };
  throw new ApiError('Enter a valid 10-digit mobile number or email.');
}
function sendOtp(d) {
  const { id, phone } = normId(d), cache = CacheService.getScriptCache();
  const rk = 'rl:' + sha(id), n = +(cache.get(rk) || 0);
  if (n >= 5) throw new ApiError('Too many codes requested. Try again in 15 minutes.');
  cache.put(rk, String(n + 1), 900);
  const otp = String(parseInt(sha(Utilities.getUuid() + salt()).slice(0, 8), 16) % 900000 + 100000);
  const brand = prop('BRAND', 'Farside Paris');
  if (phone) sendWhatsApp(id, otp);
  else MailApp.sendEmail({ to: id, name: brand, subject: brand + ' sign-in code', body: 'Your ' + brand + ' sign-in code is ' + otp + '. It is valid for 10 minutes. If you did not ask for it, ignore this email.' });
  cache.put('otp:' + sha(id), JSON.stringify({ h: sha(id + otp + salt()), t: 0 }), 600);
  return {};
}
function verifyOtp(d) {
  const { id } = normId(d), cache = CacheService.getScriptCache(), k = 'otp:' + sha(id), raw = cache.get(k);
  if (!raw) throw new ApiError('Code expired. Request a new one.');
  const o = JSON.parse(raw);
  if (o.t >= 5) { cache.remove(k); throw new ApiError('Too many wrong attempts. Request a new code.'); }
  if (o.h !== sha(id + String(d.otp || '').trim() + salt())) { o.t++; cache.put(k, JSON.stringify(o), 600); throw new ApiError('Wrong code. Try again.'); }
  cache.remove(k);
  const token = sha(Utilities.getUuid() + Utilities.getUuid() + salt());
  add('Sessions', { hash: sha(token), identity: id, expires: new Date(Date.now() + 30 * 864e5).toISOString() });
  return { token, email: id };
}
function auth(d) {
  const t = String(d.token || '');
  if (t.length < 32) throw new ApiError('Please sign in again.', 'auth');
  const sh = sheet('Sessions'), f = sh.createTextFinder(sha(t)).matchEntireCell(true).findNext();
  if (!f) throw new ApiError('Please sign in again.', 'auth');
  const r = sh.getRange(f.getRow(), 1, 1, 3).getValues()[0];
  if (new Date(r[2]) < new Date()) { sh.deleteRow(f.getRow()); throw new ApiError('Session expired. Please sign in again.', 'auth'); }
  return String(r[1]);
}
function sendWhatsApp(phone, otp) {
  const tok = prop('WA_TOKEN'), pid = prop('WA_PHONE_ID'), tpl = prop('WA_TEMPLATE');
  if (!tok || !pid || !tpl) throw new ApiError('Mobile sign-in is not set up yet. Please use your email.');
  const r = UrlFetchApp.fetch('https://graph.facebook.com/v20.0/' + pid + '/messages', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true,
    payload: JSON.stringify({
      messaging_product: 'whatsapp', to: '91' + phone, type: 'template',
      template: { name: tpl, language: { code: prop('WA_LANG', 'en') }, components: [
        { type: 'body', parameters: [{ type: 'text', text: otp }] },
        { type: 'button', sub_type: 'url', index: '0', parameters: [{ type: 'text', text: otp }] }] }
    })
  });
  if (r.getResponseCode() >= 300) { console.error(r.getContentText()); throw new ApiError('Could not send the WhatsApp code. Try email instead.'); }
}

/* ---------- pricing is always recomputed on the server ---------- */
function priceCart(items, code, noStock) {
  if (!Array.isArray(items) || !items.length || items.length > 20) throw new ApiError('Your cart is empty.');
  const map = {}; rows('Products').forEach(p => map[String(p.id)] = p);
  let sub = 0;
  const lines = items.map(it => {
    const p = map[String(it.id)];
    if (!p || isHidden(p) || !(+p.price > 0)) throw new ApiError('An item in your cart is no longer available.');
    const sizes = split(p.sizes), colors = split(p.colors), size = String(it.size || ''), color = String(it.color || ''), qty = Math.floor(+it.qty);
    if (!(qty >= 1 && qty <= 10)) throw new ApiError('Invalid quantity.');
    const st = p.stock === '' || p.stock == null ? null : +p.stock;
    if (!noStock && st !== null && qty > st) throw new ApiError(st > 0 ? 'Only ' + st + ' left of ' + p.name + '.' : p.name + ' is sold out.');
    if (sizes.length && !sizes.includes(size)) throw new ApiError('Pick a valid size for ' + p.name + '.');
    if (colors.length && color && !colors.includes(color)) throw new ApiError('Pick a valid colour for ' + p.name + '.');
    sub += +p.price * qty;
    return { id: String(p.id), name: p.name, size, color, qty, price: +p.price };
  });
  const ship = sub >= +prop('FREE_ABOVE', 999) ? 0 : +prop('SHIPPING', 60);
  const cp = code ? findCoupon(code, sub) : null, disc = cp ? cp.disc : 0;
  return { lines, sub, ship, disc, coupon: cp ? cp.code : '', total: sub - disc + ship };
}
function cleanCustomer(c) {
  c = c || {};
  const o = { name: String(c.name || '').trim().slice(0, 80), phone: String(c.phone || '').trim(), address: String(c.address || '').trim().slice(0, 300), pincode: String(c.pincode || '').trim() };
  if (o.name.length < 2) throw new ApiError('Enter your full name.');
  if (!/^[6-9]\d{9}$/.test(o.phone)) throw new ApiError('Enter a valid 10-digit mobile number.');
  if (!/^\d{6}$/.test(o.pincode)) throw new ApiError('Enter a 6-digit pincode.');
  if (o.address.length < 10) throw new ApiError('Enter your full delivery address with house number and area.');
  return o;
}

/* ---------- payments + orders ---------- */
const rzAuth = () => ({ Authorization: 'Basic ' + Utilities.base64Encode(prop('RAZORPAY_KEY_ID') + ':' + prop('RAZORPAY_KEY_SECRET')) });
function createPayment(d) {
  auth(d);
  if (d.pincode) needDelivery(d.pincode, 'online');
  const t = priceCart(d.items, d.coupon);
  if (!prop('RAZORPAY_KEY_ID') || !prop('RAZORPAY_KEY_SECRET')) throw new ApiError('Online payment is not set up yet. Choose cash on delivery.');
  const r = UrlFetchApp.fetch('https://api.razorpay.com/v1/orders', {
    method: 'post', contentType: 'application/json', headers: rzAuth(), muteHttpExceptions: true,
    payload: JSON.stringify({ amount: t.total * 100, currency: 'INR', receipt: 'fp-' + Date.now() })
  });
  if (r.getResponseCode() >= 300) { console.error(r.getContentText()); throw new ApiError('Could not start payment. Try again.'); }
  const j = JSON.parse(r.getContentText());
  return { key: prop('RAZORPAY_KEY_ID'), amount: j.amount, orderId: j.id };
}
function verifyPayment(p, total) {
  const oid = String(p.razorpay_order_id || ''), pid = String(p.razorpay_payment_id || ''), sig = String(p.razorpay_signature || '');
  if (!oid || !pid || !sig) throw new ApiError('Payment details are missing.');
  const expect = hex(Utilities.computeHmacSha256Signature(oid + '|' + pid, prop('RAZORPAY_KEY_SECRET', '')));
  if (expect !== sig) throw new ApiError('Payment could not be verified. Contact us with payment ID ' + pid + '.');
  const r = UrlFetchApp.fetch('https://api.razorpay.com/v1/orders/' + encodeURIComponent(oid), { headers: rzAuth(), muteHttpExceptions: true });
  if (r.getResponseCode() >= 300 || JSON.parse(r.getContentText()).amount !== total * 100)
    throw new ApiError('Payment amount did not match your cart. Contact us with payment ID ' + pid + '.');
  return { pid, oid };
}
function placeOrder(d) {
  const id = auth(d), cust = cleanCustomer(d.customer), method = d.method === 'cod' ? 'cod' : 'online', t = priceCart(d.items, d.coupon, method === 'online');
  needDelivery(cust.pincode, method);
  let pay = { pid: '', oid: '' };
  if (method === 'online') pay = verifyPayment(d.payment || {}, t.total);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let orderId;
  try {
    if (pay.pid) { const ex = rows('Orders').find(o => String(o.paymentId) === pay.pid); if (ex) return { orderId: ex.orderId }; }
    const n = +prop('ORDER_SEQ', 1000) + 1;
    PropertiesService.getScriptProperties().setProperty('ORDER_SEQ', String(n));
    orderId = 'FP-' + n;
    takeStock(t.lines);
    add('Orders', {
      orderId, date: new Date(), identity: id, email: id.indexOf('@') > 0 ? id : '', name: cust.name, phone: cust.phone, address: cust.address, pincode: cust.pincode,
      items: JSON.stringify(t.lines), summary: t.lines.map(l => l.qty + ' x ' + l.name + ' (' + [l.size, l.color].filter(Boolean).join('/') + ')').join(', '),
      subtotal: t.sub, shipping: t.ship, total: t.total, payment: method === 'cod' ? 'Cash on delivery' : 'Paid online',
      paymentId: pay.pid, razorpayOrderId: pay.oid, status: 'New', discount: t.disc, coupon: t.coupon
    });
  } finally { lock.releaseLock(); }
  try { notify(orderId, id, cust, t, method); } catch (e) { console.error(e); }
  return { orderId };
}
function notify(orderId, id, cust, t, method) {
  const brand = prop('BRAND', 'Farside Paris');
  const lines = t.lines.map(l => '- ' + l.qty + ' x ' + l.name + ' (' + [l.size, l.color].filter(Boolean).join('/') + ') Rs ' + l.price * l.qty).join('\n');
  const body = 'Order ' + orderId + '\n' + lines + '\nShipping: Rs ' + t.ship + (t.disc ? '\nDiscount (' + t.coupon + '): -Rs ' + t.disc : '') + '\nTotal: Rs ' + t.total + '\nPayment: ' + (method === 'cod' ? 'Cash on delivery' : 'Paid online');
  const owner = prop('OWNER_EMAIL', Session.getEffectiveUser().getEmail());
  MailApp.sendEmail(owner, 'New order ' + orderId + ' (Rs ' + t.total + ')', body + '\n\nCustomer: ' + cust.name + ', ' + cust.phone + '\n' + cust.address + ' - ' + cust.pincode);
  if (id.indexOf('@') > 0) MailApp.sendEmail({ to: id, name: brand, subject: 'Your ' + brand + ' order ' + orderId, body: 'Thank you for your order.\n\n' + body + '\n\nWe will message you when it ships.' });
}
function myOrders(d) {
  const id = auth(d);
  return { orders: rows('Orders').filter(o => String(o.identity) === id).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 30)
    .map(o => ({ orderId: o.orderId, status: o.status, date: new Date(o.date).toISOString(), payment: o.payment, summary: o.summary, total: +o.total, tracking: o.tracking || '' })) };
}

/* ---------- reviews ---------- */
function addReview(d) {
  const id = auth(d), pid = String(d.productId || ''), rating = Math.round(+d.rating), text = String(d.text || '').trim().slice(0, 600);
  if (!rows('Products').some(p => String(p.id) === pid)) throw new ApiError('Product not found.');
  if (!(rating >= 1 && rating <= 5)) throw new ApiError('Tap the stars to rate this tee.');
  if (text.length < 5) throw new ApiError('Write at least a few words.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (rows('Reviews').some(r => String(r.identity) === id && String(r.productId) === pid)) throw new ApiError('You already reviewed this tee.');
    const mine = rows('Orders').filter(o => String(o.identity) === id);
    const verified = mine.some(o => String(o.items).indexOf('"id":"' + pid + '"') > -1);
    const name = mine.length ? String(mine[mine.length - 1].name).split(' ')[0] : (id.indexOf('@') > 0 ? id.split('@')[0] : 'Customer');
    add('Reviews', { id: 'r' + Date.now(), productId: pid, identity: id, name, rating, text, verified: verified ? 'yes' : 'no', date: new Date() });
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('products');
  return {};
}

/* optional: add a daily time-driven trigger for this to clear expired sign-ins */
function cleanup() {
  const sh = sheet('Sessions'), v = sh.getDataRange().getValues(), now = new Date();
  for (let i = v.length - 1; i >= 1; i--) if (new Date(v[i][2]) < now) sh.deleteRow(i + 1);
}


/* ---------- coupons + stock ---------- */
function findCoupon(code, sub) {
  const c = rows('Coupons').find(r => String(r.code).toUpperCase() === String(code).trim().toUpperCase() && String(r.status || 'active').toLowerCase() === 'active');
  if (!c) throw new ApiError('That coupon code is not valid.');
  if (sub < +c.minOrder) throw new ApiError('This coupon needs an order of at least Rs ' + c.minOrder + '.');
  return { code: String(c.code).toUpperCase(), disc: Math.min(sub, Math.round(String(c.type).toLowerCase() === 'percent' ? sub * +c.value / 100 : +c.value)) };
}
function checkCoupon(d) {
  if (!String(d.coupon || '').trim()) throw new ApiError('Enter a coupon code.');
  const t = priceCart(d.items, d.coupon, true);
  return { code: t.coupon, discount: t.disc };
}
function takeStock(lines) {
  const sh = sheet('Products'), v = sh.getDataRange().getValues(), ci = v[0].indexOf('id'), cs = v[0].indexOf('stock');
  if (cs < 0) return;
  lines.forEach(l => { for (let i = 1; i < v.length; i++) if (String(v[i][ci]) === l.id && v[i][cs] !== '') { v[i][cs] = Math.max(0, +v[i][cs] - l.qty); sh.getRange(i + 1, cs + 1).setValue(v[i][cs]); } });
  CacheService.getScriptCache().remove('products');
}

/* ---------- owner admin (admin.html) ---------- */
function admin(d) {
  const cache = CacheService.getScriptCache(), n = +(cache.get('adm') || 0), key = prop('ADMIN_KEY', '');
  if (key.length < 8) throw new ApiError('Set ADMIN_KEY (8+ characters) in Script Properties first.');
  if (n >= 10) throw new ApiError('Too many wrong keys. Try again in 15 minutes.');
  if (String(d.key || '') !== key) { cache.put('adm', String(n + 1), 900); throw new ApiError('Wrong admin key.'); }
}
function adminList(d) {
  admin(d);
  const orders = rows('Orders').reverse().slice(0, 300).map(o => ({
    orderId: o.orderId, date: new Date(o.date).toISOString(), name: o.name, phone: o.phone, email: o.email, address: o.address, pincode: o.pincode,
    summary: o.summary, subtotal: +o.subtotal, shipping: +o.shipping, discount: +o.discount || 0, coupon: o.coupon, total: +o.total,
    payment: o.payment, paymentId: o.paymentId, status: o.status, tracking: o.tracking || ''
  }));
  return { orders, products: rows('Products'), charts: readCharts() };
}
function adminSaveProduct(d) {
  admin(d);
  const p = d.product || {}, name = String(p.name || '').trim();
  if (!name || !(+p.price > 0)) throw new ApiError('Name and a price above 0 are required.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('Products'), v = sh.getDataRange().getValues(), h = v[0], id = String(p.id || '').trim() || 'p' + Date.now();
    const rec = Object.assign({}, p, { id, name }), row = h.map(k => safe(rec[k] === undefined ? '' : rec[k]));
    const i = v.findIndex((r, j) => j > 0 && String(r[h.indexOf('id')]) === id);
    if (i > 0) sh.getRange(i + 1, 1, 1, h.length).setValues([row]); else sh.appendRow(row);
    return { id };
  } finally { lock.releaseLock(); CacheService.getScriptCache().remove('products'); }
}
function adminUpload(d) {
  admin(d);
  const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(d.data || ''));
  if (!m) throw new ApiError('Send a JPG, PNG or WebP photo.');
  const it = DriveApp.getFoldersByName('Store images'), dir = it.hasNext() ? it.next() : DriveApp.createFolder('Store images');
  const f = dir.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'p-' + Date.now() + '.' + m[1].split('/')[1]));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://lh3.googleusercontent.com/d/' + f.getId() + '=w1000' };
}
function adminSetStatus(d) {
  admin(d);
  const status = String(d.status || ''), tracking = String(d.tracking || '').slice(0, 200);
  if (['New', 'Packed', 'Shipped', 'Delivered', 'Cancelled'].indexOf(status) < 0) throw new ApiError('Pick a valid status.');
  const sh = sheet('Orders'), f = sh.createTextFinder(String(d.orderId || '')).matchEntireCell(true).findNext();
  if (!f) throw new ApiError('Order not found.');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], r = f.getRow(), o = sh.getRange(r, 1, 1, h.length).getValues()[0], g = k => String(o[h.indexOf(k)] || '');
  sh.getRange(r, h.indexOf('status') + 1).setValue(status);
  sh.getRange(r, h.indexOf('tracking') + 1).setValue(safe(tracking));
  const to = g('email'), brand = prop('BRAND', 'Farside Paris');
  if (to.indexOf('@') > 0 && (status !== g('status') || tracking !== g('tracking')))
    try { MailApp.sendEmail({ to, name: brand, subject: 'Your ' + brand + ' order ' + d.orderId + ' is ' + status, body: 'Hi ' + g('name') + ',\n\nYour order ' + d.orderId + ' is now: ' + status + '.' + (tracking ? '\nTracking: ' + tracking : '') + '\n\nThank you for shopping with ' + brand + '.' }); } catch (e) { console.error(e); }
  return {};
}

/* ---------- size charts + pincode delivery ---------- */
const safeRows = n => sheet(n) ? rows(n) : [];
function readCharts() {
  const o = {};
  safeRows('SizeCharts').forEach(r => { const n = String(r.chart).trim(), t = String(r.table).trim(); if (n && t) o[n] = t; });
  return o;
}
function getCharts() {
  const cache = CacheService.getScriptCache(), hit = cache.get('charts');
  if (hit) return JSON.parse(hit);
  const o = readCharts();
  try { cache.put('charts', JSON.stringify(o), 60); } catch (e) {}
  return o;
}
function adminSaveChart(d) {
  admin(d);
  const name = String(d.name || '').trim().slice(0, 60), table = String(d.table || '').trim().slice(0, 4000);
  if (!name) throw new ApiError('Give the chart a name.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet('SizeCharts'), v = sh.getDataRange().getValues(), i = v.findIndex((r, j) => j > 0 && String(r[0]).trim().toLowerCase() === name.toLowerCase());
    if (!table) { if (i > 0) sh.deleteRow(i + 1); }
    else if (i > 0) sh.getRange(i + 1, 2).setValue(safe(table));
    else sh.appendRow([safe(name), safe(table)]);
  } finally { lock.releaseLock(); CacheService.getScriptCache().remove('charts'); }
  return {};
}
function pinInfo(pin) {
  pin = String(pin || '').trim();
  if (!/^[1-9]\d{5}$/.test(pin)) throw new ApiError('Enter a valid 6-digit pincode.');
  const rs = safeRows('Delivery');
  if (!rs.length) return { deliverable: true, days: '4-8', cod: true };
  let best = null, def = null;
  rs.forEach(r => {
    const p = String(r.pincode).trim();
    if (p === '*') def = r;
    else if (/^\d{1,6}$/.test(p) && pin.indexOf(p) === 0 && (!best || p.length > String(best.pincode).trim().length)) best = r;
  });
  const m = best || def;
  if (!m || String(m.status || '').toLowerCase() === 'blocked') return { deliverable: false };
  return { deliverable: true, days: String(m.days || '4-8'), cod: String(m.cod || 'yes').toLowerCase() !== 'no' };
}
function needDelivery(pin, method) {
  const x = pinInfo(pin);
  if (!x.deliverable) throw new ApiError('Sorry, we do not deliver to pincode ' + pin + ' yet.');
  if (method === 'cod' && !x.cod) throw new ApiError('Cash on delivery is not available for this pincode. Please pay online.');
}
