/**
 * Easy Orders Lite — Backend (Google Apps Script)
 * ------------------------------------------------
 * الشيت ده هو قاعدة البيانات:
 *   Orders   → كل الأوردرات اللي جاية من صفحات الهبوط
 *   Users    → يوزرات العملاء (باسوورد متشفّر)
 *   Products → أسعار المنتجات (السيرفر بيحسب السعر من هنا، مش من الصفحة)
 *
 * الإعداد: من قايمة "⚙️ إدارة النظام" في الشيت → "تجهيز الشيتات" مرة واحدة،
 * وبعدين Deploy → New deployment → Web app (Execute as: Me, Access: Anyone).
 */

const CFG = {
  TZ: 'Africa/Cairo',
  TOKEN_SECONDS: 6 * 60 * 60,       // مدة الجلسة: 6 ساعات (أقصى حد للـ Cache)
  DUPLICATE_MINUTES: 30,            // نفس الرقم + نفس المنتج خلال 30 دقيقة = أوردر مكرر
  MAX_LOGIN_FAILS: 5,               // محاولات دخول غلط قبل القفل المؤقت
  ORDER_PREFIX: 'EO-',
  STATUSES: ['جديد', 'مؤكد', 'تم الشحن', 'تم التسليم', 'ملغي', 'لم يرد', 'مرتجع'],
};

const HEADERS = {
  Orders: ['order_id', 'created_at', 'date', 'client_id', 'product_id', 'product', 'qty',
    'unit_price', 'shipping', 'total', 'name', 'phone', 'phone2', 'governorate', 'address',
    'notes', 'status', 'admin_note', 'updated_at', 'updated_by',
    'utm_source', 'utm_campaign', 'utm_content', 'page_url'],
  Users: ['username', 'password_hash', 'salt', 'client_id', 'display_name', 'role', 'active'],
  Products: ['client_id', 'product_id', 'name', 'price', 'shipping', 'active'],
};

/* ============================ HTTP ============================ */

function doGet() {
  return json_({ ok: true, service: 'easy-orders-lite', time: now_() });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      case 'order':  return json_(createOrder_(body));
      case 'login':  return json_(login_(body));
      case 'orders': return json_(listOrders_(body));
      case 'update': return json_(updateOrder_(body));
      case 'logout': return json_(logout_(body));
      default:       return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: 'server_error', message: String(err && err.message || err) });
  }
}

/* ============================ Orders ============================ */

function createOrder_(b) {
  const name = clean_(b.name, 80);
  const phone = normalizePhone_(b.phone);
  const phone2 = b.phone2 ? normalizePhone_(b.phone2) : '';
  const governorate = clean_(b.governorate, 40);
  const address = clean_(b.address, 300);
  const clientId = clean_(b.client_id, 40);
  const productId = clean_(b.product_id, 60);
  const qty = Math.max(1, Math.min(20, parseInt(b.qty, 10) || 1));

  const errors = {};
  if (name.length < 3) errors.name = 'الاسم مطلوب';
  if (!/^01[0125]\d{8}$/.test(phone)) errors.phone = 'رقم موبايل غير صحيح';
  if (phone2 && !/^01[0125]\d{8}$/.test(phone2)) errors.phone2 = 'رقم موبايل غير صحيح';
  if (!governorate) errors.governorate = 'اختار المحافظة';
  if (address.length < 5) errors.address = 'العنوان مطلوب';
  if (!clientId || !productId) errors.product = 'منتج غير معروف';
  if (Object.keys(errors).length) return { ok: false, error: 'validation', fields: errors };

  // السعر من شيت Products — مش من المتصفح
  const product = findProduct_(clientId, productId);
  if (!product) return { ok: false, error: 'validation', fields: { product: 'المنتج غير متاح' } };
  const unit = Number(product.price) || 0;
  const shipping = Number(product.shipping) || 0;
  const total = unit * qty + shipping;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_('Orders');
    const dup = findRecentDuplicate_(sh, phone, clientId, productId);
    if (dup) return { ok: true, duplicate: true, order_id: dup, total: total };

    const props = PropertiesService.getScriptProperties();
    const seq = Number(props.getProperty('order_seq') || 1000) + 1;
    props.setProperty('order_seq', String(seq));
    const orderId = CFG.ORDER_PREFIX + seq;

    const d = new Date();
    const row = {
      order_id: orderId,
      created_at: fmt_(d, 'yyyy-MM-dd HH:mm:ss'),
      date: fmt_(d, 'yyyy-MM-dd'),
      client_id: clientId,
      product_id: productId,
      product: product.name,
      qty: qty,
      unit_price: unit,
      shipping: shipping,
      total: total,
      name: name,
      phone: phone,
      phone2: phone2,
      governorate: governorate,
      address: address,
      notes: clean_(b.notes, 300),
      status: CFG.STATUSES[0],
      admin_note: '',
      updated_at: '',
      updated_by: '',
      utm_source: clean_(b.utm_source, 80),
      utm_campaign: clean_(b.utm_campaign, 120),
      utm_content: clean_(b.utm_content, 120),
      page_url: clean_(b.page_url, 300),
    };
    sh.appendRow(HEADERS.Orders.map(h => asText_(row[h])));
    return { ok: true, order_id: orderId, total: total };
  } finally {
    lock.releaseLock();
  }
}

function listOrders_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'unauthorized' };

  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const from = isDate_(b.from) ? b.from : today;
  const to = isDate_(b.to) ? b.to : from;

  const sh = sheet_('Orders');
  const values = sh.getDataRange().getDisplayValues();
  const head = values.shift() || [];
  const idx = indexOf_(head);
  const out = [];
  for (let i = values.length - 1; i >= 0; i--) {   // الأحدث الأول
    const r = values[i];
    const date = r[idx.date];
    if (date < from || date > to) continue;
    if (s.role !== 'admin' && r[idx.client_id] !== s.client_id) continue;
    if (s.role === 'admin' && b.client_id && r[idx.client_id] !== b.client_id) continue;
    const o = {};
    head.forEach((h, j) => { o[h] = r[j]; });
    out.push(o);
  }
  return { ok: true, from: from, to: to, statuses: CFG.STATUSES, orders: out,
           user: { name: s.display_name, role: s.role, client_id: s.client_id } };
}

function updateOrder_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'unauthorized' };
  if (b.status && CFG.STATUSES.indexOf(b.status) === -1) return { ok: false, error: 'bad_status' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_('Orders');
    const values = sh.getDataRange().getDisplayValues();
    const idx = indexOf_(values[0]);
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][idx.order_id] !== String(b.order_id)) continue;
      if (s.role !== 'admin' && values[i][idx.client_id] !== s.client_id) return { ok: false, error: 'forbidden' };
      const rowNo = i + 1;
      if (b.status) sh.getRange(rowNo, idx.status + 1).setValue(b.status);
      if (typeof b.admin_note === 'string') sh.getRange(rowNo, idx.admin_note + 1).setValue(asText_(clean_(b.admin_note, 300)));
      sh.getRange(rowNo, idx.updated_at + 1).setValue(fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss'));
      sh.getRange(rowNo, idx.updated_by + 1).setValue(s.username);
      return { ok: true };
    }
    return { ok: false, error: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

function findRecentDuplicate_(sh, phone, clientId, productId) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const start = Math.max(2, last - 200);              // آخر 200 أوردر كفاية
  const values = sh.getRange(start, 1, last - start + 1, HEADERS.Orders.length).getDisplayValues();
  const idx = indexOf_(HEADERS.Orders);
  const limit = new Date(Date.now() - CFG.DUPLICATE_MINUTES * 60000);
  const limitStr = fmt_(limit, 'yyyy-MM-dd HH:mm:ss');
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (r[idx.created_at] < limitStr) break;
    if (r[idx.phone] === phone && r[idx.client_id] === clientId && r[idx.product_id] === productId) return r[idx.order_id];
  }
  return null;
}

function findProduct_(clientId, productId) {
  const values = sheet_('Products').getDataRange().getDisplayValues();
  const idx = indexOf_(values.shift() || []);
  for (const r of values) {
    if (r[idx.client_id] === clientId && r[idx.product_id] === productId && truthy_(r[idx.active])) {
      return { name: r[idx.name], price: r[idx.price], shipping: r[idx.shipping] };
    }
  }
  return null;
}

/* ============================ Auth ============================ */

function login_(b) {
  const username = clean_(b.username, 60).toLowerCase();
  const password = String(b.password || '');
  if (!username || !password) return { ok: false, error: 'bad_credentials' };

  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + username;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= CFG.MAX_LOGIN_FAILS) return { ok: false, error: 'locked' };

  const values = sheet_('Users').getDataRange().getDisplayValues();
  const idx = indexOf_(values.shift() || []);
  for (const r of values) {
    if (String(r[idx.username]).toLowerCase() !== username) continue;
    if (!truthy_(r[idx.active])) break;
    if (hash_(r[idx.salt], password) !== r[idx.password_hash]) break;
    const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
    const s = { username: username, client_id: r[idx.client_id], display_name: r[idx.display_name] || username,
                role: r[idx.role] === 'admin' ? 'admin' : 'client' };
    cache.put('tok_' + token, JSON.stringify(s), CFG.TOKEN_SECONDS);
    cache.remove(failKey);
    return { ok: true, token: token, user: { name: s.display_name, role: s.role, client_id: s.client_id } };
  }
  cache.put(failKey, String(fails + 1), 15 * 60);
  return { ok: false, error: 'bad_credentials' };
}

function logout_(b) {
  if (b.token) CacheService.getScriptCache().remove('tok_' + b.token);
  return { ok: true };
}

function session_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('tok_' + token);
  return raw ? JSON.parse(raw) : null;
}

function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/** تستخدمها من القايمة أو من المحرر: إضافة/تحديث يوزر */
function upsertUser(username, password, clientId, displayName, role) {
  username = String(username).trim().toLowerCase();
  const sh = sheet_('Users');
  const salt = Utilities.getUuid();
  const row = [username, hash_(salt, password), salt, clientId, displayName || username, role || 'client', 'TRUE'];
  const values = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === username) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row.map(asText_)]);
      return 'updated';
    }
  }
  sh.appendRow(row.map(asText_));
  return 'created';
}

/* ============================ Sheet menu ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ إدارة النظام')
    .addItem('1) تجهيز الشيتات (أول مرة)', 'setup')
    .addItem('إضافة عميل / تغيير باسوورد', 'menuAddUser')
    .addItem('إضافة منتج', 'menuAddProduct')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#0f2e22').setFontColor('#ffffff');
    }
    sh.getRange('A:Z').setNumberFormat('@');           // كله نص: يحافظ على صفر الموبايل والتاريخ
    sh.setRightToLeft(true);
  });
  // منتج Black Med كبداية
  if (!findProduct_('barimed', 'blackmed')) {
    sheet_('Products').appendRow(['barimed', 'blackmed', 'بلاك ميد', '450', '0', 'TRUE']);
  }
  const s = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  // getUi() مش متاح لما الدالة تشتغل من محرر الكود — نتجاهل التنبيه ساعتها
  try { SpreadsheetApp.getUi().alert('تم تجهيز الشيتات ✅\nالخطوة الجاية: إضافة عميل من نفس القايمة.'); } catch (e) {}
  return 'setup done';
}

function menuAddUser() {
  const ui = SpreadsheetApp.getUi();
  const ask = (q) => {
    const r = ui.prompt(q, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('cancelled');
    return r.getResponseText().trim();
  };
  try {
    const u = ask('اسم المستخدم (إنجليزي بدون مسافات):');
    const p = ask('الباسوورد (8 حروف على الأقل):');
    if (p.length < 8) return ui.alert('الباسوورد قصير — لازم 8 حروف على الأقل.');
    const c = ask('كود العميل client_id (مثلاً barimed):');
    const n = ask('اسم العميل اللي هيظهر في اللوحة:');
    const role = ui.alert('هل ده أدمن يشوف كل العملاء؟', ui.ButtonSet.YES_NO) === ui.Button.YES ? 'admin' : 'client';
    ui.alert('تم: ' + upsertUser(u, p, c, n, role));
  } catch (e) { if (String(e.message) !== 'cancelled') throw e; }
}

function menuAddProduct() {
  const ui = SpreadsheetApp.getUi();
  const vals = [];
  for (const q of ['كود العميل client_id:', 'كود المنتج product_id (إنجليزي):', 'اسم المنتج:', 'السعر:', 'الشحن (0 لو مجاني):']) {
    const r = ui.prompt(q, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    vals.push(r.getResponseText().trim());
  }
  vals.push('TRUE');
  sheet_('Products').appendRow(vals);
  ui.alert('تم إضافة المنتج ✅');
}

/* ============================ Helpers ============================ */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '" — شغّل setup من القايمة');
  return sh;
}
function indexOf_(head) { const o = {}; head.forEach((h, i) => { o[h] = i; }); return o; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fmt_(d, p) { return Utilities.formatDate(d, CFG.TZ, p); }
function now_() { return fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss'); }
function isDate_(s) { return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s); }
function truthy_(v) { return ['true', '1', 'yes', 'نعم'].indexOf(String(v).trim().toLowerCase()) !== -1; }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max || 200); }
// يمنع الشيت من تنفيذ أي قيمة كـ formula (=, +, -, @)
function asText_(v) { v = v == null ? '' : String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function normalizePhone_(v) {
  let s = String(v || '')
    .replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[۰-۹]/g, d => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)))
    .replace(/\D/g, '');
  if (s.indexOf('0020') === 0) s = s.slice(4);
  else if (s.indexOf('20') === 0 && s.length === 12) s = s.slice(2);
  if (s.length === 10 && s[0] === '1') s = '0' + s;
  return s;
}
