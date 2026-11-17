/**
 * CLAIM TRACKER v2 (backend API untuk versi GitHub Pages)
 * Google Apps Script. Data kekal dalam Google Sheet + folder Drive yang sama seperti v1.
 *
 * Paparan app (index.html) dihos di GitHub Pages dan bercakap dengan skrip ini melalui doPost.
 * Log masuk: "Sign in with Google" di app -> token Google disahkan di sini -> sesi app (60 hari).
 * Hanya emel dalam tab "Pengguna" boleh guna app.
 *
 * Script properties yang perlu ada:
 *   SHEET_ID, ROOT_ID   (dicipta oleh setup() dalam v1)
 *   GOOGLE_CLIENT_ID    (OAuth Client ID jenis Web, dari Google Cloud Console)
 *   APP_URL             (alamat GitHub Pages, cth https://nama.github.io/claim-tracker/)
 *
 * Deploy > New deployment > Web app:
 *   Execute as: Me
 *   Who has access: Anyone
 * Fail Aset.html (logo borang sebenar) kekal dalam projek Apps Script ini sahaja, tidak dimuat naik ke GitHub.
 */

var APP_NAME = 'Claim Tracker';

// layout: 'pm' = ikut borang PM sebenar (juga dipakai FullStag), 'perseni' = ikut borang Perseni sebenar
var ENTITI = [
  { kod: 'PM',  nama: 'Produksi Mayasari', warna: '#F47B20', layout: 'pm',      logo: 'pmLogo' },
  { kod: 'FSE', nama: 'FullStag Empire',   warna: '#1F2937', layout: 'pm',      logo: 'fseLogo' },
  { kod: 'PRS', nama: 'Perseni',           warna: '#FF3333', layout: 'perseni', logo: 'prsHeader' }
];

var KATEGORI_LALAI = ['Makan', 'Minyak', 'Tol', 'Parkir', 'Pengangkutan/Grab',
  'Peralatan/Props', 'Percetakan', 'Perisian/Langganan', 'Lain-lain'];

var STATUS = { BELUM: 'Belum claim', DAH: 'Dah claim', BAYAR: 'Dibayar' };

var H_REKOD = ['ID', 'Dicipta', 'EmelClaimant', 'Claimant', 'Jenis', 'Tarikh', 'Entiti', 'Kategori',
  'Butiran', 'NoResit', 'Projek', 'Amaun', 'Dari', 'Ke', 'KmSehala', 'DuaHala', 'Kadar', 'ResitFileId', 'ResitUrl',
  'Status', 'NoRujukan', 'TarikhHantar', 'TarikhBayar', 'Nota', 'Dikemaskini', 'OlehEmel'];
var H_LOG = ['NoRujukan', 'Dijana', 'EmelClaimant', 'Claimant', 'Entiti', 'Dari', 'Hingga', 'BilRekod',
  'Jumlah', 'PdfFileId', 'PdfUrl', 'OlehEmel'];
var H_PENGGUNA = ['Emel', 'Nama', 'Kod', 'Peranan', 'NamaPenuh', 'Jawatan'];
var H_TETAPAN = ['Kunci', 'Nilai', 'Nota'];

/* ------------------------------------------------------------------ */
/* Pemasangan                                                          */
/* ------------------------------------------------------------------ */

/** Jalankan SEKALI dari editor Apps Script. Selamat dijalankan semula (guna folder/Sheet sedia ada). */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var me = Session.getEffectiveUser().getEmail();
  Logger.log('1/6 Akaun: ' + me);

  var fname = APP_NAME + ' (Resit & Borang)';
  var it = DriveApp.getFoldersByName(fname);
  var root = it.hasNext() ? it.next() : DriveApp.createFolder(fname);
  Logger.log('2/6 Folder: ' + root.getUrl());

  var ss, fit = root.getFilesByName(APP_NAME + ' - Data');
  while (fit.hasNext()) { var f = fit.next(); if (f.getMimeType() === MimeType.GOOGLE_SHEETS) { ss = SpreadsheetApp.openById(f.getId()); break; } }
  if (!ss) { ss = SpreadsheetApp.create(APP_NAME + ' - Data'); DriveApp.getFileById(ss.getId()).moveTo(root); }
  Logger.log('3/6 Sheet: ' + ss.getUrl());

  function tab(name, headers, first) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = first ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
    if (!String(sh.getRange(1, 1).getValue())) {
      sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }
  var rekod = tab('Rekod', H_REKOD, !ss.getSheetByName('Rekod'));
  rekod.getRange(2, H_REKOD.indexOf('Amaun') + 1, rekod.getMaxRows() - 1, 1).setNumberFormat('0.00');
  Logger.log('4/6 Tab Rekod siap');

  var tet = tab('Tetapan', H_TETAPAN);
  if (tet.getLastRow() < 2) tet.getRange(2, 1, 6, 3).setValues([
    ['kadar_km', '0.65', 'RM sekilometer. Jangan tukar tanpa boss.'],
    ['lihat_semua', 'YA', 'YA = semua pengguna nampak semua rekod. TIDAK = nampak rekod sendiri sahaja.'],
    ['kategori', KATEGORI_LALAI.join(', '), 'Senarai kategori, dipisah koma.'],
    ['perseni_pelulus_nama', '', 'Nama pelulus dalam borang Perseni.'],
    ['perseni_pelulus_jawatan', '', 'Jawatan pelulus dalam borang Perseni.'],
    ['tandatangan_taip', 'YA', 'YA = nama penuntut ditaip gaya tulisan tangan di ruang tandatangan.']
  ]);
  tab('LogBorang', H_LOG);
  var pg = tab('Pengguna', H_PENGGUNA);
  if (pg.getLastRow() < 2) pg.getRange(2, 1, 1, H_PENGGUNA.length).setValues([[me, me.split('@')[0], 'AD', 'admin', '', '']]);
  Logger.log('5/6 Tab Tetapan, LogBorang, Pengguna siap');

  props.setProperties({ SHEET_ID: ss.getId(), ROOT_ID: root.getId() });
  Logger.log('6/6 Siap! Sekarang buat Deploy > New deployment.');
}

/* ------------------------------------------------------------------ */
/* Web app                                                             */
/* ------------------------------------------------------------------ */

function doGet() {
  var url = PropertiesService.getScriptProperties().getProperty('APP_URL') || '';
  var h = '<div style="font:16px sans-serif;padding:24px;text-align:center">' +
    '<h2>Claim Tracker dah pindah</h2>' +
    (url ? '<p><a target="_top" href="' + url + '">Buka app baharu</a></p>' : '<p>Buka alamat app GitHub Pages.</p>') + '</div>';
  return HtmlService.createHtmlOutput(h).setTitle(APP_NAME);
}

/* ------------------------------------------------------------------ */
/* API (dipanggil oleh app GitHub Pages melalui fetch POST)           */
/* ------------------------------------------------------------------ */

var CURRENT_EMAIL_ = null;
var SESI_HARI_ = 60;

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fns = {
      apiInit: apiInit, apiSaveRecord: apiSaveRecord, apiDeleteRecord: apiDeleteRecord, apiSetStatus: apiSetStatus,
      apiGetReceipt: apiGetReceipt, apiNextRef: apiNextRef, apiSaveClaim: apiSaveClaim, apiExportCsv: apiExportCsv,
      apiAddUser: apiAddUser, apiSetSetting: apiSetSetting, apiUpdateMe: apiUpdateMe, apiAset: apiAset, apiLogout: apiLogout
    };
    var args = Array.isArray(req.args) ? req.args : [];
    if (req.fn === 'apiLogin') {
      out = { ok: true, data: login_(args[0]) };
    } else {
      if (!fns.hasOwnProperty(req.fn)) throw new Error('Fungsi tidak sah.');
      CURRENT_EMAIL_ = sesi_(req.session);
      CURRENT_SESSION_ = req.session;
      out = { ok: true, data: fns[req.fn].apply(null, args) };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err), auth: !!(err && err.auth) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

var CURRENT_SESSION_ = null;

function authErr_(msg) { var e = new Error(msg); e.auth = true; return e; }

function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/** Sahkan ID token Google, semak emel dalam senarai Pengguna, cipta sesi app. */
function login_(idToken) {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('GOOGLE_CLIENT_ID');
  if (!clientId) throw new Error('GOOGLE_CLIENT_ID belum ditetapkan dalam Script properties.');
  if (!idToken || typeof idToken !== 'string') throw authErr_('Token log masuk tiada.');
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw authErr_('Log masuk Google tidak sah. Cuba lagi.');
  var t = JSON.parse(res.getContentText());
  if (t.aud !== clientId) throw authErr_('Token bukan untuk app ini.');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(t.iss) < 0) throw authErr_('Token tidak sah.');
  if (String(t.email_verified) !== 'true') throw authErr_('Emel Google belum disahkan.');
  if (Number(t.exp) * 1000 < Date.now()) throw authErr_('Token tamat tempoh. Cuba lagi.');
  var email = String(t.email || '').toLowerCase();
  if (!users_().some(function (u) { return u.email === email; }))
    throw authErr_('Emel ' + email + ' tiada dalam senarai pengguna. Minta boss tambah dalam Profil.');
  var token = Utilities.getUuid() + Utilities.getUuid();
  var exp = Date.now() + SESI_HARI_ * 864e5;
  props.setProperty('S_' + hash_(token), JSON.stringify({ e: email, x: exp }));
  bersihSesi_();
  return { session: token, email: email, exp: exp };
}

function sesi_(token) {
  if (!token || typeof token !== 'string') throw authErr_('Sila log masuk.');
  var key = 'S_' + hash_(token);
  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty(key);
  if (!v) throw authErr_('Sesi tamat. Sila log masuk semula.');
  var o = JSON.parse(v);
  if (o.x < Date.now()) { props.deleteProperty(key); throw authErr_('Sesi tamat. Sila log masuk semula.'); }
  return o.e;
}

function bersihSesi_() {
  var props = PropertiesService.getScriptProperties(), all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('S_') !== 0) return;
    try { if (JSON.parse(all[k]).x < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

function apiLogout() {
  if (CURRENT_SESSION_) PropertiesService.getScriptProperties().deleteProperty('S_' + hash_(CURRENT_SESSION_));
  return true;
}

/** Logo borang sebenar (dari Aset.html dalam projek ini), hanya untuk pengguna berdaftar. */
function apiAset() {
  auth_();
  var src = HtmlService.createHtmlOutputFromFile('Aset').getContent();
  var o = { pmLogo: '', prsHeader: '', prsFooter: '', fseLogo: '' }, re = /(\w+)\s*:\s*'(data:image\/[^']*)'/g, m;
  while ((m = re.exec(src))) o[m[1]] = m[2];
  return o;
}

/* ------------------------------------------------------------------ */
/* Pembantu dalaman                                                    */
/* ------------------------------------------------------------------ */

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('App belum dipasang. Jalankan setup() dahulu.');
  return SpreadsheetApp.openById(id);
}
function root_() {
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_ID'));
}
function tz_() { return Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd'T'HH:mm:ss"); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function users_() {
  var v = ss_().getSheetByName('Pengguna').getDataRange().getValues();
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return { email: String(r[0]).trim().toLowerCase(), nama: str_(r[1]), kod: str_(r[2]), admin: str_(r[3]).toLowerCase() === 'admin',
      namaPenuh: str_(r[4]), jawatan: str_(r[5]) };
  });
}

/** Pastikan pemanggil ada dalam senarai Pengguna. */
function auth_() {
  var email = String(CURRENT_EMAIL_ || '').toLowerCase();
  var u = users_().filter(function (x) { return x.email === email; })[0];
  if (!u) throw new Error('Akses ditolak untuk ' + (email || 'pengguna tidak dikenali') + '.');
  return u;
}
function admin_() {
  var u = auth_();
  if (!u.admin) throw new Error('Hanya admin boleh buat ini.');
  return u;
}

/** Dalam mod peribadi (lihat_semua = TIDAK), hanya pemilik rekod boleh ubah rekod itu. */
function pemilik_(me, emel) {
  if ((settings_().lihat_semua || 'YA').toUpperCase() === 'YA') return;
  if (String(emel || '').toLowerCase() !== me.email) throw new Error('Rekod ini bukan milik anda.');
}

function settings_() {
  var v = ss_().getSheetByName('Tetapan').getDataRange().getValues();
  var o = {};
  v.slice(1).forEach(function (r) { if (r[0]) o[String(r[0])] = str_(r[1]); });
  return o;
}

function readTable_(name, headers) {
  var sh = ss_().getSheetByName(name);
  var v = sh.getDataRange().getValues();
  return v.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r, i) {
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = str_(r[j]); });
    return o;
  });
}

function findRow_(sh, id) {
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function folderPath_(parent, names) {
  var f = parent;
  names.forEach(function (n) {
    var it = f.getFoldersByName(n);
    f = it.hasNext() ? it.next() : f.createFolder(n);
  });
  return f;
}

function entiti_(kod) {
  var e = ENTITI.filter(function (x) { return x.kod === kod; })[0];
  if (!e) throw new Error('Entiti tidak sah: ' + kod);
  return e;
}

function round2_(n) { return Math.round(Number(n) * 100) / 100; }

/* ------------------------------------------------------------------ */
/* Fungsi API                                                          */
/* ------------------------------------------------------------------ */

function apiInit() {
  var me = auth_();
  var s = settings_();
  var lihatSemua = (s.lihat_semua || 'YA').toUpperCase() === 'YA';
  var rekod = readTable_('Rekod', H_REKOD);
  var log = readTable_('LogBorang', H_LOG);
  if (!lihatSemua) { // mod peribadi: setiap orang (termasuk admin) nampak rekod sendiri sahaja
    rekod = rekod.filter(function (r) { return r.EmelClaimant === me.email; });
    log = log.filter(function (r) { return r.EmelClaimant === me.email; });
  }
  return {
    me: me,
    users: lihatSemua ? users_() : [me],
    allUsers: me.admin ? users_() : [me],
    entiti: ENTITI,
    kategori: (s.kategori || KATEGORI_LALAI.join(',')).split(',').map(function (x) { return x.trim(); }).filter(String),
    kadar: Number(s.kadar_km || 0.65),
    lihatSemua: lihatSemua,
    pelulus: { nama: s.perseni_pelulus_nama || '', jawatan: s.perseni_pelulus_jawatan || '' },
    tandatanganTaip: (s.tandatangan_taip || 'YA').toUpperCase() === 'YA',
    rekod: rekod,
    log: log.reverse(),
    sheetUrl: me.admin ? ss_().getUrl() : '',
    today: today_()
  };
}

/**
 * Simpan rekod baharu atau kemas kini.
 * @param {Object} r rekod (tanpa ID = baharu)
 * @param {Object=} file {name, mimeType, data(base64)} resit
 */
function apiSaveRecord(r, file) {
  var me = auth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName('Rekod');
    var users = users_();
    var claimant = users.filter(function (u) { return u.email === String(r.EmelClaimant || me.email).toLowerCase(); })[0];
    if (!claimant) throw new Error('Claimant tidak sah.');
    pemilik_(me, claimant.email);
    var e = entiti_(r.Entiti);
    var jenis = r.Jenis === 'Perbatuan' ? 'Perbatuan' : 'Resit';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(r.Tarikh || '')) throw new Error('Tarikh tidak sah.');

    var existing = null, row = -1;
    if (r.ID) {
      row = findRow_(sh, r.ID);
      if (row < 0) throw new Error('Rekod tidak dijumpai.');
      existing = {};
      var vals = sh.getRange(row, 1, 1, H_REKOD.length).getValues()[0];
      H_REKOD.forEach(function (h, j) { existing[h] = str_(vals[j]); });
      pemilik_(me, existing.EmelClaimant);
      if (existing.Status !== STATUS.BELUM && !r._bukaKunci)
        throw new Error('Rekod ini sudah "' + existing.Status + '" dan dikunci.');
    }

    var o = existing ? existing : {
      ID: 'R' + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100),
      Dicipta: now_(), Status: STATUS.BELUM, NoRujukan: '', TarikhHantar: '', TarikhBayar: '',
      ResitFileId: '', ResitUrl: ''
    };
    o.EmelClaimant = claimant.email;
    o.Claimant = claimant.nama;
    o.Jenis = jenis;
    o.Tarikh = r.Tarikh;
    o.Entiti = e.kod;
    o.Kategori = jenis === 'Perbatuan' ? 'Perbatuan' : String(r.Kategori || 'Lain-lain');
    o.Butiran = String(r.Butiran || '').slice(0, 500);
    o.Projek = String(r.Projek || '').slice(0, 200);
    o.NoResit = String(r.NoResit || '').slice(0, 100);
    o.Nota = String(r.Nota || '').slice(0, 500);

    if (jenis === 'Perbatuan') {
      var km = Number(r.KmSehala);
      if (!(km > 0)) throw new Error('Km sehala mesti lebih 0.');
      var kadar = Number(settings_().kadar_km || 0.65);
      var dua = r.DuaHala === false || r.DuaHala === 'TIDAK' ? 'TIDAK' : 'YA';
      o.Dari = String(r.Dari || ''); o.Ke = String(r.Ke || '');
      o.KmSehala = String(km); o.DuaHala = dua; o.Kadar = String(kadar);
      o.Amaun = round2_(km * (dua === 'YA' ? 2 : 1) * kadar).toFixed(2);
    } else {
      var amt = Number(r.Amaun);
      if (!(amt > 0)) throw new Error('Amaun mesti lebih 0.');
      o.Amaun = round2_(amt).toFixed(2);
      o.Dari = ''; o.Ke = ''; o.KmSehala = ''; o.DuaHala = ''; o.Kadar = '';
    }

    if (file && file.data) {
      var d = o.Tarikh.split('-');
      var folder = folderPath_(root_(), ['Resit', d[0], d[1], e.kod]);
      var ext = (file.name && file.name.indexOf('.') > 0) ? file.name.split('.').pop() : (file.mimeType === 'application/pdf' ? 'pdf' : 'jpg');
      var fname = [o.Tarikh, e.kod, claimant.kod, o.Amaun, o.ID].join('_') + '.' + ext;
      var blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType || 'image/jpeg', fname);
      var f = folder.createFile(blob);
      if (o.ResitFileId) { try { DriveApp.getFileById(o.ResitFileId).setTrashed(true); } catch (err) {} }
      o.ResitFileId = f.getId();
      o.ResitUrl = f.getUrl();
    }
    if (jenis === 'Resit' && !o.ResitFileId) throw new Error('Resit wajib untuk jenis Resit.');

    o.Dikemaskini = now_();
    o.OlehEmel = me.email;
    var out = H_REKOD.map(function (h) { return o[h] === undefined ? '' : o[h]; });
    if (row > 0) sh.getRange(row, 1, 1, out.length).setValues([out]);
    else sh.appendRow(out);
    return o;
  } finally {
    lock.releaseLock();
  }
}

function apiDeleteRecord(id) {
  var me = auth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName('Rekod');
    var row = findRow_(sh, id);
    if (row < 0) throw new Error('Rekod tidak dijumpai.');
    var vals = sh.getRange(row, 1, 1, H_REKOD.length).getValues()[0];
    pemilik_(me, vals[H_REKOD.indexOf('EmelClaimant')]);
    if (str_(vals[H_REKOD.indexOf('Status')]) !== STATUS.BELUM) throw new Error('Hanya rekod "Belum claim" boleh dipadam.');
    var fid = str_(vals[H_REKOD.indexOf('ResitFileId')]);
    if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (err) {} } // Trash: boleh pulih 30 hari
    sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Tukar status beberapa rekod. status: 'Dibayar' | 'Dah claim' | 'Belum claim' */
function apiSetStatus(ids, status, tarikh) {
  var me = auth_();
  if ([STATUS.BELUM, STATUS.DAH, STATUS.BAYAR].indexOf(status) < 0) throw new Error('Status tidak sah.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName('Rekod');
    var cS = H_REKOD.indexOf('Status') + 1, cB = H_REKOD.indexOf('TarikhBayar') + 1,
        cR = H_REKOD.indexOf('NoRujukan') + 1, cH = H_REKOD.indexOf('TarikhHantar') + 1, cU = H_REKOD.indexOf('Dikemaskini') + 1;
    ids.forEach(function (id) {
      var row = findRow_(sh, id);
      if (row < 0) return;
      pemilik_(me, sh.getRange(row, H_REKOD.indexOf('EmelClaimant') + 1).getValue());
      sh.getRange(row, cS).setValue(status);
      if (status === STATUS.BAYAR) sh.getRange(row, cB).setValue(tarikh || today_());
      if (status === STATUS.DAH) sh.getRange(row, cB).setValue('');
      if (status === STATUS.BELUM) { sh.getRange(row, cB).setValue(''); sh.getRange(row, cR).setValue(''); sh.getRange(row, cH).setValue(''); }
      sh.getRange(row, cU).setValue(now_());
    });
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Ambil fail resit (base64) untuk dilampirkan dalam PDF borang. */
function apiGetReceipt(fileId) {
  var me = auth_();
  var lihatSemua = (settings_().lihat_semua || 'YA').toUpperCase() === 'YA';
  var ok = readTable_('Rekod', H_REKOD).some(function (r) {
    return r.ResitFileId && r.ResitFileId === String(fileId) && (lihatSemua || r.EmelClaimant === me.email);
  });
  if (!ok) throw new Error('Resit tidak dijumpai.');
  var f = DriveApp.getFileById(fileId);
  var b = f.getBlob();
  return { name: f.getName(), mimeType: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

/** No rujukan seterusnya, cth PM-202611-FR-01 */
function apiNextRef(entiti, emelClaimant, tarikhHingga) {
  auth_();
  var u = users_().filter(function (x) { return x.email === String(emelClaimant).toLowerCase(); })[0];
  if (!u) throw new Error('Claimant tidak sah.');
  var ym = String(tarikhHingga || today_()).slice(0, 7).replace('-', '');
  var prefix = entiti_(entiti).kod + '-' + ym + '-' + u.kod + '-';
  var n = readTable_('LogBorang', H_LOG).filter(function (r) { return r.NoRujukan.indexOf(prefix) === 0; }).length;
  return prefix + ('0' + (n + 1)).slice(-2);
}

/**
 * Simpan PDF borang yang dijana di pelayar, log, dan tanda rekod "Dah claim".
 * meta: {noRujukan, entiti, emelClaimant, dari, hingga, ids[], jumlah}
 */
function apiSaveClaim(meta, pdfBase64) {
  var me = auth_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var e = entiti_(meta.entiti);
    var u = users_().filter(function (x) { return x.email === String(meta.emelClaimant).toLowerCase(); })[0];
    if (!u) throw new Error('Claimant tidak sah.');
    pemilik_(me, u.email);
    var sh = ss_().getSheetByName('Rekod');
    var cS = H_REKOD.indexOf('Status') + 1, cR = H_REKOD.indexOf('NoRujukan') + 1,
        cH = H_REKOD.indexOf('TarikhHantar') + 1, cU = H_REKOD.indexOf('Dikemaskini') + 1;

    // sahkan semua rekod masih Belum claim, entiti & claimant sama
    var rows = meta.ids.map(function (id) {
      var row = findRow_(sh, id);
      if (row < 0) throw new Error('Rekod ' + id + ' tidak dijumpai.');
      var v = sh.getRange(row, 1, 1, H_REKOD.length).getValues()[0];
      if (str_(v[cS - 1]) !== STATUS.BELUM) throw new Error('Rekod ' + id + ' sudah diclaim.');
      if (str_(v[H_REKOD.indexOf('Entiti')]) !== e.kod) throw new Error('Rekod ' + id + ' bukan entiti ' + e.kod + '.');
      if (str_(v[H_REKOD.indexOf('EmelClaimant')]).toLowerCase() !== u.email) throw new Error('Rekod ' + id + ' bukan milik claimant ini.');
      return row;
    });

    var folder = folderPath_(root_(), ['Borang', meta.hingga.slice(0, 4), e.kod]);
    var name = meta.noRujukan + '_' + e.nama.replace(/\s+/g, '') + '_' + u.nama + '.pdf';
    var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', name));
    var tarikh = today_();
    rows.forEach(function (row) {
      sh.getRange(row, cS).setValue(STATUS.DAH);
      sh.getRange(row, cR).setValue(meta.noRujukan);
      sh.getRange(row, cH).setValue(tarikh);
      sh.getRange(row, cU).setValue(now_());
    });
    ss_().getSheetByName('LogBorang').appendRow([meta.noRujukan, now_(), u.email, u.nama, e.kod, meta.dari, meta.hingga,
      String(rows.length), Number(meta.jumlah).toFixed(2), f.getId(), f.getUrl(), me.email]);
    return { url: f.getUrl(), id: f.getId(), name: name };
  } finally {
    lock.releaseLock();
  }
}

/** Eksport CSV ke Drive (folder Eksport). filters: {dari, hingga, entiti?, emelClaimant?} */
function apiExportCsv(filters) {
  var me = auth_();
  var s = settings_();
  var rows = readTable_('Rekod', H_REKOD).filter(function (r) {
    if (filters.dari && r.Tarikh < filters.dari) return false;
    if (filters.hingga && r.Tarikh > filters.hingga) return false;
    if (filters.entiti && r.Entiti !== filters.entiti) return false;
    if (filters.emelClaimant && r.EmelClaimant !== filters.emelClaimant) return false;
    if ((s.lihat_semua || 'YA').toUpperCase() !== 'YA' && r.EmelClaimant !== me.email) return false;
    return true;
  });
  var cols = ['Tarikh', 'Claimant', 'Entiti', 'Jenis', 'Kategori', 'Butiran', 'NoResit', 'Projek', 'Dari', 'Ke', 'KmSehala',
    'DuaHala', 'Kadar', 'Amaun', 'Status', 'NoRujukan', 'TarikhHantar', 'TarikhBayar', 'ResitUrl', 'ID'];
  var esc = function (v) { v = String(v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  var csv = '﻿' + [cols.join(',')].concat(rows.map(function (r) {
    return cols.map(function (c) { return esc(r[c]); }).join(',');
  })).join('\r\n');
  var folder = folderPath_(root_(), ['Eksport']);
  var name = 'ClaimTracker_' + (filters.dari || 'awal') + '_' + (filters.hingga || today_()) + (filters.entiti ? '_' + filters.entiti : '') + '.csv';
  var f = folder.createFile(Utilities.newBlob(csv, 'text/csv', name));
  return { url: f.getUrl(), name: name, count: rows.length };
}

/** Admin: tambah pengguna (cth Amalina). Kongsi Sheet & folder dengan emel itu. */
function apiAddUser(email, nama, kod, namaPenuh, jawatan) {
  admin_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Emel tidak sah.');
  kod = String(kod || '').trim().toUpperCase().slice(0, 4);
  if (!kod) throw new Error('Kod singkatan diperlukan (cth AM).');
  if (users_().some(function (u) { return u.email === email; })) throw new Error('Pengguna sudah ada.');
  if (users_().some(function (u) { return u.kod === kod; })) throw new Error('Kod ' + kod + ' sudah digunakan.');
  var ss = ss_();
  DriveApp.getFileById(ss.getId()).addEditor(email);
  root_().addEditor(email);
  ss.getSheetByName('Pengguna').appendRow([email, String(nama || '').trim(), kod, 'ahli',
    String(namaPenuh || '').trim().toUpperCase(), String(jawatan || '').trim().toUpperCase()]);
  return users_();
}

/** Admin: tukar tetapan (kadar_km, lihat_semua, kategori). */
function apiSetSetting(key, value) {
  admin_();
  if (['kadar_km', 'lihat_semua', 'kategori', 'perseni_pelulus_nama', 'perseni_pelulus_jawatan', 'tandatangan_taip'].indexOf(key) < 0)
    throw new Error('Tetapan tidak dibenarkan.');
  if (key === 'kadar_km' && !(Number(value) > 0)) throw new Error('Kadar tidak sah.');
  if (key === 'perseni_pelulus_nama' || key === 'perseni_pelulus_jawatan') value = String(value).trim().toUpperCase();
  if (key === 'lihat_semua' || key === 'tandatangan_taip') value = String(value).toUpperCase() === 'YA' ? 'YA' : 'TIDAK';
  var sh = ss_().getSheetByName('Tetapan');
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(String(value)); return true; }
  }
  sh.appendRow([key, String(value), '']);
  return true;
}

/** Setiap pengguna boleh kemas kini profil sendiri (nama penuh & jawatan untuk borang). */
function apiUpdateMe(nama, namaPenuh, jawatan) {
  var me = auth_();
  var sh = ss_().getSheetByName('Pengguna');
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === me.email) {
      if (nama) sh.getRange(i + 1, 2).setValue(String(nama).trim());
      sh.getRange(i + 1, 5).setValue(String(namaPenuh || '').trim().toUpperCase());
      sh.getRange(i + 1, 6).setValue(String(jawatan || '').trim().toUpperCase());
      return users_();
    }
  }
  throw new Error('Pengguna tidak dijumpai.');
}


