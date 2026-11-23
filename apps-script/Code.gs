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
      apiAddUser: apiAddUser, apiSetSetting: apiSetSetting, apiUpdateMe: apiUpdateMe, apiAset: apiAset, apiLogout: apiLogout,
      apiScanResit: apiScanResit
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

    // Elak rekod double: hantaran semula borang yang sama (kunci sama) pulangkan rekod yang sudah disimpan.
    var cache = CacheService.getScriptCache();
    var kunci = !r.ID && /^[A-Za-z0-9_-]{8,64}$/.test(String(r.KunciHantar || '')) ? 'K_' + me.email + '_' + r.KunciHantar : '';
    if (kunci) {
      var dahId = cache.get(kunci);
      var dahRow = dahId ? findRow_(sh, dahId) : -1;
      if (dahRow > 0) {
        var dah = {}, dv = sh.getRange(dahRow, 1, 1, H_REKOD.length).getValues()[0];
        H_REKOD.forEach(function (h, j) { dah[h] = str_(dv[j]); });
        dah._ulang = true;
        return dah;
      }
    }

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
    if (kunci) cache.put(kunci, o.ID, 21600); // ingat 6 jam
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


/* ------------------------------------------------------------------ */
/* Baca resit (OCR Google Drive)                                      */
/* ------------------------------------------------------------------ */

/**
 * Baca gambar/PDF resit dengan OCR Google Drive, kemudian cuba kenal pasti merchant, jumlah, tarikh dan no. resit.
 * Hasilnya hanya CADANGAN; pengguna semak dan betulkan sebelum simpan.
 * file: {mimeType, data(base64)}
 */
function apiScanResit(file) {
  auth_();
  if (!file || !file.data) throw new Error('Tiada fail resit.');
  if (file.data.length > 14e6) throw new Error('Fail terlalu besar untuk dibaca.');
  var mime = file.mimeType === 'application/pdf' ? 'application/pdf' : (file.mimeType || 'image/jpeg');
  var blob = Utilities.newBlob(Utilities.base64Decode(file.data), mime, 'resit-ocr');
  var meta = { name: 'tmp-ocr-' + Date.now(), mimeType: 'application/vnd.google-apps.document' };
  var rootId = PropertiesService.getScriptProperties().getProperty('ROOT_ID');
  if (rootId) meta.parents = [rootId];
  var doc = Drive.Files.create(meta, blob, { ocrLanguage: 'en', fields: 'id' });
  var text = '';
  try {
    var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc.id + '/export?mimeType=text%2Fplain', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('OCR gagal (' + res.getResponseCode() + ').');
    text = res.getContentText();
  } finally {
    try { Drive.Files.remove(doc.id); } catch (e) { try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e2) {} }
  }
  var out = parseResit_(text, today_());
  out.teks = String(text).slice(0, 3000);
  return out;
}

/** Tafsir teks OCR resit. Pulangkan {merchant, amaun, tarikh(yyyy-MM-dd), noResit}; medan kosong jika tidak pasti. */
function parseResit_(text, today) {
  var lines = String(text || '').replace(/\r/g, '').split('\n')
    .map(function (l) { return l.replace(/[ \t]+/g, ' ').replace(/\s{2,}/g, ' ').trim(); })
    .filter(function (l) { return l.length; });
  var MONEY = /(?:RM\s*)?(\d{1,3}(?:,\d{3})+|\d+)[.,](\d{2})(?!\d)/gi;
  function monies(l) {
    var out = [], m; MONEY.lastIndex = 0;
    while ((m = MONEY.exec(l))) out.push(Number(m[1].replace(/,/g, '') + '.' + m[2]));
    return out;
  }
  function onlyMoney(l) { return /^(RM\s*)?[\d,]+[.,]\d{2}$/i.test(l.replace(/\s+/g, ' ').trim()) || /^RM\s*$/i.test(l); }
  var EXCL = /change|baki|\bcash\b|tunai|tender|bayaran tunai|rounding|pembundaran|discount|diskaun|saving/i;
  var SUB = /sub\s*-?\s*total|total\s*(items?|qty|quantity|disc|tax|sst|gst|saving|point)/i;
  var P1 = /net\s*total|grand\s*total|total\s*amount|amount\s*due|jumlah\s*besar|jumlah\s*bersih|total\s*payable|total\s*\(?rm\)?|rounded\s*total|total\s*incl|total\s*sales|amaun\s*perlu\s*dibayar/i;
  var P2 = /\btotal\b|\bjumlah\b|\bamount\b/i;

  function amountAt(i) {
    var v = monies(lines[i]);
    if (v.length) return v[v.length - 1];
    for (var k = i + 1; k < Math.min(lines.length, i + 4); k++) {
      if (P2.test(lines[k]) || EXCL.test(lines[k])) break;
      var w = monies(lines[k]);
      if (w.length && onlyMoney(lines[k])) return w[w.length - 1];
      if (!onlyMoney(lines[k])) break;
    }
    return null;
  }
  var amaun = null;
  [P1, P2].some(function (re) {
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (!re.test(l) || SUB.test(l) || EXCL.test(l)) continue;
      var a = amountAt(i);
      if (a !== null && a > 0) { amaun = a; return true; }
    }
    return false;
  });
  if (amaun === null) {
    var best = 0;
    lines.forEach(function (l) { if (!EXCL.test(l) && !SUB.test(l)) monies(l).forEach(function (v) { if (v > best && v < 100000) best = v; }); });
    if (best > 0) amaun = best;
  }

  // tarikh
  var MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, mei: 5, jun: 6, jul: 7, aug: 8, ogo: 8, sep: 9, oct: 10, okt: 10, nov: 11, dec: 12, dis: 12 };
  var tarikh = '';
  var t0 = new Date(today + 'T00:00:00Z').getTime();
  function ok(y, m, d) {
    if (y < 100) y += 2000;
    if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 2000)) return '';
    var t = Date.UTC(y, m - 1, d);
    if (t > t0 + 2 * 864e5 || t < t0 - 730 * 864e5) return '';
    return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
  }
  for (var i = 0; i < lines.length && !tarikh; i++) {
    var l = lines[i], m;
    if ((m = /(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/.exec(l))) tarikh = ok(+m[1], +m[2], +m[3]);
    if (!tarikh && (m = /\b(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})\b/.exec(l))) tarikh = ok(+m[3], +m[2], +m[1]);
    if (!tarikh && (m = /\b(\d{1,2})[\s\-\/.]*([A-Za-z]{3})[A-Za-z]*[\s\-\/.,]*(\d{2,4})\b/.exec(l)) && MON[m[2].toLowerCase()]) tarikh = ok(+m[3], MON[m[2].toLowerCase()], +m[1]);
  }

  // no. resit
  var noResit = '';
  var NOKEY = /(receipt|resit|invoice|inv|bill|doc(?:ument)?|order|check|cheque|trans(?:action)?|ref|slip)\s*(no\.?|num(?:ber)?|#)|no\.?\s*(resit|invois|bil)|^(r(?:ct|cpt)?|no|bill)\s*[:#]/i;
  for (var j = 0; j < lines.length && !noResit; j++) {
    if (!NOKEY.test(lines[j])) continue;
    var after = lines[j].split(/[:#]/).slice(1).join(':').trim();
    var mm = /[A-Z0-9][A-Z0-9\-\/]{2,}/i.exec(after);
    if (mm && /\d/.test(mm[0])) noResit = mm[0];
  }

  // merchant: baris bermakna pertama di bahagian atas
  var SKIP = /receipt|resit|invoice|invois|tax|cukai|welcome|selamat|terima kasih|thank|tel\b|tel:|phone|fax|whatsapp|gst|sst|reg\b|reg\.|co\.?\s*no|www|http|@|^\W+$|cashier|table|date|tarikh|time|masa/i;
  var ADDR = /\bjalan\b|\bjln\b|\blot\b|\btaman\b|\bseksyen\b|\bsection\b|\bno\.?\s*\d|\b\d{5}\b|floor|aras|kuala lumpur|selangor|w\.p\./i;
  var LEGAL = /sdn\.?\s*bhd|bhd\b|enterprise|\bplt\b|holdings|trading|\(\d+.*\)/i;
  function letters(l) { return (l.match(/[A-Za-z]/g) || []).length; }
  function digits(l) { return (l.match(/\d/g) || []).length; }
  var merchant = '', top = lines.slice(0, 8);
  for (var q = 0; q < top.length; q++) {
    var c = top[q];
    if (letters(c) < 3 || digits(c) > letters(c) * 0.4 || SKIP.test(c) || ADDR.test(c) || monies(c).length) continue;
    merchant = c;
    var nx = top[q + 1];
    if (!LEGAL.test(c) && nx && !digits(nx) && letters(nx) >= 3 && nx.split(' ').length <= 3 && c.split(' ').length <= 4 &&
        !LEGAL.test(nx) && !SKIP.test(nx) && !ADDR.test(nx)) merchant += ' ' + nx;
    break;
  }
  if (merchant && merchant === merchant.toUpperCase())
    merchant = merchant.toLowerCase().replace(/(^|[\s(\-\/])([a-z])/g, function (a, b, ch) { return b + ch.toUpperCase(); });
  merchant = merchant.replace(/\s+/g, ' ').slice(0, 120);

  return { merchant: merchant, amaun: amaun === null ? '' : amaun.toFixed(2), tarikh: tarikh, noResit: noResit };
}
