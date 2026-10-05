/**
 * Porto X - Stok Opname (SO) backend
 * Dipasang di Apps Script yang terikat ke Spreadsheet MASTER (Extensions > Apps Script).
 * Tab di master: SO, Operator, Barcode (dibuat otomatis oleh setupAdmin).
 */
const TZ = 'Asia/Jakarta';
const LOG_HEAD = ['No', 'Waktu', 'Nama', 'Barcode', 'Jenis', 'SKU', 'Nama Produk', 'Size', 'Qty', 'Catatan', 'ID', 'Status', 'Keterangan'];

/* ------------------------------ Admin ------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SO Admin')
    .addItem('Setup awal (jalankan sekali)', 'setupAdmin')
    .addItem('Buat spreadsheet untuk baris SO yang belum punya link', 'fillMissingLinks')
    .addItem('Refresh data barcode', 'clearCache')
    .addToUi();
}

function setupAdmin() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('MASTER_ID', ss.getId());
  const defs = {
    'SO': ['Tanggal Cek', 'PIC Gudang', 'Kode SO', 'Link Spreadsheet SO', 'Status'],
    'Operator': ['Nama'],
    'Barcode': ['Barcode', 'SKU', 'Nama Produk', 'Size', 'Qty', 'Jenis']
  };
  Object.keys(defs).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, defs[n].length).setValues([defs[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const so = ss.getSheetByName('SO');
  so.getRange('A2:A1000').setNumberFormat('dd/MM/yyyy');
  so.getRange('E2:E1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Aktif', 'Nonaktif'], true).build());
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onSoEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSoEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('Setup selesai. Isi tab Operator dan Barcode, lalu deploy sebagai Web App.');
}

// Trigger: admin mengisi Kode SO (kolom C) -> tanggal, spreadsheet SO, dan status terisi otomatis.
function onSoEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'SO' || e.range.getColumn() !== 3 || e.range.getRow() < 2) return;
  for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) prepareRow(sh, r);
}

function fillMissingLinks() {
  const sh = master().getSheetByName('SO');
  for (let r = 2; r <= sh.getLastRow(); r++) prepareRow(sh, r);
}

function prepareRow(sh, r) {
  const kode = String(sh.getRange(r, 3).getValue()).trim();
  if (!kode) return;
  const dupe = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues()
    .filter((v, i) => i + 2 !== r && String(v[0]).trim().toLowerCase() === kode.toLowerCase()).length;
  if (dupe) { sh.getRange(r, 4).setValue('KODE SO DOBEL - ganti kode'); return; }
  if (!sh.getRange(r, 1).getValue()) sh.getRange(r, 1).setValue(new Date());
  const link = String(sh.getRange(r, 4).getValue());
  if (!link || link.indexOf('http') !== 0) sh.getRange(r, 4).setValue(createSoSpreadsheet(kode));
  if (!sh.getRange(r, 5).getValue()) sh.getRange(r, 5).setValue('Aktif');
}

function createSoSpreadsheet(kode) {
  const ss = SpreadsheetApp.create('SO ' + kode);
  ss.setSpreadsheetTimeZone(TZ);
  const log = ss.getSheets()[0].setName('Log');
  log.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]).setFontWeight('bold');
  log.setFrozenRows(1);
  ['D:D', 'F:F', 'H:H'].forEach(a => log.getRange(a).setNumberFormat('@'));   // barcode, SKU, size sebagai teks
  log.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  const rek = ss.insertSheet('Rekap');
  rek.getRange('A1').setFormula(
    `=IFERROR(QUERY(Log!A2:M,"select F, G, H, sum(I), count(I) where F is not null and (L is null or L <> 'Batal') group by F, G, H order by F label F 'SKU', G 'Nama Produk', H 'Size', sum(I) 'Total Qty', count(I) 'Jumlah Scan'",0),"Belum ada data")`);
  rek.getRange('G1').setValue('Total Pairs').setFontWeight('bold');
  rek.getRange('G2').setFormula('=SUMIF(Log!L2:L,"<>Batal",Log!I2:I)');
  rek.setFrozenRows(1);
  return ss.getUrl();
}

function clearCache() {
  CacheService.getScriptCache().remove('map');
  try { SpreadsheetApp.getUi().alert('Data barcode akan dimuat ulang pada scan berikutnya.'); } catch (e) {}
}

/* ------------------------------ Helper ------------------------------ */

function master() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('MASTER_ID'));
}

function getMap() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('map');
  if (hit) return JSON.parse(hit);
  const rows = master().getSheetByName('Barcode').getDataRange().getValues().slice(1);
  const map = {};
  rows.forEach(r => {
    const code = String(r[0]).trim();
    if (code) map[code] = [String(r[1]).trim(), String(r[2]).trim(), String(r[3]).trim(), Number(r[4]) || 1, String(r[5]).trim()];
  });
  try { cache.put('map', JSON.stringify(map), 600); } catch (e) {}
  return map;
}

function getOperators() {
  const sh = master().getSheetByName('Operator');
  return sh.getLastRow() < 2 ? [] :
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function findSo(kode) {
  const v = master().getSheetByName('SO').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][2]).trim().toLowerCase() === String(kode).trim().toLowerCase()) {
      return { pic: v[i][1], link: String(v[i][3]), aktif: String(v[i][4]).trim() === 'Aktif' };
    }
  }
  return null;
}

function checkSo(kode) {
  const so = findSo(kode);
  if (!so) return { error: 'Kode SO tidak ditemukan' };
  if (!so.aktif) return { error: 'SO sedang nonaktif. Hubungi admin.' };
  if (so.link.indexOf('http') !== 0) return { error: 'Spreadsheet SO belum dibuat. Hubungi admin.' };
  return { so };
}

const safe = s => /^[=+\-@]/.test(String(s)) ? "'" + s : String(s);

/* ------------------------------ API ------------------------------ */

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = { init: apiInit, start: apiStart, scan: apiScan, note: apiNote, cancel: apiCancel }[req.action];
    out = fn ? fn(req) : { ok: false, error: 'Aksi tidak dikenal' };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiInit() {
  return { ok: true, operators: getOperators() };
}

function apiStart(req) {
  if (getOperators().indexOf(req.nama) < 0) return { ok: false, error: 'Nama tidak terdaftar' };
  const c = checkSo(req.kode);
  if (c.error) return { ok: false, error: c.error };
  return { ok: true, pic: c.so.pic, map: getMap() };
}

function apiScan(req) {
  const c = checkSo(req.kode);
  if (c.error) return { ok: false, error: c.error };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const log = SpreadsheetApp.openByUrl(c.so.link).getSheetByName('Log');
    const last = log.getLastRow();
    const ids = new Set(last > 1 ? log.getRange(2, 11, last - 1, 1).getValues().map(r => r[0]) : []);
    const map = getMap();
    let no = last > 1 ? last - 1 : 0;
    const rows = [];
    req.scans.forEach(s => {
      if (ids.has(s.id)) return;                       // sudah tersimpan (kirim ulang)
      const code = String(s.code).trim();
      const m = map[code];
      rows.push([++no, new Date(s.t), req.nama, code,
        m ? m[4] : 'Tidak Dikenal', m ? m[0] : 'TIDAK DIKENAL', m ? m[1] : safe(code),
        m ? m[2] : '', m ? m[3] : 1, safe(s.note || ''), s.id,
        s.cancel ? 'Batal' : 'Aktif', s.cancel ? 'Dibatalkan oleh ' + req.nama + ' sebelum terkirim' : '']);
    });
    if (rows.length) log.getRange(last + 1, 1, rows.length, LOG_HEAD.length).setValues(rows);
    return { ok: true, saved: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function apiNote(req) {
  const c = checkSo(req.kode);
  if (c.error) return { ok: false, error: c.error };
  const log = SpreadsheetApp.openByUrl(c.so.link).getSheetByName('Log');
  const f = log.createTextFinder(req.id).matchEntireCell(true).findNext();
  if (!f) return { ok: false, error: 'Scan belum tersimpan' };
  log.getRange(f.getRow(), 10).setValue(safe(req.note || ''));
  return { ok: true };
}

// Operator hanya bisa membatalkan scan miliknya sendiri. Baris tidak dihapus: ditandai "Batal" dan Rekap mengabaikannya.
function apiCancel(req) {
  const c = checkSo(req.kode);
  if (c.error) return { ok: false, error: c.error };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const log = SpreadsheetApp.openByUrl(c.so.link).getSheetByName('Log');
    const f = log.createTextFinder(req.id).matchEntireCell(true).findNext();
    if (!f) return { ok: false, error: 'Scan belum tersimpan' };
    const row = f.getRow();
    if (String(log.getRange(row, 3).getValue()) !== req.nama) return { ok: false, error: 'Hanya pemilik scan yang bisa membatalkan' };
    log.getRange(row, 12, 1, 2).setValues([['Batal', 'Dibatalkan oleh ' + req.nama + ' ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss')]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
