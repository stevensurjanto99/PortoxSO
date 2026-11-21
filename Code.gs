/**
 * Porto X - Stok Opname (SO) backend  (versi dengan Upload Stok Existing + Banding)
 * Dipasang di Apps Script yang terikat ke Spreadsheet MASTER (Extensions > Apps Script).
 * Tab di master: SO, Operator, Barcode (dibuat otomatis oleh setupAdmin).
 * File tambahan yang dibutuhkan: Upload.html (sidebar upload stok).
 */
const TZ = 'Asia/Jakarta';
const LOG_HEAD = ['No', 'Waktu', 'Nama', 'Barcode', 'Jenis', 'SKU', 'Nama Produk', 'Size', 'Qty', 'Catatan', 'ID'];
const UNKNOWN = 'TIDAK DIKENAL';
const UNKNOWN_NOTE = 'Barcode tidak ada di master';

/* ------------------------------ Admin ------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SO Admin')
    .addItem('Setup awal (jalankan sekali)', 'setupAdmin')
    .addItem('Buat spreadsheet untuk baris SO yang belum punya link', 'fillMissingLinks')
    .addItem('Upload stok existing (Excel)', 'showUpload')
    .addItem('Refresh data barcode', 'clearCache')
    .addItem('Cek konfigurasi (diagnosa)', 'diagnose')
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
  // Barcode, SKU, Nama, Size, Jenis = teks (cegah barcode jadi 8.99E+12 / "35/36" jadi tanggal)
  const bc = ss.getSheetByName('Barcode');
  ['A:D', 'F:F'].forEach(a => bc.getRange(a).setNumberFormat('@'));
  ss.getSheetByName('Operator').getRange('A:A').setNumberFormat('@');
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
  const link = String(sh.getRange(r, 4).getValue());
  const hasLink = link.indexOf('http') === 0;
  const dupe = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues()
    .filter((v, i) => i + 2 !== r && String(v[0]).trim().toLowerCase() === kode.toLowerCase()).length;
  // FIX: baris yang sudah punya link tidak boleh ditimpa tulisan "DOBEL"
  if (dupe && !hasLink) { sh.getRange(r, 4).setValue('KODE SO DOBEL - ganti kode'); return; }
  if (!sh.getRange(r, 1).getValue()) sh.getRange(r, 1).setValue(new Date());
  if (!hasLink) sh.getRange(r, 4).setValue(createSoSpreadsheet(kode));
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
  ensureLogHighlight(log);
  const rek = ss.insertSheet('Rekap');
  rek.getRange('A1').setFormula(
    `=IFERROR(QUERY(Log!A2:K,"select F, G, H, sum(I), count(I) where F is not null group by F, G, H order by F label F 'SKU', G 'Nama Produk', H 'Size', sum(I) 'Total Qty', count(I) 'Jumlah Scan'",0),"Belum ada data")`);
  rek.getRange('G1').setValue('Total Pairs').setFontWeight('bold');
  rek.getRange('G2').setFormula('=SUM(Log!I2:I)');
  rek.setFrozenRows(1);
  return ss.getUrl();
}

// Highlight oranye untuk baris scan yang barcode-nya tidak dikenal
function ensureLogHighlight(log) {
  const f = `=$F2="${UNKNOWN}"`;
  const exists = log.getConditionalFormatRules().some(r => {
    const c = r.getBooleanCondition();
    return c && c.getCriteriaValues()[0] === f;
  });
  if (exists) return;
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f).setBackground('#fce5cd')
    .setRanges([log.getRange('A2:K5000')]).build();
  log.setConditionalFormatRules(log.getConditionalFormatRules().concat(rule));
}

function clearCache() {
  cacheRemoveBig('map');
  try { SpreadsheetApp.getUi().alert('Data barcode akan dimuat ulang pada scan berikutnya.'); } catch (e) {}
}

/* ------------------------------ Diagnosa ------------------------------ */

function diagnose() {
  const out = [];
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_ID');
  out.push('MASTER_ID: ' + (id ? 'ada' : 'KOSONG -> jalankan "Setup awal"'));
  try {
    const ops = getOperators();
    out.push('Operator terbaca: ' + ops.length + (ops.length ? ' (' + ops.join(', ') + ')' : ' -> cek tab Operator, nama harus di kolom A'));
  } catch (e) { out.push('Operator ERROR: ' + e); }
  try {
    const map = getMap();
    const keys = Object.keys(map);
    out.push('Barcode terbaca: ' + keys.length);
    if (keys.length) out.push('Contoh: ' + keys[0] + ' -> ' + JSON.stringify(map[keys[0]]));
  } catch (e) { out.push('Barcode ERROR: ' + e); }
  try {
    const v = master().getSheetByName('SO').getDataRange().getValues().slice(1);
    out.push('Baris SO: ' + v.length + ', aktif: ' + v.filter(r => String(r[4]).trim() === 'Aktif').length);
  } catch (e) { out.push('SO ERROR: ' + e); }
  out.push('Trigger onSoEdit: ' + (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onSoEdit') ? 'ada' : 'TIDAK ADA'));
  SpreadsheetApp.getUi().alert(out.join('\n'));
}

/* ------------------------------ Helper ------------------------------ */

function master() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_ID');
  if (!id) throw new Error('MASTER_ID belum diset. Jalankan menu SO Admin > Setup awal.');
  return SpreadsheetApp.openById(id);
}

// Cache besar: dipecah per chunk (batas CacheService 100 KB per key)
function cachePutBig(key, str) {
  const cache = CacheService.getScriptCache();
  const size = 30000, n = Math.ceil(str.length / size), obj = {};
  for (let i = 0; i < n; i++) obj[key + '_' + i] = str.substr(i * size, size);
  obj[key + '_n'] = String(n);
  try { cache.putAll(obj, 600); } catch (e) {}
}
function cacheGetBig(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key + '_n'));
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const got = cache.getAll(keys);
  let s = '';
  for (let i = 0; i < n; i++) {
    if (got[key + '_' + i] == null) return null;
    s += got[key + '_' + i];
  }
  return s;
}
function cacheRemoveBig(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key + '_n')) || 0;
  const keys = [key + '_n'];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  cache.removeAll(keys);
}

function getMap() {
  const hit = cacheGetBig('map');
  if (hit) return JSON.parse(hit);
  const rows = master().getSheetByName('Barcode').getDataRange().getValues().slice(1);
  const map = {};
  rows.forEach(r => {
    const code = String(r[0]).trim();
    if (code) map[code] = [String(r[1]).trim(), String(r[2]).trim(), String(r[3]).trim(), Number(r[4]) || 1, String(r[5]).trim()];
  });
  cachePutBig('map', JSON.stringify(map));
  return map;
}

// Nama operator: kolom A tab Operator. Header "Nama" (jika ada) dilewati; baris 1 tanpa header tetap terbaca.
function getOperators() {
  const sh = master().getSheetByName('Operator');
  if (!sh) throw new Error('Tab "Operator" tidak ditemukan');
  const names = sh.getDataRange().getValues().map(r => String(r[0]).trim()).filter(Boolean);
  if (names.length && names[0].toLowerCase() === 'nama') names.shift();
  return names;
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
    const fn = { init: apiInit, start: apiStart, scan: apiScan, note: apiNote, upsertBarcodes: apiUpsertBarcodes }[req.action];
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
  const nama = String(req.nama || '').trim().toLowerCase();
  if (getOperators().map(n => n.toLowerCase()).indexOf(nama) < 0) return { ok: false, error: 'Nama tidak terdaftar' };
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
      ids.add(s.id);                                   // FIX: cegah ID kembar di request yang sama
      const code = String(s.code).trim();
      const m = map[code];
      const note = safe(s.note || '') || (m ? '' : UNKNOWN_NOTE);
      rows.push([++no, new Date(s.t), req.nama, code,
        m ? m[4] : 'Tidak Dikenal', m ? m[0] : UNKNOWN, m ? m[1] : safe(code),
        m ? m[2] : '', m ? m[3] : 1, note, s.id]);
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

/* ------------------------------ Sinkron barcode dari aplikasi label ------------------------------ */
/**
 * Script Properties (Project Settings): LABEL_SYNC_KEY = kunci rahasia (sama dengan di aplikasi label).
 * Urutan kolom tab Barcode mengikuti getMap(): Barcode, SKU, Nama Produk, Size, Qty, Jenis.
 */
const JENIS_LABEL = { satuan: 'Satuan', master: 'Master' };   // sesuaikan dengan isi kolom Jenis yang sudah ada

function apiUpsertBarcodes(req) {
  const key = PropertiesService.getScriptProperties().getProperty('LABEL_SYNC_KEY');
  if (!key || req.key !== key) return { ok: false, error: 'kunci salah' };

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = master().getSheetByName('Barcode');
    const last = sh.getLastRow();
    const pos = {};                                          // barcode -> nomor baris
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => {
        const c = String(r[0]).trim();
        if (c) pos[c] = i + 2;
      });
    }
    let added = 0, updated = 0, next = last + 1;
    const fresh = [];
    (req.items || []).forEach(it => {
      const code = String(it.barcode || '').trim();
      if (!code) return;
      const row = [code, String(it.sku || '').trim(), String(it.nama || '').trim(),
        String(it.size || '').trim(), Number(it.qty) || 1,
        JENIS_LABEL[String(it.jenis).toLowerCase()] || String(it.jenis || '')];
      if (pos[code]) {
        sh.getRange(pos[code], 1, 1, 6).setNumberFormats([['@', '@', '@', '@', '0', '@']]).setValues([row]);
        updated++;
      } else {
        pos[code] = next + fresh.length;
        fresh.push(row);
        added++;
      }
    });
    if (fresh.length) {
      const rg = sh.getRange(next, 1, fresh.length, 6);
      rg.setNumberFormats(fresh.map(() => ['@', '@', '@', '@', '0', '@'])).setValues(fresh);   // teks: cegah "35/36" jadi tanggal
    }
    SpreadsheetApp.flush();
    cacheRemoveBig('map');                                   // supaya scanner langsung melihat barcode baru
    return { ok: true, added: added, updated: updated };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ Upload stok existing + Banding ------------------------------ */

function showUpload() {
  const html = HtmlService.createHtmlOutputFromFile('Upload').setTitle('Upload Stok Existing');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Dipanggil sidebar: daftar Kode SO yang sudah punya spreadsheet
function getSoList() {
  const v = master().getSheetByName('SO').getDataRange().getValues().slice(1);
  return v.filter(r => String(r[2]).trim() && String(r[3]).indexOf('http') === 0)
    .map(r => ({ kode: String(r[2]).trim(), status: String(r[4]).trim() }));
}

// Dipanggil sidebar: payload = { fileName, rows: [[sku, nama, qty], ...] }
function uploadStock(kode, payload) {
  const so = findSo(kode);
  if (!so) throw new Error('Kode SO tidak ditemukan');
  if (so.link.indexOf('http') !== 0) throw new Error('Spreadsheet SO belum dibuat');

  // gabungkan SKU kembar (jumlahkan)
  const agg = {}, order = [];
  (payload.rows || []).forEach(r => {
    const sku = String(r[0] || '').trim();
    if (!sku) return;
    if (!agg[sku]) { agg[sku] = [sku, String(r[1] || '').trim(), 0]; order.push(sku); }
    agg[sku][2] += Number(r[2]) || 0;
  });
  const data = order.map(k => agg[k]);
  if (!data.length) throw new Error('Tidak ada baris SKU yang valid di file');

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = SpreadsheetApp.openByUrl(so.link);
    let st = ss.getSheetByName('Stok Existing');
    if (!st) st = ss.insertSheet('Stok Existing');
    st.clear();                                              // upload ulang = timpa
    st.getRange(1, 1, 1, 3).setValues([['SKU', 'Nama Produk (Odoo)', 'Available Stock']]).setFontWeight('bold');
    st.getRange(2, 1, data.length, 1).setNumberFormat('@');
    st.getRange(2, 1, data.length, 3).setValues(data);
    st.setFrozenRows(1);
    st.autoResizeColumns(1, 3);
    const info = 'Stok Excel: ' + (payload.fileName || '-') + '  |  diupload ' +
      Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + '  |  ' + data.length + ' SKU';
    ensureBanding(ss, info);
    ensureLogHighlight(ss.getSheetByName('Log'));
    SpreadsheetApp.flush();
    return { ok: true, skus: data.length, total: data.reduce((a, r) => a + r[2], 0), url: ss.getUrl() };
  } finally {
    lock.releaseLock();
  }
}

function ensureBanding(ss, info) {
  let sh = ss.getSheetByName('Banding');
  const fresh = !sh;
  if (!sh) sh = ss.insertSheet('Banding');
  const keep = fresh ? true : sh.getRange('A1').getValue();   // pertahankan status checkbox
  sh.getRange('A1:K4').clearContent();
  sh.getRange('A5:G').clearContent();
  sh.getRange('A1').insertCheckboxes().setValue(keep === false ? false : true);
  sh.getRange('B1').setValue('<- Centang: bandingkan SELURUH stok Excel (SKU belum discan = 0). Tidak dicentang: hanya SKU yang sudah discan.')
    .setFontStyle('italic');
  sh.getRange('A2').setValue(info);
  sh.getRange('A3').setFormula('="Total Scan: "&SUM(D5:D)&"   |   Total Stok Excel (yang dibandingkan): "&SUM(E5:E)&"   |   Selisih: "&SUM(F5:F)');
  sh.getRange('A4:G4').setValues([['SKU', 'Nama Produk', 'Size', 'Qty Scan', 'Stok Excel', 'Selisih (Scan - Excel)', 'Status']])
    .setFontWeight('bold').setBackground('#d9d9d9');

  sh.getRange('A5').setFormula(
    `=IFERROR(LET(` +
    `s,IFERROR(QUERY(Log!A2:K,"select F where F is not null and F <> '${UNKNOWN}' group by F",0),""),` +
    `e,IFERROR(FILTER('Stok Existing'!A2:A,(('Stok Existing'!A2:A<>"")*($A$1=TRUE))=1),""),` +
    `u,UNIQUE(VSTACK(s,e)),` +
    `SORT(FILTER(u,u<>""))),"")`);
  sh.getRange('B5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",IFERROR(VLOOKUP($A$5:$A,'Stok Existing'!A:B,2,FALSE),IFERROR(VLOOKUP($A$5:$A,Log!F:G,2,FALSE),""))))`);
  sh.getRange('C5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",IFERROR(VLOOKUP($A$5:$A,Log!F:H,3,FALSE),IFERROR(REGEXEXTRACT($B$5:$B,"\\((\\d+/\\d+),"),""))))`);
  sh.getRange('D5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",SUMIF(Log!F:F,$A$5:$A,Log!I:I)))`);
  sh.getRange('E5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",SUMIF('Stok Existing'!A:A,$A$5:$A,'Stok Existing'!C:C)))`);
  sh.getRange('F5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",$D$5:$D-$E$5:$E))`);
  sh.getRange('G5').setFormula(
    `=ARRAYFORMULA(IF($A$5:$A="","",IF(COUNTIF('Stok Existing'!A:A,$A$5:$A)=0,"Tidak ada di Excel",IF($F$5:$F=0,"Cocok",IF($D$5:$D=0,"Belum discan",IF($F$5:$F<0,"Kurang","Lebih"))))))`);

  // Barcode tidak dikenal: blok terpisah + catatan + highlight
  sh.getRange('I3').setValue('Scan barcode TIDAK DIKENAL').setFontWeight('bold');
  sh.getRange('I4').setFormula(
    `=IFERROR(QUERY(Log!A2:K,"select D, count(D) where F = '${UNKNOWN}' group by D label D 'Barcode', count(D) 'Jumlah Scan'",1),"Tidak ada")`);
  sh.getRange('K4').setValue('Catatan').setFontWeight('bold');
  sh.getRange('K5').setFormula(`=ARRAYFORMULA(IF($I$5:$I="","","${UNKNOWN_NOTE} - cek/tambahkan di tab Barcode"))`);
  sh.getRange('I4:K4').setFontWeight('bold').setBackground('#fce5cd');

  sh.getRange('A5:A').setNumberFormat('@');
  sh.getRange('C5:C').setNumberFormat('@');
  sh.getRange('I5:I').setNumberFormat('@');
  sh.getRange('D5:F').setNumberFormat('0');
  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 360); sh.setColumnWidth(3, 60);
  sh.setColumnWidths(4, 3, 110); sh.setColumnWidth(7, 140);
  sh.setColumnWidth(9, 150); sh.setColumnWidth(10, 100); sh.setColumnWidth(11, 330);

  const rng = sh.getRange('A5:G3000');
  const rule = (f, bg) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(bg).setRanges([rng]).build();
  sh.setConditionalFormatRules([
    rule('=$G5="Cocok"', '#d9ead3'),
    rule('=$G5="Kurang"', '#f4cccc'),
    rule('=$G5="Belum discan"', '#f4cccc'),
    rule('=$G5="Lebih"', '#fff2cc'),
    rule('=$G5="Tidak ada di Excel"', '#d9d2e9'),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I5<>""')
      .setBackground('#fce5cd').setRanges([sh.getRange('I5:K500')]).build()
  ]);
}
