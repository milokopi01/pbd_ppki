/**
 * =====================================================================
 *  e-PBD : SISTEM PELAPORAN PENTAKSIRAN BILIK DARJAH (Google Apps Script)
 *  Backend API — membaca & menulis data terus dari Google Sheet.
 *  TIADA data disimpan dalam kod ini. Semua data (termasuk kata laluan
 *  admin) berada dalam Google Sheet: 01_TETAPAN, 02_MURID, 03_MATA_PELAJARAN,
 *  04_MARKAH, 05_TAFSIRAN_TP, 06_PENYATAAN_KEMAHIRAN, 07_BAHASA_LABEL,
 *  08_RUJUKAN, 09_PANDUAN, 10_LOG, 99_ARKIB_ASAL.
 *
 *  CARA PASANG:
 *  1. Upload DATA_PELAPORAN_PBD.xlsx ke Google Drive > Buka dengan Google Sheets
 *     (File > Save as Google Sheets).
 *  2. Dalam Google Sheet: Extensions > Apps Script. Padam kod asal, tampal fail ini.
 *  3. Deploy > New deployment > Type: Web app
 *       Execute as : Me
 *       Who has access : Anyone
 *  4. Salin Web App URL (berakhir /exec) dan tampal di baris ke-5 index.html.
 *  5. Kelas baharu / tahun baharu: buka menu "Kelas & Tahun" dalam aplikasi.
 *     Sheet MURID & MARKAH baharu dicipta secara automatik; data tahun lepas kekal.
 *  PENTING: selepas tampal kod baharu, buat Deploy > Manage deployments > Edit > New version.
 *  (Jika skrip TIDAK dibuka dari dalam Google Sheet, isi SPREADSHEET_ID di bawah.)
 * =====================================================================
 */
const SPREADSHEET_ID = ''; // kosongkan jika skrip dibuka dari Extensions > Apps Script

const SH = {
  KELAS: '00_KELAS',
  TETAPAN: '01_TETAPAN', MURID: '02_MURID', MP: '03_MATA_PELAJARAN', MARKAH: '04_MARKAH',
  TAFSIRAN: '05_TAFSIRAN_TP', KEMAHIRAN: '06_PENYATAAN_KEMAHIRAN', BAHASA: '07_BAHASA_LABEL',
  RUJUKAN: '08_RUJUKAN', PANDUAN: '09_PANDUAN', LOG: '10_LOG'
};
const MARKAH_HEADERS = ['ID_MURID','KOD_MP','K1','K2','K3','K4','K5','K6','K7','K8','K9','K10','K11','K12','TP_KESELURUHAN','ULASAN_GURU_MP','DIKEMASKINI'];

const KELAS_HEADERS = ['ID_KELAS','SESI','TAHUN','NAMA_KELAS','NAMA_GURU_KELAS','SHEET_MURID','SHEET_MARKAH','STATUS','DICIPTA'];
// Tetapan yang berbeza bagi setiap kelas (disimpan dalam 00_KELAS, bukan 01_TETAPAN)
const KELAS_KEYS = { SESI: 'SESI', TAHUN: 'TAHUN', KELAS: 'NAMA_KELAS', NAMA_GURU_KELAS: 'NAMA_GURU_KELAS' };

function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak dijumpai. Pastikan fail data betul.');
  return sh;
}
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- Baca sheet sebagai senarai objek ---------- */
function readTable_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 1) return [];
  const headers = values[0].map(String);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(c => c === '' || c === null)) continue;
    const o = {};
    headers.forEach((h, j) => {
      let v = r[j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[h] = v;
    });
    o._row = i + 1;
    rows.push(o);
  }
  return rows;
}
function settings_() {
  const o = {};
  readTable_(SH.TETAPAN).forEach(r => { o[String(r.KUNCI).trim()] = r.NILAI; });
  return o;
}
function checkPass_(pw) {
  const real = String(settings_().ADMIN_PASSWORD || '').trim();
  if (!real || String(pw || '').trim() !== real) throw new Error('Kata laluan admin salah. Akses edit ditolak.');
}
function log_(action, detail) {
  try { sheet_(SH.LOG).appendRow([new Date(), action, String(detail || '').slice(0, 500)]); } catch (e) {}
}

/* ---------- GET ---------- */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'ping') return out_({ ok: true, msg: 'e-PBD API aktif' });
    if (action === 'getAll') return out_({ ok: true, data: getAll_(e.parameter.kelas) });
    return out_({ ok: false, error: 'Tindakan tidak dikenali: ' + action });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

/* ====================== KELAS & TAHUN (SESI) ======================
 * Setiap kelas bagi setiap tahun mempunyai 2 sheet sendiri:
 *   "MURID 2026 2 AL BIRUNI"  dan  "MARKAH 2026 2 AL BIRUNI"
 * Senarai semua kelas (termasuk tahun-tahun lepas) disimpan dalam 00_KELAS.
 * Data tahun lepas tidak dipadam, jadi sejarah PBD sentiasa boleh dibuka semula.
 */
function kelasSheet_() {
  const ss = ss_(); let sh = ss.getSheetByName(SH.KELAS);
  if (!sh) {
    sh = ss.insertSheet(SH.KELAS, 0);
    sh.getRange(1, 1, 1, KELAS_HEADERS.length).setValues([KELAS_HEADERS]).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() < 2) { // pindahan automatik: daftar data sedia ada sebagai kelas pertama
    const t = settings_();
    if (ss.getSheetByName(SH.MURID)) {
      const sesi = String(t.SESI || new Date().getFullYear()), tahun = String(t.TAHUN || ''), nama = String(t.KELAS || 'KELAS 1');
      sh.appendRow([kelasId_(sesi, tahun, nama), sesi, tahun, nama, t.NAMA_GURU_KELAS || '', SH.MURID, SH.MARKAH, 'AKTIF', new Date()]);
    }
  }
  return sh;
}
function kelasId_(sesi, tahun, nama) { return (String(sesi) + '-' + String(tahun) + '-' + String(nama)).toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, ''); }
function senaraiKelas_() { kelasSheet_(); return readTable_(SH.KELAS); }
function kelas_(id) {
  const all = senaraiKelas_();
  if (!all.length) throw new Error('Tiada kelas berdaftar. Sila tambah kelas dahulu.');
  if (!id) { // lalai: kelas aktif terbaharu
    const sorted = all.slice().sort((a, b) => (+b.SESI || 0) - (+a.SESI || 0));
    return sorted.find(k => String(k.STATUS).toUpperCase() !== 'ARKIB') || sorted[0];
  }
  const k = all.find(x => String(x.ID_KELAS) === String(id));
  if (!k) throw new Error('Kelas "' + id + '" tidak dijumpai.');
  return k;
}
function safeName_(s) { return String(s).replace(/[\[\]\*\?\/\\:']/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90); }

function createKelas_(d) {
  const sesi = String(d.SESI || '').trim(), tahun = String(d.TAHUN || '').trim(), nama = String(d.NAMA_KELAS || '').trim().toUpperCase();
  if (!sesi || !nama) throw new Error('Sesi (tahun kalendar) dan nama kelas wajib diisi.');
  const id = kelasId_(sesi, tahun, nama);
  if (senaraiKelas_().some(k => k.ID_KELAS === id)) throw new Error('Kelas ' + tahun + ' ' + nama + ' bagi sesi ' + sesi + ' sudah wujud.');
  const ss = ss_(); const label = safeName_(sesi + ' ' + tahun + ' ' + nama);
  const shM = ss.insertSheet('MURID ' + label), shK = ss.insertSheet('MARKAH ' + label);
  // tajuk lajur murid disalin dari 02_MURID supaya struktur sama
  const tmpl = ss.getSheetByName(SH.MURID);
  const mh = tmpl ? headers_(tmpl) : ['ID_MURID','BIL','NAMA_MURID','NAMA_RINGKAS','NO_MYKID','JANTINA','ULASAN_GURU_KELAS','AKTIF','CATATAN'];
  [[shM, mh], [shK, MARKAH_HEADERS]].forEach(([sh, h]) => { sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff'); sh.setFrozenRows(1); });
  // pilihan: salin murid dari kelas lain (contoh: naik tahun) — rekod TP tidak disalin
  if (d.SALIN_DARI) {
    const src = kelas_(d.SALIN_DARI);
    const rows = readTable_(src.SHEET_MURID).filter(r => String(r.AKTIF || 'YA').toUpperCase() !== 'TIDAK');
    if (rows.length) shM.getRange(2, 1, rows.length, mh.length).setValues(rows.map(r => mh.map(h => h === 'ULASAN_GURU_KELAS' ? '' : (r[h] === undefined ? '' : r[h]))));
  }
  kelasSheet_().appendRow([id, sesi, tahun, nama, d.NAMA_GURU_KELAS || '', shM.getName(), shK.getName(), 'AKTIF', new Date()]);
  return id;
}
function setStatusKelas_(id, status) {
  const sh = kelasSheet_(); const k = kelas_(id); const hd = headers_(sh);
  sh.getRange(k._row, hd.indexOf('STATUS') + 1).setValue(status);
  return id;
}
function updateKelas_(id, data) { // simpan tetapan khusus kelas
  const sh = kelasSheet_(); const k = kelas_(id); const hd = headers_(sh);
  Object.keys(KELAS_KEYS).forEach(key => { if (data[key] !== undefined) sh.getRange(k._row, hd.indexOf(KELAS_KEYS[key]) + 1).setValue(data[key]); });
}

function getAll_(kelasId) {
  const k = kelas_(kelasId);
  const set = settings_();
  delete set.ADMIN_PASSWORD; // kata laluan tidak dihantar ke pelayar
  set.SESI = k.SESI; set.TAHUN = k.TAHUN; set.KELAS = k.NAMA_KELAS; set.NAMA_GURU_KELAS = k.NAMA_GURU_KELAS;
  return {
    tetapan: set,
    kelasAktif: k,
    senaraiKelas: senaraiKelas_(),
    murid: readTable_(k.SHEET_MURID),
    mp: readTable_(SH.MP),
    markah: readTable_(k.SHEET_MARKAH),
    tafsiran: readTable_(SH.TAFSIRAN),
    kemahiran: readTable_(SH.KEMAHIRAN),
    bahasa: readTable_(SH.BAHASA),
    panduan: readTable_(SH.PANDUAN),
    serverTime: new Date().toISOString()
  };
}

/* ---------- POST (semua tindakan edit memerlukan kata laluan) ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const a = body.action;
    if (a === 'login') { checkPass_(body.password); log_('LOGIN', 'Admin log masuk'); return out_({ ok: true }); }
    checkPass_(body.password);
    let res; let kid = body.kelas;
    const K = () => kelas_(kid);
    switch (a) {
      case 'createKelas':   res = kid = createKelas_(body.data); break;
      case 'arkibKelas':    res = setStatusKelas_(body.id, body.status || 'ARKIB'); break;
      case 'saveTetapan':   updateKelas_(K().ID_KELAS, body.data); res = saveTetapan_(body.data); break;
      case 'saveMurid':     res = upsert_(K().SHEET_MURID, 'ID_MURID', body.data, 'M'); break;
      case 'deleteMurid':   res = deleteMurid_(K(), body.id); break;
      case 'saveMP':        res = upsert_(SH.MP, 'KOD_MP', body.data, null); break;
      case 'deleteMP':      res = deleteMP_(body.id); break;
      case 'saveMarkah':    res = saveMarkah_(K().SHEET_MARKAH, body.kod, body.rows); break;
      case 'saveUlasanKelas': res = saveUlasanKelas_(K().SHEET_MURID, body.rows); break;
      case 'saveTafsiran':  res = saveTafsiran_(body.data); break;
      default: throw new Error('Tindakan tidak dikenali: ' + a);
    }
    log_(a, JSON.stringify(body.data || body.id || body.kod || '').slice(0, 300));
    return out_({ ok: true, result: res, data: getAll_(kid) });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }

function saveTetapan_(data) {
  const sh = sheet_(SH.TETAPAN);
  const vals = sh.getDataRange().getValues();
  Object.keys(data || {}).forEach(k => {
    if (KELAS_KEYS[k]) return; // disimpan dalam 00_KELAS
    if (k === 'ADMIN_PASSWORD' && !String(data[k]).trim()) return;
    let found = false;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === k) { sh.getRange(i + 1, 2).setValue(data[k]); found = true; break; }
    }
    if (!found) sh.appendRow([k, data[k], '']);
  });
  return true;
}

function upsert_(name, key, data, prefix) {
  const sh = sheet_(name);
  const hd = headers_(sh);
  const vals = sh.getDataRange().getValues();
  if (!data[key]) {
    if (!prefix) throw new Error(key + ' wajib diisi.');
    let max = 0;
    for (let i = 1; i < vals.length; i++) {
      const n = parseInt(String(vals[i][0]).replace(/\D/g, ''), 10); if (n > max) max = n;
    }
    data[key] = prefix + String(max + 1).padStart(3, '0');
  }
  const keyCol = hd.indexOf(key);
  let rowIdx = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][keyCol]) === String(data[key])) { rowIdx = i; break; }
  const row = hd.map((h, j) => (data[h] !== undefined ? data[h] : (rowIdx > -1 ? vals[rowIdx][j] : '')));
  if (rowIdx > -1) sh.getRange(rowIdx + 1, 1, 1, hd.length).setValues([row]);
  else sh.appendRow(row);
  return data[key];
}

function deleteRowsWhere_(name, col, val) {
  const sh = sheet_(name);
  const hd = headers_(sh); const c = hd.indexOf(col);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][c]) === String(val)) sh.deleteRow(i + 1);
}
function deleteMurid_(k, id) { deleteRowsWhere_(k.SHEET_MURID, 'ID_MURID', id); deleteRowsWhere_(k.SHEET_MARKAH, 'ID_MURID', id); return id; }
function deleteMP_(kod) { // mata pelajaran dikongsi semua kelas; rekod TP lama dikekalkan sebagai sejarah
  deleteRowsWhere_(SH.MP, 'KOD_MP', kod); return kod; }

function saveMarkah_(sheetName, kod, rows) {
  if (!kod) throw new Error('Kod mata pelajaran tiada.');
  const sh = sheet_(sheetName);
  const hd = headers_(sh);
  const vals = sh.getDataRange().getValues();
  const idx = {};
  for (let i = 1; i < vals.length; i++) if (String(vals[i][1]) === String(kod)) idx[String(vals[i][0])] = i + 1;
  const now = new Date();
  const append = [];
  (rows || []).forEach(r => {
    const line = hd.map(h => {
      if (h === 'ID_MURID') return r.ID_MURID;
      if (h === 'KOD_MP') return kod;
      if (h === 'DIKEMASKINI') return now;
      return r[h] === undefined || r[h] === null ? '' : r[h];
    });
    if (idx[r.ID_MURID]) sh.getRange(idx[r.ID_MURID], 1, 1, hd.length).setValues([line]);
    else append.push(line);
  });
  if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, hd.length).setValues(append);
  return (rows || []).length;
}

function saveUlasanKelas_(sheetName, rows) {
  const sh = sheet_(sheetName); const hd = headers_(sh);
  const cId = hd.indexOf('ID_MURID'), cU = hd.indexOf('ULASAN_GURU_KELAS');
  const vals = sh.getDataRange().getValues();
  (rows || []).forEach(r => {
    for (let i = 1; i < vals.length; i++) if (String(vals[i][cId]) === String(r.ID_MURID)) { sh.getRange(i + 1, cU + 1).setValue(r.ULASAN_GURU_KELAS || ''); break; }
  });
  return true;
}

function saveTafsiran_(d) {
  const sh = sheet_(SH.TAFSIRAN); const hd = headers_(sh);
  const vals = sh.getDataRange().getValues();
  const cS = hd.indexOf('SET'), cK = hd.indexOf('KOD_MP');
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][cS]) === String(d.SET) && String(vals[i][cK]) === String(d.KOD_MP)) {
      ['TP1','TP2','TP3','TP4','TP5','TP6'].forEach(t => sh.getRange(i + 1, hd.indexOf(t) + 1).setValue(d[t] || ''));
      return true;
    }
  }
  sh.appendRow(hd.map(h => d[h] !== undefined ? d[h] : ''));
  return true;
}

/* Jalankan sekali dari editor untuk menguji sambungan */
function testSambungan() { Logger.log(JSON.stringify(getAll_()).slice(0, 1500)); }
