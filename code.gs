/**
 * =====================================================================
 *  e-PBD PPKI : SISTEM PELAPORAN PENTAKSIRAN BILIK DARJAH (Google Apps Script)
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
 *  6. BERBILANG SEKOLAH (v5): setiap guru WAJIB daftar (nama, emel, kata laluan,
 *     no telefon, kod sekolah). Data diasingkan mengikut KOD_SEKOLAH. Semua data
 *     lama dipindah automatik ke sekolah SK BENUT (JBA6041) pada panggilan pertama.
 *     (Boleh juga jalankan migrasiSKBenut() secara manual dari editor.)
 *  7. MASTER ADMIN (v6): di skrin log masuk tekan "Log masuk Master Admin" dan masukkan
 *     kata laluan MASTER_PASSWORD (di bawah). Master Admin boleh:
 *       - melihat SEMUA akaun guru berserta kata laluan, menukar kata laluan,
 *         menyahaktif atau memadam akaun;
 *       - membuka data mana-mana sekolah (pilih kod sekolah di menu sisi);
 *       - HANYA Master Admin boleh mengubah Tafsiran TP.
 *     Guru biasa log masuk dengan EMEL + KATA LALUAN yang didaftarkan dan hanya
 *     melihat data sekolah (kod sekolah) masing-masing.
 *     Lupa kata laluan / pertanyaan: WhatsApp Master Admin 01110661077.
 *     NOTA: kata laluan guru disimpan dalam lajur KATA_LALUAN sheet 00_GURU supaya
 *     boleh dilihat oleh Master Admin. Hadkan akses Google Sheet kepada Master Admin sahaja.
 *     Akaun lama: kata laluan direkod automatik selepas guru log masuk sekali.
 *  8. MATA PELAJARAN & TAFSIRAN TP SERAGAM (v7): senarai mata pelajaran (kod, susunan,
 *     nama) dan Tafsiran TP adalah SAMA untuk semua sekolah dan disimpan sebagai baris
 *     KOD_SEKOLAH = "SEMUA" dalam 03_MATA_PELAJARAN & 05_TAFSIRAN_TP. Hanya Master Admin
 *     boleh menambah/memadam mata pelajaran, menukar kod/susunan/nama, dan mengubah
 *     Tafsiran TP — perubahan terus dilihat oleh semua sekolah.
 *     Guru setiap sekolah hanya boleh kemaskini: nama guru MP, bidang/komponen,
 *     status aktif dan catatan (disimpan pada baris kod sekolah masing-masing).
 *     Data asal SK BENUT disalin automatik sebagai set "SEMUA" pada panggilan pertama.
 *  9. TAHAP 1 / TAHAP 2 (v8): setiap kelas mempunyai lajur TAHAP dalam 00_KELAS (1 = Tahun 1-3,
 *     2 = Tahun 4-6). Mata pelajaran, Tafsiran TP & Penyataan Kemahiran ditapis ikut tahap kelas
 *     yang sedang dibuka supaya tidak tertukar. Tahap boleh ditukar di Tetapan atau semasa cipta kelas.
 *  PENTING: selepas tampal kod baharu, buat Deploy > Manage deployments > Edit > New version.
 *  (Jika skrip TIDAK dibuka dari dalam Google Sheet, isi SPREADSHEET_ID di bawah.)
 * =====================================================================
 */
const SPREADSHEET_ID = ''; // kosongkan jika skrip dibuka dari Extensions > Apps Script

/* ---------- MASTER ADMIN ---------- */
const MASTER_PASSWORD = '202020';      // kata laluan Master Admin
const MASTER_EMEL = 'MASTER_ADMIN';    // pengenal sesi Master Admin (bukan emel sebenar)
const WASAP_ADMIN = '01110661077';     // WhatsApp Master Admin (lupa kata laluan / pertanyaan)
const KOD_GLOBAL = 'SEMUA';            // baris mata pelajaran & tafsiran TP yang dikongsi semua sekolah
const KOMP_KEY = 'KOMPONEN (dipisah | )';
// lajur mata pelajaran yang boleh diubah oleh guru (berbeza bagi setiap sekolah)
const MP_SEKOLAH = ['NAMA_GURU_MP', KOMP_KEY, 'AKTIF', 'CATATAN'];

const SH = {
  KELAS: '00_KELAS',
  TETAPAN: '01_TETAPAN', MURID: '02_MURID', MP: '03_MATA_PELAJARAN', MARKAH: '04_MARKAH',
  TAFSIRAN: '05_TAFSIRAN_TP', KEMAHIRAN: '06_PENYATAAN_KEMAHIRAN', BAHASA: '07_BAHASA_LABEL',
  RUJUKAN: '08_RUJUKAN', PANDUAN: '09_PANDUAN', LOG: '10_LOG', GURU: '00_GURU'
};
const KOD_ASAL = 'JBA6041';          // sekolah asal (SK BENUT) — semua data lama dipindah ke kod ini
const GURU_HEADERS = ['KOD_SEKOLAH','NAMA_GURU','EMEL','NO_TELEFON','KATA_LALUAN_HASH','GARAM','TOKEN','STATUS','DIDAFTAR','LOG_MASUK_AKHIR','KATA_LALUAN'];
// sheet yang dikongsi tetapi diasingkan baris demi baris mengikut KOD_SEKOLAH
const SHEET_BERSEKOLAH = ['00_KELAS', '01_TETAPAN', '03_MATA_PELAJARAN', '05_TAFSIRAN_TP'];
const MARKAH_HEADERS = ['ID_MURID','KOD_MP','K1','K2','K3','K4','K5','K6','K7','K8','K9','K10','K11','K12','TP_KESELURUHAN','ULASAN_GURU_MP','DIKEMASKINI'];

const KELAS_HEADERS = ['ID_KELAS','SESI','TAHUN','NAMA_KELAS','NAMA_GURU_KELAS','SHEET_MURID','SHEET_MARKAH','STATUS','DICIPTA','KOD_SEKOLAH','TAHAP'];
// Tetapan yang berbeza bagi setiap kelas (disimpan dalam 00_KELAS, bukan 01_TETAPAN)
const KELAS_KEYS = { SESI: 'SESI', TAHUN: 'TAHUN', KELAS: 'NAMA_KELAS', NAMA_GURU_KELAS: 'NAMA_GURU_KELAS', TAHAP: 'TAHAP' };

/* ====================== TAHAP 1 / TAHAP 2 (v8) ======================
 * Setiap kelas ditetapkan sebagai TAHAP 1 (Tahun 1-3) atau TAHAP 2 (Tahun 4-6) dalam lajur
 * TAHAP sheet 00_KELAS. Mata pelajaran (03), Tafsiran TP (05) & Penyataan Kemahiran (06)
 * juga mempunyai lajur TAHAP, jadi setiap kelas hanya melihat mata pelajaran & pentaksiran
 * tahapnya sendiri. Baris dengan TAHAP kosong dianggap terpakai untuk kedua-dua tahap. */
function normTahap_(v) { const s = String(v == null ? '' : v).toUpperCase().replace(/TAHAP|\s/g, '').replace(/\.0+$/, ''); return (s === '1' || s === '2') ? s : ''; }
function tahapDariTahun_(t) { const n = parseInt(String(t == null ? '' : t).replace(/[^0-9]/g, ''), 10); return (n >= 4 && n <= 6) ? '2' : '1'; }
function tahapKelas_(k) { return k ? (normTahap_(k.TAHAP) || tahapDariTahun_(k.TAHUN)) : '1'; }
function ikutTahap_(r, tahap) { const t = normTahap_(r.TAHAP); return !t || t === String(tahap); }
function pastiLajur_(name, col) {
  const sh = ss_().getSheetByName(name); if (!sh || sh.getLastColumn() < 1) return;
  if (headers_(sh).indexOf(col) > -1) return;
  const c = sh.getLastColumn() + 1;
  sh.getRange(1, c).setValue(col).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff');
  if (name === SH.KELAS && sh.getLastRow() > 1) { // isi tahap kelas sedia ada berdasarkan tahun
    const hd = headers_(sh), cT = hd.indexOf('TAHUN'); const vals = sh.getRange(2, 1, sh.getLastRow() - 1, hd.length).getValues();
    sh.getRange(2, c, vals.length, 1).setValues(vals.map(v => [tahapDariTahun_(cT > -1 ? v[cT] : '')]));
  }
}
function migrasiTahap() { [SH.KELAS, SH.MP, SH.TAFSIRAN, SH.KEMAHIRAN].forEach(n => pastiLajur_(n, 'TAHAP')); }

function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Ralat data sistem. Sila WhatsApp Master Admin ' + WASAP_ADMIN + '.');
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
/* Data rujukan (jarang berubah) disimpan dalam cache pelayan 6 jam -> muatan lebih pantas */
function readCached_(name) {
  const c = CacheService.getScriptCache(), k = 'TBL_' + name;
  try { const v = c.get(k); if (v) return JSON.parse(v); } catch (e) {}
  const rows = readTable_(name);
  try { const s = JSON.stringify(rows); if (s.length < 95000) c.put(k, s, 21600); } catch (e) {}
  return rows;
}
const sama_ = (a, b) => String(a || '').trim().toUpperCase() === String(b || '').trim().toUpperCase();
function readSekolah_(name, kod) { return readTable_(name).filter(r => sama_(r.KOD_SEKOLAH, kod)); }
function appendObj_(sh, obj) { const hd = headers_(sh); sh.appendRow(hd.map(h => obj[h] !== undefined ? obj[h] : '')); }
function settings_(kod) {
  const o = {};
  readSekolah_(SH.TETAPAN, kod).forEach(r => { o[String(r.KUNCI).trim()] = r.NILAI; });
  return o;
}
function log_(action, detail, guru) {
  try { sheet_(SH.LOG).appendRow([new Date(), action, (guru ? '[' + guru.KOD_SEKOLAH + ' ' + guru.EMEL + '] ' : '') + String(detail || '').slice(0, 500)]); } catch (e) {}
}

/* ====================== MIGRASI KE BERBILANG SEKOLAH ======================
 * Tambah lajur KOD_SEKOLAH pada sheet dikongsi dan isi baris lama dengan JBA6041.
 * Selamat dijalankan berulang kali (baris yang sudah ada kod tidak diubah). */
function pastiLajurKod_(name) {
  const sh = ss_().getSheetByName(name); if (!sh) return;
  if (sh.getLastColumn() < 1) return;
  let hd = headers_(sh); let c = hd.indexOf('KOD_SEKOLAH');
  if (c < 0) {
    c = hd.length;
    sh.getRange(1, c + 1).setValue('KOD_SEKOLAH').setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff');
  }
  const n = sh.getLastRow() - 1; if (n < 1) return;
  const rng = sh.getRange(2, c + 1, n, 1); const vals = rng.getValues();
  const data = sh.getRange(2, 1, n, Math.max(1, c)).getValues();
  let ubah = false;
  for (let i = 0; i < n; i++) {
    const kosong = data[i].every(v => v === '' || v === null);
    if (!kosong && !String(vals[i][0]).trim()) { vals[i][0] = KOD_ASAL; ubah = true; }
  }
  if (ubah) rng.setValues(vals);
}
function guruSheet_() {
  const ss = ss_(); let sh = ss.getSheetByName(SH.GURU);
  if (!sh) {
    sh = ss.insertSheet(SH.GURU, 0);
    sh.getRange(1, 1, 1, GURU_HEADERS.length).setValues([GURU_HEADERS]).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < 1) {
    sh.getRange(1, 1, 1, GURU_HEADERS.length).setValues([GURU_HEADERS]).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff');
  } else { // tambah lajur baharu (cth: KATA_LALUAN) pada sheet sedia ada
    const hd = headers_(sh);
    GURU_HEADERS.forEach(h => { if (hd.indexOf(h) < 0) { hd.push(h); sh.getRange(1, hd.length).setValue(h).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff'); } });
  }
  return sh;
}
function migrasi_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('MIGRASI_SEKOLAH_V5') !== 'SIAP') { migrasiSKBenut(); p.setProperty('MIGRASI_SEKOLAH_V5', 'SIAP'); }
  if (p.getProperty('MIGRASI_GLOBAL_V7') !== 'SIAP') { migrasiGlobal(); p.setProperty('MIGRASI_GLOBAL_V7', 'SIAP'); }
  if (p.getProperty('MIGRASI_TAHAP_V8') !== 'SIAP') { migrasiTahap(); p.setProperty('MIGRASI_TAHAP_V8', 'SIAP'); }
}
/* v7: cipta set mata pelajaran & tafsiran TP seragam ("SEMUA") dari data SK BENUT.
 * Boleh dijalankan secara manual; tidak berbuat apa-apa jika set SEMUA sudah wujud. */
function migrasiGlobal() {
  [SH.MP, SH.TAFSIRAN].forEach(name => {
    const sh = ss_().getSheetByName(name); if (!sh) return;
    pastiLajurKod_(name);
    if (readSekolah_(name, KOD_GLOBAL).length) return;
    const hd = headers_(sh);
    const rows = readSekolah_(name, KOD_ASAL).map(r => hd.map(h => h === 'KOD_SEKOLAH' ? KOD_GLOBAL
      : (name === SH.MP && (h === 'NAMA_GURU_MP' || h === 'CATATAN')) ? '' : (r[h] === undefined ? '' : r[h])));
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, hd.length).setValues(rows);
  });
  log_('MIGRASI', 'Set mata pelajaran & tafsiran TP seragam (SEMUA) dicipta');
}
/* Boleh dijalankan secara manual dari editor Apps Script */
function migrasiSKBenut() {
  guruSheet_();
  kelasSheet_(); // pastikan 00_KELAS wujud (kelas asal didaftar dahulu)
  SHEET_BERSEKOLAH.forEach(pastiLajurKod_);
  // pastikan NAMA_SEKOLAH ada untuk SK BENUT
  const t = settings_(KOD_ASAL);
  if (!String(t.NAMA_SEKOLAH || '').trim()) setTetapan_(KOD_ASAL, 'NAMA_SEKOLAH', 'SK BENUT');
  log_('MIGRASI', 'Semua data lama ditetapkan kepada ' + KOD_ASAL);
}

/* Sekolah baharu: salin tetapan umum dari sekolah asal.
 * (Mata pelajaran & tafsiran TP tidak perlu disalin — semua sekolah guna set SEMUA.) */
function sediaSekolah_(kod) {
  if (readSekolah_(SH.TETAPAN, kod).length) return;
  const KEKAL = ['PELAPORAN', 'JENIS_SEKOLAH', 'BAHASA_LAPORAN', 'TAJUK_SIJIL', 'JAWATAN_PENTADBIR'];
  const shT = sheet_(SH.TETAPAN);
  const rowsT = readSekolah_(SH.TETAPAN, KOD_ASAL).filter(r => String(r.KUNCI).trim() && String(r.KUNCI).trim() !== 'ADMIN_PASSWORD');
  rowsT.forEach(r => { const o = Object.assign({}, r); delete o._row; o.KOD_SEKOLAH = kod; if (KEKAL.indexOf(String(r.KUNCI).trim()) < 0) o.NILAI = ''; appendObj_(shT, o); });
  if (!rowsT.length) appendObj_(shT, { KUNCI: 'NAMA_SEKOLAH', NILAI: '', KOD_SEKOLAH: kod });
}

/* ====================== MATA PELAJARAN SERAGAM ======================
 * Kod, susunan & nama dari set SEMUA (Master Admin). Nama guru MP, komponen,
 * aktif & catatan dari baris sekolah itu sendiri (jika ada). */
function mpSekolah_(kod, tahap) {
  tahap = normTahap_(tahap) || '1';
  const sek = {}; readSekolah_(SH.MP, kod).filter(r => ikutTahap_(r, tahap)).forEach(r => { const key = String(r.KOD_MP).trim().toUpperCase(); if (!sek[key] || normTahap_(r.TAHAP)) sek[key] = r; });
  return readSekolah_(SH.MP, KOD_GLOBAL).filter(g => ikutTahap_(g, tahap)).map(g => {
    const o = Object.assign({}, g); delete o._row;
    const s = sek[String(g.KOD_MP).trim().toUpperCase()];
    o.NAMA_GURU_MP = s ? s.NAMA_GURU_MP : ''; o.CATATAN = s ? s.CATATAN : '';
    if (s) { o[KOMP_KEY] = s[KOMP_KEY]; o.AKTIF = s.AKTIF; }
    o.KOD_SEKOLAH = kod; o.TAHAP = normTahap_(g.TAHAP) || tahap;
    return o;
  });
}
function saveMP_(d, kod, master, tahap) {
  d = d || {}; tahap = normTahap_(tahap) || '1';
  const kmp = String(d.KOD_MP || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!kmp) throw new Error('Kod mata pelajaran wajib diisi.');
  const cari = () => { const c = readSekolah_(SH.MP, KOD_GLOBAL).filter(r => String(r.KOD_MP).trim().toUpperCase() === kmp && ikutTahap_(r, tahap)); return c.find(r => normTahap_(r.TAHAP)) || c[0]; };
  let glob = cari();
  if (master) { // bahagian seragam: kod, susunan, nama (dilihat semua sekolah)
    const g = { KOD_MP: kmp };
    if (d.SUSUNAN !== undefined) g.SUSUNAN = d.SUSUNAN;
    if (d.NAMA_MATA_PELAJARAN !== undefined) g.NAMA_MATA_PELAJARAN = String(d.NAMA_MATA_PELAJARAN).trim();
    if (!glob) { g[KOMP_KEY] = d[KOMP_KEY] || ''; g.AKTIF = d.AKTIF || 'Y'; } // lalai untuk sekolah lain
    if (!String(g.NAMA_MATA_PELAJARAN || (glob && glob.NAMA_MATA_PELAJARAN) || '').trim()) throw new Error('Nama mata pelajaran wajib diisi.');
    upsertSekolah_(SH.MP, 'KOD_MP', g, KOD_GLOBAL, tahap);
    glob = cari();
  } else if (!glob) throw new Error('Hanya Master Admin boleh menambah mata pelajaran baharu.');
  // bahagian sekolah: nama guru, komponen, aktif, catatan
  const s = { KOD_MP: kmp, NAMA_MATA_PELAJARAN: glob.NAMA_MATA_PELAJARAN, SUSUNAN: glob.SUSUNAN };
  MP_SEKOLAH.forEach(k => { if (d[k] !== undefined) s[k] = d[k]; });
  upsertSekolah_(SH.MP, 'KOD_MP', s, kod, tahap);
  return kmp;
}

/* ====================== GURU: DAFTAR & LOG MASUK ====================== */
function hash_(garam, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, garam + '|' + pw, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
const normEmel_ = e => String(e || '').trim().toLowerCase();
function cariGuru_(emel) {
  const em = normEmel_(emel); if (!em) return null;
  return readTable_(SH.GURU).find(g => normEmel_(g.EMEL) === em) || null;
}
function setGuru_(g, data) {
  const sh = guruSheet_(); const hd = headers_(sh);
  Object.keys(data).forEach(k => { const c = hd.indexOf(k); if (c > -1) sh.getRange(g._row, c + 1).setValue(data[k]); g[k] = data[k]; });
}
function awamGuru_(g) { return { NAMA_GURU: g.NAMA_GURU, EMEL: g.EMEL, NO_TELEFON: String(g.NO_TELEFON || ''), KOD_SEKOLAH: String(g.KOD_SEKOLAH).toUpperCase(), MASTER: !!g._master }; }
function daftar_(d) {
  d = d || {};
  const nama = String(d.NAMA_GURU || '').trim().toUpperCase().replace(/\s+/g, ' ');
  const emel = normEmel_(d.EMEL);
  const pw = String(d.PASSWORD || '');
  const tel = String(d.NO_TELEFON || '').trim().toUpperCase().replace(/[^0-9+\- ]/g, '');
  const kod = String(d.KOD_SEKOLAH || '').trim().toUpperCase().replace(/\s+/g, '');
  if (nama.length < 3) throw new Error('Nama guru wajib diisi.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emel)) throw new Error('Emel tidak sah.');
  if (pw.length < 6) throw new Error('Kata laluan sekurang-kurangnya 6 aksara.');
  if (tel.replace(/\D/g, '').length < 9) throw new Error('No telefon tidak sah.');
  if (!/^[A-Z]{3}[0-9]{4}$/.test(kod)) throw new Error('Kod sekolah tidak sah. Contoh: JBA6041 (3 huruf + 4 nombor).');
  if (cariGuru_(emel)) throw new Error('Emel ini sudah didaftarkan. Sila log masuk.');
  sediaSekolah_(kod);
  const garam = Utilities.getUuid();
  appendObj_(guruSheet_(), { KOD_SEKOLAH: kod, NAMA_GURU: nama, EMEL: emel, NO_TELEFON: "'" + tel, KATA_LALUAN_HASH: hash_(garam, pw), GARAM: garam, TOKEN: Utilities.getUuid(), STATUS: 'AKTIF', DIDAFTAR: new Date(), KATA_LALUAN: "'" + pw });
  const g = cariGuru_(emel);
  log_('DAFTAR', nama, g);
  return g;
}
function login_(emel, pw) {
  const g = cariGuru_(emel);
  if (!g || hash_(String(g.GARAM), String(pw || '')) !== String(g.KATA_LALUAN_HASH)) throw new Error('Emel atau kata laluan salah. Lupa kata laluan? WhatsApp Master Admin ' + WASAP_ADMIN + '.');
  if (String(g.STATUS || 'AKTIF').toUpperCase() !== 'AKTIF') throw new Error('Akaun anda tidak aktif. Sila WhatsApp Master Admin ' + WASAP_ADMIN + '.');
  if (!String(g.TOKEN || '').trim()) setGuru_(g, { TOKEN: Utilities.getUuid() });
  const ubah = { LOG_MASUK_AKHIR: new Date() };
  if (String(g.KATA_LALUAN || '') !== String(pw)) ubah.KATA_LALUAN = "'" + pw; // rekod untuk paparan Master Admin
  setGuru_(g, ubah);
  return g;
}
function sahGuru_(emel, token, sekolah) {
  if (isMasterEmel_(emel)) return sahMaster_(token, sekolah);
  const g = cariGuru_(emel);
  if (!g || !String(token || '').trim() || String(g.TOKEN) !== String(token)) throw new Error('Sesi tamat. Sila log masuk semula (kata laluan).');
  if (String(g.STATUS || 'AKTIF').toUpperCase() !== 'AKTIF') throw new Error('Akaun anda tidak aktif. Sila WhatsApp Master Admin ' + WASAP_ADMIN + '.');
  return g;
}
function tukarKataLaluan_(g, pw) {
  if (String(pw).length < 6) throw new Error('Kata laluan baharu sekurang-kurangnya 6 aksara.');
  const garam = Utilities.getUuid(), token = Utilities.getUuid();
  setGuru_(g, { GARAM: garam, KATA_LALUAN_HASH: hash_(garam, pw), TOKEN: token, KATA_LALUAN: "'" + pw });
  return token;
}

/* ====================== MASTER ADMIN ====================== */
function isMasterEmel_(e) { return String(e || '').trim().toUpperCase() === MASTER_EMEL; }
function sigMaster_() { return hash_('MASTER', MASTER_PASSWORD); }
function loginMaster_(pw) {
  const c = CacheService.getScriptCache(); const n = +(c.get('MASTER_GAGAL') || 0);
  if (n >= 5) throw new Error('Terlalu banyak cubaan salah. Cuba lagi selepas 10 minit.');
  if (String(pw || '') !== MASTER_PASSWORD) { c.put('MASTER_GAGAL', String(n + 1), 600); throw new Error('Kata laluan Master Admin salah.'); }
  c.remove('MASTER_GAGAL');
  const p = PropertiesService.getScriptProperties();
  let tok = p.getProperty('MASTER_TOKEN');
  // token baharu jika kata laluan Master Admin ditukar (sesi lama tamat)
  if (!tok || p.getProperty('MASTER_SIG') !== sigMaster_()) { tok = Utilities.getUuid(); p.setProperties({ MASTER_TOKEN: tok, MASTER_SIG: sigMaster_() }); }
  return tok;
}
function masterGuru_(sekolah) {
  let kod = String(sekolah || '').trim().toUpperCase();
  if (!/^[A-Z]{3}[0-9]{4}$/.test(kod)) kod = KOD_ASAL;
  return { _master: true, KOD_SEKOLAH: kod, NAMA_GURU: 'MASTER ADMIN', EMEL: MASTER_EMEL, NO_TELEFON: WASAP_ADMIN };
}
function sahMaster_(token, sekolah) {
  const p = PropertiesService.getScriptProperties();
  if (!String(token || '').trim() || p.getProperty('MASTER_TOKEN') !== String(token) || p.getProperty('MASTER_SIG') !== sigMaster_())
    throw new Error('Sesi tamat. Sila log masuk semula (kata laluan).');
  return masterGuru_(sekolah);
}
/* Senarai semua sekolah (kod, nama, bilangan guru) */
function senaraiSekolah_() {
  const nama = {}, guru = {}; nama[KOD_ASAL] = nama[KOD_ASAL] || '';
  readTable_(SH.TETAPAN).forEach(r => {
    const k = String(r.KOD_SEKOLAH || '').trim().toUpperCase(); if (!k) return;
    if (nama[k] === undefined) nama[k] = '';
    if (String(r.KUNCI).trim() === 'NAMA_SEKOLAH') nama[k] = String(r.NILAI || '');
  });
  readTable_(SH.GURU).forEach(g => {
    const k = String(g.KOD_SEKOLAH || '').trim().toUpperCase(); if (!k) return;
    if (nama[k] === undefined) nama[k] = '';
    guru[k] = (guru[k] || 0) + 1;
  });
  return Object.keys(nama).sort().map(k => ({ KOD: k, NAMA: nama[k], GURU: guru[k] || 0 }));
}
function senaraiKodSekolah_() { return senaraiSekolah_().map(x => x.KOD); }
/* Semua akaun guru berserta kata laluan — untuk Master Admin sahaja */
function senaraiAkaun_() {
  return readTable_(SH.GURU).map(g => ({
    KOD_SEKOLAH: String(g.KOD_SEKOLAH || '').toUpperCase(), NAMA_GURU: g.NAMA_GURU, EMEL: g.EMEL,
    NO_TELEFON: String(g.NO_TELEFON || ''), KATA_LALUAN: String(g.KATA_LALUAN === undefined ? '' : g.KATA_LALUAN),
    STATUS: String(g.STATUS || 'AKTIF').toUpperCase(), DIDAFTAR: g.DIDAFTAR, LOG_MASUK_AKHIR: g.LOG_MASUK_AKHIR
  }));
}
function kemaskiniAkaun_(emel, d) {
  const g = cariGuru_(emel); if (!g) throw new Error('Akaun ' + emel + ' tidak dijumpai.');
  d = d || {}; const ubah = {};
  const pw = String(d.KATA_LALUAN || '').trim();
  if (pw) {
    if (pw.length < 6) throw new Error('Kata laluan sekurang-kurangnya 6 aksara.');
    const garam = Utilities.getUuid();
    Object.assign(ubah, { GARAM: garam, KATA_LALUAN_HASH: hash_(garam, pw), KATA_LALUAN: "'" + pw, TOKEN: Utilities.getUuid() }); // sesi lama guru tamat
  }
  if (d.NAMA_GURU !== undefined && String(d.NAMA_GURU).trim()) ubah.NAMA_GURU = String(d.NAMA_GURU).trim().toUpperCase().replace(/\s+/g, ' ');
  if (d.NO_TELEFON !== undefined && String(d.NO_TELEFON).trim()) ubah.NO_TELEFON = "'" + String(d.NO_TELEFON).trim().replace(/[^0-9+\- ]/g, '');
  if (d.KOD_SEKOLAH !== undefined && String(d.KOD_SEKOLAH).trim()) {
    const kod = String(d.KOD_SEKOLAH).trim().toUpperCase().replace(/\s+/g, '');
    if (!/^[A-Z]{3}[0-9]{4}$/.test(kod)) throw new Error('Kod sekolah tidak sah. Contoh: JBA6041.');
    if (!sama_(kod, g.KOD_SEKOLAH)) { sediaSekolah_(kod); ubah.KOD_SEKOLAH = kod; }
  }
  if (d.STATUS) ubah.STATUS = String(d.STATUS).toUpperCase() === 'AKTIF' ? 'AKTIF' : 'TIDAK AKTIF';
  setGuru_(g, ubah);
  return g.EMEL;
}
function padamAkaun_(emel) {
  const g = cariGuru_(emel); if (!g) throw new Error('Akaun ' + emel + ' tidak dijumpai.');
  guruSheet_().deleteRow(g._row);
  return emel;
}

/* ---------- GET ---------- */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getAll';
    if (action === 'ping') return out_({ ok: true, msg: 'e-PBD PPKI API aktif' });
    if (action === 'getAll') return out_({ ok: false, error: 'Sila log masuk melalui aplikasi e-PBD PPKI.' });
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
    if (ss.getSheetByName(SH.MURID) && !PropertiesService.getScriptProperties().getProperty('KELAS_ASAL_SIAP')) {
      const t = {}; readTable_(SH.TETAPAN).forEach(r => { if (!String(r.KOD_SEKOLAH || '').trim() || sama_(r.KOD_SEKOLAH, KOD_ASAL)) t[String(r.KUNCI).trim()] = r.NILAI; });
      const sesi = String(t.SESI || new Date().getFullYear()), tahun = String(t.TAHUN || ''), nama = String(t.KELAS || 'KELAS 1');
      appendObj_(sh, { ID_KELAS: kelasId_(sesi, tahun, nama), SESI: sesi, TAHUN: tahun, NAMA_KELAS: nama, NAMA_GURU_KELAS: t.NAMA_GURU_KELAS || '', SHEET_MURID: SH.MURID, SHEET_MARKAH: SH.MARKAH, STATUS: 'AKTIF', DICIPTA: new Date(), KOD_SEKOLAH: KOD_ASAL });
    }
    PropertiesService.getScriptProperties().setProperty('KELAS_ASAL_SIAP', '1');
  }
  return sh;
}
function kelasId_(sesi, tahun, nama, kod) { return ((kod ? String(kod) + '-' : '') + String(sesi) + '-' + String(tahun) + '-' + String(nama)).toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_|_$/g, ''); }
function senaraiKelas_(kod) { kelasSheet_(); return readSekolah_(SH.KELAS, kod); }
function kelas_(id, kod, bolehKosong) {
  const all = senaraiKelas_(kod);
  if (!all.length) { if (bolehKosong) return null; throw new Error('Tiada kelas berdaftar. Sila tambah kelas dahulu (menu Kelas & Tahun).'); }
  if (!id || (bolehKosong && !all.some(x => String(x.ID_KELAS) === String(id)))) { // lalai: kelas aktif terbaharu
    const sorted = all.slice().sort((a, b) => (+b.SESI || 0) - (+a.SESI || 0));
    return sorted.find(k => String(k.STATUS).toUpperCase() !== 'ARKIB') || sorted[0];
  }
  const k = all.find(x => String(x.ID_KELAS) === String(id));
  if (!k) throw new Error('Kelas "' + id + '" tidak dijumpai.');
  return k;
}
function safeName_(s) { return String(s).replace(/[\[\]\*\?\/\\:']/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90); }

function createKelas_(d, kod) {
  const sesi = String(d.SESI || '').trim(), tahun = String(d.TAHUN || '').trim(), nama = String(d.NAMA_KELAS || '').trim().toUpperCase();
  if (!sesi || !nama) throw new Error('Sesi (tahun kalendar) dan nama kelas wajib diisi.');
  const id = sama_(kod, KOD_ASAL) ? kelasId_(sesi, tahun, nama) : kelasId_(sesi, tahun, nama, kod);
  if (readTable_(SH.KELAS).some(k => k.ID_KELAS === id)) throw new Error('Kelas ' + tahun + ' ' + nama + ' bagi sesi ' + sesi + ' sudah wujud.');
  const ss = ss_(); const label = safeName_((sama_(kod, KOD_ASAL) ? '' : kod + ' ') + sesi + ' ' + tahun + ' ' + nama);
  if (ss.getSheetByName('MURID ' + label) || ss.getSheetByName('MARKAH ' + label)) throw new Error('Kelas ini sudah wujud.');
  const shM = ss.insertSheet('MURID ' + label), shK = ss.insertSheet('MARKAH ' + label);
  // tajuk lajur murid disalin dari 02_MURID supaya struktur sama
  const tmpl = ss.getSheetByName(SH.MURID);
  const mh = tmpl ? headers_(tmpl) : ['ID_MURID','BIL','NAMA_MURID','NAMA_RINGKAS','NO_MYKID','JANTINA','ULASAN_GURU_KELAS','AKTIF','CATATAN'];
  [[shM, mh], [shK, MARKAH_HEADERS]].forEach(([sh, h]) => { sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0f5257').setFontColor('#ffffff'); sh.setFrozenRows(1); });
  // pilihan: salin murid dari kelas lain (contoh: naik tahun) — rekod TP tidak disalin
  if (d.SALIN_DARI) {
    const src = kelas_(d.SALIN_DARI, kod);
    const rows = readTable_(src.SHEET_MURID).filter(r => String(r.AKTIF || 'YA').toUpperCase() !== 'TIDAK');
    if (rows.length) shM.getRange(2, 1, rows.length, mh.length).setValues(rows.map(r => mh.map(h => h === 'ULASAN_GURU_KELAS' ? '' : (r[h] === undefined ? '' : r[h]))));
  }
  appendObj_(kelasSheet_(), { ID_KELAS: id, SESI: sesi, TAHUN: tahun, NAMA_KELAS: nama, NAMA_GURU_KELAS: String(d.NAMA_GURU_KELAS || '').toUpperCase(), SHEET_MURID: shM.getName(), SHEET_MARKAH: shK.getName(), STATUS: 'AKTIF', DICIPTA: new Date(), KOD_SEKOLAH: kod, TAHAP: normTahap_(d.TAHAP) || tahapDariTahun_(tahun) });
  return id;
}
function setStatusKelas_(id, status, kod) {
  const sh = kelasSheet_(); const k = kelas_(id, kod); const hd = headers_(sh);
  sh.getRange(k._row, hd.indexOf('STATUS') + 1).setValue(status);
  return id;
}
function updateKelas_(id, data, kod) { // simpan tetapan khusus kelas
  const sh = kelasSheet_(); const k = kelas_(id, kod); const hd = headers_(sh);
  if (data.TAHAP !== undefined) data.TAHAP = normTahap_(data.TAHAP) || tahapDariTahun_(data.TAHUN !== undefined ? data.TAHUN : k.TAHUN);
  Object.keys(KELAS_KEYS).forEach(key => { const c = hd.indexOf(KELAS_KEYS[key]); if (data[key] !== undefined && c > -1) sh.getRange(k._row, c + 1).setValue(data[key]); });
}

/* Kemaskini maklumat kelas (sesi, tahun, tahap, nama kelas, guru kelas) — mana-mana guru berdaftar sekolah ini */
function editKelas_(id, d, kod) {
  const k = kelas_(id, kod);
  const nama = String(d.NAMA_KELAS != null ? d.NAMA_KELAS : k.NAMA_KELAS).trim().toUpperCase();
  const sesi = String(d.SESI != null ? d.SESI : k.SESI).trim();
  if (!sesi || !nama) throw new Error('Sesi dan nama kelas wajib diisi.');
  updateKelas_(id, { SESI: sesi, TAHUN: String(d.TAHUN != null ? d.TAHUN : k.TAHUN).trim(), KELAS: nama,
    NAMA_GURU_KELAS: String(d.NAMA_GURU_KELAS != null ? d.NAMA_GURU_KELAS : k.NAMA_GURU_KELAS).trim().toUpperCase(), TAHAP: d.TAHAP }, kod);
  return id;
}
/* Padam kelas: buang baris dalam 00_KELAS serta sheet MURID & MARKAH kelas itu (tidak boleh dibatalkan) */
function padamKelas_(id, kod) {
  const sh = kelasSheet_(); const k = kelas_(id, kod); const ss = ss_();
  [k.SHEET_MURID, k.SHEET_MARKAH].forEach(n => {
    if (!n || n === SH.MURID || n === SH.MARKAH) return; // sheet templat asal tidak dipadam
    const s = ss.getSheetByName(n); if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  sh.deleteRow(k._row);
  return id;
}

function getAll_(kelasId, guru) {
  const kod = guru.KOD_SEKOLAH;
  const k = kelas_(kelasId, kod, true);
  const set = settings_(kod);
  delete set.ADMIN_PASSWORD; // kata laluan tidak dihantar ke pelayar
  const tahap = tahapKelas_(k);
  if (k) { set.SESI = k.SESI; set.TAHUN = k.TAHUN; set.KELAS = k.NAMA_KELAS; set.NAMA_GURU_KELAS = k.NAMA_GURU_KELAS; }
  else { set.SESI = set.SESI || new Date().getFullYear(); set.KELAS = set.KELAS || ''; }
  set.TAHAP = tahap;
  set.KOD_SEKOLAH = kod;
  const o = {
    tetapan: set,
    guru: awamGuru_(guru),
    kelasAktif: k,
    senaraiKelas: senaraiKelas_(kod),
    murid: k ? readTable_(k.SHEET_MURID) : [],
    tahap: tahap, // TAHAP 1 / 2 kelas aktif — mata pelajaran & pentaksiran ditapis ikut tahap ini
    mp: mpSekolah_(kod, tahap),
    markah: k ? readTable_(k.SHEET_MARKAH) : [],
    tafsiran: readSekolah_(SH.TAFSIRAN, KOD_GLOBAL).filter(r => ikutTahap_(r, tahap)), // seragam untuk semua sekolah
    kemahiran: readCached_(SH.KEMAHIRAN).filter(r => ikutTahap_(r, tahap)),
    bahasa: readCached_(SH.BAHASA),
    panduan: readCached_(SH.PANDUAN),
    serverTime: new Date().toISOString()
  };
  if (guru._master) { o.senaraiSekolah = senaraiSekolah_(); o.akaun = senaraiAkaun_(); }
  return o;
}

/* ---------- POST (semua tindakan edit memerlukan kata laluan) ---------- */
/* ====================== KEHADIRAN SERENTAK (v9) ======================
 * Setiap pelayar menghantar "hadir" setiap ~20 saat dengan halaman & kelas yang sedang dibuka.
 * Disimpan dalam CacheService (bukan Sheet) supaya pantas dan tidak mengunci data.
 * Pelayar menerima senarai guru lain yang sedang membuka halaman yang SAMA. */
const HADIR_TTL = 60; // saat — dianggap sudah keluar jika tiada isyarat
function hadir_(guru, body) {
  const kod = String(guru.KOD_SEKOLAH).toUpperCase(), c = CacheService.getScriptCache(), key = 'HADIR_' + kod;
  const sid = String(body.sid || '').slice(0, 40); if (!sid) return [];
  const page = String(body.page || '').slice(0, 120), now = Date.now();
  const lk = LockService.getScriptLock(); let ada = false; try { ada = lk.tryLock(3000); } catch (e) {}
  try {
    let m = {}; try { m = JSON.parse(c.get(key) || '{}'); } catch (e) {}
    Object.keys(m).forEach(s => { if (now - m[s].t > HADIR_TTL * 1000) delete m[s]; });
    if (page) m[sid] = { n: String(guru.NAMA_GURU || guru.EMEL || 'Guru'), e: String(guru.EMEL || ''), p: page, t: now };
    else delete m[sid];
    try { c.put(key, JSON.stringify(m), 600); } catch (e) {}
    return Object.keys(m).filter(s => s !== sid && page && m[s].p === page).map(s => ({ nama: m[s].n, sejak: m[s].t }));
  } finally { if (ada) lk.releaseLock(); }
}

function doPost(e) {
  try { // isyarat kehadiran: tanpa kunci utama supaya tidak melambatkan guru lain yang sedang menyimpan
    const b0 = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b0.action === 'semakEmel') { // semakan pantas: emel sudah berdaftar? (tanpa kunci, tanpa migrasi)
      const em = normEmel_(b0.emel);
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return out_({ ok: true, sah: false, ada: false });
      return out_({ ok: true, sah: true, ada: !!cariGuru_(em) });
    }
    if (b0.action === 'hadir') { const g0 = sahGuru_(b0.emel, b0.token, b0.sekolah); return out_({ ok: true, lain: hadir_(g0, b0) }); }
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const a = body.action;
    migrasi_();
    if (a === 'daftar') { const g = daftar_(body.data); return out_({ ok: true, emel: g.EMEL, token: g.TOKEN, guru: awamGuru_(g), data: getAll_('', g) }); }
    if (a === 'login') { const g = login_(body.emel, body.password); log_('LOGIN', 'Guru log masuk', g); return out_({ ok: true, emel: g.EMEL, token: g.TOKEN, guru: awamGuru_(g), data: getAll_(body.kelas, g) }); }
    if (a === 'loginMaster') { const tok = loginMaster_(body.password); const g = masterGuru_(body.sekolah); log_('LOGIN_MASTER', 'Master Admin log masuk', g); return out_({ ok: true, emel: MASTER_EMEL, token: tok, guru: awamGuru_(g), data: getAll_(body.kelas, g) }); }
    const guru = sahGuru_(body.emel, body.token, body.sekolah);
    const M = () => { if (!guru._master) throw new Error('Hanya Master Admin dibenarkan membuat perubahan ini.'); };
    const kod = String(guru.KOD_SEKOLAH).toUpperCase();
    if (a === 'getAll') return out_({ ok: true, data: getAll_(body.kelas, guru) });
    let res; let kid = body.kelas; let tokenBaru = null;
    const K = () => kelas_(kid, kod);
    const TH = d => normTahap_(d && d.TAHAP) || tahapKelas_(kelas_(kid, kod, true)); // tahap kelas aktif
    switch (a) {
      case 'createKelas':   res = kid = createKelas_(body.data, kod); break;
      case 'arkibKelas':    res = setStatusKelas_(body.id, body.status || 'ARKIB', kod); break;
      case 'editKelas':     res = editKelas_(body.id, body.data || {}, kod); break;
      case 'padamKelas':    res = padamKelas_(body.id, kod); if (String(kid) === String(body.id)) kid = ''; break;
      case 'saveTetapan': {
        const d = Object.assign({}, body.data || {});
        if (!guru._master && String(d.ADMIN_PASSWORD || '').trim()) tokenBaru = tukarKataLaluan_(guru, String(d.ADMIN_PASSWORD).trim());
        delete d.ADMIN_PASSWORD; delete d.KOD_SEKOLAH;
        const k = kelas_(kid, kod, true); if (k) updateKelas_(k.ID_KELAS, d, kod);
        delete d.TAHAP; // tahap disimpan pada kelas sahaja
        res = saveTetapan_(d, kod); break;
      }
      case 'saveMurid':     res = upsert_(K().SHEET_MURID, 'ID_MURID', body.data, 'M'); break;
      case 'deleteMurid':   res = deleteMurid_(K(), body.id); break;
      case 'saveMP':        res = saveMP_(body.data, kod, guru._master, TH(body.data)); break;
      case 'deleteMP':      M(); res = deleteMP_(body.id, KOD_GLOBAL, TH(body.data)); break; // padam untuk semua sekolah
      case 'saveMarkah':    res = saveMarkah_(K().SHEET_MARKAH, body.kod, body.rows); break;
      case 'saveUlasanKelas': res = saveUlasanKelas_(K().SHEET_MURID, body.rows); break;
      case 'saveTafsiran': { // HANYA Master Admin
        M();
        const d = Object.assign({}, body.data || {}); delete d.SEMUA_SEKOLAH;
        if (!String(d.SET || '').trim() || !String(d.KOD_MP || '').trim()) throw new Error('Set dan Kod MP wajib diisi.');
        res = saveTafsiran_(d, KOD_GLOBAL, TH(d)); break; // seragam: semua sekolah nampak tafsiran yang sama
      }
      case 'kemaskiniAkaun': M(); res = kemaskiniAkaun_(body.id, body.data); break;
      case 'padamAkaun':     M(); res = padamAkaun_(body.id); break;
      default: throw new Error('Tindakan tidak dikenali: ' + a);
    }
    log_(a, JSON.stringify(body.data && (body.data.ADMIN_PASSWORD || body.data.KATA_LALUAN) ? '(maklumat sulit disembunyikan) ' + (body.id || '') : (body.data || body.id || body.kod || '')).slice(0, 300), guru);
    return out_({ ok: true, result: res, token: tokenBaru, data: getAll_(kid, guru) });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }

function setTetapan_(kod, k, v) {
  const sh = sheet_(SH.TETAPAN); const hd = headers_(sh);
  const cK = hd.indexOf('KUNCI') > -1 ? hd.indexOf('KUNCI') : 0, cN = hd.indexOf('NILAI') > -1 ? hd.indexOf('NILAI') : 1, cS = hd.indexOf('KOD_SEKOLAH');
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][cK]).trim() === k && sama_(vals[i][cS], kod)) { sh.getRange(i + 1, cN + 1).setValue(v); return; }
  }
  const o = { KOD_SEKOLAH: kod }; o[hd[cK]] = k; o[hd[cN]] = v; appendObj_(sh, o);
}
function saveTetapan_(data, kod) {
  Object.keys(data || {}).forEach(k => {
    if (KELAS_KEYS[k]) return; // disimpan dalam 00_KELAS
    if (k === 'ADMIN_PASSWORD' || k === 'KOD_SEKOLAH') return;
    setTetapan_(kod, k, data[k]);
  });
  return true;
}
/* upsert bagi sheet dikongsi (mata pelajaran) — padanan KUNCI + KOD_SEKOLAH */
function upsertSekolah_(name, key, data, kod, tahap) {
  data = Object.assign({}, data || {}); data.KOD_SEKOLAH = kod;
  if (!data[key]) throw new Error(key + ' wajib diisi.');
  const sh = sheet_(name); const hd = headers_(sh); const vals = sh.getDataRange().getValues();
  const cK = hd.indexOf(key), cS = hd.indexOf('KOD_SEKOLAH');
  const cT = hd.indexOf('TAHAP'); tahap = normTahap_(tahap);
  let rowIdx = -1, kosong = -1;
  for (let i = 1; i < vals.length; i++) if (String(vals[i][cK]) === String(data[key]) && sama_(vals[i][cS], kod)) {
    if (!tahap || cT < 0) { rowIdx = i; break; }
    const t = normTahap_(vals[i][cT]);
    if (t === tahap) { rowIdx = i; break; }
    if (!t && kosong < 0) kosong = i;
  }
  if (rowIdx < 0 && kosong > -1) rowIdx = kosong; // baris tanpa tahap (dikongsi kedua-dua tahap) — kekal tanpa tahap
  else if (tahap) data.TAHAP = tahap;
  const row = hd.map((h, j) => (data[h] !== undefined ? data[h] : (rowIdx > -1 ? vals[rowIdx][j] : '')));
  if (rowIdx > -1) sh.getRange(rowIdx + 1, 1, 1, hd.length).setValues([row]);
  else sh.appendRow(row);
  return data[key];
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
function deleteMP_(kodMp, kodSekolah, tahap) { // mata pelajaran dikongsi semua kelas sekolah ini; rekod TP lama dikekalkan sebagai sejarah
  const sh = sheet_(SH.MP); const hd = headers_(sh); const c = hd.indexOf('KOD_MP'), cS = hd.indexOf('KOD_SEKOLAH'), cT = hd.indexOf('TAHAP');
  const vals = sh.getDataRange().getValues(); tahap = normTahap_(tahap);
  for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][c]) === String(kodMp) && sama_(vals[i][cS], kodSekolah) && (!tahap || cT < 0 || ikutTahap_({ TAHAP: vals[i][cT] }, tahap))) sh.deleteRow(i + 1);
  return kodMp; }

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
  if (cId < 0 || cU < 0) throw new Error('Lajur ID_MURID atau ULASAN_GURU_KELAS tiada dalam Google Sheet.');
  if (!Array.isArray(rows) || !rows.length) throw new Error('Tiada ulasan untuk disimpan.');
  const vals = sh.getDataRange().getValues();
  const updates = rows.map(r => {
    if (!r || !String(r.ID_MURID || '').trim() || typeof r.ULASAN_GURU_KELAS !== 'string') throw new Error('Maklumat ulasan tidak sah.');
    const matches = vals.slice(1).map((line, i) => String(line[cId]) === String(r.ID_MURID) ? i + 2 : 0).filter(Boolean);
    if (matches.length !== 1) throw new Error('Rekod murid tidak dijumpai atau bertindih. Muat semula sebelum menyimpan.');
    return { row: matches[0], value: r.ULASAN_GURU_KELAS };
  });
  updates.forEach(u => sh.getRange(u.row, cU + 1).setValue(u.value.startsWith('=') ? "'" + u.value : u.value));
  SpreadsheetApp.flush();
  return updates.length;
}

function saveTafsiran_(d, kod, tahap) {
  d = Object.assign({}, d || {}); d.KOD_SEKOLAH = kod; tahap = normTahap_(tahap) || normTahap_(d.TAHAP) || '1'; d.TAHAP = tahap;
  const sh = sheet_(SH.TAFSIRAN); const hd = headers_(sh);
  const vals = sh.getDataRange().getValues();
  const cS = hd.indexOf('SET'), cK = hd.indexOf('KOD_MP'), cSek = hd.indexOf('KOD_SEKOLAH'), cT = hd.indexOf('TAHAP');
  const padan = i => String(vals[i][cS]) === String(d.SET) && String(vals[i][cK]) === String(d.KOD_MP) && sama_(vals[i][cSek], kod);
  let tepat = -1; if (cT > -1) for (let i = 1; i < vals.length; i++) if (padan(i) && normTahap_(vals[i][cT]) === tahap) { tepat = i; break; }
  for (let i = 1; i < vals.length; i++) {
    if (tepat > -1 ? i === tepat : (padan(i) && (cT < 0 || !normTahap_(vals[i][cT])))) {
      ['MATA_PELAJARAN','TP1','TP2','TP3','TP4','TP5','TP6'].forEach(t => { const c = hd.indexOf(t); if (c > -1 && (t !== 'MATA_PELAJARAN' || String(d[t] || '').trim())) sh.getRange(i + 1, c + 1).setValue(d[t] || ''); });
      return true;
    }
  }
  sh.appendRow(hd.map(h => d[h] !== undefined ? d[h] : ''));
  return true;
}

/* Jalankan sekali dari editor untuk menguji sambungan */
function testSambungan() { migrasiSKBenut(); Logger.log(JSON.stringify(getAll_('', { KOD_SEKOLAH: KOD_ASAL, NAMA_GURU: 'UJIAN', EMEL: '' })).slice(0, 1500)); }
