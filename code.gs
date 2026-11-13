/** CHOPPER TEAM - Google Apps Script (Extensions > Apps Script > tempel semua > Deploy > Web app) */
const SH = {UNIT:'Unit', KD:'Kedatangan Kendaraan', JD:'Jadwal'};
const TZ = 'Asia/Jakarta', FIRST = 4;           // data mulai baris 4
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActive(), p = e.parameter || {};
    const today = parseInt(Utilities.formatDate(new Date(), TZ, 'd'));
    const day = parseInt(p.day) || today, A = getAtt(ss, day);
    return out({ok:true, today, day, now:Utilities.formatDate(new Date(), TZ, 'HH:mm'),
      units:getUnits(ss), need:getNeed(ss), arrivals:getArrivals(ss), schedule:getSchedule(ss, day),
      att:A.st, swap:A.sw, ot:A.ot, trouble:getTrouble(ss)});
  } catch (err) { return out({ok:false, error:String(err)}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActive();
    if (d.action === 'att') {
      if (d.status && ['Masuk','Sakit','Ijin','Mangkir'].indexOf(d.status) < 0) throw new Error('Status tidak valid');
      saveAtt(d); return out({ok:true});
    }
    if (d.action === 'lembur') { saveLembur(d); return out({ok:true}); }
    if (d.action === 'swap') { saveSwap(d); return out({ok:true}); }
    if (d.action === 'trouble') { addTrouble(ss, d); return out({ok:true}); }
    if (d.action === 'fix') { fixTrouble(ss, d); return out({ok:true}); }
    const sh = ss.getSheetByName(SH.KD);
    const r = parseInt(d.row);
    if (!(r >= FIRST)) throw new Error('Baris tidak valid');
    if (!sh.getRange(r, 5).getDisplayValue()) throw new Error('Baris kosong');
    const cur = sh.getRange(r, 7).getDisplayValue();
    if (d.unit && d.unit !== cur && openTrouble(ss)[d.unit]) throw new Error('Unit ' + d.unit + ' sedang trouble');
    ensureHeader(sh);
    const rg = sh.getRange(r, 7, 1, 4);
    rg.setNumberFormat('@');
    rg.setValues([[d.unit || '', d.mulai || '', d.selesai || '', durasi(d.mulai, d.selesai)]]);
    sh.getRange(r, 10).setNumberFormat('0');
    if (['', 'Shift 1', 'Shift 2'].indexOf(d.shift || '') < 0) throw new Error('Shift tidak valid');
    sh.getRange(r, 11).setNumberFormat('@').setValue(d.shift || '');
    return out({ok:true});
  } catch (err) { return out({ok:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}

function durasi(a, b) {
  const m = t => { const x = String(t || '').match(/^(\d+)[:.](\d+)/); return x ? +x[1] * 60 + +x[2] : null; };
  const s = m(a), f = m(b);
  if (s === null || f === null) return '';
  let d = f - s; if (d < 0) d += 1440;
  return d;
}

function ensureHeader(sh) {
  const h = sh.getRange(3, 7, 1, 4);
  if (h.getDisplayValue() === '') {
    h.setValues([['Unit Chopper', 'Jam Mulai', 'Jam Selesai', 'Durasi (mnt)']])
     .setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }
  const k = sh.getRange(3, 11);
  if (k.getDisplayValue() === '') k.setValue('Shift').setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
}

function getUnits(ss) {
  const sh = ss.getSheetByName(SH.UNIT), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 1).getDisplayValues().map(r => r[0].trim()).filter(String);
}

function getNeed(ss) {   // kebutuhan pakan per hari = sheet Unit, sel C2
  return parseInt(String(ss.getSheetByName(SH.UNIT).getRange(2, 3).getDisplayValue()).replace(/\D/g, '')) || 0;
}

function getArrivals(ss) {
  const sh = ss.getSheetByName(SH.KD); ensureHeader(sh);
  const n = sh.getLastRow() - FIRST + 1;
  if (n < 1) return [];
  const v = sh.getRange(FIRST, 1, n, 11).getDisplayValues(), res = [];
  v.forEach((r, i) => {
    if (!r[4] || /total/i.test(r[0] + r[2])) return;
    res.push({row:FIRST + i, no:r[0], tgl:r[1], sup:r[2], pic:r[3], plat:r[4], netto:r[5],
              unit:r[6], mulai:r[7], selesai:r[8], durasi:r[9], shift:r[10]});
  });
  return res;
}

function getSchedule(ss, day) {
  const sh = ss.getSheetByName(SH.JD), col = 3 + day;       // D = tanggal 1
  const n = sh.getLastRow() - FIRST + 1;
  if (n < 1 || col > sh.getLastColumn()) return {dayName:'', list:[]};
  const dayName = sh.getRange(3, col).getDisplayValue();
  const v = sh.getRange(FIRST, 1, n, col).getDisplayValues();
  const list = v.filter(r => r[1]).map(r => ({nama:r[1], jabatan:r[2], shift:String(r[col - 1]).trim().toUpperCase()}));
  return {dayName, month:sh.getRange(1, 4).getDisplayValue() || sh.getRange(1, 16).getDisplayValue(), list};
}

/* ===== KEHADIRAN (Masuk / Sakit / Ijin / Mangkir + Tukar Shift) -> sheet "Kehadiran" dibuat otomatis ===== */
function dayKey(day) {
  const n = new Date(), y = Utilities.formatDate(n, TZ, 'yyyy'), m = Utilities.formatDate(n, TZ, 'MM');
  return y + '-' + m + '-' + ('0' + day).slice(-2);
}
function attSheet(ss) {
  let sh = ss.getSheetByName('Kehadiran');
  if (!sh) { sh = ss.insertSheet('Kehadiran'); sh.getRange('A:G').setNumberFormat('@'); sh.setFrozenRows(1); }
  if (sh.getRange(1, 7).getDisplayValue() === '')
    sh.getRange(1, 1, 1, 7).setValues([['Tanggal', 'Nama', 'Status', 'Dicatat', 'Tukar Shift Dengan', 'Lembur (jam)', 'Deskripsi Lembur']])
      .setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold');
  return sh;
}
function getAtt(ss, day) {
  const sh = attSheet(ss), n = sh.getLastRow() - 1, k = dayKey(day), st = {}, sw = {}, ot = {};
  if (n < 1) return {st, sw, ot};
  sh.getRange(2, 1, n, 7).getDisplayValues().forEach(r => {
    if (r[0] !== k) return;
    if (r[2]) st[r[1]] = r[2];
    if (r[4]) sw[r[1]] = r[4];
    if (r[5]) ot[r[1]] = {jam:r[5], desc:r[6]};
  });
  return {st, sw, ot};
}
// ubah status / tukar shift / lembur satu orang; baris dihapus bila semuanya kosong
function upsertAtt(day, nama, patch) {
  const sh = attSheet(SpreadsheetApp.getActive()), k = dayKey(day), n = sh.getLastRow() - 1;
  let row = 0;
  if (n > 0) sh.getRange(2, 1, n, 2).getDisplayValues().some((r, i) => { if (r[0] === k && r[1] === nama) { row = i + 2; return true; } });
  const cur = row ? sh.getRange(row, 1, 1, 7).getDisplayValues()[0] : [k, nama, '', '', '', '', ''];
  if ('status' in patch) cur[2] = patch.status || '';
  if ('partner' in patch) cur[4] = patch.partner || '';
  if ('jam' in patch) { cur[5] = patch.jam || ''; cur[6] = patch.jam ? patch.desc : ''; }
  cur[3] = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  if (!cur[2] && !cur[4] && !cur[5]) { if (row) sh.deleteRow(row); return; }
  if (row) sh.getRange(row, 1, 1, 7).setNumberFormat('@').setValues([cur]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 7).setNumberFormat('@').setValues([cur]);
}
function saveLembur(d) {   // jam: bilangan bulat 1-12 (kelipatan 1 jam); 0 = hapus lembur
  const j = parseInt(d.jam) || 0, desc = String(d.desc || '').trim();
  if (j < 0 || j > 12) throw new Error('Lembur harus 1-12 jam');
  if (j > 0 && !desc) throw new Error('Deskripsi lembur wajib diisi');
  upsertAtt(d.day, d.nama, {jam: j, desc: desc});
}
function saveAtt(d) { upsertAtt(d.day, d.nama, {status: d.status}); }
function saveSwap(d) {
  const p = d.partner || '';
  if (p) {
    if (p === d.nama) throw new Error('Pilih operator lain');
    const sh = attSheet(SpreadsheetApp.getActive()), n = sh.getLastRow() - 1, k = dayKey(d.day);
    if (n > 0) sh.getRange(2, 1, n, 5).getDisplayValues().forEach(r => {
      if (r[0] === k && r[1] !== d.nama && r[4] && (r[1] === p || r[4] === p || r[4] === d.nama))
        throw new Error('Operator sudah punya tukar shift lain');
    });
  }
  upsertAtt(d.day, d.nama, {partner: p});
}

/* ===== TROUBLE UNIT -> sheet "Trouble" dibuat otomatis ===== */
function troubleSheet(ss) {
  let sh = ss.getSheetByName('Trouble');
  if (!sh) {
    sh = ss.insertSheet('Trouble');
    sh.getRange('A:I').setNumberFormat('@');
    sh.getRange(1, 1, 1, 9).setValues([['Tanggal', 'Jam Lapor', 'Unit', 'Masalah', 'Status', 'Nama Mekanik', 'Start Perbaikan', 'Selesai Perbaikan', 'Durasi (mnt)']])
      .setBackground('#c62828').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function getTrouble(ss) {   // semua yang masih trouble + 10 riwayat selesai terakhir
  const sh = troubleSheet(ss), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const all = sh.getRange(2, 1, n, 9).getDisplayValues()
    .map((r, i) => ({row:i + 2, tgl:r[0], jam:r[1], unit:r[2], masalah:r[3], status:r[4], mekanik:r[5], start:r[6], selesai:r[7], durasi:r[8]}))
    .filter(t => t.unit);
  return all.filter(t => t.status !== 'Selesai').concat(all.filter(t => t.status === 'Selesai').slice(-10)).sort((a, b) => a.row - b.row);
}
function openTrouble(ss) {
  const m = {}; getTrouble(ss).forEach(t => { if (t.status !== 'Selesai') m[t.unit] = 1; }); return m;
}
function addTrouble(ss, d) {
  const m = String(d.masalah || '').trim();
  if (getUnits(ss).indexOf(d.unit) < 0) throw new Error('Unit tidak valid');
  if (!m) throw new Error('Deskripsi masalah wajib diisi');
  if (openTrouble(ss)[d.unit]) throw new Error('Unit ini sudah dilaporkan trouble');
  const sh = troubleSheet(ss), now = new Date();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 9).setNumberFormat('@')
    .setValues([[Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), Utilities.formatDate(now, TZ, 'HH:mm'), d.unit, m, 'Trouble', '', '', '', '']]);
}
function fixTrouble(ss, d) {
  const sh = troubleSheet(ss), r = parseInt(d.row);
  if (!(r >= 2)) throw new Error('Baris tidak valid');
  const cur = sh.getRange(r, 1, 1, 9).getDisplayValues()[0];
  if (!cur[2]) throw new Error('Data trouble tidak ditemukan');
  if (cur[4] === 'Selesai') throw new Error('Perbaikan sudah selesai');
  const mk = String(d.mekanik || '').trim(), st = String(d.start || '').trim(), se = String(d.selesai || '').trim();
  if (!mk) throw new Error('Nama mekanik wajib diisi');
  if (!st) throw new Error('Start perbaikan wajib diisi');
  let dur = '';
  if (se) {
    dur = Math.round((Utilities.parseDate(se, TZ, 'yyyy-MM-dd HH:mm') - Utilities.parseDate(st, TZ, 'yyyy-MM-dd HH:mm')) / 60000);
    if (dur < 0) throw new Error('Selesai harus setelah start');
  }
  sh.getRange(r, 5, 1, 5).setNumberFormat('@').setValues([[se ? 'Selesai' : 'Trouble', mk, st, se, dur]]);
}

/* ===== ARSIP OTOMATIS TIAP TENGAH MALAM ===== */
// Jalankan SEKALI secara manual: pilih fungsi setupTrigger > Run > izinkan akses
function setupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'archiveDaily').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('archiveDaily').timeBased().atHour(0).nearMinute(5).everyDays(1).inTimezone(TZ).create();
}
function archiveDaily() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive(), kd = ss.getSheetByName(SH.KD), ar = ss.getSheetByName('Arsip');
    const n = kd.getLastRow() - FIRST + 1; if (n < 1) return 0;
    const v = kd.getRange(FIRST, 1, n, 11).getDisplayValues(), an = ar.getLastRow() - 1, seen = {}, add = [];
    if (an > 0) ar.getRange(2, 1, an, 5).getDisplayValues().forEach(r => seen[r[1] + '|' + r[0] + '|' + r[4]] = 1);
    v.forEach(r => {
      if (!r[4] || /total/i.test(r[0] + r[2])) return;
      const k = r[1] + '|' + r[0] + '|' + r[4];
      if (seen[k]) return; seen[k] = 1; add.push(r);          // anti dobel
    });
    if (add.length) ar.getRange(ar.getLastRow() + 1, 1, add.length, 11).setNumberFormat('@').setValues(add);
    kd.getRange(FIRST, 7, n, 5).clearContent();                // kosongkan Unit/Jam/Durasi/Shift utk hari baru
    return add.length;
  } finally { lock.releaseLock(); }
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Chopper').addItem('Arsipkan sekarang', 'archiveNow').addToUi();
}
function archiveNow() { SpreadsheetApp.getUi().alert(archiveDaily() + ' baris dipindah ke Arsip'); }
