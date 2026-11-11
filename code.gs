/** CHOPPER TEAM - Google Apps Script (Extensions > Apps Script > tempel semua > Deploy > Web app) */
const SH = {UNIT:'Unit', KD:'Kedatangan Kendaraan', JD:'Jadwal'};
const TZ = 'Asia/Jakarta', FIRST = 4;           // data mulai baris 4
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActive(), p = e.parameter || {};
    const today = parseInt(Utilities.formatDate(new Date(), TZ, 'd'));
    const day = parseInt(p.day) || today;
    return out({ok:true, today, day, now:Utilities.formatDate(new Date(), TZ, 'HH:mm'),
      units:getUnits(ss), arrivals:getArrivals(ss), schedule:getSchedule(ss, day), att:getAtt(ss, day)});
  } catch (err) { return out({ok:false, error:String(err)}); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'att') {
      if (d.status && ['Masuk','Sakit','Mangkir'].indexOf(d.status) < 0) throw new Error('Status tidak valid');
      saveAtt(d); return out({ok:true});
    }
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.KD);
    const r = parseInt(d.row);
    if (!(r >= FIRST)) throw new Error('Baris tidak valid');
    if (!sh.getRange(r, 5).getDisplayValue()) throw new Error('Baris kosong');
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

/* ===== KEHADIRAN (Masuk / Sakit / Mangkir) -> sheet "Kehadiran" dibuat otomatis ===== */
function dayKey(day) {
  const n = new Date(), y = Utilities.formatDate(n, TZ, 'yyyy'), m = Utilities.formatDate(n, TZ, 'MM');
  return y + '-' + m + '-' + ('0' + day).slice(-2);
}
function attSheet(ss) {
  let sh = ss.getSheetByName('Kehadiran');
  if (!sh) {
    sh = ss.insertSheet('Kehadiran');
    sh.getRange(1, 1, 1, 4).setValues([['Tanggal', 'Nama', 'Status', 'Dicatat']])
      .setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@'); sh.setFrozenRows(1);
  }
  return sh;
}
function getAtt(ss, day) {
  const sh = attSheet(ss), n = sh.getLastRow() - 1, k = dayKey(day), m = {};
  if (n < 1) return m;
  sh.getRange(2, 1, n, 3).getDisplayValues().forEach(r => { if (r[0] === k) m[r[1]] = r[2]; });
  return m;
}
function saveAtt(d) {
  const sh = attSheet(SpreadsheetApp.getActive()), k = dayKey(d.day), n = sh.getLastRow() - 1;
  let row = 0;
  if (n > 0) sh.getRange(2, 1, n, 2).getDisplayValues().some((r, i) => { if (r[0] === k && r[1] === d.nama) { row = i + 2; return true; } });
  if (!d.status) { if (row) sh.deleteRow(row); return; }
  const t = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  if (row) sh.getRange(row, 3, 1, 2).setValues([[d.status, t]]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[k, d.nama, d.status, t]]);
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
