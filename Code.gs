/**
 * เจ๊แตง รถเกี่ยวข้าว – ระบบสำรองข้อมูล (Google Apps Script)
 *
 * วิธีใช้: เปิด Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางโค้ดนี้ทั้งหมด
 * แล้วกด ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > ประเภท "เว็บแอป"
 *   - เรียกใช้ในฐานะ: ฉัน
 *   - ผู้มีสิทธิ์เข้าถึง: ทุกคน
 * คัดลอก URL (ลงท้าย /exec) ไปวางในหน้า "สำรองข้อมูล" ของเว็บแอป
 */

// ตั้งรหัสผ่านของตัวเอง (ตัวอักษรอังกฤษ/ตัวเลข) แล้วใช้รหัสเดียวกันในหน้าตั้งค่าของเว็บแอป
const SECRET = 'CHANGE_ME';

const SHEET_JOBS = 'งานเกี่ยว';
const SHEET_EXP = 'ค่าใช้จ่าย';
const SHEET_LOG = 'ประวัติการสำรอง';

const JOB_HEAD = ['รหัส', 'วันที่', 'ชื่อลูกค้า', 'ไร่', 'งาน', 'ตารางวา', 'รวม (ไร่)', 'ราคา/ไร่',
  'ยอดตามราคา', 'เก็บจริง (บาท)', 'รับเงินแล้ว', 'วันที่รับเงิน', 'หมายเหตุ', 'แก้ไขล่าสุด'];
const EXP_HEAD = ['รหัส', 'วันที่', 'ประเภท', 'จำนวนเงิน (บาท)', 'หมายเหตุ', 'แก้ไขล่าสุด'];

const TYPE_TH = { fuel: 'ค่าน้ำมัน', maint: 'ค่าบำรุงรักษา', other: 'อื่นๆ' };
const TYPE_KEY = { 'ค่าน้ำมัน': 'fuel', 'ค่าบำรุงรักษา': 'maint', 'อื่นๆ': 'other' };

/* ---------- เว็บแอปเรียกเข้ามา ---------- */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    checkSecret_(body.secret);
    if (body.action !== 'backup') throw new Error('unknown action');

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const jobs = body.jobs || [];
      const exps = body.expenses || [];
      writeJobs_(jobs);
      writeExpenses_(exps);
      appendLog_(jobs.length, exps.length);
    } finally {
      lock.releaseLock();
    }
    return out_({ ok: true, jobs: (body.jobs || []).length, expenses: (body.expenses || []).length });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    checkSecret_(p.secret);
    if (p.action === 'ping') {
      return out_({ ok: true, name: ss_().getName() });
    }
    if (p.action === 'restore') {
      return out_({ ok: true, jobs: readJobs_(), expenses: readExpenses_() });
    }
    return out_({ ok: true, message: 'เจ๊แตง backup API พร้อมใช้งาน' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- เขียนข้อมูล ---------- */

function writeJobs_(jobs) {
  const rows = jobs.map(function (j) {
    const total = n_(j.rai) + n_(j.ngan) / 4 + n_(j.wa) / 400;
    return [
      j.id, j.date, j.customer, blank_(j.rai), blank_(j.ngan), blank_(j.wa),
      Math.round(total * 1000) / 1000, n_(j.rate), Math.round(total * n_(j.rate)),
      n_(j.charged), !!j.paid, j.paidDate || '', j.note || '', j.updatedAt || ''
    ];
  });
  // คอลัมน์ที่เป็นข้อความล้วน (กัน Sheets แปลงวันที่เอง): id=1, date=2, paidDate=12, updatedAt=14
  writeTable_(SHEET_JOBS, JOB_HEAD, rows, [1, 2, 12, 14]);
}

function writeExpenses_(exps) {
  const rows = exps.map(function (x) {
    return [x.id, x.date, TYPE_TH[x.type] || 'อื่นๆ', n_(x.amount), x.note || '', x.updatedAt || ''];
  });
  writeTable_(SHEET_EXP, EXP_HEAD, rows, [1, 2, 6]);
}

function writeTable_(name, head, rows, textCols) {
  const sh = sheet_(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#e6197f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (rows.length) {
    textCols.forEach(function (c) {
      sh.getRange(2, c, rows.length, 1).setNumberFormat('@');
    });
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  }
  sh.autoResizeColumns(1, head.length);
}

function appendLog_(jobCount, expCount) {
  const sh = sheet_(SHEET_LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['เวลา', 'จำนวนงาน', 'จำนวนรายการค่าใช้จ่าย']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#ffb800');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), jobCount, expCount]);
  // เก็บประวัติแค่ 200 บรรทัดล่าสุด
  const extra = sh.getLastRow() - 201;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* ---------- อ่านข้อมูลกลับ ---------- */

function readJobs_() {
  return rows_(SHEET_JOBS).map(function (r) {
    return {
      id: String(r[0]), date: d_(r[1]), customer: String(r[2] || ''),
      rai: r[3], ngan: r[4], wa: r[5],
      rate: num_(r[7]), charged: num_(r[9]),
      paid: r[10] === true || String(r[10]).toUpperCase() === 'TRUE',
      paidDate: d_(r[11]), note: String(r[12] || ''), updatedAt: String(r[13] || '')
    };
  });
}

function readExpenses_() {
  return rows_(SHEET_EXP).map(function (r) {
    return {
      id: String(r[0]), date: d_(r[1]), type: TYPE_KEY[r[2]] || 'other',
      amount: num_(r[3]), note: String(r[4] || ''), updatedAt: String(r[5] || '')
    };
  });
}

function rows_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(function (r) { return r[0] !== ''; });
}

/* ---------- ตัวช่วย ---------- */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  return ss_().getSheetByName(name) || ss_().insertSheet(name);
}

function checkSecret_(s) {
  if (SECRET && s !== SECRET) throw new Error('รหัสผ่านไม่ถูกต้อง');
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function n_(v) { const x = parseFloat(v); return isFinite(x) ? x : 0; }
function num_(v) { return n_(v); }
function blank_(v) { return v === '' || v == null ? '' : n_(v); }

function d_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v ? String(v) : '';
}
