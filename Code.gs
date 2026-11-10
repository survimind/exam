/**
 * ห้องสอบ · ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 *
 * ข้อมูลทั้งหมดเก็บใน Google Sheet ที่ผูกกับสคริปต์นี้ (มี 3 แท็บ: exams, keys, attempts)
 * - เฉลย (keys) อยู่ฝั่งเซิร์ฟเวอร์เท่านั้น นักเรียนดึงไม่ได้
 * - รหัสผู้สอน เก็บใน Script Properties ชื่อ TEACHER_TOKEN ไม่ได้เขียนไว้ในโค้ด
 */

var GRACE_MS = 60 * 1000;          // เผื่อเวลาให้นักเรียนที่เน็ตหลุดตอนหมดเวลา
var TAB = { exams: 'exams', keys: 'keys', attempts: 'attempts' };

/* ---------- ทางเข้า ---------- */

function doGet() {
  return json_({ ok: true, service: 'exam-room', time: Date.now() });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    return json_(route_(body));
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function route_(b) {
  switch (b.action) {
    case 'exams':   return { ok: true, exams: listExams_() };          // ข้อสอบทั้งหมด (ไม่มีเฉลย)
    case 'status':  return { ok: true, attempts: myAttempts_(b) };      // ผลของนักเรียนคนนั้นเอง
    case 'start':   return startAttempt_(b);
    case 'save':    return saveAttempt_(b, false);
    case 'submit':  return saveAttempt_(b, true);
    case 'teacher': return teacherOp_(b);
    default:        return { ok: false, error: 'unknown_action' };
  }
}

/* ---------- ตัวช่วยเรื่องชีต ---------- */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['id', 'label', 'json', 'updatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function readAll_(name) {
  var vals = sheet_(name).getDataRange().getValues(), out = {};
  for (var i = 1; i < vals.length; i++) {
    var id = String(vals[i][0]);
    if (!id) continue;
    try { out[id] = JSON.parse(vals[i][2]); } catch (e) {}
  }
  return out;
}
function findRow_(sh, id) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}
function put_(name, id, obj, label) {
  var sh = sheet_(name), row = findRow_(sh, id);
  var vals = [id, label || '', JSON.stringify(obj), obj.updatedAt || Date.now()];
  if (row) sh.getRange(row, 1, 1, 4).setValues([vals]);
  else sh.appendRow(vals);
}
function del_(name, id) {
  var sh = sheet_(name), row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
}

/* ---------- ข้อสอบ ---------- */

function listExams_() {
  var ex = readAll_(TAB.exams), out = [];
  Object.keys(ex).forEach(function (id) { var e = ex[id]; e.id = id; out.push(e); });
  out.sort(function (a, b) { return (b.updatedAt || 0) - (a.updatedAt || 0); });
  return out;
}
function listKeys_() {
  return readAll_(TAB.keys); // { examId: { answers: {qid: index}, updatedAt } }
}

/* ---------- นักเรียน ---------- */

function normName_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function attKey_(examId, no) { return examId + '|' + no; }
function checkStudent_(b) {
  var no = String(b.no == null ? '' : b.no).trim();
  var name = normName_(b.name);
  if (!/^\d{1,3}$/.test(no)) return { error: 'bad_no' };
  if (!name) return { error: 'bad_name' };
  return { no: no, name: name };
}

function myAttempts_(b) {
  var who = checkStudent_(b);
  if (who.error) return {};
  var all = readAll_(TAB.attempts), out = {};
  Object.keys(all).forEach(function (k) {
    var a = all[k];
    if (a.no === who.no && normName_(a.name) === who.name) out[a.examId] = a;
  });
  return out;
}

function startAttempt_(b) {
  var who = checkStudent_(b);
  if (who.error) return { ok: false, error: who.error };
  var ex = readAll_(TAB.exams)[b.examId];
  if (!ex) return { ok: false, error: 'no_exam' };

  var key = attKey_(b.examId, who.no);
  var a = readAll_(TAB.attempts)[key];
  if (a && normName_(a.name) !== who.name) return { ok: false, error: 'name_mismatch' };
  if (a && a.status === 'submitted') return { ok: false, error: 'submitted', attempt: a };

  var now = Date.now();
  if (a) {
    // กลับมาทำต่อ: ถ้าหมดเวลาไปแล้วให้ส่งให้อัตโนมัติ
    if (now > a.endAt + GRACE_MS) return finalize_(a, 'timeout', now, key, who);
    if (!ex.open) return { ok: false, error: 'closed' };
    a.events.push({ t: 'rejoin', at: now });
    a.status = 'doing';
  } else {
    if (!ex.open) return { ok: false, error: 'closed' };
    a = {
      examId: b.examId, title: ex.title, no: who.no, name: who.name,
      proctor: !!ex.proctor, startedAt: now,
      endAt: now + (Number(ex.minutes) || 30) * 60000,
      status: 'doing', answers: {}, events: [{ t: 'start', at: now }], updatedAt: now
    };
  }
  a.updatedAt = now;
  put_(TAB.attempts, key, a, who.name + ' · เลขที่ ' + who.no + ' · ' + ex.title);
  return { ok: true, attempt: a };
}

function saveAttempt_(b, isSubmit) {
  var who = checkStudent_(b);
  if (who.error) return { ok: false, error: who.error };
  var ex = readAll_(TAB.exams)[b.examId];
  if (!ex) return { ok: false, error: 'no_exam' };

  var key = attKey_(b.examId, who.no);
  var a = readAll_(TAB.attempts)[key];
  if (!a) return { ok: false, error: 'reset' };                 // ยังไม่เริ่ม หรือครูให้ทำใหม่แล้ว
  if (normName_(a.name) !== who.name) return { ok: false, error: 'name_mismatch' };
  if (a.status === 'submitted') return { ok: true, attempt: a };  // ส่งไปแล้ว ไม่รับคำตอบเพิ่ม

  var now = Date.now();
  a.answers = cleanAnswers_(b.answers, ex);
  a.events = cleanEvents_(b.events);
  if (isSubmit) return finalize_(a, b.kind === 'timeout' ? 'timeout' : 'submit', now, key, who);
  if (now > a.endAt + GRACE_MS) return finalize_(a, 'timeout', now, key, who);
  a.status = b.status === 'away' ? 'away' : 'doing';
  a.updatedAt = now;
  put_(TAB.attempts, key, a, who.name + ' · เลขที่ ' + who.no + ' · ' + ex.title);
  return { ok: true, attempt: a };
}

function finalize_(a, kind, now, key, who) {
  var last = a.events[a.events.length - 1];
  if (!last || (last.t !== 'submit' && last.t !== 'timeout')) a.events.push({ t: kind, at: now });
  a.status = 'submitted';
  a.submittedAt = now;
  a.updatedAt = now;
  put_(TAB.attempts, key, a, who.name + ' · เลขที่ ' + who.no + ' · ' + a.title);
  return { ok: true, attempt: a };
}

// รับเฉพาะคำตอบของข้อที่มีอยู่จริง และดัชนีตัวเลือกต้องอยู่ในช่วง
function cleanAnswers_(ans, ex) {
  var out = {};
  if (!ans || typeof ans !== 'object') return out;
  (ex.questions || []).forEach(function (q) {
    var v = ans[q.id];
    if (typeof v === 'number' && v >= 0 && v < (q.choices || []).length && v === Math.floor(v)) out[q.id] = v;
  });
  return out;
}
// เก็บเหตุการณ์เฉพาะฟิลด์ที่รู้จัก และไม่เกิน 300 รายการล่าสุด
function cleanEvents_(evs) {
  if (!Array.isArray(evs)) return [];
  var out = [];
  evs.slice(-300).forEach(function (e) {
    if (!e || typeof e.t !== 'string' || typeof e.at !== 'number') return;
    var c = { t: e.t.slice(0, 20), at: e.at };
    if (e.r != null) c.r = String(e.r).slice(0, 10);
    if (e.d != null) c.d = Number(e.d) || 0;
    out.push(c);
  });
  return out;
}

/* ---------- ผู้สอน (ต้องมีรหัส) ---------- */

function checkToken_(t) {
  var real = PropertiesService.getScriptProperties().getProperty('TEACHER_TOKEN');
  return !!real && String(t || '') === real;
}

function teacherOp_(b) {
  if (!checkToken_(b.token)) return { ok: false, error: 'unauthorized' };
  switch (b.op) {
    case 'list':
      return { ok: true, exams: listExams_(), keys: listKeys_(), attempts: values_(readAll_(TAB.attempts)) };
    case 'saveExam':
      return saveExam_(b);
    case 'deleteExam':
      var id = cleanId_(b.examId);
      if (!id) return { ok: false, error: 'bad_id' };
      del_(TAB.exams, id);
      del_(TAB.keys, id);
      return { ok: true };
    case 'toggleOpen':
      var eid = cleanId_(b.examId), ex = readAll_(TAB.exams)[eid];
      if (!ex) return { ok: false, error: 'no_exam' };
      ex.open = !!b.open;
      ex.updatedAt = Date.now();
      put_(TAB.exams, eid, ex, ex.title);
      return { ok: true };
    case 'reset':
      var rid = cleanId_(b.examId), rno = String(b.no || '').trim();
      if (!rid || !rno) return { ok: false, error: 'bad_request' };
      del_(TAB.attempts, attKey_(rid, rno));
      return { ok: true };
    default:
      return { ok: false, error: 'unknown_op' };
  }
}

function saveExam_(b) {
  var e = b.exam || {};
  var id = cleanId_(e.id);
  if (!id) return { ok: false, error: 'bad_id' };
  var title = String(e.title || '').trim();
  if (!title) return { ok: false, error: 'no_title' };
  var minutes = Math.round(Number(e.minutes));
  if (!(minutes >= 1 && minutes <= 600)) return { ok: false, error: 'bad_minutes' };
  var answers = b.answers || {};
  var qs = (e.questions || []).map(function (q, i) {
    var cs = (q.choices || []).map(function (c) { return String(c || '').trim(); }).filter(Boolean);
    if (!String(q.text || '').trim()) throw new Error('ข้อที่ ' + (i + 1) + ' ยังไม่มีโจทย์');
    if (cs.length < 2) throw new Error('ข้อที่ ' + (i + 1) + ' ต้องมีตัวเลือกอย่างน้อย 2 ตัว');
    return { id: String(q.id), text: String(q.text).trim(), choices: cs, pts: Number(q.pts) || 0 };
  });
  if (!qs.length) return { ok: false, error: 'no_questions' };

  var keyAns = {};
  qs.forEach(function (q) {
    var a = answers[q.id];
    if (typeof a !== 'number' || a < 0 || a >= q.choices.length) throw new Error('ข้อ ' + q.id + ' ยังไม่มีเฉลยที่ถูกต้อง');
    keyAns[q.id] = a;
  });

  var now = Date.now();
  put_(TAB.exams, id, {
    title: title, minutes: minutes, proctor: !!e.proctor, shuffle: !!e.shuffle,
    open: !!e.open, questions: qs, updatedAt: now
  }, title);
  put_(TAB.keys, id, { answers: keyAns, updatedAt: now }, title);
  return { ok: true, id: id };
}

function cleanId_(v) {
  var s = String(v == null ? '' : v);
  return /^[A-Za-z0-9_-]{1,60}$/.test(s) ? s : '';
}
function values_(obj) { return Object.keys(obj).map(function (k) { return obj[k]; }); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ตั้งค่าครั้งแรก (รันจากเมนู Run ในตัวแก้ไข) ---------- */

// ใส่ข้อสอบ "แบบทดสอบ บทที่ 1" จากข้อมูลเดิม เปิดรับ = false เพื่อให้ครูตรวจเฉลยก่อน
// เฉลยนี้ "เดา" จากเนื้อหา PDF ยังไม่ใช่เฉลยจริง ครูต้องตรวจก่อนเปิดรับ
function seedFinance1() {
  var now = Date.now();
  var questions = [
    ['q01','การเงิน (Finance) หมายถึงข้อใด',['การเก็บเงินไว้ในบ้านเพื่อใช้ยามจำเป็น','การศึกษาวิธีจัดหา แบ่งสรร และใช้ทรัพยากรทางการเงิน','การจ่ายเงินซื้อสินค้าเพื่อใช้ในชีวิตประจำวัน','การกู้ยืมเงินจากธนาคารมาใช้จ่าย','การลงทุนซื้อหุ้นเพื่อหวังกำไร']],
    ['q02','"เป็ด 1 ตัวราคา 500 บาท หมู 1 ตัวราคา 1,000 บาท ทำให้รู้ว่าหมู 1 ตัวมีค่าเท่ากับเป็ด 2 ตัว" ข้อความนี้แสดงหน้าที่ใดของเงิน',['เป็นสื่อกลางในการแลกเปลี่ยน','เป็นเครื่องรักษามูลค่า','เป็นมาตรฐานการชำระหนี้','เป็นเครื่องวัดมูลค่า','เป็นเครื่องมือบริหารความเสี่ยง']],
    ['q03','บริษัทหักเงินเดือนพนักงานทุกเดือนเข้ากองทุนสำรองเลี้ยงชีพ การออมลักษณะนี้เรียกว่าอะไร',['การออมโดยวิธีบังคับ','การยกเว้นรายจ่ายที่ไม่จำเป็น','การประหยัดการใช้จ่ายในรายได้พิเศษ','กำไรจากการขายหลักทรัพย์','การฝากเงินประจำกับธนาคารด้วยตนเอง']],
    ['q04','ข้อใดเป็นการลงทุนทางการเงิน',['การสร้างโรงงานแห่งใหม่','การซื้อเครื่องจักรเพื่อใช้ในการผลิต','การก่อสร้างระบบสาธารณูปโภคของภาครัฐ','การซื้อหุ้นสามัญหรือพันธบัตร','การซื้อที่ดินเพื่อสร้างอาคารสำนักงาน']],
    ['q05','การซื้อตู้เย็นไว้ใช้ในบ้านเพื่อความสะดวกสบาย โดยไม่หวังกำไร จัดเป็นการลงทุนประเภทใด',['การลงทุนเพื่อการบริโภค','การลงทุนในธุรกิจ','การลงทุนในหลักทรัพย์','การลงทุนเพื่อเก็งกำไร','การลงทุนทางการเงินระยะยาว']],
    ['q06','นางสาวมณีต้องการลงทุนที่ให้ผลตอบแทนสม่ำเสมอทุกงวด เพื่อนำไปใช้จ่ายและวางแผนการเงินได้แน่นอน จุดมุ่งหมายในการลงทุนของนางสาวมณีตรงกับข้อใด',['ความปลอดภัยของเงินต้น','การงอกเงยของเงินทุน','เสถียรภาพของรายได้','การเก็งกำไรระยะสั้น','การสร้างสภาพคล่อง']],
    ['q07','รัฐบาลเปลี่ยนนโยบายการจัดเก็บภาษี และเกิดความไม่มั่นคงทางการเมือง ทำให้นักลงทุนชะลอการลงทุน สถานการณ์นี้เกิดจากปัจจัยใด',['ปัจจัยด้านเศรษฐกิจ','ปัจจัยด้านการเมือง','ปัจจัยด้านผลตอบแทน','ปัจจัยด้านแหล่งข้อมูลข่าวสาร','ปัจจัยด้านสังคมและวัฒนธรรม']],
    ['q08','นักลงทุนตัดสินใจซื้อหุ้นหลังติดตามข่าวผลประกอบการของบริษัทผ่านสื่อต่าง ๆ สะท้อนปัจจัยที่กำหนดการลงทุนข้อใด',['ปัจจัยด้านเศรษฐกิจ','ปัจจัยด้านการเมือง','ปัจจัยด้านความเสี่ยงในการลงทุน','ปัจจัยด้านผลตอบแทน','ปัจจัยด้านแหล่งข้อมูลข่าวสาร']],
    ['q09','ข้อใดเป็นตัวอย่างของการลงทุนในอสังหาริมทรัพย์',['การซื้อทองคำแท่งเก็บไว้','การซื้อพันธบัตรรัฐบาล','การซื้อหุ้นในตลาดหลักทรัพย์','การซื้อคอนโดมิเนียมเพื่อปล่อยเช่า','การซื้อสกุลเงินดิจิทัล']],
    ['q10','ข้อใดเป็นข้อดีของการลงทุนในหุ้น',['ซื้อขายได้สะดวกในช่วงที่ตลาดหลักทรัพย์เปิดทำการ','ราคาไม่มีความผันผวนเลย','ไม่มีความเสี่ยงใด ๆ','ได้รับดอกเบี้ยคงที่ทุกเดือน','ไม่ต้องใช้เงินลงทุนเลย']],
    ['q11','เพราะเหตุใดการลงทุนในอสังหาริมทรัพย์จึงไม่เหมาะกับผู้ที่ต้องการแปลงการลงทุนเป็นเงินสดอย่างรวดเร็ว',['ให้ผลตอบแทนต่ำกว่าเงินฝากธนาคาร','ซื้อขายเปลี่ยนมือได้ช้า','ไม่สามารถนำไปปล่อยเช่าได้','ราคาเปลี่ยนแปลงทุกวินาที','ไม่มีกฎหมายรองรับการซื้อขาย']],
    ['q12','ผลตอบแทนจากการลงทุนในค่าเงินต่างประเทศได้มาจากสิ่งใด',['ดอกเบี้ยที่รัฐบาลจ่ายให้ทุกปี','เงินปันผลจากบริษัท','ส่วนต่างของอัตราแลกเปลี่ยนจากการซื้อขายสกุลเงิน','ค่าเช่ารายเดือน','สิทธิลดหย่อนภาษี']],
    ['q13','ข้อใดเป็นความเสี่ยงจากการลงทุนในเงินดิจิทัล',['ยังไม่ได้รับการรับรองจากรัฐบาล และไม่มีทรัพย์สินค้ำประกัน','ใช้ซื้อสินค้าได้ทั้งในโลกเสมือนและโลกจริง','มีสกุลเงินให้เลือกลงทุนมากกว่า 10,000 สกุล','ได้รับความสนใจในฐานะเงินบนโลกออนไลน์','เป็นที่ยอมรับจากธนาคารกลางทั่วโลก']],
    ['q14','นายเอกนำเงินไปฝากกับธนาคารพาณิชย์ จากนั้นธนาคารนำเงินฝากไปปล่อยกู้ให้ธุรกิจ การเคลื่อนย้ายเงินทุนลักษณะนี้เป็นแบบใด',['การเคลื่อนย้ายเงินทุนโดยตรง','การเคลื่อนย้ายเงินทุนผ่านตลาดหลักทรัพย์','การเคลื่อนย้ายเงินทุนผ่านวาณิชธนกิจ','การเคลื่อนย้ายเงินทุนโดยอ้อมผ่านตัวกลางทางการเงิน','การเคลื่อนย้ายเงินทุนโดยตรงผ่านบริษัทนายหน้า']],
    ['q15','ธุรกิจนำเข้าสินค้าทำสัญญาซื้อเงินตราต่างประเทศล่วงหน้ากับธนาคารพาณิชย์ เพื่อป้องกันความเสี่ยงจากการที่เงินบาทอ่อนค่า แสดงบทบาทใดของระบบการเงิน',['ส่งเสริมการออม','ช่องทางในการบริหารความเสี่ยง','บริการด้านการชำระราคา','เครื่องมือในการบริหารเศรษฐกิจของภาครัฐ','สนับสนุนการระดมเงินทุนและสร้างสภาพคล่อง']]
  ];
  var answers = { q01:1,q02:3,q03:0,q04:3,q05:0,q06:2,q07:1,q08:4,q09:3,q10:0,q11:1,q12:2,q13:0,q14:3,q15:1 };
  var qs = questions.map(function (q) {
    return { id: q[0], text: q[1], choices: q[2], pts: 1 };
  });
  var title = 'แบบทดสอบ บทที่ 1 ความรู้เบื้องต้นเกี่ยวกับการเงินและการลงทุน';
  put_(TAB.exams, 'finance1', { title: title, minutes: 30, proctor: true, shuffle: false, open: false, questions: qs, updatedAt: now }, title);
  put_(TAB.keys, 'finance1', { answers: answers, updatedAt: now }, title);
}
