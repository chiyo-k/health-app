/**
 * マンホール商会 健康記録アプリ（Google Apps Script）
 *
 * 従業員：社員番号＋暗証番号でログインし、血圧・体重（身長）を入力する。
 * 経理担当者：管理者としてログインし、判定期（4月・10月）ごとの一覧・手当対象者を確認する。
 *
 * 判定基準は厚生労働省「標準的な健診・保健指導プログラム」の特定健診の判定値を参考にしている。
 * 社内基準を変える場合は CRITERIA の値だけを書き換える。
 */

var CRITERIA = {
  // 血圧：保健指導判定値（収縮期130mmHg以上 または 拡張期85mmHg以上）に該当しなければクリア
  bpSystolicBelow: 130,
  bpDiastolicBelow: 85,
  // 血圧：受診勧奨判定値（収縮期140mmHg以上 または 拡張期90mmHg以上）は受診を促すメッセージを出す
  bpReferSystolic: 140,
  bpReferDiastolic: 90,
  // BMI：18.5以上25未満（普通体重）ならクリア
  bmiMin: 18.5,
  bmiBelow: 25,
  // 体重：前回の判定測定から3%以上減っていればクリア
  weightLossPct: 3
};

// 判定を行う月（4月・10月）
var JUDGE_MONTHS = [4, 10];

var SHEET_EMPLOYEES = '従業員';
var SHEET_RECORDS = '記録';
var SHEET_RESULTS = '判定結果';

var EMP_HEADERS = ['社員番号', '氏名', '身長(cm)', '暗証番号', '管理者', '在籍'];
var REC_HEADERS = ['登録日時', '社員番号', '氏名', '測定日', '判定期', '収縮期血圧', '拡張期血圧', '体重(kg)', '身長(cm)', 'BMI', 'メモ'];
var RESULT_HEADERS = ['判定期', '社員番号', '氏名', '測定日', '血圧', '血圧判定', 'BMI', 'BMI判定', '体重(kg)', '前回体重(kg)', '減少率(%)', '体重判定', 'クリア数', '手当対象'];

var SESSION_HOURS = 6;
var MAX_LOGIN_FAILS = 5;

/* ------------------------------------------------------------------ */
/* Web アプリの入口                                                    */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('マンホール商会 健康記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('健康記録アプリ')
    .addItem('初期設定（シート作成）', 'setup')
    .addItem('暗証番号を暗号化する', 'hashAllPins')
    .addItem('今期の判定結果を書き出す', 'menuExportCurrentPeriod')
    .addToUi();
}

/** 初回のみ実行：必要なシートと見出しを作る */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_EMPLOYEES, EMP_HEADERS);
  ensureSheet_(ss, SHEET_RECORDS, REC_HEADERS);
  ensureSheet_(ss, SHEET_RESULTS, RESULT_HEADERS);
  var emp = ss.getSheetByName(SHEET_EMPLOYEES);
  if (emp.getLastRow() === 1) {
    emp.appendRow(['0001', '見本 太郎', 170, '1234', '管理者', '在籍']);
  }
  // 社員番号・暗証番号の先頭の0が消えないよう文字列扱いにする
  emp.getRange('A:A').setNumberFormat('@');
  emp.getRange('D:D').setNumberFormat('@');
  // 判定期（2026-04 など）が日付に自動変換されないよう文字列扱いにする
  ss.getSheetByName(SHEET_RECORDS).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_RECORDS).getRange('E:E').setNumberFormat('@');
  ss.getSheetByName(SHEET_RESULTS).getRange('A:B').setNumberFormat('@');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F0E8');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ------------------------------------------------------------------ */
/* 判定ロジック（シートに依存しない純粋関数）                            */
/* ------------------------------------------------------------------ */

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function calcBmi(weightKg, heightCm) {
  if (!weightKg || !heightCm) return null;
  var m = heightCm / 100;
  return round1_(weightKg / (m * m));
}

/** 測定日（yyyy-mm-dd）から判定期を返す。4月・10月以外は ''（任意の記録） */
function periodOf(dateStr) {
  var parts = String(dateStr).split('-');
  var y = Number(parts[0]);
  var m = Number(parts[1]);
  if (JUDGE_MONTHS.indexOf(m) === -1) return '';
  return y + '-' + (m < 10 ? '0' + m : String(m));
}

/** 判定期の表示名（例：2026-10 → 2026年10月） */
function periodLabel(period) {
  var p = String(period).split('-');
  return p[0] + '年' + Number(p[1]) + '月';
}

/** 現在（または指定日）時点で最新の判定期 */
function currentPeriod(today) {
  var d = today || new Date();
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  if (m >= 10) return y + '-10';
  if (m >= 4) return y + '-04';
  return (y - 1) + '-10';
}

/** 1つ前の判定期 */
function previousPeriod(period) {
  var p = String(period).split('-');
  var y = Number(p[0]);
  return Number(p[1]) === 10 ? y + '-04' : (y - 1) + '-10';
}

/**
 * 3項目を判定する。
 * rec: { systolic, diastolic, weight, height }
 * prevWeight: 前回判定期の体重（なければ null）
 */
function judge(rec, prevWeight) {
  var bmi = calcBmi(rec.weight, rec.height);

  var bpOk = rec.systolic < CRITERIA.bpSystolicBelow && rec.diastolic < CRITERIA.bpDiastolicBelow;
  var bpRefer = rec.systolic >= CRITERIA.bpReferSystolic || rec.diastolic >= CRITERIA.bpReferDiastolic;

  var bmiOk = bmi !== null && bmi >= CRITERIA.bmiMin && bmi < CRITERIA.bmiBelow;

  var lossPct = null;
  var weightOk = null; // null = 前回の記録がなく判定なし
  if (prevWeight) {
    lossPct = round1_((prevWeight - rec.weight) / prevWeight * 100);
    weightOk = lossPct >= CRITERIA.weightLossPct;
  }

  var clearCount = (bpOk ? 1 : 0) + (bmiOk ? 1 : 0) + (weightOk ? 1 : 0);
  return {
    bmi: bmi,
    bpOk: bpOk,
    bpRefer: bpRefer,
    bmiOk: bmiOk,
    lossPct: lossPct,
    weightOk: weightOk,
    clearCount: clearCount,
    eligible: clearCount >= 1
  };
}

/**
 * 1人分の記録一覧から判定期ごとの代表記録（その期で測定日が最も新しいもの）を返す。
 * records は { date, period, createdAt, ... } の配列。
 */
function latestByPeriod(records) {
  var map = {};
  records.forEach(function (r) {
    if (!r.period) return;
    var cur = map[r.period];
    if (!cur || r.date > cur.date || (r.date === cur.date && r.createdAt > cur.createdAt)) {
      map[r.period] = r;
    }
  });
  return map;
}

/** 指定の判定期について、1人分の判定結果を作る。その期の記録がなければ null */
function judgeForPeriod(records, period) {
  var byPeriod = latestByPeriod(records);
  var rec = byPeriod[period];
  if (!rec) return null;
  var prev = byPeriod[previousPeriod(period)];
  var result = judge(rec, prev ? prev.weight : null);
  result.record = rec;
  result.prevWeight = prev ? prev.weight : null;
  return result;
}

/* ------------------------------------------------------------------ */
/* シートの読み書き                                                     */
/* ------------------------------------------------------------------ */

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function fmtDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return str_(v);
}

/** 判定期のセルを yyyy-MM の文字列にする（スプレッドシートが日付に変換していても読めるように） */
function fmtPeriod_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return str_(v);
}

function readEmployees_() {
  var values = sheet_(SHEET_EMPLOYEES).getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!str_(r[0])) continue;
    list.push({
      row: i + 1,
      id: str_(r[0]),
      name: str_(r[1]),
      height: Number(r[2]) || null,
      pin: str_(r[3]),
      admin: str_(r[4]) !== '',
      active: str_(r[5]) !== '退職'
    });
  }
  return list;
}

function findEmployee_(id) {
  var list = readEmployees_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === str_(id)) return list[i];
  }
  return null;
}

function readRecords_(employeeId) {
  var values = sheet_(SHEET_RECORDS).getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var id = str_(r[1]);
    if (!id || (employeeId && id !== employeeId)) continue;
    var date = fmtDate_(r[3]);
    list.push({
      createdAt: r[0] instanceof Date ? r[0].getTime() : 0,
      employeeId: id,
      name: str_(r[2]),
      date: date,
      period: fmtPeriod_(r[4]) || periodOf(date),
      systolic: Number(r[5]),
      diastolic: Number(r[6]),
      weight: Number(r[7]),
      height: Number(r[8]),
      bmi: Number(r[9]) || null,
      memo: str_(r[10])
    });
  }
  return list;
}

/* ------------------------------------------------------------------ */
/* 暗証番号とセッション                                                 */
/* ------------------------------------------------------------------ */

function hashPin_(id, pin) {
  var salt = PropertiesService.getScriptProperties().getProperty('PIN_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('PIN_SALT', salt);
  }
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + id + ':' + pin);
  return 'h:' + Utilities.base64Encode(bytes);
}

/** 従業員シートに平文で入力された暗証番号を暗号化（ハッシュ化）する */
function hashAllPins() {
  var sh = sheet_(SHEET_EMPLOYEES);
  readEmployees_().forEach(function (e) {
    if (e.pin && e.pin.indexOf('h:') !== 0) {
      sh.getRange(e.row, 4).setValue(hashPin_(e.id, e.pin));
    }
  });
}

function pinMatches_(emp, pin) {
  if (!emp.pin) return false;
  if (emp.pin.indexOf('h:') === 0) return emp.pin === hashPin_(emp.id, pin);
  // 経理担当者が平文で登録した直後：照合できたらその場で暗号化しておく
  if (emp.pin === pin) {
    sheet_(SHEET_EMPLOYEES).getRange(emp.row, 4).setValue(hashPin_(emp.id, pin));
    return true;
  }
  return false;
}

function newSession_(emp) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, emp.id, SESSION_HOURS * 3600);
  return token;
}

/** トークンから従業員を取り出す。無効ならエラー */
function requireUser_(token) {
  var id = token ? CacheService.getScriptCache().get('s:' + token) : null;
  var emp = id ? findEmployee_(id) : null;
  if (!emp || !emp.active) throw new Error('ログインの有効期限が切れました。もう一度ログインしてください。');
  return emp;
}

function requireAdmin_(token) {
  var emp = requireUser_(token);
  if (!emp.admin) throw new Error('管理者のみ利用できます。');
  return emp;
}

/* ------------------------------------------------------------------ */
/* 画面から呼ばれる API（google.script.run）                            */
/* ------------------------------------------------------------------ */

function apiLogin(id, pin) {
  id = str_(id);
  pin = str_(pin);
  var cache = CacheService.getScriptCache();
  var failKey = 'f:' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_LOGIN_FAILS) {
    throw new Error('暗証番号を続けて間違えたため、30分間ログインできません。');
  }
  var emp = findEmployee_(id);
  if (!emp || !emp.active || !pinMatches_(emp, pin)) {
    cache.put(failKey, String(fails + 1), 1800);
    throw new Error('社員番号または暗証番号が違います。');
  }
  cache.remove(failKey);
  return { token: newSession_(emp), me: profile_(emp) };
}

function apiLogout(token) {
  if (token) CacheService.getScriptCache().remove('s:' + token);
  return true;
}

function profile_(emp) {
  return { id: emp.id, name: emp.name, height: emp.height, admin: emp.admin };
}

/** 本人の記録・最新の判定結果・基準をまとめて返す */
function apiMyData(token) {
  var emp = requireUser_(token);
  var records = readRecords_(emp.id);
  records.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.createdAt - b.createdAt; });

  var period = currentPeriod();
  var shown = period;
  var result = judgeForPeriod(records, period);
  var submitted = !!result;
  if (!result) {
    // 今期が未提出なら前の判定期の結果を表示する
    shown = previousPeriod(period);
    result = judgeForPeriod(records, shown);
  }
  if (result) result.period = shown;

  return {
    me: profile_(emp),
    records: records.map(stripRecord_),
    currentPeriod: period,
    currentPeriodLabel: periodLabel(period),
    currentSubmitted: submitted,
    result: result ? publicResult_(result) : null,
    criteria: CRITERIA,
    judgeMonths: JUDGE_MONTHS
  };
}

function stripRecord_(r) {
  return {
    date: r.date, period: r.period, systolic: r.systolic, diastolic: r.diastolic,
    weight: r.weight, height: r.height, bmi: r.bmi, memo: r.memo
  };
}

function publicResult_(res) {
  return {
    period: res.period,
    periodLabel: periodLabel(res.period),
    record: stripRecord_(res.record),
    prevWeight: res.prevWeight,
    bmi: res.bmi,
    bpOk: res.bpOk,
    bpRefer: res.bpRefer,
    bmiOk: res.bmiOk,
    lossPct: res.lossPct,
    weightOk: res.weightOk,
    clearCount: res.clearCount,
    eligible: res.eligible
  };
}

function validNumber_(v, min, max, label) {
  var n = Number(v);
  if (!isFinite(n) || n < min || n > max) {
    throw new Error(label + 'の値を確認してください（' + min + '〜' + max + '）。');
  }
  return n;
}

/** 記録を1件追加する */
function apiSaveRecord(token, input) {
  var emp = requireUser_(token);
  var date = str_(input.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('測定日を入力してください。');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (date > today) throw new Error('未来の日付は登録できません。');

  var systolic = validNumber_(input.systolic, 60, 260, '収縮期血圧（上）');
  var diastolic = validNumber_(input.diastolic, 30, 180, '拡張期血圧（下）');
  if (diastolic >= systolic) throw new Error('血圧の上と下が逆になっていないか確認してください。');
  var weight = round1_(validNumber_(input.weight, 20, 250, '体重'));
  var height = round1_(validNumber_(input.height || emp.height, 100, 230, '身長'));
  var bmi = calcBmi(weight, height);
  var period = periodOf(date);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_(SHEET_RECORDS).appendRow([
      new Date(), emp.id, emp.name, date, period, systolic, diastolic, weight, height, bmi, str_(input.memo).slice(0, 200)
    ]);
    // 身長が変わっていれば従業員シートも更新する
    if (height !== emp.height) {
      sheet_(SHEET_EMPLOYEES).getRange(emp.row, 3).setValue(height);
    }
  } finally {
    lock.releaseLock();
  }
  return apiMyData(token);
}

/** 暗証番号の変更 */
function apiChangePin(token, oldPin, newPin) {
  var emp = requireUser_(token);
  if (!pinMatches_(emp, str_(oldPin))) throw new Error('今の暗証番号が違います。');
  if (!/^\d{4,8}$/.test(str_(newPin))) throw new Error('新しい暗証番号は4〜8桁の数字にしてください。');
  sheet_(SHEET_EMPLOYEES).getRange(emp.row, 4).setValue(hashPin_(emp.id, str_(newPin)));
  return true;
}

/* ------------------------------------------------------------------ */
/* 管理者（経理担当者）向け                                             */
/* ------------------------------------------------------------------ */

/** 判定期ごとの全員の一覧 */
function buildPeriodSummary_(period) {
  var employees = readEmployees_().filter(function (e) { return e.active; });
  var all = readRecords_();
  var byEmp = {};
  all.forEach(function (r) { (byEmp[r.employeeId] = byEmp[r.employeeId] || []).push(r); });

  var rows = employees.map(function (e) {
    var res = judgeForPeriod(byEmp[e.id] || [], period);
    if (!res) return { id: e.id, name: e.name, submitted: false };
    res.period = period;
    var out = publicResult_(res);
    out.id = e.id;
    out.name = e.name;
    out.submitted = true;
    return out;
  });

  var submitted = rows.filter(function (r) { return r.submitted; });
  return {
    period: period,
    periodLabel: periodLabel(period),
    rows: rows,
    stats: {
      total: rows.length,
      submitted: submitted.length,
      eligible: submitted.filter(function (r) { return r.eligible; }).length,
      bpOk: submitted.filter(function (r) { return r.bpOk; }).length,
      bmiOk: submitted.filter(function (r) { return r.bmiOk; }).length,
      weightOk: submitted.filter(function (r) { return r.weightOk; }).length,
      bpRefer: submitted.filter(function (r) { return r.bpRefer; }).length
    }
  };
}

/** 記録がある判定期の一覧（新しい順）。今期は記録がなくても含める */
function listPeriods_() {
  var set = {};
  set[currentPeriod()] = true;
  readRecords_().forEach(function (r) { if (r.period) set[r.period] = true; });
  return Object.keys(set).sort().reverse();
}

function apiAdminSummary(token, period) {
  requireAdmin_(token);
  var periods = listPeriods_();
  return { periods: periods, summary: buildPeriodSummary_(period || periods[0]), criteria: CRITERIA };
}

/** 1人分の全記録（管理者が推移を確認する用） */
function apiAdminEmployeeRecords(token, employeeId) {
  requireAdmin_(token);
  var emp = findEmployee_(employeeId);
  if (!emp) throw new Error('社員が見つかりません。');
  var records = readRecords_(emp.id).map(stripRecord_);
  records.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return { employee: profile_(emp), records: records };
}

/** 判定結果シートへ書き出す（同じ判定期の行は置き換える） */
function exportPeriod_(period) {
  var summary = buildPeriodSummary_(period);
  var sh = sheet_(SHEET_RESULTS);
  var values = sh.getDataRange().getValues();
  // 既存の同じ判定期の行を下から削除
  for (var i = values.length - 1; i >= 1; i--) {
    if (fmtPeriod_(values[i][0]) === period) sh.deleteRow(i + 1);
  }
  var mark = function (v) { return v === null ? '判定なし' : v ? '○' : '×'; };
  var rows = summary.rows.map(function (r) {
    if (!r.submitted) {
      return [period, r.id, r.name, '未提出', '', '', '', '', '', '', '', '', 0, '×'];
    }
    return [
      period, r.id, r.name, r.record.date,
      r.record.systolic + '/' + r.record.diastolic, mark(r.bpOk),
      r.bmi, mark(r.bmiOk),
      r.record.weight, r.prevWeight === null ? '' : r.prevWeight,
      r.lossPct === null ? '' : r.lossPct, mark(r.weightOk),
      r.clearCount, r.eligible ? '○' : '×'
    ];
  });
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, RESULT_HEADERS.length).setValues(rows);
  }
  return summary.stats;
}

function apiAdminExport(token, period) {
  requireAdmin_(token);
  return exportPeriod_(period);
}

function menuExportCurrentPeriod() {
  var period = currentPeriod();
  var stats = exportPeriod_(period);
  SpreadsheetApp.getUi().alert(
    periodLabel(period) + 'の判定結果を「' + SHEET_RESULTS + '」シートに書き出しました。\n' +
    '提出 ' + stats.submitted + '/' + stats.total + '人、手当対象 ' + stats.eligible + '人'
  );
}
