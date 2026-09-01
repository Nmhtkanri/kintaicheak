/**
 * 2027年度 健康診断申込（Google Apps Script 試験版）
 * 個人情報と個別トークンは管理者専用Apps Script側だけで設定する。
 */

const PILOT_SPREADSHEET_ID = '__SET_IN_PRIVATE_SCRIPT__';
const PILOT_TARGETS = [];

const PILOT_CONFIG = Object.freeze({
  fiscalYear: 2027,
  examDateMin: '2027-04-01',
  examDateMax: '2028-03-31',
  targetSheetName: '対象者',
  responseSheetName: '回答',
  auditSheetName: '監査ログ',
  otherClinic: 'その他',
  clinicOptions: Object.freeze([
    'Myメディカルクリニック（大手町）',
    'Myメディカルクリニック（渋谷）',
    'Myメディカルクリニック（新宿）',
    'Myメディカルクリニック（横浜）',
    'Myメディカルクリニック（田町）',
    'Myメディカルクリニック（八重洲）',
    '医療法人社団同友会 春日クリニック',
    'その他',
  ]),
  healthOptions: Object.freeze([
    '基本健診',
    '1日人間ドック',
    '1日人間ドック・胃カメラ',
    '1日人間ドック・バリウム',
    '婦人科検診',
  ]),
  dependentRelationships: Object.freeze(['妻', '夫']),
});

const TARGET_HEADERS = Object.freeze([
  '個別トークン', '社員番号', '氏名', '社用メール', '前年度健診機関',
  '前年度健診内容', '申込状態', '受付番号', '回答日時', '初回アクセス日時',
]);

const TARGET_COLUMNS = Object.freeze({
  status: 7,
  receiptId: 8,
  answeredAt: 9,
  firstAccessAt: 10,
});

const RESPONSE_HEADERS = Object.freeze([
  '回答日時', '受付番号', '年度', '社員番号', '氏名', '社用メール',
  '申込区分', '選択健診機関', '確定健診機関', '健診内容', '健診オプション',
  'その他健診予定日', '被扶養者申込', '続柄', '被扶養者氏名', '備考',
]);

function setupPilot() {
  validatePrivateConfiguration_();
  const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
  const targetSheet = getOrCreateSheet_(spreadsheet, PILOT_CONFIG.targetSheetName);
  const responseSheet = getOrCreateSheet_(spreadsheet, PILOT_CONFIG.responseSheetName);
  const auditSheet = getOrCreateSheet_(spreadsheet, PILOT_CONFIG.auditSheetName);
  ensureHeaders_(targetSheet, TARGET_HEADERS);
  ensureHeaders_(responseSheet, RESPONSE_HEADERS);
  ensureHeaders_(auditSheet, ['日時', 'イベント', '社員番号', '詳細']);

  const existingTokens = new Set();
  if (targetSheet.getLastRow() > 1) {
    targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 1)
      .getDisplayValues().forEach((row) => existingTokens.add(row[0]));
  }
  const rows = PILOT_TARGETS
    .filter((person) => !existingTokens.has(person.token))
    .map((person) => [
      person.token, person.employeeId, person.name, person.email,
      person.previousClinic || '', person.previousCourse || '',
      '未回答', '', '', '',
    ]);
  if (rows.length) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }
  [targetSheet, responseSheet, auditSheet].forEach(formatSheet_);
  appendAudit_('SETUP', '', `対象者 ${PILOT_TARGETS.length}名`);
  return {spreadsheetUrl: spreadsheet.getUrl(), targetCount: PILOT_TARGETS.length};
}

function doGet(e) {
  const token = cleanText_(e && e.parameter ? e.parameter.token : '', 200);
  ensurePilotInitialized_();
  const employee = findEmployeeByToken_(token);
  if (!employee) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<body style="font-family:sans-serif;padding:40px;line-height:1.7"><h1>申込URLが無効です</h1>' +
      '<p>健康診断担当者へご連絡ください。<br>Invalid application link. Please contact the administrator.</p></body></html>'
    ).setTitle('健康診断申込');
  }
  try {
    recordFirstAccess_(employee);
  } catch (error) {
    console.error(`初回アクセス日時を記録できませんでした: ${error}`);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.employee = employee;
  template.token = token;
  template.config = PILOT_CONFIG;
  return template.evaluate()
    .setTitle(`${PILOT_CONFIG.fiscalYear}年度 健康診断申込`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ensurePilotInitialized_() {
  validatePrivateConfiguration_();
  const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
  const targetSheet = spreadsheet.getSheetByName(PILOT_CONFIG.targetSheetName);
  if (!targetSheet || targetSheet.getLastRow() < PILOT_TARGETS.length + 1) {
    setupPilot();
  }
}

function submitApplication(token, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const employee = findEmployeeByToken_(cleanText_(token, 200));
    if (!employee) throw new Error('申込URLが無効です。 / Invalid application link.');
    if (employee.status === '回答済') {
      throw new Error(`このURLでは既に申込済みです。受付番号: ${employee.receiptId}`);
    }
    const normalized = validatePayload_(employee, payload || {});
    const now = new Date();
    const receiptId = createReceiptId_(employee.employeeId, now);
    const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
    const responseSheet = spreadsheet.getSheetByName(PILOT_CONFIG.responseSheetName);
    if (!responseSheet) throw new Error('回答シートが未設定です。管理者へご連絡ください。');
    responseSheet.appendRow([
      now, receiptId, PILOT_CONFIG.fiscalYear, employee.employeeId,
      employee.name, employee.email, normalized.applicationType,
      normalized.clinicChoice, normalized.clinic, normalized.course,
      normalized.healthOptions.join('、'), normalized.otherPlannedDate,
      normalized.dependentRequested ? '希望する' : '希望しない',
      normalized.dependentRelationship, normalized.dependentName, normalized.remarks,
    ]);
    const targetSheet = spreadsheet.getSheetByName(PILOT_CONFIG.targetSheetName);
    targetSheet.getRange(employee.rowNumber, TARGET_COLUMNS.status, 1, 3)
      .setValues([['回答済', receiptId, now]]);
    appendAudit_('SUBMIT', employee.employeeId, receiptId);
    return {ok: true, receiptId, employeeId: employee.employeeId};
  } finally {
    lock.releaseLock();
  }
}

function getPilotLinks() {
  const baseUrl = ScriptApp.getService().getUrl();
  if (!baseUrl) throw new Error('先にWebアプリをデプロイしてください。');
  return PILOT_TARGETS.map((person) => ({
    employeeId: person.employeeId,
    name: person.name,
    email: person.email,
    url: `${baseUrl}?token=${encodeURIComponent(person.token)}`,
  }));
}

function validatePrivateConfiguration_() {
  if (!PILOT_SPREADSHEET_ID || PILOT_SPREADSHEET_ID === '__SET_IN_PRIVATE_SCRIPT__') {
    throw new Error('PILOT_SPREADSHEET_ID が未設定です。');
  }
  if (!Array.isArray(PILOT_TARGETS) || PILOT_TARGETS.length !== 3) {
    throw new Error('試験対象者は3名に固定してください。');
  }
  const ids = new Set();
  const emails = new Set();
  const tokens = new Set();
  PILOT_TARGETS.forEach((person) => {
    if (!/^\d{7}$/.test(String(person.employeeId || ''))) throw new Error('社員番号の形式が不正です。');
    if (!person.name || !/@nmht\.co\.jp$/i.test(String(person.email || ''))) throw new Error('対象者情報が不正です。');
    if (!person.token || String(person.token).length < 32) throw new Error('個別トークンが短すぎます。');
    ids.add(String(person.employeeId));
    emails.add(String(person.email).toLowerCase());
    tokens.add(String(person.token));
  });
  if (ids.size !== 3 || emails.size !== 3 || tokens.size !== 3) {
    throw new Error('社員番号・メール・個別トークンは3名すべて一意にしてください。');
  }
}

function findEmployeeByToken_(token) {
  if (!token) return null;
  validatePrivateConfiguration_();
  const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(PILOT_CONFIG.targetSheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const found = sheet.getRange(2, 1, sheet.getLastRow() - 1, TARGET_HEADERS.length)
    .createTextFinder(token).matchEntireCell(true).findNext();
  if (!found || found.getColumn() !== 1) return null;
  const rowNumber = found.getRow();
  const row = sheet.getRange(rowNumber, 1, 1, TARGET_HEADERS.length).getDisplayValues()[0];
  return {
    rowNumber,
    employeeId: row[1], name: row[2], email: row[3],
    previousClinic: row[4], previousCourse: row[5],
    hasPrevious: Boolean(row[4] && row[5]),
    status: row[TARGET_COLUMNS.status - 1],
    receiptId: row[TARGET_COLUMNS.receiptId - 1],
    answeredAt: row[TARGET_COLUMNS.answeredAt - 1],
    firstAccessAt: row[TARGET_COLUMNS.firstAccessAt - 1],
  };
}

function recordFirstAccess_(employee) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
    const targetSheet = spreadsheet.getSheetByName(PILOT_CONFIG.targetSheetName);
    if (!targetSheet) throw new Error('対象者シートが未設定です。');
    const firstAccessCell = targetSheet.getRange(
      employee.rowNumber,
      TARGET_COLUMNS.firstAccessAt,
    );
    if (firstAccessCell.getValue()) return;
    firstAccessCell.setValue(new Date());
    appendAudit_('FIRST_ACCESS', employee.employeeId, '有効な個別URLを初回表示');
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(employee, payload) {
  const applicationType = cleanText_(payload.applicationType, 20);
  if (!['same', 'change'].includes(applicationType)) {
    throw new Error('申込内容を選択してください。 / Select an application option.');
  }
  if (applicationType === 'same' && !employee.hasPrevious) {
    throw new Error('テスト版では前年度情報が未連携です。「変更する」を選択してください。');
  }
  if (payload.agreement !== true) {
    throw new Error('確認欄にチェックしてください。 / Please confirm the details.');
  }

  let clinicChoice = '';
  let clinic = employee.previousClinic;
  let course = employee.previousCourse;
  let healthOptions = [];
  let otherPlannedDate = '';
  if (applicationType === 'change') {
    clinicChoice = cleanText_(payload.clinicChoice, 100);
    if (!PILOT_CONFIG.clinicOptions.includes(clinicChoice)) {
      throw new Error('健診機関を選択してください。 / Select a clinic.');
    }
    if (clinicChoice === PILOT_CONFIG.otherClinic) {
      clinic = cleanText_(payload.customClinic, 100) || 'その他（医療機関未定）';
      course = 'その他';
      otherPlannedDate = cleanText_(payload.otherPlannedDate, 10);
      if (otherPlannedDate &&
          (otherPlannedDate < PILOT_CONFIG.examDateMin || otherPlannedDate > PILOT_CONFIG.examDateMax)) {
        throw new Error('健診予定日は受診期間内で入力してください。 / Enter a date within the examination period.');
      }
    } else {
      const requested = Array.isArray(payload.healthOptions) ? payload.healthOptions : [];
      const invalid = requested.some((item) => !PILOT_CONFIG.healthOptions.includes(cleanText_(item, 100)));
      if (invalid) throw new Error('選択できない健診オプションが含まれています。');
      healthOptions = PILOT_CONFIG.healthOptions.filter((item) => requested.includes(item));
      if (!healthOptions.length) throw new Error('健診オプションを1つ以上選択してください。');
      clinic = clinicChoice;
      course = healthOptions.join('、');
    }
  }

  const dependentRequested = payload.dependentRequested === true;
  let dependentRelationship = cleanText_(payload.dependentRelationship, 10);
  let dependentName = cleanText_(payload.dependentName, 100);
  if (dependentRequested) {
    if (!PILOT_CONFIG.dependentRelationships.includes(dependentRelationship) || !dependentName) {
      throw new Error('被扶養者の続柄（妻・夫）と氏名を入力してください。');
    }
  } else {
    dependentRelationship = '';
    dependentName = '';
  }
  return {
    applicationType, clinicChoice, clinic, course, healthOptions,
    otherPlannedDate, dependentRequested, dependentRelationship,
    dependentName, remarks: cleanText_(payload.remarks, 500),
  };
}

function createReceiptId_(employeeId, dateValue) {
  const stamp = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  return `HC-${PILOT_CONFIG.fiscalYear}-${employeeId}-${stamp}`;
}

function cleanText_(value, maxLength) {
  return value === null || value === undefined ? '' : String(value).trim().slice(0, maxLength);
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  headers.forEach((header, index) => {
    if (existing[index] && existing[index] !== header) {
      throw new Error(`${sheet.getName()}シートの${index + 1}列目が想定外です: ${existing[index]}`);
    }
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function formatSheet_(sheet) {
  if (sheet.getLastColumn() > 0) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setBackground('#116548').setFontColor('#ffffff').setFontWeight('bold');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  }
}

function appendAudit_(eventName, employeeId, detail) {
  const spreadsheet = SpreadsheetApp.openById(PILOT_SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(PILOT_CONFIG.auditSheetName);
  if (sheet) sheet.appendRow([new Date(), eventName, employeeId, detail]);
}
