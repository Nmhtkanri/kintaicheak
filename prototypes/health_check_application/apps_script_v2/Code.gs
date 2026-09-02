/**
 * 健康診断申込 v2（Google Apps Script・年度別スプレッドシートに紐づけて使う）
 *
 * シート構成（設定 / 選択肢 / 対象者 / 回答 / 監査ログ）と列並びは Operation Hub の
 * services/health_apply/schema.py と同じ。SCHEMA_VERSION が設定シートと違えば止まる。
 *
 * 役割分担:
 *   - Hub  … 対象者シートの 1〜14列（社員番号・氏名・メール・前年度情報）を追記する
 *   - GAS  … 案内メール送信（トークン生成・15列以降）、個別URLの表示、回答の追記
 * 生トークンはメールのURLにだけ載せ、シートには SHA-256 のハッシュだけを置く。
 * 回答は上書きせず追記（回答版を増やす）。訂正はメニュー「再回答を許可」で受け付ける。
 */

const SCHEMA_VERSION = '2027.1';

const SHEETS = Object.freeze({
  settings: '設定', options: '選択肢', targets: '対象者', responses: '回答', audit: '監査ログ',
});

const SETTINGS_HEADERS = Object.freeze(['キー', '値', '備考']);
const OPTION_HEADERS = Object.freeze(['区分', 'コード', '表示名', '有効', '並び順', '別名', '備考']);
const TARGET_HEADERS = Object.freeze([
  '年度', '社員番号', '氏名', '社用メール', '在籍区分',
  '前年度情報元', '前年度健診機関コード', '前年度健診機関名',
  '前年度健診種別コード', '前年度健診種別名', '前年度追加検査', '前年度健診機関(原文)',
  '登録日時', '登録者',
  'トークンハッシュ', '送信日時', '送信回数', '初回アクセス日時',
  '申込状態', '受付番号', '回答版', '回答日時', '備考',
]);
const RESPONSE_HEADERS = Object.freeze([
  '回答日時', '受付番号', '年度', '社員番号', '回答版', '氏名', '社用メール',
  '申込区分', '健診機関コード', '健診機関名', 'その他医療機関名',
  '健診種別コード', '健診種別名', '追加検査', 'その他健診予定日',
  '被扶養者申込', '続柄', '被扶養者氏名', '備考', 'トークンハッシュ', '回答元',
]);
const AUDIT_HEADERS = Object.freeze(['日時', 'イベント', '実行元', '実行者', '年度', '社員番号', '詳細']);

const REQUIRED_SETTING_KEYS = Object.freeze([
  'スキーマ版', '年度', '前年度', '受付開始', '受付終了', '受診期間開始', '受診期間終了', '回答受付',
]);

const KIND = Object.freeze({ institution: '機関', examType: '種別', extra: '追加検査', relationship: '続柄' });
const OTHER_INSTITUTION_CODE = 'OTHER';
const SOURCE_NONE = 'なし';
const STATUS = Object.freeze({ unsent: '未送信', sent: '送信済', answered: '回答済', reanswer: '再回答待ち', invalid: '無効' });
const ACTOR = 'AppsScript';

/** 列名 → 1始まりの列番号 */
const TARGET_COL = Object.freeze(TARGET_HEADERS.reduce((m, h, i) => { m[h] = i + 1; return m; }, {}));

// ---------------------------------------------------------------------------
// 共通
// ---------------------------------------------------------------------------

function cleanText_(value, maxLength) {
  const s = value === null || value === undefined ? '' : String(value).trim();
  return maxLength ? s.slice(0, maxLength) : s;
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
}

function formatDate_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function hashToken_(token) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  return digest.map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function newToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function receiptId_(fiscalYear, employeeId, version) {
  return `HC-${fiscalYear}-${employeeId}-${pad2_(version)}`;
}

function spreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sheet = spreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error(`「${name}」シートがありません。メニューの「シートを初期化」を実行してください。`);
  return sheet;
}

function readRows_(name, headers) {
  const sheet = sheet_(name);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, headers.length).getDisplayValues();
  return values
    .map((row, i) => {
      const obj = { rowNumber: i + 2 };
      headers.forEach((h, j) => { obj[h] = cleanText_(row[j]); });
      return obj;
    })
    .filter((obj) => headers.some((h) => obj[h] !== ''));
}

function verifyHeaders_(name, headers) {
  const sheet = sheet_(name);
  const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map((v) => cleanText_(v));
  headers.forEach((h, i) => {
    if (actual[i] !== h) {
      throw new Error(`「${name}」シートの${i + 1}列目が想定外です: 「${actual[i]}」（期待: 「${h}」）`);
    }
  });
}

// ---------------------------------------------------------------------------
// 設定・選択肢
// ---------------------------------------------------------------------------

function readSettings_() {
  const kv = {};
  readRows_(SHEETS.settings, SETTINGS_HEADERS).forEach((row) => {
    if (row['キー']) kv[row['キー']] = row['値'];
  });
  return kv;
}

function requireSettings_(kv) {
  const missing = REQUIRED_SETTING_KEYS.filter((k) => !kv[k]);
  if (missing.length) throw new Error(`設定シートに値がありません: ${missing.join('、')}`);
  if (kv['スキーマ版'] !== SCHEMA_VERSION) {
    throw new Error(`設定シートのスキーマ版が違います: 「${kv['スキーマ版']}」（このスクリプトは ${SCHEMA_VERSION}）`);
  }
  return kv;
}

function parseActive_(value) {
  const s = cleanText_(value).toLowerCase();
  if (!s) return true;
  return !['0', 'false', 'no', '無効', '×', 'x'].includes(s);
}

function readOptions_() {
  const byKind = {};
  Object.keys(KIND).forEach((k) => { byKind[KIND[k]] = []; });
  readRows_(SHEETS.options, OPTION_HEADERS).forEach((row) => {
    const kind = row['区分'];
    if (!byKind[kind]) return;
    byKind[kind].push({
      code: row['コード'], name: row['表示名'], active: parseActive_(row['有効']),
      order: row['並び順'] === '' ? 9999 : Number(row['並び順']),
    });
  });
  Object.keys(byKind).forEach((kind) => {
    byKind[kind].sort((a, b) => (a.order - b.order) || a.name.localeCompare(b.name, 'ja'));
  });
  return byKind;
}

function activeOptions_(options, kind) {
  return (options[kind] || []).filter((o) => o.active);
}

function optionByCode_(options, kind, code, activeOnly) {
  return (options[kind] || []).find((o) => o.code === cleanText_(code) && (!activeOnly || o.active)) || null;
}

/** 受付期間内かどうか。回答受付=0 なら常に false。日付は yyyy-MM-dd で比較する。 */
function isAccepting_(kv, now) {
  if (cleanText_(kv['回答受付']) !== '1') return false;
  const today = formatDate_(now || new Date());
  const from = cleanText_(kv['受付開始']);
  const to = cleanText_(kv['受付終了']);
  if (from && today < from) return false;
  if (to && today > to) return false;
  return true;
}

// ---------------------------------------------------------------------------
// 対象者
// ---------------------------------------------------------------------------

function findTargetByHash_(hash) {
  if (!hash) return null;
  return readRows_(SHEETS.targets, TARGET_HEADERS).find((row) => row['トークンハッシュ'] === hash) || null;
}

function updateTarget_(rowNumber, values) {
  const sheet = sheet_(SHEETS.targets);
  Object.keys(values).forEach((header) => {
    const col = TARGET_COL[header];
    if (!col) throw new Error(`対象者シートに列がありません: ${header}`);
    sheet.getRange(rowNumber, col).setValue(values[header]);
  });
}

function previousOf_(target) {
  const source = target['前年度情報元'] || SOURCE_NONE;
  return {
    hasPrevious: source !== SOURCE_NONE && !!target['前年度健診機関コード'] && !!target['前年度健診種別コード'],
    source,
    institutionCode: target['前年度健診機関コード'],
    institutionName: target['前年度健診機関名'],
    examTypeCode: target['前年度健診種別コード'],
    examTypeName: target['前年度健診種別名'],
    extraCodes: (target['前年度追加検査'] || '').split(/[;；、,]/).map((s) => s.trim()).filter(Boolean),
    raw: target['前年度健診機関(原文)'],
  };
}

function recordFirstAccess_(target) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cell = sheet_(SHEETS.targets).getRange(target.rowNumber, TARGET_COL['初回アクセス日時']);
    if (cleanText_(cell.getValue())) return;
    cell.setValue(formatDateTime_(new Date()));
    appendAudit_('FIRST_ACCESS', target['社員番号'], '有効な個別URLを初回表示（メール検査でも付くので本人閲覧の証拠にはしない）', target['年度']);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Web アプリ
// ---------------------------------------------------------------------------

function doGet(e) {
  const token = cleanText_(e && e.parameter ? e.parameter.token : '', 200);
  let kv;
  try {
    kv = requireSettings_(readSettings_());
  } catch (error) {
    return simplePage_('現在は申込を受け付けていません', `設定が整っていません。健康診断担当者へご連絡ください。<br><small>${escapeHtml_(error.message)}</small>`);
  }
  const target = findTargetByHash_(token ? hashToken_(token) : '');
  if (!target) {
    return simplePage_('申込URLが無効です', '健康診断担当者へご連絡ください。<br>Invalid application link. Please contact the administrator.');
  }
  if (!isAccepting_(kv, new Date())) {
    return simplePage_('受付期間外です', `受付期間は ${escapeHtml_(kv['受付開始'])} 〜 ${escapeHtml_(kv['受付終了'])} です。健康診断担当者へご連絡ください。`);
  }
  try {
    recordFirstAccess_(target);
  } catch (error) {
    console.error(`初回アクセス日時を記録できませんでした: ${error}`);
  }
  const options = readOptions_();
  const template = HtmlService.createTemplateFromFile('Index');
  template.token = token;
  template.settings = kv;
  template.employee = {
    employeeId: target['社員番号'], name: target['氏名'], email: target['社用メール'],
    status: target['申込状態'] || STATUS.unsent, receiptId: target['受付番号'],
  };
  template.previous = previousOf_(target);
  template.options = {
    institutions: activeOptions_(options, KIND.institution),
    examTypes: activeOptions_(options, KIND.examType),
    extras: activeOptions_(options, KIND.extra),
    relationships: activeOptions_(options, KIND.relationship),
  };
  template.canAnswer = [STATUS.sent, STATUS.reanswer].includes(template.employee.status);
  return template.evaluate()
    .setTitle(`${kv['年度']}年度 健康診断申込`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function simplePage_(title, bodyHtml) {
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + `<body style="font-family:sans-serif;padding:40px;line-height:1.7"><h1>${escapeHtml_(title)}</h1><p>${bodyHtml}</p></body></html>`
  ).setTitle('健康診断申込');
}

function escapeHtml_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * 回答の受付。画面からは選択肢の**コード**だけを受け取り、表示名はサーバ側で選択肢シートから引く。
 * 本人特定はトークン（→ハッシュ→対象者行）だけで行い、画面から来た社員番号などは信用しない。
 */
function submitApplication(token, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const kv = requireSettings_(readSettings_());
    if (!isAccepting_(kv, new Date())) throw new Error('受付期間外です。健康診断担当者へご連絡ください。');
    const target = findTargetByHash_(cleanText_(token, 200) ? hashToken_(cleanText_(token, 200)) : '');
    if (!target) throw new Error('申込URLが無効です。 / Invalid application link.');
    const status = target['申込状態'] || STATUS.unsent;
    if (status === STATUS.answered) {
      throw new Error(`このURLでは既に申込済みです。受付番号: ${target['受付番号']}。訂正したい場合は健康診断担当者へご連絡ください。`);
    }
    if (![STATUS.sent, STATUS.reanswer].includes(status)) {
      throw new Error('このURLは現在使えません。健康診断担当者へご連絡ください。');
    }
    const options = readOptions_();
    const normalized = validatePayload_(target, payload || {}, options, kv);
    const now = new Date();
    const version = nextVersion_(target);
    const receiptId = receiptId_(kv['年度'], target['社員番号'], version);
    const row = buildResponseRow_(kv, target, normalized, version, receiptId, formatDateTime_(now));
    sheet_(SHEETS.responses).appendRow(row);
    updateTarget_(target.rowNumber, {
      '申込状態': STATUS.answered, '受付番号': receiptId, '回答版': String(version), '回答日時': formatDateTime_(now),
    });
    appendAudit_('SUBMIT', target['社員番号'], `${receiptId} 回答版=${version} 区分=${normalized.applicationType}`, kv['年度']);
    return { ok: true, receiptId, employeeId: target['社員番号'], version };
  } finally {
    lock.releaseLock();
  }
}

function nextVersion_(target) {
  const current = parseInt(target['回答版'], 10);
  return (isNaN(current) ? 0 : current) + 1;
}

function validatePayload_(target, payload, options, kv) {
  const applicationType = cleanText_(payload.applicationType, 20);
  if (!['same', 'change'].includes(applicationType)) {
    throw new Error('申込内容を選択してください。 / Select an application option.');
  }
  if (payload.agreement !== true) {
    throw new Error('確認欄にチェックしてください。 / Please confirm the details.');
  }
  const previous = previousOf_(target);
  let institution;
  let examType;
  let extras = [];
  let otherInstitution = '';
  let otherPlannedDate = '';

  if (applicationType === 'same') {
    if (!previous.hasPrevious) {
      throw new Error('前年度の情報が無いため「前年度と同じ」は選べません。「変更する」を選択してください。');
    }
    institution = optionByCode_(options, KIND.institution, previous.institutionCode, false)
      || { code: previous.institutionCode, name: previous.institutionName };
    examType = optionByCode_(options, KIND.examType, previous.examTypeCode, false)
      || { code: previous.examTypeCode, name: previous.examTypeName };
    extras = previous.extraCodes.map((code) => optionByCode_(options, KIND.extra, code, false) || { code, name: code });
  } else {
    institution = optionByCode_(options, KIND.institution, payload.clinicCode, true);
    if (!institution) throw new Error('健診機関を選択してください。 / Select a clinic.');
    if (institution.code === OTHER_INSTITUTION_CODE) {
      otherInstitution = cleanText_(payload.customClinic, 100);
      otherPlannedDate = cleanText_(payload.otherPlannedDate, 10);
      if (otherPlannedDate) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(otherPlannedDate)) throw new Error('健診予定日の形式が正しくありません。');
        if (otherPlannedDate < kv['受診期間開始'] || otherPlannedDate > kv['受診期間終了']) {
          throw new Error(`健診予定日は受診期間（${kv['受診期間開始']}〜${kv['受診期間終了']}）内で入力してください。`);
        }
      }
    }
    examType = optionByCode_(options, KIND.examType, payload.courseCode, true);
    if (!examType) throw new Error('健診種別を選択してください。 / Select a course.');
    const requested = Array.isArray(payload.extraCodes) ? payload.extraCodes : [];
    extras = requested.map((code) => {
      const opt = optionByCode_(options, KIND.extra, code, true);
      if (!opt) throw new Error('選択できない追加検査が含まれています。');
      return opt;
    });
  }

  const dependentRequested = payload.dependentRequested === true;
  let relationship = '';
  let dependentName = '';
  if (dependentRequested) {
    const rel = optionByCode_(options, KIND.relationship, payload.dependentRelationship, true);
    dependentName = cleanText_(payload.dependentName, 100);
    if (!rel || !dependentName) throw new Error('被扶養者の続柄と氏名を入力してください。');
    relationship = rel.code;
  }
  return {
    applicationType, institution, otherInstitution, otherPlannedDate, examType, extras,
    dependentRequested, relationship, dependentName, remarks: cleanText_(payload.remarks, 500),
  };
}

function buildResponseRow_(kv, target, n, version, receiptId, answeredAt) {
  const values = {
    '回答日時': answeredAt,
    '受付番号': receiptId,
    '年度': kv['年度'],
    '社員番号': target['社員番号'],
    '回答版': String(version),
    '氏名': target['氏名'],
    '社用メール': target['社用メール'],
    '申込区分': n.applicationType,
    '健診機関コード': n.institution.code,
    '健診機関名': n.institution.name,
    'その他医療機関名': n.otherInstitution,
    '健診種別コード': n.examType.code,
    '健診種別名': n.examType.name,
    '追加検査': n.extras.map((o) => o.code).join(';'),
    'その他健診予定日': n.otherPlannedDate,
    '被扶養者申込': n.dependentRequested ? '1' : '0',
    '続柄': n.relationship,
    '被扶養者氏名': n.dependentName,
    '備考': n.remarks,
    'トークンハッシュ': target['トークンハッシュ'],
    '回答元': 'Web',
  };
  return RESPONSE_HEADERS.map((h) => values[h]);
}

// ---------------------------------------------------------------------------
// 管理メニュー（スプレッドシートを開いた人だけが使う）
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi().createMenu('健康診断申込')
    .addItem('案内メールを送る（未送信のみ）', 'sendInvitations')
    .addItem('選択した行を再送（URLを作り直す）', 'resendSelected')
    .addItem('選択した行の再回答を許可', 'reopenSelected')
    .addSeparator()
    .addItem('シートを初期化（ヘッダー・書式）', 'setupWorkbook')
    .addToUi();
}

function setupWorkbook() {
  const ss = spreadsheet_();
  const specs = [
    [SHEETS.settings, SETTINGS_HEADERS], [SHEETS.options, OPTION_HEADERS], [SHEETS.targets, TARGET_HEADERS],
    [SHEETS.responses, RESPONSE_HEADERS], [SHEETS.audit, AUDIT_HEADERS],
  ];
  specs.forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      verifyHeaders_(name, headers);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#116548').setFontColor('#ffffff');
  });
  // コード列は書式なしテキスト（前ゼロ・英字入りの HPM コードを数値にさせない）
  const opt = ss.getSheetByName(SHEETS.options);
  opt.getRange(1, 2, opt.getMaxRows(), 1).setNumberFormat('@');
  const tgt = ss.getSheetByName(SHEETS.targets);
  [TARGET_COL['社員番号'], TARGET_COL['前年度健診機関コード'], TARGET_COL['前年度健診種別コード'], TARGET_COL['回答版']]
    .forEach((col) => tgt.getRange(1, col, tgt.getMaxRows(), 1).setNumberFormat('@'));
  const res = ss.getSheetByName(SHEETS.responses);
  [4, 5, 9, 12].forEach((col) => res.getRange(1, col, res.getMaxRows(), 1).setNumberFormat('@'));
  const settings = ss.getSheetByName(SHEETS.settings);
  if (settings.getLastRow() < 2) {
    const rows = [['スキーマ版', SCHEMA_VERSION, '']].concat(
      REQUIRED_SETTING_KEYS.filter((k) => k !== 'スキーマ版').map((k) => [k, '', '']),
      [['WebアプリURL', '', 'デプロイ後の /exec のURL'], ['案内メール件名', '', ''],
       ['案内メール本文', '', '{氏名} {URL} {受付終了} を置換'], ['担当者連絡先', '', '']]
    );
    settings.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  appendAudit_('SETUP', '', `シートを初期化（スキーマ版 ${SCHEMA_VERSION}）`, readSettings_()['年度'] || '');
}

/** 送る対象: 送信日時が空で、申込状態が未送信（または空）の行。 */
function targetsToSend_(targets) {
  return targets.filter((t) => !t['送信日時'] && (!t['申込状態'] || t['申込状態'] === STATUS.unsent) && !!t['社用メール']);
}

function buildInvitationEmail_(kv, target, url) {
  const replace = (s) => String(s || '')
    .replace(/\{氏名\}/g, target['氏名'])
    .replace(/\{URL\}/g, url)
    .replace(/\{受付終了\}/g, kv['受付終了'] || '')
    .replace(/\{受付開始\}/g, kv['受付開始'] || '')
    .replace(/\{年度\}/g, kv['年度'] || '');
  return {
    to: target['社用メール'],
    subject: replace(kv['案内メール件名'] || `【健康診断】${kv['年度']}年度の申込のお願い`),
    body: replace(kv['案内メール本文'] || `{氏名} さん\n\n下記の個別URLから健康診断の申込をお願いします。\n{URL}\n\n受付期限: {受付終了}`),
  };
}

function sendInvitations() {
  const ui = SpreadsheetApp.getUi();
  const kv = requireSettings_(readSettings_());
  if (!kv['WebアプリURL']) throw new Error('設定シートの WebアプリURL が空です。');
  const rows = targetsToSend_(readRows_(SHEETS.targets, TARGET_HEADERS));
  if (!rows.length) { ui.alert('未送信の対象者はいません。'); return; }
  const phrase = `SEND ${kv['年度']} ${rows.length}`;
  const answer = ui.prompt('案内メールの送信', `${rows.length}名に個別URLを送ります。確認のため「${phrase}」と入力してください。`, ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK || cleanText_(answer.getResponseText()) !== phrase) {
    ui.alert('送信を中止しました（確認語が一致しません）。');
    return;
  }
  const result = sendTo_(kv, rows, 'SEND');
  ui.alert(`送信 ${result.sent}件 / 失敗 ${result.failed}件。失敗行は送信日時が空のままなので、もう一度実行すると再送されます。`);
}

function resendSelected() {
  const ui = SpreadsheetApp.getUi();
  const kv = requireSettings_(readSettings_());
  const rows = selectedTargets_();
  if (!rows.length) { ui.alert('対象者シートで再送したい行を選択してから実行してください。'); return; }
  const answer = ui.prompt('再送', `${rows.length}名の個別URLを作り直して再送します（古いURLは無効になります）。「RESEND ${rows.length}」と入力してください。`, ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK || cleanText_(answer.getResponseText()) !== `RESEND ${rows.length}`) {
    ui.alert('再送を中止しました。');
    return;
  }
  const result = sendTo_(kv, rows, 'RESEND');
  ui.alert(`再送 ${result.sent}件 / 失敗 ${result.failed}件`);
}

function selectedTargets_() {
  const sheet = sheet_(SHEETS.targets);
  const range = spreadsheet_().getActiveRange();
  if (!range || range.getSheet().getName() !== SHEETS.targets) return [];
  const first = Math.max(range.getRow(), 2);
  const last = range.getLastRow();
  return readRows_(SHEETS.targets, TARGET_HEADERS).filter((t) => t.rowNumber >= first && t.rowNumber <= last);
}

function sendTo_(kv, rows, eventName) {
  let sent = 0;
  let failed = 0;
  rows.forEach((target) => {
    const token = newToken_();
    const url = `${kv['WebアプリURL']}?token=${encodeURIComponent(token)}`;
    const mail = buildInvitationEmail_(kv, target, url);
    try {
      MailApp.sendEmail({ to: mail.to, subject: mail.subject, body: mail.body, name: '健康診断担当' });
      const count = parseInt(target['送信回数'], 10);
      updateTarget_(target.rowNumber, {
        'トークンハッシュ': hashToken_(token),
        '送信日時': formatDateTime_(new Date()),
        '送信回数': String((isNaN(count) ? 0 : count) + 1),
        '申込状態': target['申込状態'] === STATUS.reanswer ? STATUS.reanswer : STATUS.sent,
      });
      appendAudit_(eventName, target['社員番号'], `案内メール送信 → ${mail.to}`, kv['年度']);
      sent += 1;
    } catch (error) {
      appendAudit_(`${eventName}_FAILED`, target['社員番号'], String(error), kv['年度']);
      failed += 1;
    }
  });
  return { sent, failed };
}

function reopenSelected() {
  const ui = SpreadsheetApp.getUi();
  const kv = requireSettings_(readSettings_());
  const rows = selectedTargets_().filter((t) => t['申込状態'] === STATUS.answered);
  if (!rows.length) { ui.alert('「回答済」の行を選択してから実行してください。'); return; }
  rows.forEach((t) => {
    updateTarget_(t.rowNumber, { '申込状態': STATUS.reanswer });
    appendAudit_('REOPEN', t['社員番号'], `再回答を許可（直前の受付番号 ${t['受付番号']}）`, kv['年度']);
  });
  ui.alert(`${rows.length}名を「再回答待ち」にしました。同じURLから再回答できます（回答は追記され、回答版が増えます）。`);
}

function appendAudit_(eventName, employeeId, detail, fiscalYear) {
  const sheet = spreadsheet_().getSheetByName(SHEETS.audit);
  if (!sheet) return;
  let who = '';
  try { who = Session.getActiveUser().getEmail() || ''; } catch (e) { who = ''; }
  sheet.appendRow([formatDateTime_(new Date()), eventName, ACTOR, who || 'web', fiscalYear || '', employeeId || '', detail || '']);
}
