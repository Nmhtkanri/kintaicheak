// 健康診断申込 Apps Script v2（Code.gs）を node で動かすための偽 GAS 環境。
// 標準入力に {code_path, workbook:{シート名:[[...],...]}, scenario:"JS", now?:"ISO", mail_fail?:"addr"} を
// 受け取り、scenario を評価した結果と、実行後のシート・送信メール・監査を JSON で標準出力に返す。
// 実データは一切使わない（テスト側が架空の行を渡す）。
'use strict';
const fs = require('fs');
const crypto = require('crypto');

const input = JSON.parse(fs.readFileSync(0, 'utf8'));
const source = fs.readFileSync(input.code_path, 'utf8');

function makeSheet(name, rows) {
  const sheet = {
    name,
    rows: (rows || []).map((r) => r.slice()),
    getName() { return name; },
    getLastRow() { return this.rows.length; },
    getMaxRows() { return Math.max(this.rows.length, 50); },
    getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); },
    setFrozenRows() { return this; },
    appendRow(values) { this.rows.push(values.map((v) => (v === undefined || v === null ? '' : v))); return this; },
    getRange(row, col, numRows, numCols) {
      const s = this;
      numRows = numRows || 1; numCols = numCols || 1;
      const ensure = (r, c) => { while (s.rows.length < r) s.rows.push([]); const arr = s.rows[r - 1]; while (arr.length < c) arr.push(''); return arr; };
      return {
        getDisplayValues() {
          const out = [];
          for (let r = 0; r < numRows; r += 1) {
            const rr = s.rows[row - 1 + r] || [];
            const line = [];
            for (let c = 0; c < numCols; c += 1) line.push(rr[col - 1 + c] === undefined ? '' : String(rr[col - 1 + c]));
            out.push(line);
          }
          return out;
        },
        getValues() { return this.getDisplayValues(); },
        getValue() { const rr = s.rows[row - 1] || []; return rr[col - 1] === undefined ? '' : rr[col - 1]; },
        setValue(v) { ensure(row, col)[col - 1] = v; return this; },
        setValues(vals) { vals.forEach((line, r) => line.forEach((v, c) => { ensure(row + r, col + c)[col + c - 1] = v; })); return this; },
        setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; },
        getRow() { return row; }, getLastRow() { return row + numRows - 1; }, getSheet() { return s; },
      };
    },
  };
  return sheet;
}

const sheets = {};
Object.keys(input.workbook || {}).forEach((name) => { sheets[name] = makeSheet(name, input.workbook[name]); });
const ss = {
  getSheetByName(name) { return sheets[name] || null; },
  insertSheet(name) { sheets[name] = makeSheet(name, []); return sheets[name]; },
  getActiveRange() { return input.active_range ? sheets[input.active_range.sheet].getRange(input.active_range.row, 1, input.active_range.rows, 1) : null; },
};
const alerts = [];
const ui = {
  ButtonSet: { OK_CANCEL: 'OK_CANCEL' }, Button: { OK: 'OK' },
  alert(msg) { alerts.push(String(msg)); },
  prompt() { return { getSelectedButton() { return 'OK'; }, getResponseText() { return input.prompt_answer || ''; } }; },
  createMenu() { const m = { addItem() { return m; }, addSeparator() { return m; }, addToUi() {} }; return m; },
};
global.SpreadsheetApp = { getActiveSpreadsheet() { return ss; }, getUi() { return ui; } };

function jst(date) { return new Date(date.getTime() + 9 * 3600 * 1000); }
global.Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' },
  computeDigest(_alg, text) { return Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b)); },
  formatDate(date, _tz, fmt) {
    const d = jst(date);
    const iso = d.toISOString();
    if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
    return iso.slice(0, 19);
  },
  getUuid() { return crypto.randomUUID(); },
};
global.LockService = { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } };
global.Session = { getActiveUser() { return { getEmail() { return 'tester@example.invalid'; } }; } };
const mails = [];
global.MailApp = {
  sendEmail(msg) {
    if (input.mail_fail && msg.to === input.mail_fail) throw new Error('mail failed: ' + msg.to);
    mails.push(msg);
  },
};
global.HtmlService = {
  createHtmlOutput(html) { return { html, setTitle() { return this; } }; },
  createTemplateFromFile() { return { evaluate() { return { setTitle() { return this; }, addMetaTag() { return this; } }; } }; },
};
global.console = { error() {}, log() {} };
if (input.now) {
  const fixed = new Date(input.now);
  const RealDate = Date;
  global.Date = class extends RealDate {
    constructor(...args) { super(...(args.length ? args : [fixed.getTime()])); }
    static now() { return fixed.getTime(); }
  };
}

// strict mode の直接 eval は関数宣言も const も eval の中に閉じるので、
// Code.gs と scenario を1つの eval にまとめて評価する（最後の式が結果になる）。
let result = null;
let error = null;
try {
  result = eval(source + '\n;(' + input.scenario + ')');
} catch (e) {
  error = String(e && e.message ? e.message : e);
}
const out = { result, error, alerts, mails, sheets: {} };
Object.keys(sheets).forEach((name) => { out.sheets[name] = sheets[name].rows; });
process.stdout.write(JSON.stringify(out));
