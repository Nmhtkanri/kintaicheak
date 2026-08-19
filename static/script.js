'use strict';

// =============================================================================
// ドラッグ&ドロップ + ファイル選択
// =============================================================================
function setupDropZone(dropZoneId, inputId, selectedId, multiple) {
    const dropZone = document.getElementById(dropZoneId);
    const input = document.getElementById(inputId);
    const selectedArea = document.getElementById(selectedId);

    if (!dropZone || !input) return;

    dropZone.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'LABEL') {
            input.click();
        }
    });

    input.addEventListener('change', () => updateSelected(input.files, selectedArea));

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files || []);
        const dt = new DataTransfer();
        const allowMultiple = multiple && input.multiple;
        const seen = new Set();

        const addFile = (file) => {
            if (!file) return;
            const key = `${file.name}\n${file.size}\n${file.lastModified}\n${file.type}`;
            if (seen.has(key)) return;
            seen.add(key);
            dt.items.add(file);
        };

        if (allowMultiple) {
            Array.from(input.files || []).forEach(addFile);
            files.forEach(addFile);
        } else {
            addFile(files[0]);
        }

        input.files = dt.files;
        updateSelected(input.files, selectedArea);
    });
}

function updateSelected(files, container) {
    container.innerHTML = '';
    if (!files || files.length === 0) return;
    for (const f of files) {
        const span = document.createElement('span');
        span.className = 'selected-file-item';
        span.textContent = f.name;
        container.appendChild(span);
    }
}

setupDropZone('jinjer-drop-zone', 'jinjer-input', 'jinjer-selected', false);
setupDropZone('timesheet-drop-zone', 'timesheet-input', 'timesheet-selected', true);
setupDropZone('shift-files-drop-zone', 'shift-files-input', 'shift-files-selected', true);
setupDropZone('hh-drop-zone', 'hh-file-input', 'hh-selected', false);

// =============================================================================
// グローバル状態
// =============================================================================
const form = document.getElementById('upload-form');
const runBtn = document.getElementById('run-btn');
const progressArea = document.getElementById('progress-area');
const progressBar = document.getElementById('progress-bar');
const progressMessage = document.getElementById('progress-message');
const errorArea = document.getElementById('error-area');
const resultArea = document.getElementById('result-area');
const csvExportArea = document.getElementById('csv-export-area');

let progressStep = 0;

// 凡例レビューに必要な状態
let pendingSessionId = null;
let pendingCodeSheets = [];
let pendingMode = 'match';  // match | csv_export
let pendingTemplates = [];  // jinjer スケジュール雛形の一覧（プルダウン用）

// メール下書きモードの状態。
// 起動時のモードがメールの場合、applyModeUI の初回実行（このすぐ下）から
// loadMailTemplates が参照するため、必ずここ（初回実行より前）で宣言しておくこと。
let mailPlans = [];
let mailTemplates = [];
let mailTemplatesLoaded = false;
let mailDefaultCc = '';

// =============================================================================
// モード切替UI
// =============================================================================
function getCurrentMode() {
    const checked = document.querySelector('input[name="mode"]:checked');
    return checked ? checked.value : 'match';
}

// 固定タブバー右側に出す、現在モードの一言説明
const MODE_HINTS = {
    match:      '請求勤怠と jinjer の差異を洗い出す',
    csv_export: 'シフト表を読み取って jinjer へ登録する',
    keiri:      'freee 取引インポート用の4CSVを作る',
    invoice:    '請求書PDFを確認してfreee売上取引CSVを作る',
    sharoushi:  '社労士へ渡す給与CSVを作る／保険料一覧表PDFの標準報酬をjinjerへ投入する',
    shaho:      '標準報酬月額の検算と保険料突合（jinjerには書きません）',
    expense:    'テレワーク・出社日数と経費を集計する',
    mail:       '下書きのみ作成・送信はしません',
    health_hpm: '健診ExcelからHPM取込用CSVを作る（取込は手動）',
};

// 進捗バー／エラー表示がどのモードのものかを覚えておく。
// タブでモードを切り替えても消さず、そのモードに戻ったときに再表示するため。
let runningMode = null;
let errorMode = null;

function applyModeUI(mode) {
    const jinjerSection = document.getElementById('jinjer-section');
    const timesheetSection = document.getElementById('timesheet-section');
    const shiftFilesSection = document.getElementById('shift-files-section');
    const targetYmSection = document.getElementById('target-ym-section');
    const jinjerRequiredTag = document.getElementById('jinjer-required-tag');
    const jinjerOptionalTag = document.getElementById('jinjer-optional-tag');
    const jinjerInput = document.getElementById('jinjer-input');
    const timesheetInput = document.getElementById('timesheet-input');
    const shiftFilesInput = document.getElementById('shift-files-input');
    const targetYearInput = document.getElementById('target-year');
    const targetMonthInput = document.getElementById('target-month');
    const settingsSection = document.getElementById('settings-section');
    const runBtn = document.getElementById('run-btn');
    const matchStepHeader = document.getElementById('match-step-header');
    const scheduleStepHeader = document.getElementById('schedule-step-header');
    const monthlyCompareCard = document.getElementById('monthly-compare-card');
    const monthlyExportCard = document.getElementById('monthly-export-card');
    const batchCompareCard = document.getElementById('batch-compare-card');
    const expenseCard = document.getElementById('expense-card');
    const kiCard = document.getElementById('ki-card');
    const keiriCard = document.getElementById('keiri-card');
    const invoiceCard = document.getElementById('invoice-card');
    const sharoushiCard = document.getElementById('sharoushi-card');
    const shahoImportCard = document.getElementById('shaho-import-card');
    const shahoCard = document.getElementById('shaho-card');
    const mailCard = document.getElementById('mail-card');
    const healthCard = document.getElementById('health-card');
    const sseCard = document.getElementById('sse-card');

    const isSchedule = mode === 'csv_export';
    const isExpense = mode === 'expense';
    const isKeiri = mode === 'keiri';
    const isInvoice = mode === 'invoice';
    const isSharoushi = mode === 'sharoushi';
    const isShaho = mode === 'shaho';
    const isMail = mode === 'mail';
    const isHealthHpm = mode === 'health_hpm';
    const isMatch = !isSchedule && !isExpense && !isKeiri && !isSharoushi
        && !isInvoice && !isShaho && !isMail && !isHealthHpm;

    // 突合アップロードフォーム本体は常に隠す（モード選択だけ残す。突合は⚡一括/手順2-3で行う）
    if (jinjerSection) jinjerSection.style.display = 'none';
    if (timesheetSection) timesheetSection.style.display = 'none';
    if (jinjerInput) jinjerInput.disabled = true;
    if (timesheetInput) timesheetInput.disabled = true;
    if (matchStepHeader) matchStepHeader.style.display = 'none';
    if (settingsSection) settingsSection.style.display = 'none';

    // スケジュールアップロード専用UI（csv_export モードのみ表示・有効化）
    if (targetYmSection) targetYmSection.style.display = isSchedule ? '' : 'none';
    if (shiftFilesSection) shiftFilesSection.style.display = isSchedule ? '' : 'none';
    if (shiftFilesInput) shiftFilesInput.disabled = !isSchedule;
    if (targetYearInput) targetYearInput.disabled = !isSchedule;
    if (targetMonthInput) targetMonthInput.disabled = !isSchedule;
    if (scheduleStepHeader) scheduleStepHeader.style.display = isSchedule ? '' : 'none';
    // スケジュール開始時刻のピンポイント修正カードもスケジュールモードのみ
    if (sseCard) sseCard.style.display = isSchedule ? '' : 'none';

    // ①フォームの実行ボタンはスケジュールモードのみ（スケジュールCSV作成）
    if (runBtn) {
        if (isSchedule) { runBtn.style.display = ''; runBtn.textContent = 'スケジュールCSVを作成'; }
        else runBtn.style.display = 'none';
    }

    // 勤怠チェックの導線（⚡一括＋手順2/手順3）は勤怠チェックモードのみ
    if (batchCompareCard) batchCompareCard.style.display = isMatch ? '' : 'none';
    if (monthlyCompareCard) monthlyCompareCard.style.display = isMatch ? '' : 'none';
    if (monthlyExportCard) monthlyExportCard.style.display = isMatch ? '' : 'none';

    // 経費チェックモード: 承認前作業＋経費集計の2カードを表示
    if (expenseCard) expenseCard.style.display = isExpense ? '' : 'none';
    if (kiCard) kiCard.style.display = isExpense ? '' : 'none';
    // 経理モード: 仕訳CSV生成カードのみ表示
    if (keiriCard) keiriCard.style.display = isKeiri ? '' : 'none';
    // 請求書モード: PDF確認とfreee CSV生成カードのみ表示
    if (invoiceCard) invoiceCard.style.display = isInvoice ? '' : 'none';
    // 社労士モード: 社労士CSV生成カードと、保険料一覧表PDFの投入カード
    if (sharoushiCard) sharoushiCard.style.display = isSharoushi ? '' : 'none';
    if (shahoImportCard) shahoImportCard.style.display = isSharoushi ? '' : 'none';
    // 標準報酬チェック: 検算カードのみ表示
    if (shahoCard) shahoCard.style.display = isShaho ? '' : 'none';
    // メール下書きモード: メールカードのみ表示（初回表示時にテンプレ一覧を読み込む）
    if (mailCard) {
        mailCard.style.display = isMail ? '' : 'none';
        if (isMail) loadMailTemplates();
    }
    // 健康診断HPMモード: 健診カードのみ表示
    if (healthCard) healthCard.style.display = isHealthHpm ? '' : 'none';

    // アップロードフォーム（＝スケジュールモードの入力カード）はスケジュールモードのみ表示
    if (form) form.style.display = isSchedule ? '' : 'none';

    // タブバー右側の一言説明
    const modeBarHint = document.getElementById('mode-bar-hint');
    if (modeBarHint) modeBarHint.textContent = MODE_HINTS[mode] || '';

    // 出力エリアの片付け。
    // data-mode を持つ出力エリアは「そのモードのときだけ、実行済みなら表示」する。
    // 中身は消さないので、モードを戻せば結果がそのまま再表示される（再実行は不要）。
    // 進捗・エラーも同じ扱い：所属モードに戻ったときだけ出す（実行中の切替で見失わない）。
    // カードの内側にある結果（keiri/expense/mail/bc/qe）はカードごと隠れるのでここでは触らない。
    if (progressArea) progressArea.style.display = (runningMode === mode) ? 'block' : 'none';
    if (errorArea) errorArea.style.display = (errorMode === mode) ? 'block' : 'none';
    document.querySelectorAll('[data-mode]').forEach(el => {
        const show = (el.dataset.mode === mode) && (el.dataset.hasResult === '1');
        el.style.display = show ? 'block' : 'none';
    });
}

document.querySelectorAll('input[name="mode"]').forEach(radio => {
    radio.addEventListener('change', () => {
        applyModeUI(getCurrentMode());
        // タブで切り替えたら、そのモードの先頭（入力カード）が見える位置へ戻す
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
});
applyModeUI(getCurrentMode());

// =============================================================================
// フォーム送信
// =============================================================================
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    pendingMode = getCurrentMode();
    startProcessing();

    const formData = new FormData(form);
    if (!formData.has('mode')) formData.append('mode', pendingMode);
    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData,
            headers: { 'Accept': 'text/event-stream' }
        });

        if (!response.ok) {
            const data = await response.json();
            showError(data.errors ? data.errors.join('<br>') : 'エラーが発生しました');
            stopProcessing();
            return;
        }

        await consumeSSEResponse(response);

    } catch (err) {
        showError('通信エラーが発生しました: ' + err.message);
        stopProcessing();
    }
});


// handler を差し替えられるようにしてある。既定は勤怠チェック用の進捗バー処理だが、
// 健診モードのように専用の進捗表示を持つ画面では別のハンドラを渡す。
async function consumeSSEResponse(response, handler = processSSEPart) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const parts = buffer.split('\n\n');
        buffer = parts.pop();
        for (const part of parts) {
            handler(part);
        }
    }
}


function parseSSEPart(part) {
    const lines = part.split('\n');
    let eventType = null;
    let data = null;
    for (const line of lines) {
        if (line.startsWith('event: ')) eventType = line.slice(7).trim();
        if (line.startsWith('data: ')) {
            try { data = JSON.parse(line.slice(6)); } catch { }
        }
    }
    return { eventType, data };
}


function processSSEPart(part) {
    const { eventType, data } = parseSSEPart(part);
    if (!eventType || !data) return;

    if (eventType === 'progress') {
        progressStep = Math.min(progressStep + 12, 80);
        progressBar.style.width = progressStep + '%';
        progressMessage.textContent = data.message || '処理中...';
    } else if (eventType === 'code_review_needed') {
        // 凡例レビュー画面を出す
        progressArea.style.display = 'none';
        pendingSessionId = data.session_id;
        pendingCodeSheets = data.code_sheets || [];
        pendingMode = data.mode || pendingMode || 'match';
        pendingTemplates = data.available_templates || [];
        openLegendModal(pendingCodeSheets);
        stopProcessing(false);
    } else if (eventType === 'done') {
        progressBar.style.width = '100%';
        progressMessage.textContent = '完了！';
        setTimeout(() => {
            progressArea.style.display = 'none';
            showResult(data);
        }, 400);
        stopProcessing(false);
    } else if (eventType === 'csv_export_done') {
        progressBar.style.width = '100%';
        progressMessage.textContent = '完了！';
        setTimeout(() => {
            progressArea.style.display = 'none';
            showCsvExportResult(data);
        }, 400);
        stopProcessing(false);
    } else if (eventType === 'error') {
        showError(data.message || 'エラーが発生しました');
        stopProcessing();
    }
}

// =============================================================================
// 進捗・結果表示
// =============================================================================
function startProcessing() {
    progressStep = 5;
    runningMode = getCurrentMode();
    errorMode = null;
    progressBar.style.width = progressStep + '%';
    progressMessage.textContent = '処理を開始しています...';
    progressArea.style.display = 'block';
    errorArea.style.display = 'none';
    resultArea.style.display = 'none';
    resultArea.dataset.hasResult = '0';
    if (csvExportArea) {
        csvExportArea.style.display = 'none';
        csvExportArea.dataset.hasResult = '0';
    }
    runBtn.disabled = true;
    runBtn.textContent = '処理中...';
}

function stopProcessing(hideProgress = true) {
    runningMode = null;
    runBtn.disabled = false;
    // モードに応じたボタンラベルに戻す
    runBtn.textContent = (getCurrentMode() === 'csv_export') ? 'スケジュールCSVを作成' : 'チェック実行';
    if (hideProgress) progressArea.style.display = 'none';
}

function showError(msg) {
    errorMode = getCurrentMode();
    errorArea.innerHTML = msg;
    errorArea.style.display = 'block';
    resultArea.style.display = 'none';
    resultArea.dataset.hasResult = '0';
}

function showResult(data) {
    const { summary, table, excel_filename, unsubmitted, warnings, new_template_filename, new_template_count } = data;

    document.getElementById('cnt-ok').textContent = summary.ok;
    document.getElementById('cnt-ng').textContent = summary.ng;
    document.getElementById('cnt-caution').textContent = summary.caution;
    document.getElementById('cnt-missing').textContent = summary.missing;

    const allOkMsg = document.getElementById('all-ok-msg');
    const ngTableArea = document.getElementById('ng-table-area');
    const warningArea = document.getElementById('result-warning-area');

    if (warnings && warnings.length > 0) {
        warningArea.innerHTML = warnings.map(msg => `<div>${escapeHtml(msg)}</div>`).join('');
        warningArea.style.display = 'block';
    } else {
        warningArea.innerHTML = '';
        warningArea.style.display = 'none';
    }

    if (summary.ng === 0 && summary.caution === 0 && summary.missing === 0) {
        allOkMsg.style.display = 'block';
        ngTableArea.style.display = 'none';
    } else if (summary.ng > 0 || summary.caution > 0) {
        allOkMsg.style.display = 'none';
        ngTableArea.style.display = 'block';
        renderTable(table);
    } else {
        allOkMsg.style.display = 'none';
        ngTableArea.style.display = 'none';
    }

    // 未提出者リスト表示
    const unsubArea = document.getElementById('unsubmitted-area');
    const unsubList = document.getElementById('unsubmitted-list');
    if (unsubmitted && unsubmitted.length > 0) {
        unsubList.innerHTML = '';
        unsubmitted.forEach(name => {
            const tag = document.createElement('span');
            tag.className = 'unsubmitted-tag';
            tag.textContent = name;
            unsubList.appendChild(tag);
        });
        unsubArea.style.display = 'block';
    } else {
        unsubArea.style.display = 'none';
    }

    const downloadLink = document.getElementById('download-link');
    if (excel_filename) {
        downloadLink.href = '/download/' + encodeURIComponent(excel_filename);
        downloadLink.style.display = 'inline-block';
    }

    // 新規雛形 CSV
    const newTplLink = document.getElementById('new-template-link');
    const newTplMsg = document.getElementById('new-template-msg');
    if (new_template_filename) {
        newTplLink.href = '/download/' + encodeURIComponent(new_template_filename);
        newTplLink.style.display = 'inline-block';
        newTplMsg.textContent =
            `✨ jinjer の既存雛形に該当しない記号が ${new_template_count} 件ありました。` +
            `新規雛形 CSV を生成したので、jinjer にそのままインポートできます。`;
        newTplMsg.style.display = 'block';
    } else {
        newTplLink.style.display = 'none';
        newTplMsg.style.display = 'none';
    }

    // 実行中に別モードへ切り替えられていても結果は保持し、そのモードに戻ったときに出す
    resultArea.dataset.hasResult = '1';
    if (getCurrentMode() === 'match') {
        resultArea.style.display = 'block';
        resultArea.scrollIntoView({ behavior: 'smooth' });
    }
}

function renderTable(rows) {
    const tbody = document.getElementById('result-tbody');
    tbody.innerHTML = '';

    const judgeClass = {
        'NG': 'judge-ng',
        '要確認': 'judge-caution',
        'OK': 'judge-ok',
        'データ欠損': 'judge-missing',
    };

    for (const row of rows) {
        const tr = document.createElement('tr');
        const cells = [
            row.name, row.date,
            row.sheet_start, row.jinjer_start, row.start_diff,
            row.sheet_end, row.jinjer_end, row.end_diff,
            row.sheet_total_work, row.jinjer_total_work, row.total_work_diff,
            row.judgment, row.detail
        ];
        cells.forEach((val, idx) => {
            const td = document.createElement('td');
            td.textContent = val !== null && val !== undefined ? val : '';
            if (idx === 11) {
                td.className = judgeClass[row.judgment] || '';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    }
}

// =============================================================================
// 凡例レビューモーダル
// =============================================================================
const legendModal = document.getElementById('legend-modal');
const legendModalClose = document.getElementById('legend-modal-close');
const legendCancelBtn = document.getElementById('legend-cancel-btn');
const legendConfirmBtn = document.getElementById('legend-confirm-btn');
const legendSheetsContainer = document.getElementById('legend-sheets-container');

function openLegendModal(sheets) {
    legendSheetsContainer.innerHTML = '';
    sheets.forEach((sheet, sheetIdx) => {
        legendSheetsContainer.appendChild(renderLegendSheet(sheet, sheetIdx));
    });
    legendModal.style.display = 'flex';
    // スケジュールアップロードモードでは、CSV を作る前に氏名→従業員IDを照合しておく。
    // 同姓（吉田 等）で自動確定できない人はここで気づいて選べる＝やり直しが要らない。
    checkScheduleNames();
}

// =============================================================================
// 氏名 → jinjer 従業員ID の事前照合
// =============================================================================

// 1行分の照合結果を描画する。ambiguous のときは候補プルダウンを出す。
function renderNameStatus(rowEl, result) {
    const td = rowEl.querySelector('.legend-employee-status');
    if (!td) return;
    td.innerHTML = '';
    rowEl.dataset.nameStatus = (result && result.status) || '';

    if (!result) return;

    const badge = document.createElement('span');
    badge.className = 'name-status name-status-' + result.status;

    if (result.status === 'checking') {
        badge.textContent = '照合中…';
        badge.className = 'name-status name-status-checking';
        td.appendChild(badge);
        return;
    }
    if (result.status === 'ok') {
        const via = result.via_alias ? '（エイリアス）' : '';
        badge.textContent = `✅ ${result.official_name || ''}(${result.employee_id})${via}`;
        badge.title = result.via_alias
            ? '氏名エイリアス表で従業員IDを確定しました'
            : 'jinjer の従業員IDが一意に決まりました';
        td.appendChild(badge);
        return;
    }
    if (result.status === 'error') {
        badge.textContent = '⚠️ 照合できません';
        badge.title = result.message || '';
        td.appendChild(badge);
        return;
    }
    if (result.status === 'ambiguous') {
        badge.textContent = '⚠️ 同姓が複数';
        badge.title = '姓だけでは決められません。下の候補から選ぶか、フルネームを入力してください。';
        td.appendChild(badge);

        const sel = document.createElement('select');
        sel.className = 'name-candidate-select';
        const blank = document.createElement('option');
        blank.value = '';
        blank.textContent = '候補から選ぶ…';
        sel.appendChild(blank);
        (result.candidates || []).forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.full_name;
            opt.textContent = `${c.full_name} (${c.employee_id})`;
            sel.appendChild(opt);
        });
        sel.addEventListener('change', () => {
            if (!sel.value) return;
            const inp = rowEl.querySelector('input[data-field="employee_name"]');
            if (!inp) return;
            inp.value = sel.value;   // フルネームにすれば既存の解決ロジックで一意に決まる
            recheckOneName(rowEl);
        });
        td.appendChild(sel);
        return;
    }
    // unknown
    badge.textContent = '❓ jinjerに該当者なし';
    badge.title = 'jinjer に登録されている氏名（漢字）と一致するように入力してください。';
    td.appendChild(badge);
}

function employeeRowsOf(sheetEl) {
    return Array.from(sheetEl.querySelectorAll('tr.legend-employee-row'));
}

async function postNameCheck(sheets) {
    const response = await fetch('/resolve_schedule_names', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: pendingSessionId, sheets: sheets }),
    });
    return await response.json();
}

// モーダル内の全氏名をまとめて照合する
async function checkScheduleNames() {
    if (pendingMode !== 'csv_export') return;

    const sheetEls = Array.from(legendSheetsContainer.querySelectorAll('.legend-sheet'));
    const payloadSheets = [];
    sheetEls.forEach(sheetEl => {
        const original = pendingCodeSheets[parseInt(sheetEl.dataset.sheetIdx, 10)];
        if (!original) return;
        const rows = employeeRowsOf(sheetEl);
        rows.forEach(r => renderNameStatus(r, { status: 'checking' }));
        payloadSheets.push({
            filename: original.filename,
            names: rows.map(r => {
                const inp = r.querySelector('input[data-field="employee_name"]');
                return inp ? inp.value.trim() : '';
            }),
        });
    });
    if (payloadSheets.length === 0) return;

    let data;
    try {
        data = await postNameCheck(payloadSheets);
    } catch (e) {
        sheetEls.forEach(el => employeeRowsOf(el).forEach(
            r => renderNameStatus(r, { status: 'error', message: String(e) })));
        return;
    }
    if (!data || !data.success) {
        const message = (data && data.error) || '照合に失敗しました';
        sheetEls.forEach(el => employeeRowsOf(el).forEach(
            r => renderNameStatus(r, { status: 'error', message: message })));
        return;
    }

    let payloadIdx = 0;
    sheetEls.forEach(sheetEl => {
        const original = pendingCodeSheets[parseInt(sheetEl.dataset.sheetIdx, 10)];
        if (!original) return;
        const res = data.sheets[payloadIdx++];
        if (!res) return;
        const rows = employeeRowsOf(sheetEl);
        rows.forEach((r, i) => renderNameStatus(r, (res.names || [])[i]));
    });
    (data.warnings || []).forEach(w => console.warn('[氏名エイリアス]', w));
}

// 1行だけ再照合する（氏名を直した／候補を選んだとき）
async function recheckOneName(rowEl) {
    if (pendingMode !== 'csv_export') return;
    const sheetEl = rowEl.closest('.legend-sheet');
    if (!sheetEl) return;
    const original = pendingCodeSheets[parseInt(sheetEl.dataset.sheetIdx, 10)];
    if (!original) return;
    const inp = rowEl.querySelector('input[data-field="employee_name"]');
    const name = inp ? inp.value.trim() : '';

    renderNameStatus(rowEl, { status: 'checking' });
    try {
        const data = await postNameCheck([{ filename: original.filename, names: [name] }]);
        if (data && data.success && data.sheets[0]) {
            renderNameStatus(rowEl, data.sheets[0].names[0]);
        } else {
            renderNameStatus(rowEl, {
                status: 'error', message: (data && data.error) || '照合に失敗しました' });
        }
    } catch (e) {
        renderNameStatus(rowEl, { status: 'error', message: String(e) });
    }
}

// 未照合（同姓未選択・該当者なし）の氏名を集める
function collectUnresolvedNames() {
    const unresolved = [];
    legendSheetsContainer.querySelectorAll('.legend-sheet').forEach(sheetEl => {
        employeeRowsOf(sheetEl).forEach(r => {
            const status = r.dataset.nameStatus;
            if (status === 'ambiguous' || status === 'unknown') {
                const inp = r.querySelector('input[data-field="employee_name"]');
                unresolved.push((inp ? inp.value.trim() : '') || '(未入力)');
            }
        });
    });
    return unresolved;
}

function closeLegendModal() {
    legendModal.style.display = 'none';
    pendingSessionId = null;
    pendingCodeSheets = [];
}

legendModalClose.addEventListener('click', closeLegendModal);
legendCancelBtn.addEventListener('click', closeLegendModal);
legendModal.addEventListener('click', (e) => {
    if (e.target === legendModal) closeLegendModal();
});

function renderLegendSheet(sheet, sheetIdx) {
    const wrap = document.createElement('div');
    wrap.className = 'legend-sheet';
    wrap.dataset.sheetIdx = sheetIdx;
    sheet.__sheetIdx = sheetIdx;

    const header = document.createElement('div');
    header.className = 'legend-sheet-header';
    const empCount = (sheet.employees && sheet.employees.length) || 0;
    header.innerHTML = `
        <span>📄 ${escapeHtml(sheet.filename || '勤務表')}</span>
        <span class="legend-sheet-header-meta">従業員 ${empCount}人</span>
    `;
    wrap.appendChild(header);

    // 対象年月
    const ymRow = document.createElement('div');
    ymRow.className = 'legend-ym-row';
    ymRow.innerHTML = `
        <label>対象年月:
            <input type="number" data-ym="year" value="${sheet.year || ''}" placeholder="2026" min="2000" max="2099" style="width:80px">
            年
            <input type="number" data-ym="month" value="${sheet.month || ''}" placeholder="4" min="1" max="12" style="width:60px">
            月
        </label>
        <span class="legend-ym-hint">※年月が読み取れない場合のみ入力</span>
    `;
    wrap.appendChild(ymRow);

    // 従業員氏名（手入力で修正可）
    wrap.appendChild(renderEmployeesEditor(sheet, sheetIdx));

    // 雛形マッチ結果（あれば表示用に使う）
    const tmMatched = (sheet.template_match && sheet.template_match.matched) || [];
    const tmMatchedMap = {};
    tmMatched.forEach(m => { tmMatchedMap[m.code] = m; });

    const table = document.createElement('table');
    table.className = 'legend-table';
    table.innerHTML = `
        <thead>
            <tr>
                <th>記号</th>
                <th>種別</th>
                <th>出勤</th>
                <th>退勤</th>
                <th>休憩(分)</th>
                <th>休?</th>
                <th>jinjer 雛形</th>
                <th></th>
            </tr>
        </thead>
        <tbody></tbody>
    `;
    const tbody = table.querySelector('tbody');

    const legend = sheet.legend || [];
    legend.forEach((entry, rowIdx) => {
        tbody.appendChild(renderLegendRow(entry, sheetIdx, rowIdx, tmMatchedMap));
    });

    wrap.appendChild(table);

    const actions = document.createElement('div');
    actions.className = 'legend-actions';

    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'btn-mini btn-mini-add';
    addBtn.textContent = '＋ 記号を追加';
    addBtn.addEventListener('click', () => {
        const newEntry = { code: '', label: '', start_time: '', end_time: '', break_minutes: 0, is_off: false };
        const idx = tbody.children.length;
        tbody.appendChild(renderLegendRow(newEntry, sheetIdx, idx, tmMatchedMap));
        // sheet.legend にも反映するため、最後に取得時に DOM から読む
    });
    actions.appendChild(addBtn);
    wrap.appendChild(actions);

    return wrap;
}

function renderEmployeesEditor(sheet, sheetIdx) {
    const wrap = document.createElement('div');
    wrap.className = 'legend-employees';

    const title = document.createElement('div');
    title.className = 'legend-employees-title';
    title.innerHTML = '👥 従業員氏名 <span class="legend-employees-hint">jinjer に登録されている氏名（漢字）と一致するように入力してください</span>';
    wrap.appendChild(title);

    const tbl = document.createElement('table');
    tbl.className = 'legend-employees-table';
    tbl.innerHTML = `
        <thead>
            <tr>
                <th style="width:34%">氏名</th>
                <th style="width:38%">jinjer 照合</th>
                <th style="width:20%">シフト件数</th>
                <th></th>
            </tr>
        </thead>
        <tbody></tbody>
    `;
    const tbody = tbl.querySelector('tbody');

    const employees = sheet.employees || [];
    employees.forEach((emp, empIdx) => {
        appendEmployeeRows(tbody, emp, empIdx);
    });

    if (employees.length === 0) {
        const tr = document.createElement('tr');
        tr.className = 'legend-employees-empty';
        tr.innerHTML = `<td colspan="4">⚠️ 画像から従業員を抽出できませんでした。下の「+ 従業員を追加」から手動で追加してください（ただしシフトは画像から取得できないため、スケジュールCSVは作成できません）。</td>`;
        tbody.appendChild(tr);
    }

    wrap.appendChild(tbl);

    const actions = document.createElement('div');
    actions.className = 'legend-employees-actions';

    const addEmpBtn = document.createElement('button');
    addEmpBtn.type = 'button';
    addEmpBtn.className = 'btn-mini btn-mini-add';
    addEmpBtn.textContent = '＋ 従業員を追加';
    addEmpBtn.addEventListener('click', () => addEmployeeToSheet(sheet));
    actions.appendChild(addEmpBtn);
    wrap.appendChild(actions);

    return wrap;
}

function addEmployeeToSheet(sheet) {
    const sheetEl = legendSheetsContainer.querySelector(`.legend-sheet[data-sheet-idx="${sheet.__sheetIdx}"]`);
    if (!sheetEl) return;

    const empTbody = sheetEl.querySelector('.legend-employees-table tbody');
    if (!empTbody) return;

    // 「画像から抽出できませんでした」の placeholder 行は除去する
    const emptyRow = empTbody.querySelector('.legend-employees-empty');
    if (emptyRow) emptyRow.remove();

    if (!sheet.employees) sheet.employees = [];
    const newEmp = { name: '', shifts: [] };
    sheet.employees.push(newEmp);

    const newIdx = sheet.employees.length - 1;
    const row = appendEmployeeRows(empTbody, newEmp, newIdx);
    row.dataset.isNew = '1';

    const inp = row.querySelector('input[data-field="employee_name"]');
    if (inp) inp.focus();

    // ヘッダーの従業員人数表示を更新
    const meta = sheetEl.querySelector('.legend-sheet-header-meta');
    if (meta) meta.textContent = `従業員 ${sheet.employees.length}人`;
}

// 従業員1人につき「本行(氏名・件数・削除)」＋「詳細行(日別シフト編集)」の2行を返す。
// 詳細行は本行の直下に並べ、本行の _detailRow から参照できるようにする。
function renderEmployeeRow(emp, empIdx) {
    const tr = document.createElement('tr');
    tr.className = 'legend-employee-row';
    tr.dataset.empIdx = empIdx;

    // 氏名
    const nameTd = document.createElement('td');
    const nameInp = document.createElement('input');
    nameInp.type = 'text';
    nameInp.placeholder = '例: 田村桃子';
    const rawName = (emp && emp.name) ? String(emp.name) : '';
    // Claude が "不明" を入れていた場合は空欄にして、ユーザーに必ず手入力させる
    nameInp.value = (rawName === '不明') ? '' : rawName;
    nameInp.dataset.field = 'employee_name';
    nameInp.className = 'legend-employee-name-input';
    nameTd.appendChild(nameInp);
    tr.appendChild(nameTd);

    // jinjer 照合結果（従業員ID が引けるか）。CSV を作る前にここで気づけるようにする。
    const statusTd = document.createElement('td');
    statusTd.className = 'legend-employee-status';
    tr.appendChild(statusTd);
    // 氏名を直したら、その行だけ再照合する
    nameInp.addEventListener('change', () => recheckOneName(tr));

    // 詳細行（日別シフト編集）を先に作る
    const detail = document.createElement('tr');
    detail.className = 'legend-employee-detail';
    detail.dataset.empIdx = empIdx;
    detail.style.display = 'none';
    const detailTd = document.createElement('td');
    detailTd.colSpan = 4;
    const editor = buildShiftEditor((emp && emp.shifts) ? emp.shifts : []);
    detailTd.appendChild(editor);
    detail.appendChild(detailTd);
    tr._detailRow = detail;

    // シフト件数 ＋ 編集トグル
    const cntTd = document.createElement('td');
    cntTd.className = 'legend-employee-shift-count';
    const toggleBtn = document.createElement('button');
    toggleBtn.type = 'button';
    toggleBtn.className = 'btn-mini legend-shift-toggle';
    const updateToggleLabel = () => {
        const n = editor.querySelectorAll('.legend-shift-row').length;
        const open = detail.style.display !== 'none';
        toggleBtn.textContent = `🗓 ${n}日分を${open ? '閉じる' : '編集'}`;
    };
    toggleBtn.addEventListener('click', () => {
        detail.style.display = (detail.style.display === 'none') ? '' : 'none';
        updateToggleLabel();
    });
    editor._updateCount = updateToggleLabel;  // 行の追加/削除時に件数表示を更新
    cntTd.appendChild(toggleBtn);
    tr.appendChild(cntTd);
    updateToggleLabel();

    // 削除（本行＋詳細行をまとめて除外）
    const delTd = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'btn-mini btn-mini-danger';
    delBtn.textContent = '🗑';
    delBtn.title = 'この従業員を除外';
    delBtn.addEventListener('click', () => {
        if (tr._detailRow) tr._detailRow.remove();
        tr.remove();
    });
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    return { main: tr, detail: detail };
}

// 日別シフト編集エリア（日付＋記号の行を並べる）
function buildShiftEditor(shifts) {
    const box = document.createElement('div');
    box.className = 'legend-shift-editor';

    const hint = document.createElement('div');
    hint.className = 'legend-shift-hint';
    hint.textContent = '日付と記号を直接修正できます。誤読・空欄の記号はここで直してください（記号は凡例の記号を入力）。';
    box.appendChild(hint);

    const rows = document.createElement('div');
    rows.className = 'legend-shift-rows';
    (shifts || []).forEach(s => rows.appendChild(buildShiftRow(s)));
    box.appendChild(rows);

    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.className = 'btn-mini btn-mini-add';
    addBtn.textContent = '＋ 日を追加';
    addBtn.addEventListener('click', () => {
        rows.appendChild(buildShiftRow({ date: '', code: '', comment: null }));
        if (box._updateCount) box._updateCount();
    });
    box.appendChild(addBtn);

    return box;
}

function buildShiftRow(shift) {
    const row = document.createElement('div');
    row.className = 'legend-shift-row';
    if (shift && shift.comment) row.dataset.comment = shift.comment;

    const dInp = document.createElement('input');
    dInp.type = 'text';
    dInp.className = 'legend-shift-date';
    dInp.dataset.shiftField = 'date';
    dInp.placeholder = 'YYYY-MM-DD';
    if (shift && shift.date) dInp.value = shift.date;
    row.appendChild(dInp);

    const cInp = document.createElement('input');
    cInp.type = 'text';
    cInp.className = 'legend-shift-code';
    cInp.dataset.shiftField = 'code';
    cInp.placeholder = '記号';
    if (shift && shift.code != null) cInp.value = shift.code;
    row.appendChild(cInp);

    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'btn-mini btn-mini-danger';
    del.textContent = '🗑';
    del.title = 'この日を削除';
    del.addEventListener('click', () => {
        const box = row.closest('.legend-shift-editor');
        row.remove();
        if (box && box._updateCount) box._updateCount();
    });
    row.appendChild(del);

    return row;
}

// 本行＋詳細行を tbody に追加し、本行を返す（focus / dataset 用）
function appendEmployeeRows(tbody, emp, empIdx) {
    const { main, detail } = renderEmployeeRow(emp, empIdx);
    tbody.appendChild(main);
    tbody.appendChild(detail);
    return main;
}

function renderLegendRow(entry, sheetIdx, rowIdx, tmMatchedMap) {
    const tr = document.createElement('tr');
    tr.dataset.rowIdx = rowIdx;
    if (entry.is_off) tr.classList.add('legend-row-off');

    const codeTd = document.createElement('td');
    codeTd.appendChild(makeInput(entry.code || '', 'code'));
    tr.appendChild(codeTd);

    const labelTd = document.createElement('td');
    labelTd.appendChild(makeInput(entry.label || '', 'label'));
    tr.appendChild(labelTd);

    const startTd = document.createElement('td');
    startTd.appendChild(makeInput(entry.start_time || '', 'start_time', '12:30'));
    tr.appendChild(startTd);

    const endTd = document.createElement('td');
    endTd.appendChild(makeInput(entry.end_time || '', 'end_time', '21:00'));
    tr.appendChild(endTd);

    const breakTd = document.createElement('td');
    const breakInput = document.createElement('input');
    breakInput.type = 'number';
    breakInput.min = '0';
    breakInput.step = '5';
    breakInput.value = entry.break_minutes || 0;
    breakInput.dataset.field = 'break_minutes';
    breakTd.appendChild(breakInput);
    tr.appendChild(breakTd);

    const offTd = document.createElement('td');
    const offChk = document.createElement('input');
    offChk.type = 'checkbox';
    offChk.checked = !!entry.is_off;
    offChk.dataset.field = 'is_off';
    offChk.addEventListener('change', () => {
        if (offChk.checked) tr.classList.add('legend-row-off');
        else tr.classList.remove('legend-row-off');
    });
    offTd.appendChild(offChk);
    tr.appendChild(offTd);

    const tplTd = document.createElement('td');
    const matched = tmMatchedMap[entry.code];
    if (entry.is_off) {
        const lbl = document.createElement('span');
        lbl.className = 'legend-template-tag';
        lbl.textContent = '休扱い';
        tplTd.appendChild(lbl);
    } else if (pendingTemplates && pendingTemplates.length) {
        // jinjer 雛形をプルダウンで選択（初期値＝自動マッチ）。選んだ雛形IDがCSVに入る。
        tplTd.appendChild(buildTemplateSelect(entry, matched));
    } else {
        // 雛形マスタが読めない場合は従来のタグ表示
        const tag = document.createElement('span');
        if (matched) {
            tag.className = 'legend-template-tag';
            tag.textContent = `No${matched.template_no} ${matched.template_name}`;
        } else if (entry.start_time) {
            tag.className = 'legend-template-tag unmatched';
            tag.textContent = '新規雛形候補';
        }
        if (tag.textContent) tplTd.appendChild(tag);
    }
    tr.appendChild(tplTd);

    const delTd = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'btn-mini btn-mini-danger';
    delBtn.textContent = '🗑';
    delBtn.title = 'この行を削除';
    delBtn.addEventListener('click', () => tr.remove());
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    return tr;
}

function makeInput(value, field, placeholder = '') {
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.value = value;
    inp.dataset.field = field;
    if (placeholder) inp.placeholder = placeholder;
    return inp;
}

// "09:00:00" → "09:00"（時刻表示用に秒を落とす）
function fmtTplTime(s) {
    if (!s) return '';
    const parts = String(s).split(':');
    return (parts.length >= 2) ? `${parts[0]}:${parts[1]}` : String(s);
}

// jinjer 雛形選択プルダウンを作る。選択値（雛形ID）が CSV の各セルに入る。
// 初期選択: 凡例エントリの template_id > 自動マッチ結果 > （自動マッチ）
function buildTemplateSelect(entry, matched) {
    const sel = document.createElement('select');
    sel.className = 'legend-template-select';
    sel.dataset.field = 'template_id';

    const auto = document.createElement('option');
    auto.value = '';
    auto.textContent = '（自動マッチ）';
    sel.appendChild(auto);

    pendingTemplates.forEach(t => {
        const opt = document.createElement('option');
        opt.value = (t.id != null) ? String(t.id) : '';
        const times = (t.start || t.end) ? `（${fmtTplTime(t.start)}〜${fmtTplTime(t.end)}）` : '';
        opt.textContent = `${t.name || t.id}${times}`;
        sel.appendChild(opt);
    });

    let initial = (entry && entry.template_id != null && entry.template_id !== '')
        ? String(entry.template_id) : '';
    if (!initial && matched && matched.template_id != null) {
        initial = String(matched.template_id);
    }
    sel.value = (initial && pendingTemplates.some(t => String(t.id) === initial)) ? initial : '';
    return sel;
}

// 詳細行(日別シフト編集)から {date, code, comment} のリストを読み取る
function readShiftsFromDetail(detailRow) {
    const shifts = [];
    if (!detailRow) return shifts;
    detailRow.querySelectorAll('.legend-shift-row').forEach(rowEl => {
        const dInp = rowEl.querySelector('[data-shift-field="date"]');
        const cInp = rowEl.querySelector('[data-shift-field="code"]');
        const date = dInp ? dInp.value.trim() : '';
        const code = cInp ? (cInp.value || '').trim() : '';
        if (!date && !code) return;  // 空行は無視
        shifts.push({
            date: date,
            code: code,
            comment: rowEl.dataset.comment || null,
        });
    });
    return shifts;
}

function collectLegendFromUI() {
    const sheets = [];
    legendSheetsContainer.querySelectorAll('.legend-sheet').forEach(sheetEl => {
        const sheetIdx = parseInt(sheetEl.dataset.sheetIdx, 10);
        const original = pendingCodeSheets[sheetIdx];
        if (!original) return;

        const legend = [];
        // .legend-table の tbody のみを対象にする（従業員テーブルと混ざらないように）
        const legendTable = sheetEl.querySelector('.legend-table');
        if (legendTable) {
            legendTable.querySelectorAll('tbody tr').forEach(tr => {
                const obj = { code: '', label: '', start_time: '', end_time: '', break_minutes: 0, is_off: false, template_id: '' };
                tr.querySelectorAll('input').forEach(inp => {
                    const f = inp.dataset.field;
                    if (!f) return;
                    if (f === 'is_off') obj[f] = inp.checked;
                    else if (f === 'break_minutes') obj[f] = parseInt(inp.value, 10) || 0;
                    else obj[f] = inp.value.trim();
                });
                // 選択した jinjer 雛形ID（プルダウン）。空＝自動マッチ。
                const tplSel = tr.querySelector('select[data-field="template_id"]');
                if (tplSel) obj.template_id = tplSel.value || '';
                if (obj.code) legend.push(obj);
            });
        }

        // 従業員：UI で編集された氏名・日別シフトを採用する。
        // 日別シフト(日付＋記号)は詳細行(_detailRow)から読み、誤読・空欄の手修正を反映する。
        // 「＋ 従業員を追加」で UI から手動追加された行は original.employees に存在しないので新規扱い。
        const employees = [];
        const empTable = sheetEl.querySelector('.legend-employees-table');
        if (empTable) {
            empTable.querySelectorAll('tbody tr').forEach(tr => {
                if (tr.classList.contains('legend-employees-empty')) return;
                if (tr.classList.contains('legend-employee-detail')) return;  // 詳細行はスキップ
                const nameInp = tr.querySelector('input[data-field="employee_name"]');
                if (!nameInp) return;
                const editedName = nameInp.value.trim();
                const empIdx = parseInt(tr.dataset.empIdx, 10);
                const originalEmp = (original.employees || [])[empIdx];
                // 日別シフトは UI（詳細行）で編集された値を採用
                const shifts = readShiftsFromDetail(tr._detailRow);
                if (originalEmp) {
                    employees.push({
                        name: editedName || originalEmp.name || '不明',
                        shifts: shifts,
                    });
                } else {
                    // UI で新規追加された従業員（画像から抽出できなかったケースなど）
                    if (!editedName) return;
                    employees.push({
                        name: editedName,
                        shifts: shifts,
                    });
                }
            });
        }

        // 対象年月を取得
        const yearInp = sheetEl.querySelector('input[data-ym="year"]');
        const monthInp = sheetEl.querySelector('input[data-ym="month"]');
        const year = yearInp && yearInp.value ? parseInt(yearInp.value, 10) : original.year;
        const month = monthInp && monthInp.value ? parseInt(monthInp.value, 10) : original.month;

        sheets.push({
            filename: original.filename,
            legend: legend,
            off_markers: original.off_markers || [],
            employees: employees,
            year: year,
            month: month,
        });
    });
    return sheets;
}

legendConfirmBtn.addEventListener('click', async () => {
    const sheets = collectLegendFromUI();
    if (sheets.length === 0) {
        alert('凡例が空です');
        return;
    }

    // スケジュールアップロードモードでは年月必須
    if (pendingMode === 'csv_export') {
        for (const s of sheets) {
            if (!s.year || !s.month) {
                alert(`「${s.filename}」の対象年月を入力してください`);
                return;
            }
            if (!s.employees || s.employees.length === 0) {
                alert(`「${s.filename}」に従業員がいません。氏名を入力してください。`);
                return;
            }
            const blank = s.employees.filter(e => !e.name || e.name === '不明');
            if (blank.length > 0) {
                alert(`「${s.filename}」に氏名が未入力の従業員が ${blank.length} 名います。jinjer に登録されている氏名を入力してください。`);
                return;
            }
        }

        // jinjer と照合できていない氏名は従業員IDが空欄のまま "未分類" CSV に落ちる。
        // 実行前にここで止めて、やり直し（アップロードからの全部やり直し）を防ぐ。
        const unresolved = collectUnresolvedNames();
        if (unresolved.length > 0) {
            const ok = confirm(
                `jinjer の従業員IDが確定していない氏名が ${unresolved.length} 名います:\n\n`
                + unresolved.join(' / ')
                + '\n\nこのまま実行すると、この方たちは従業員ID空欄の「未分類」CSVに出力され、'
                + 'jinjer に取り込めません。\n'
                + '「キャンセル」で凡例画面に戻り、候補から選ぶかフルネームを入力してください。\n\n'
                + 'このまま実行しますか？'
            );
            if (!ok) return;
        }
    }

    const endpoint = (pendingMode === 'csv_export') ? '/export_jinjer_csv' : '/resolve_and_match';

    legendModal.style.display = 'none';
    startProcessing();
    progressMessage.textContent = (pendingMode === 'csv_export')
        ? 'jinjer インポート用 CSV を生成中...'
        : '凡例を解決して突合中...';

    try {
        const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'text/event-stream',
            },
            body: JSON.stringify({
                session_id: pendingSessionId,
                sheets: sheets,
            }),
        });

        if (!response.ok) {
            const data = await response.json();
            showError(data.errors ? data.errors.join('<br>') : 'エラーが発生しました');
            stopProcessing();
            return;
        }

        await consumeSSEResponse(response);

    } catch (err) {
        showError('通信エラーが発生しました: ' + err.message);
        stopProcessing();
    }
});

// =============================================================================
// スケジュールアップロードモード結果表示
// =============================================================================
function showCsvExportResult(data) {
    if (!csvExportArea) return;

    const { csv_files, missing_ids, merges, ake_auto, ake_schedule_priority,
            ake_conflicts, new_template_filename, new_template_count } = data;

    const list = document.getElementById('csv-files-list');
    list.innerHTML = '';
    (csv_files || []).forEach(f => {
        const item = document.createElement('div');
        item.className = 'csv-file-item';
        const groupTag = f.attendance_group_name
            ? `<span class="csv-file-group-tag">🕒 ${escapeHtml(f.attendance_group_name)}</span>`
            : '';
        item.innerHTML = `
            <div>
                <div class="csv-file-info">
                    📄 ${escapeHtml(f.filename)}
                    ${groupTag}
                </div>
                <div class="csv-file-meta">
                    対象: ${escapeHtml(f.source)} / ${f.year}年${f.month}月 / ${f.rows}人分
                </div>
            </div>
            <a href="/download/${encodeURIComponent(f.filename)}" class="csv-file-download">
                ⬇ ダウンロード
            </a>
        `;
        list.appendChild(item);
    });

    // 複数ファイルあるときのアナウンス
    const splitNote = document.getElementById('csv-split-note');
    if (splitNote) {
        if ((csv_files || []).length > 1) {
            splitNote.textContent = `📦 ${csv_files.length} ファイル に自動分割しました（画面インポート用の形式では別打刻グループの従業員を混在させると全行エラーになるため）。下の🚀「APIで直接投入」を使う場合はアップロード不要で、全ファイルまとめて投入できます。`;
            splitNote.style.display = 'block';
        } else {
            splitNote.style.display = 'none';
        }
    }

    const warnArea = document.getElementById('missing-ids-warn');
    const warnList = document.getElementById('missing-ids-list');
    if (missing_ids && missing_ids.length > 0) {
        warnList.textContent = missing_ids.join(' / ');
        warnArea.style.display = 'block';
    } else {
        warnArea.style.display = 'none';
    }

    // 夜勤明け（退勤30:00以降の翌日）を自動で「休み」にしたトレース表示
    const akeArea = document.getElementById('ake-area');
    const akeList = document.getElementById('ake-list');
    if (akeArea && akeList) {
        akeList.innerHTML = '';
        const applied = ake_auto || [];
        const schedPriority = ake_schedule_priority || [];
        const conflicts = ake_conflicts || [];
        if (applied.length > 0 || schedPriority.length > 0 || conflicts.length > 0) {
            applied.forEach(a => {
                const row = document.createElement('div');
                row.className = 'merge-item';
                const m = a.month || '';
                row.innerHTML = `
                    <div class="merge-name">👤 ${escapeHtml(a.name)}</div>
                    <div class="merge-detail">
                        ${escapeHtml(String(m))}/${a.night_day} が夜勤（退勤 ${Math.floor(a.end_time_minutes / 60)}:${String(a.end_time_minutes % 60).padStart(2, '0')}）
                        → ${escapeHtml(String(m))}/${a.ake_day} を
                        <span class="merge-cell-tag">休み</span>
                        <span class="ake-before">（元: ${escapeHtml(a.before || '空欄')}）</span>
                    </div>
                `;
                akeList.appendChild(row);
            });
            schedPriority.forEach(s => {
                const row = document.createElement('div');
                row.className = 'merge-item';
                const m = s.month || '';
                row.innerHTML = `
                    <div class="merge-name">👤 ${escapeHtml(s.name)}</div>
                    <div class="merge-detail">
                        ${escapeHtml(String(m))}/${s.night_day} が夜勤だが、
                        ${escapeHtml(String(m))}/${s.ake_day} にも予定
                        <span class="merge-cell-tag">${escapeHtml(s.next_value)}</span>
                        が入っているため <strong>シフト表を優先</strong>（明け休にしない）
                    </div>
                `;
                akeList.appendChild(row);
            });
            conflicts.forEach(c => {
                const row = document.createElement('div');
                row.className = 'merge-item ake-conflict';
                row.innerHTML = `
                    <div class="merge-name">⚠️ ${escapeHtml(c.name)}</div>
                    <div class="merge-detail">
                        ${escapeHtml(String(c.month || ''))}/${c.night_day} の夜勤明け — ${escapeHtml(c.reason)}
                    </div>
                `;
                akeList.appendChild(row);
            });
            akeArea.style.display = 'block';
        } else {
            akeArea.style.display = 'none';
        }
    }

    // 深夜跨ぎ統合のトレース表示
    const mergesArea = document.getElementById('merges-area');
    const mergesList = document.getElementById('merges-list');
    if (mergesArea && mergesList) {
        mergesList.innerHTML = '';
        if (merges && merges.length > 0) {
            merges.forEach(m => {
                const row = document.createElement('div');
                row.className = 'merge-item';
                const ymPrefix = (m.year && m.month) ? `${m.year}/${m.month}` : '';
                row.innerHTML = `
                    <div class="merge-name">👤 ${escapeHtml(m.name)}</div>
                    <div class="merge-detail">
                        ${escapeHtml(ymPrefix)}/${m.day_n}（${escapeHtml(m.code1)}＝${escapeHtml(m.label1)}）
                        ＋
                        ${escapeHtml(ymPrefix)}/${m.day_n_plus_1}（${escapeHtml(m.code2)}＝${escapeHtml(m.label2)}）
                        →
                        <strong>${escapeHtml(m.merged_start)}-${escapeHtml(m.merged_end)}</strong>
                        <span class="merge-cell-tag">${escapeHtml(m.cell_value)}</span>
                    </div>
                `;
                mergesList.appendChild(row);
            });
            mergesArea.style.display = 'block';
        } else {
            mergesArea.style.display = 'none';
        }
    }

    const tplMsg = document.getElementById('csv-new-template-msg');
    if (new_template_filename) {
        tplMsg.innerHTML =
            `✨ jinjer の既存雛形に該当しない記号が ${new_template_count} 件ありました。<br>` +
            `先に新規雛形CSVを jinjer に登録してから、月次スケジュールCSVをアップロードしてください: ` +
            `<a href="/download/${encodeURIComponent(new_template_filename)}" style="color:#4a148c; font-weight:bold; text-decoration:underline">` +
            `📥 ${escapeHtml(new_template_filename)}</a>`;
        tplMsg.style.display = 'block';
    } else {
        tplMsg.style.display = 'none';
    }

    // API直接投入ブロック（既定ルート）を対象月ごとに生成
    renderScheduleApiBlocks(csv_files || []);

    // 実行中に別モードへ切り替えられていても結果は保持し、そのモードに戻ったときに出す
    csvExportArea.dataset.hasResult = '1';
    if (getCurrentMode() === 'csv_export') {
        csvExportArea.style.display = 'block';
        csvExportArea.scrollIntoView({ behavior: 'smooth' });
    }
}

function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// =============================================================================
// スケジュールAPI直接投入（既定ルート）
//   ①差分プレビュー(dry-run) → fingerprint取得 → ②投入(execute)
//   投入ジョブはサーバー側でプランを再計算し、fingerprintが一致しないと中止する
// =============================================================================
function renderScheduleApiBlocks(csvFiles) {
    const area = document.getElementById('sched-api-area');
    const blocks = document.getElementById('sched-api-blocks');
    if (!area || !blocks) return;
    blocks.innerHTML = '';
    const byMonth = new Map();
    (csvFiles || []).forEach(f => {
        if (!f.filename || !f.year || !f.month) return;
        const key = `${f.year}-${String(f.month).padStart(2, '0')}`;
        if (!byMonth.has(key)) byMonth.set(key, []);
        byMonth.get(key).push(f.filename);
    });
    if (byMonth.size === 0) {
        area.style.display = 'none';
        return;
    }
    [...byMonth.keys()].sort().forEach(month => {
        blocks.appendChild(buildScheduleApiBlock(month, byMonth.get(month)));
    });
    area.style.display = 'block';
}

function buildScheduleApiBlock(month, filenames) {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'margin-top:10px; padding:10px; border:1px solid #cfe2ff; border-radius:6px; background:#fff';
    wrap.innerHTML = `
        <div style="font-weight:bold">対象月 ${escapeHtml(month)}（グリッド ${filenames.length} ファイル）</div>
        <div style="margin:8px 0">
            <button type="button" class="sched-api-preview-btn">① 差分プレビュー（書き込みなし）</button>
            <button type="button" class="sched-api-execute-btn" disabled>② この内容で投入</button>
            <span class="sched-api-status" style="margin-left:8px; font-size:13px"></span>
        </div>
        <pre class="sched-api-log" style="display:none; max-height:220px; overflow:auto; background:#f7f7f7; padding:8px; font-size:11px; margin:6px 0"></pre>
        <div class="sched-api-result" style="display:none; margin-top:8px"></div>
    `;
    const state = { month, filenames, fingerprint: '', planRows: 0, running: false };
    const previewBtn = wrap.querySelector('.sched-api-preview-btn');
    const executeBtn = wrap.querySelector('.sched-api-execute-btn');
    previewBtn.addEventListener('click', () => runScheduleApiImport(wrap, state, false));
    executeBtn.addEventListener('click', () => {
        if (!state.fingerprint) return;
        const msg = `書込 ${state.planRows} 行を jinjer 本番へ投入します。\n` +
            `対象月: ${state.month}\n\n差分プレビューの内容（特に「要手動確認」）を確認しましたか？`;
        if (!confirm(msg)) return;
        runScheduleApiImport(wrap, state, true);
    });
    return wrap;
}

async function runScheduleApiImport(wrap, state, execute) {
    if (state.running) return;
    state.running = true;
    const previewBtn = wrap.querySelector('.sched-api-preview-btn');
    const executeBtn = wrap.querySelector('.sched-api-execute-btn');
    const statusEl = wrap.querySelector('.sched-api-status');
    const logEl = wrap.querySelector('.sched-api-log');
    const resultEl = wrap.querySelector('.sched-api-result');
    previewBtn.disabled = true;
    executeBtn.disabled = true;
    if (!execute) state.fingerprint = '';  // プレビューやり直し → 前回の承認は無効化
    statusEl.textContent = execute ? '⏳ 投入中…（ステータス確認込みで数分かかります）'
                                   : '⏳ 差分プレビュー作成中…（jinjerの現状を取得しています）';
    logEl.style.display = 'block';
    logEl.textContent = '';
    resultEl.style.display = 'none';

    const fd = new FormData();
    fd.append('csv_filenames', JSON.stringify(state.filenames));
    fd.append('month', state.month);
    fd.append('execute', execute ? '1' : '0');
    if (execute) fd.append('fingerprint', state.fingerprint);

    let jobId;
    try {
        const res = await fetch('/schedule_api_import', { method: 'POST', body: fd });
        const j = await res.json();
        if (!j.success) throw new Error((j.errors || ['開始に失敗しました']).join(' / '));
        jobId = j.job_id;
    } catch (e) {
        statusEl.textContent = `❌ ${e.message}`;
        previewBtn.disabled = false;
        state.running = false;
        return;
    }

    const timer = setInterval(async () => {
        let s;
        try {
            const res = await fetch(`/api_import_status/${jobId}`);
            s = await res.json();
        } catch (e) {
            return;  // 一時的な取得失敗は次のポーリングへ
        }
        if (s.log && s.log.length) {
            logEl.textContent = s.log.join('\n');
            logEl.scrollTop = logEl.scrollHeight;
        }
        if (!s.done) return;
        clearInterval(timer);
        state.running = false;
        previewBtn.disabled = false;
        const r = s.result;
        if (!r) {
            statusEl.textContent = '❌ 処理に失敗しました（上のログを確認してください）';
            return;
        }
        if (r.dry_run) {
            if (s.ok && r.fingerprint && r.plan_rows > 0) {
                state.fingerprint = r.fingerprint;
                state.planRows = r.plan_rows;
                executeBtn.disabled = false;
                statusEl.textContent =
                    `✅ プレビュー完了: 書込 ${r.plan_rows} 行（一致 ${r.matched_rows} 日は書きません）` +
                    (r.manual_count ? ` / 要手動確認 ${r.manual_count} 件` : '');
            } else if (s.ok) {
                statusEl.textContent = '✅ 差分なし（すべてjinjerと一致しています）' +
                    (r.manual_count ? ` / 要手動確認 ${r.manual_count} 件` : '');
            } else {
                statusEl.textContent = '⚠️ 送信前チェックNGで中止しました（下の内容とレポートを確認）';
            }
            renderScheduleApiResult(resultEl, r);
        } else {
            statusEl.textContent = s.ok
                ? `✅ 投入完了: 反映OK ${r.verified_ok} / NG ${r.verified_ng}`
                : `⚠️ 投入結果に要確認あり（レポートの「要手動確認」を見てください）`;
            state.fingerprint = '';
            executeBtn.disabled = true;
            renderScheduleApiResult(resultEl, r);
        }
    }, 3000);
}

function renderScheduleApiResult(el, r) {
    const kubunColor = (k) => {
        if (!k) return '#666';
        if (k.indexOf('休日に予定残存') >= 0) return '#b45309';       // 手動削除が必要
        if (k.indexOf('削除しないで確認') >= 0) return '#1d4ed8';     // 半休の可能性
        if (k.indexOf('検証NG') >= 0 || k.indexOf('失敗') >= 0 || k.indexOf('中止') >= 0) return '#b91c1c';
        return '#666';
    };
    let html = '';
    if (r.report_url) {
        html += `<div style="margin-bottom:8px"><a href="${escapeHtml(r.report_url)}" ` +
            `style="font-weight:bold">📥 ${r.dry_run ? '承認用レポート（Excel）' : '投入結果レポート（Excel）'}をダウンロード</a></div>`;
    }
    if (r.manual && r.manual.length) {
        html += `<div style="font-weight:bold; margin:6px 0 2px">要手動確認（${r.manual_count} 件）` +
            `${r.manual_count > r.manual.length ? `（先頭 ${r.manual.length} 件を表示）` : ''}</div>`;
        html += '<table style="border-collapse:collapse; font-size:12px; width:100%">' +
            '<tr style="background:#eee"><th style="border:1px solid #ccc; padding:2px 6px">従業員</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">日付</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">区分</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">備考</th></tr>';
        r.manual.forEach(m => {
            html += `<tr>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap">${escapeHtml(m['従業員番号'] || '')} ${escapeHtml(m['氏名'] || '')}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap">${escapeHtml(m['日付'] || '')}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap; color:${kubunColor(m['区分'])}; font-weight:bold">${escapeHtml(m['区分'] || '')}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px">${escapeHtml(m['備考'] || '')}</td></tr>`;
        });
        html += '</table>';
    }
    if (r.dry_run && r.plan_preview && r.plan_preview.length) {
        const shown = r.plan_preview.length;
        html += `<div style="font-weight:bold; margin:10px 0 2px">書込予定 ${r.plan_rows} 行` +
            `${r.plan_rows > shown ? `（先頭 ${shown} 行を表示。全量はレポート参照）` : ''}</div>`;
        html += '<div style="max-height:300px; overflow:auto">' +
            '<table style="border-collapse:collapse; font-size:12px; width:100%">' +
            '<tr style="background:#eee"><th style="border:1px solid #ccc; padding:2px 6px">従業員</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">日付</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">区分</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">グリッド</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">新予定</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">新休憩</th>' +
            '<th style="border:1px solid #ccc; padding:2px 6px">現状</th></tr>';
        r.plan_preview.forEach(p => {
            html += `<tr>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap">${escapeHtml(p.emp)} ${escapeHtml(p.name)}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap">${escapeHtml(p.date)}(${escapeHtml(p.youbi)})</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px">${escapeHtml(p.kind)}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px">${escapeHtml(p.cell)}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap; font-weight:bold">${escapeHtml(p.new)}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px; white-space:nowrap">${escapeHtml(p.breaks)}</td>` +
                `<td style="border:1px solid #ccc; padding:2px 6px">${escapeHtml(p.cur)}</td></tr>`;
        });
        html += '</table></div>';
    }
    if (!r.dry_run) {
        html += `<div style="margin-top:6px">投入 ${r.submitted_rows} 行 / 反映OK ${r.verified_ok} / NG ${r.verified_ng}</div>`;
    }
    el.innerHTML = html;
    el.style.display = html ? 'block' : 'none';
}

// =============================================================================
// 経理モード — freee 取引インポート4CSVの生成
// =============================================================================

/** 検算・要確認・差分は Markdown で返ってくる。見出し/表/箇条書きだけ最低限HTMLにする。 */
function keiriRenderMarkdown(md) {
    if (!md) return '<p class="hint">（内容なし）</p>';
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const inline = (s) => esc(s)
        .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
        .replace(/`([^`]+)`/g, '<code>$1</code>');
    const lines = md.split('\n');
    const out = [];
    let table = null;
    const flushTable = () => {
        if (!table) return;
        // 2行目が区切り(|---|)なら1行目をヘッダーにする
        const sep = table[1] && /^\|[\s:|-]+$/.test(table[1]);
        const rows = sep ? table.filter((_, i) => i !== 1) : table;
        let html = '<table class="keiri-md-table">';
        rows.forEach((line, i) => {
            const cells = line.replace(/^\|/, '').replace(/\|$/, '').split('|');
            const tag = (sep && i === 0) ? 'th' : 'td';
            html += '<tr>' + cells.map(c => '<' + tag + '>' + inline(c.trim()) + '</' + tag + '>').join('') + '</tr>';
        });
        out.push(html + '</table>');
        table = null;
    };
    for (const raw of lines) {
        const line = raw.trimEnd();
        if (line.startsWith('|')) { (table = table || []).push(line); continue; }
        flushTable();
        if (!line.trim()) continue;
        const h = line.match(/^(#{1,4})\s+(.*)$/);
        if (h) {
            const lv = Math.min(h[1].length + 1, 5);
            out.push('<h' + lv + '>' + inline(h[2]) + '</h' + lv + '>');
            continue;
        }
        const li = line.match(/^(\s*)[-*]\s+(.*)$/);
        if (li) {
            // 字下げした箇条書き（例: 住民税の相殺の内訳）は入れ子として描く
            const sub = li[1].length >= 2;
            out.push('<div class="keiri-md-li' + (sub ? ' keiri-md-sub' : '') + '">'
                + (sub ? '－' : '・') + inline(li[2]) + '</div>');
            continue;
        }
        out.push('<p>' + inline(line) + '</p>');
    }
    flushTable();
    return out.join('');
}

function keiriShowError(msgs) {
    const el = document.getElementById('keiri-error-area');
    if (!el) return;
    const list = Array.isArray(msgs) ? msgs : [msgs];
    el.innerHTML = list.map(m => '<div>' + m + '</div>').join('');
    el.style.display = list.length ? 'block' : 'none';
}

function keiriRenderFiles(data) {
    const el = document.getElementById('keiri-files');
    if (!el) return;
    let html = '<table class="keiri-md-table"><tr><th>種別</th><th>ファイル</th><th>取引数</th><th>行数</th><th></th></tr>';
    for (const f of data.files || []) {
        const url = '/keiri_download/' + data.ym + '/' + encodeURIComponent(f.filename);
        html += '<tr><td>' + f['種別'] + '</td><td>' + f.filename + '</td><td>' + f['取引数'] + '</td><td>' + f['行数'] + '</td>'
              + '<td><a class="btn btn-download" style="padding:2px 8px; font-size:11px" href="' + url + '">📥</a></td></tr>';
    }
    html += '</table>';
    html += '<div class="hint" style="margin-top:4px">出力先: <code>' + data.out_dir + '</code>';
    if (data.keihi_book) html += '<br>経費転記の材料: <code>' + data.keihi_book + '</code>';
    else html += '<br>⚠️ 経費利用履歴のブックが見つからず、経費転記は分解していません（要確認リストを見てください）';
    html += '</div>';
    el.innerHTML = html;
}

function keiriRenderDiff(data) {
    const area = document.getElementById('keiri-diff-area');
    const el = document.getElementById('keiri-diff-summary');
    if (!area || !el) return;
    if (data.diff_error) {
        el.innerHTML = '<div class="hint">' + data.diff_error + '</div>';
        area.style.display = 'block';
        return;
    }
    if (!data.diff_summary || !data.diff_summary.length) { area.style.display = 'none'; return; }
    let html = '<table class="keiri-md-table"><tr><th>種別</th><th>一致</th><th>金額差</th>'
             + '<th>生成のみ</th><th>最終のみ</th><th>生成だけの人</th><th>最終だけの人</th></tr>';
    for (const r of data.diff_summary) {
        if (r['状態'] !== '突合済み') {
            html += '<tr><td>' + r['種別'] + '</td><td colspan="6" class="hint">未突合（最終CSVがまだありません）</td></tr>';
            continue;
        }
        const bad = (n) => n ? '<td style="color:#c00; font-weight:bold">' + n + '</td>' : '<td>0</td>';
        html += '<tr><td>' + r['種別'] + '</td><td>' + r['一致'] + '</td>' + bad(r['金額差']) + bad(r['生成のみ'])
              + bad(r['最終のみ']) + bad(r['生成だけの人']) + bad(r['最終だけの人']) + '</tr>';
    }
    el.innerHTML = html + '</table>';
    area.style.display = 'block';
}

document.querySelectorAll('.keiri-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.keiri-pane').forEach(p => { p.style.display = 'none'; });
        document.querySelectorAll('.keiri-tab').forEach(b => b.classList.remove('active'));
        const pane = document.getElementById(btn.dataset.target);
        if (pane) pane.style.display = 'block';
        btn.classList.add('active');
    });
});

/** 「その他」の保留者を入力フォームで描く（無ければ枠ごと隠す）。 */
function keiriRenderSonota(data) {
    const area = document.getElementById('keiri-sonota-area');
    const rows = document.getElementById('keiri-sonota-rows');
    if (!area || !rows) return;
    const pending = data.sonota_pending || [];
    if (!pending.length) { area.style.display = 'none'; rows.innerHTML = ''; return; }

    const choices = data.sonota_choices || [];
    const dl = document.getElementById('keiri-sonota-choices');
    if (dl) {
        // 「勘定科目｜品目｜税区分」の1本の候補にする（3つセットで選べば取り違えが起きない）
        dl.innerHTML = choices.map(c =>
            '<option value="' + mailEsc(c['勘定科目'] + '｜' + c['品目'] + '｜' + c['税区分']) + '">').join('');
    }
    const pathEl = document.getElementById('keiri-sonota-path');
    if (pathEl) pathEl.textContent = '台帳: ' + (data.sonota_manual_csv || '');

    let html = '<table class="keiri-md-table"><tr><th>社員番号</th><th>氏名</th><th>金額</th>'
             + '<th>部門</th><th>勘定科目｜品目｜税区分</th><th>備考</th></tr>';
    pending.forEach((p, i) => {
        html += '<tr data-emp="' + mailEsc(p['社員番号']) + '">'
              + '<td>' + mailEsc(p['社員番号']) + '</td>'
              + '<td>' + mailEsc(p['氏名']) + '</td>'
              + '<td style="text-align:right">' + Number(p['金額']).toLocaleString() + '</td>'
              + '<td>' + mailEsc(p['部門']) + '</td>'
              + '<td><input type="text" class="keiri-sonota-combo" list="keiri-sonota-choices"'
              + ' data-idx="' + i + '" placeholder="選ぶか直接入力" style="width:280px"></td>'
              + '<td><input type="text" class="keiri-sonota-biko" data-idx="' + i + '"'
              + ' placeholder="例: 有給残6日買取分" style="width:240px"></td></tr>';
    });
    html += '</table>';
    rows.innerHTML = html;
    rows.dataset.pending = JSON.stringify(pending);
    area.style.display = 'block';
}

/** 入力欄 → 台帳へ送る行。空欄の人は送らない（＝これまでどおり保留のまま）。 */
function keiriCollectSonota() {
    const rows = document.getElementById('keiri-sonota-rows');
    if (!rows || !rows.dataset.pending) return { entries: [], errors: [] };
    const pending = JSON.parse(rows.dataset.pending);
    const entries = [], errors = [];
    rows.querySelectorAll('.keiri-sonota-combo').forEach(inp => {
        const combo = (inp.value || '').trim();
        if (!combo) return;
        const p = pending[Number(inp.dataset.idx)];
        const parts = combo.split(/[｜|]/).map(s => s.trim());
        if (parts.length !== 3 || parts.some(s => !s)) {
            errors.push(p['社員番号'] + ' ' + p['氏名']
                + '：「勘定科目｜品目｜税区分」の3つを ｜ 区切りで入れてください（入力: ' + combo + '）');
            return;
        }
        const biko = rows.querySelector('.keiri-sonota-biko[data-idx="' + inp.dataset.idx + '"]');
        entries.push({
            '社員番号': p['社員番号'], '氏名': p['氏名'], '金額': p['金額'],
            '勘定科目': parts[0], '品目': parts[1], '税区分': parts[2],
            '備考': biko ? (biko.value || '').trim() : '',
        });
    });
    return { entries, errors };
}

const keiriRunBtn = document.getElementById('keiri-run-btn');
if (keiriRunBtn) {
    const keiriRun = async () => {
        const status = document.getElementById('keiri-status');
        const month = (document.getElementById('keiri-month').value || '').trim();
        keiriShowError([]);
        if (!/^\d{4}-\d{2}$/.test(month)) {
            keiriShowError('支給月は YYYY-MM 形式で入力してください（例: 2026-08）');
            return false;
        }
        const fd = new FormData();
        fd.append('month', month);
        fd.append('min_status', document.getElementById('keiri-min-status').value);
        fd.append('master_csv', document.getElementById('keiri-master-csv').value);
        fd.append('keihi_mapping_csv', document.getElementById('keiri-keihi-mapping').value);
        fd.append('keihi_book', document.getElementById('keiri-keihi-book').value);
        fd.append('run_diff', document.getElementById('keiri-run-diff').checked ? '1' : '0');
        fd.append('refresh_statements', document.getElementById('keiri-refresh-statements').checked ? '1' : '0');
        fd.append('refresh_custom', document.getElementById('keiri-refresh-custom').checked ? '1' : '0');

        keiriRunBtn.disabled = true;
        status.textContent = '生成中…（給与明細の取得に数分かかることがあります）';
        document.getElementById('keiri-result-area').style.display = 'none';
        try {
            const res = await fetch('/keiri_run', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) {
                keiriShowError(data.errors || ['生成に失敗しました']);
                status.textContent = '';
                return false;
            }
            document.getElementById('keiri-cnt-emp').textContent = data.employees;
            document.getElementById('keiri-paid-on').textContent = data.paid_on || '—';
            const alerts = data.alerts || {};
            document.getElementById('keiri-cnt-alert').textContent =
                (alerts['備考の手入力が必要'] || 0) + (alerts['未収入金の候補'] || 0)
                + (alerts['経費転記で保留'] || 0) + (alerts['部門未知値'] || 0);
            document.getElementById('keiri-cnt-keihi').textContent = alerts['経費転記の分解'] || 0;
            keiriRenderFiles(data);
            keiriRenderDiff(data);
            keiriRenderSonota(data);
            document.getElementById('keiri-pane-yokakunin').innerHTML = keiriRenderMarkdown(data.yokakunin_md);
            document.getElementById('keiri-pane-kensan').innerHTML = keiriRenderMarkdown(data.kensan_md);
            document.getElementById('keiri-pane-diff').innerHTML = keiriRenderMarkdown(data.diff_md);
            document.getElementById('keiri-result-area').style.display = 'block';
            status.textContent = '完了（' + data.month + '）';
            return true;
        } catch (e) {
            keiriShowError('通信に失敗しました: ' + e);
            status.textContent = '';
            return false;
        } finally {
            keiriRunBtn.disabled = false;
        }
    };
    keiriRunBtn.addEventListener('click', keiriRun);

    const sonotaSaveBtn = document.getElementById('keiri-sonota-save-btn');
    if (sonotaSaveBtn) {
        sonotaSaveBtn.addEventListener('click', async () => {
            const status = document.getElementById('keiri-sonota-status');
            const month = (document.getElementById('keiri-month').value || '').trim();
            keiriShowError([]);
            const { entries, errors } = keiriCollectSonota();
            if (errors.length) { keiriShowError(errors); return; }
            if (!entries.length) {
                keiriShowError('科目が1件も入力されていません');
                return;
            }
            sonotaSaveBtn.disabled = true;
            status.textContent = '台帳に保存中…';
            try {
                const res = await fetch('/keiri_sonota_save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ month, entries }),
                });
                const data = await res.json();
                if (!data.success) {
                    keiriShowError(data.errors || ['台帳の保存に失敗しました']);
                    status.textContent = '';
                    return;
                }
                status.textContent = data.saved + '件を台帳に保存しました。仕訳を作り直しています…';
                // 台帳を読み直させるため、同じ条件でそのまま再生成する
                const ok = await keiriRun();
                status.textContent = ok ? data.saved + '件を反映しました' : '';
            } catch (e) {
                keiriShowError('通信に失敗しました: ' + e);
                status.textContent = '';
            } finally {
                sonotaSaveBtn.disabled = false;
            }
        });
    }
}

// =============================================================================
// メール下書きモード — 一覧表×テンプレート → Outlook下書き（送信機能なし）
// =============================================================================

// （状態変数 mailPlans / mailTemplates / mailTemplatesLoaded / mailDefaultCc は
//   ファイル先頭の「グローバル状態」で宣言。起動時モードがメールでも参照できるようにするため）

function mailEsc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function mailShowError(msgs) {
    const el = document.getElementById('mail-error-area');
    if (!el) return;
    const list = Array.isArray(msgs) ? msgs : [msgs];
    el.innerHTML = list.map(m => '<div>' + mailEsc(m) + '</div>').join('');
    el.style.display = list.length ? 'block' : 'none';
}

async function loadMailTemplates(force) {
    if (mailTemplatesLoaded && !force) return;
    const select = document.getElementById('mail-template-select');
    if (!select) return;
    try {
        const res = await fetch('/mail_templates');
        const data = await res.json();
        if (!data.success) { mailShowError(data.errors || []); return; }
        mailTemplates = data.templates || [];
        mailDefaultCc = data.default_cc || '';
        const ccField = document.getElementById('mail-cc');
        if (ccField && !ccField.value && mailDefaultCc) ccField.value = mailDefaultCc;
        const current = select.value;
        select.innerHTML = '<option value="">（選択すると下の欄に読み込みます）</option>'
            + mailTemplates.map(t => '<option value="' + mailEsc(t.name) + '">' + mailEsc(t.name) + '</option>').join('');
        if (current && mailTemplates.some(t => t.name === current)) select.value = current;
        mailTemplatesLoaded = true;
    } catch (e) {
        mailShowError('テンプレート一覧の取得に失敗しました: ' + e);
    }
}

const mailTemplateSelect = document.getElementById('mail-template-select');
if (mailTemplateSelect) {
    mailTemplateSelect.addEventListener('change', () => {
        const tpl = mailTemplates.find(t => t.name === mailTemplateSelect.value);
        if (!tpl) return;
        document.getElementById('mail-template-name').value = tpl.name;
        document.getElementById('mail-subject').value = tpl.subject || '';
        document.getElementById('mail-body').value = tpl.body || '';
        document.getElementById('mail-cc').value = tpl.cc || mailDefaultCc;
        document.getElementById('mail-bcc-mode').value = (tpl.bcc_mode === 'bcc') ? 'bcc' : 'to_only';
        document.getElementById('mail-importance').value = tpl.importance || 'normal';
    });
}

const mailTemplateSaveBtn = document.getElementById('mail-template-save-btn');
if (mailTemplateSaveBtn) {
    mailTemplateSaveBtn.addEventListener('click', async () => {
        mailShowError([]);
        const name = (document.getElementById('mail-template-name').value || '').trim();
        if (!name) { mailShowError('テンプレート名を入力してください'); return; }
        const fd = new FormData();
        fd.append('name', name);
        fd.append('subject', document.getElementById('mail-subject').value);
        fd.append('body', document.getElementById('mail-body').value);
        fd.append('cc', document.getElementById('mail-cc').value);
        fd.append('bcc_mode', document.getElementById('mail-bcc-mode').value);
        fd.append('importance', document.getElementById('mail-importance').value);
        try {
            const res = await fetch('/mail_templates_save', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['保存に失敗しました']); return; }
            await loadMailTemplates(true);
            document.getElementById('mail-template-select').value = name;
            document.getElementById('mail-status').textContent = 'テンプレート「' + name + '」を保存しました';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
        }
    });
}

const mailTemplateDeleteBtn = document.getElementById('mail-template-delete-btn');
if (mailTemplateDeleteBtn) {
    mailTemplateDeleteBtn.addEventListener('click', async () => {
        mailShowError([]);
        const select = document.getElementById('mail-template-select');
        const name = select ? select.value : '';
        if (!name) { mailShowError('削除するテンプレートをプルダウンで選んでください'); return; }
        if (!window.confirm('テンプレート「' + name + '」を削除します。よろしいですか？')) return;
        const fd = new FormData();
        fd.append('name', name);
        fd.append('delete', '1');
        try {
            const res = await fetch('/mail_templates_save', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['削除に失敗しました']); return; }
            await loadMailTemplates(true);
            document.getElementById('mail-status').textContent = 'テンプレート「' + name + '」を削除しました';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
        }
    });
}

function mailFormData() {
    const fd = new FormData();
    fd.append('table_path', document.getElementById('mail-table-path').value);
    fd.append('address_book', document.getElementById('mail-address-book').value);
    fd.append('subject', document.getElementById('mail-subject').value);
    fd.append('body', document.getElementById('mail-body').value);
    fd.append('cc', document.getElementById('mail-cc').value);
    fd.append('bcc_mode', document.getElementById('mail-bcc-mode').value);
    fd.append('importance', document.getElementById('mail-importance').value);
    return fd;
}

function mailSelectedIds() {
    return Array.from(document.querySelectorAll('.mail-plan-check:checked')).map(cb => cb.dataset.id);
}

function mailUpdateDraftsButton() {
    const btn = document.getElementById('mail-drafts-btn');
    if (!btn) return;
    const count = mailSelectedIds().length;
    btn.disabled = count === 0;
    btn.textContent = '✉️ 選択した ' + count + ' 件の下書きをOutlookに作成';
}

function mailRenderPlans(data) {
    const hint = document.getElementById('mail-columns-hint');
    if (hint) {
        hint.innerHTML = '差し込みに使える列: ' + (data.columns || [])
            .map(c => '<code>{{' + mailEsc(c) + '}}</code>').join(' ');
    }
    let html = '<table class="keiri-md-table"><tr><th>作成</th><th>社員番号</th><th>氏名</th>'
             + '<th>宛先</th><th>件名</th><th>状態</th><th>本文</th></tr>';
    for (const p of mailPlans) {
        const ok = p.status === 'OK';
        const check = ok
            ? '<input type="checkbox" class="mail-plan-check" data-id="' + mailEsc(p.employee_id) + '" checked>'
            : '—';
        const to = p.to.length ? mailEsc(p.to.join('; ')) : '<span style="color:#c00">なし</span>';
        const bcc = p.bcc.length ? '<br><span class="hint">BCC: ' + mailEsc(p.bcc.join('; ')) + '</span>' : '';
        const breakdown = p.breakdown ? '<br><span class="hint">' + mailEsc(p.breakdown) + '</span>' : '';
        const status = ok
            ? '<span style="color:#2e7d32">OK</span>'
            : '<span style="color:#c00; font-weight:bold">要確認</span>';
        const issues = (p.issues || []).length
            ? '<br><span class="hint" style="color:#c00">' + p.issues.map(mailEsc).join('<br>') + '</span>' : '';
        html += '<tr' + (ok ? '' : ' style="background:#fdecec"') + '>'
            + '<td style="text-align:center">' + check + '</td>'
            + '<td>' + mailEsc(p.employee_id) + '</td>'
            + '<td>' + mailEsc(p.name) + '</td>'
            + '<td>' + to + bcc + breakdown + '</td>'
            + '<td>' + mailEsc(p.subject) + '</td>'
            + '<td>' + status + issues + '</td>'
            + '<td><details><summary style="cursor:pointer">本文</summary><pre style="white-space:pre-wrap; font-size:11px; margin:4px 0">'
            + mailEsc(p.body) + '</pre></details></td></tr>';
    }
    html += '</table>';
    const el = document.getElementById('mail-plans');
    el.innerHTML = html;
    el.querySelectorAll('.mail-plan-check').forEach(cb => cb.addEventListener('change', mailUpdateDraftsButton));
    mailUpdateDraftsButton();
}

const mailPreviewBtn = document.getElementById('mail-preview-btn');
if (mailPreviewBtn) {
    mailPreviewBtn.addEventListener('click', async () => {
        const status = document.getElementById('mail-status');
        mailShowError([]);
        document.getElementById('mail-drafts-result').style.display = 'none';
        mailPreviewBtn.disabled = true;
        status.textContent = 'プレビューを作成中…';
        try {
            const res = await fetch('/mail_preview', { method: 'POST', body: mailFormData() });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['プレビューに失敗しました']); status.textContent = ''; return; }
            mailPlans = data.plans || [];
            document.getElementById('mail-cnt-total').textContent = data.counts.total;
            document.getElementById('mail-cnt-ok').textContent = data.counts.ok;
            document.getElementById('mail-cnt-warn').textContent = data.counts.warn;
            mailRenderPlans(data);
            document.getElementById('mail-result-area').style.display = 'block';
            status.textContent = 'プレビューを確認して、作成する人にチェックを入れてください（まだ何も作られていません）';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            mailPreviewBtn.disabled = false;
        }
    });
}

const mailDraftsBtn = document.getElementById('mail-drafts-btn');
if (mailDraftsBtn) {
    mailDraftsBtn.addEventListener('click', async () => {
        const status = document.getElementById('mail-drafts-status');
        mailShowError([]);
        const ids = mailSelectedIds();
        if (!ids.length) return;
        if (!window.confirm(ids.length + '件の下書きをOutlookに作成します。よろしいですか？\n（送信はされません。送信はOutlookで1通ずつ確認してから行ってください）')) return;
        const fd = mailFormData();
        fd.append('selected_ids', JSON.stringify(ids));
        mailDraftsBtn.disabled = true;
        status.textContent = 'Outlookに下書きを作成中…';
        try {
            const res = await fetch('/mail_drafts', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['下書き作成に失敗しました']); status.textContent = ''; return; }
            const failed = (data.results || []).filter(r => r.result !== '下書き作成済');
            let html = '✅ 下書き作成 ' + data.processed + '件 / スキップ ' + data.skipped + '件 / 失敗 ' + data.failed + '件';
            html += '<br>ログ: <code>' + mailEsc(data.log_path) + '</code>';
            if (failed.length) {
                html += '<br><span style="color:#c00">' + failed.map(r =>
                    mailEsc(r.employee_id + ' ' + r.name + ': ' + r.result)).join('<br>') + '</span>';
            }
            html += '<br><b>Outlook の「下書き」フォルダを開いて、内容を確認してから送信してください。</b>';
            const el = document.getElementById('mail-drafts-result');
            el.innerHTML = html;
            el.style.display = 'block';
            status.textContent = '完了';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            mailDraftsBtn.disabled = false;
            mailUpdateDraftsButton();
        }
    });
}

// --- メール台帳の更新（jinjerと同期。確認→反映の2段階） ---

const mailLedgerDiffBtn = document.getElementById('mail-ledger-diff-btn');
if (mailLedgerDiffBtn) {
    mailLedgerDiffBtn.addEventListener('click', async () => {
        const status = document.getElementById('mail-ledger-status');
        mailShowError([]);
        mailLedgerDiffBtn.disabled = true;
        status.textContent = 'jinjerから最新の従業員情報を取得中…（数十秒かかります）';
        document.getElementById('mail-ledger-apply-result').style.display = 'none';
        try {
            const fd = new FormData();
            fd.append('address_book', document.getElementById('mail-address-book').value);
            const res = await fetch('/mail_ledger_diff', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['差分の取得に失敗しました']); status.textContent = ''; return; }
            mailLedgerRender(data);
            status.textContent = '差分を確認して、反映する行にチェックを入れてください（まだ台帳は書き換えていません）';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            mailLedgerDiffBtn.disabled = false;
        }
    });
}

function mailLedgerSelected(cls) {
    return Array.from(document.querySelectorAll('.' + cls + ':checked')).map(cb => cb.dataset.id);
}

function mailLedgerUpdateApplyButton() {
    const btn = document.getElementById('mail-ledger-apply-btn');
    if (!btn) return;
    const n = mailLedgerSelected('ledger-add-check').length + mailLedgerSelected('ledger-del-check').length;
    btn.disabled = n === 0;
    btn.textContent = '📥 チェックした ' + n + ' 件を台帳に反映する';
}

function mailLedgerRender(data) {
    const adds = data.additions || [];
    const dels = data.retirees || [];
    const summary = '追加候補 <b>' + adds.length + '人</b> / 退職の削除候補 <b>' + dels.length + '人</b>'
        + ' / アドレス不一致 ' + (data.mismatches || []).length + '件'
        + ' / 台帳にあるがjinjerに無い番号 ' + (data.missing_in_jinjer || []).length + '件';
    document.getElementById('mail-ledger-summary').innerHTML = summary;
    let html = '';
    if (adds.length) {
        html += '<div style="font-weight:600; margin:6px 0 2px">追加候補（新入社員）</div>'
            + '<table class="keiri-md-table"><tr><th>反映</th><th>社員番号</th><th>氏名</th><th>社用(D列)</th><th>個人(F列)</th></tr>';
        for (const a of adds) {
            const warn = a.no_email ? ' <span style="color:#c00">jinjerにメール未登録</span>' : '';
            html += '<tr><td style="text-align:center"><input type="checkbox" class="ledger-add-check" data-id="' + mailEsc(a.id) + '" checked></td>'
                + '<td>' + mailEsc(a.id) + '</td><td>' + mailEsc(a.name) + warn + '</td>'
                + '<td>' + mailEsc(a.company_email || '—') + '</td><td>' + mailEsc(a.personal_email || '—') + '</td></tr>';
        }
        html += '</table>';
    }
    if (dels.length) {
        html += '<div style="font-weight:600; margin:10px 0 2px; color:#c00000">削除候補（jinjerで退職）</div>'
            + '<table class="keiri-md-table"><tr><th>反映</th><th>社員番号</th><th>氏名</th><th>退職日</th></tr>';
        for (const d of dels) {
            html += '<tr><td style="text-align:center"><input type="checkbox" class="ledger-del-check" data-id="' + mailEsc(d.id) + '" checked></td>'
                + '<td>' + mailEsc(d.id) + '</td><td>' + mailEsc(d.name) + '</td><td>' + mailEsc(d.retirement_date || '') + '</td></tr>';
        }
        html += '</table>';
    }
    if ((data.mismatches || []).length) {
        html += '<div style="font-weight:600; margin:10px 0 2px">アドレス不一致（報告のみ・台帳は変更しません）</div>'
            + data.mismatches.map(m => '<div class="hint">' + mailEsc(m) + '</div>').join('');
    }
    if ((data.missing_in_jinjer || []).length) {
        html += '<div style="font-weight:600; margin:10px 0 2px">台帳にあるがjinjerに見つからない番号（報告のみ・要確認）</div>'
            + data.missing_in_jinjer.map(m => '<div class="hint">' + mailEsc(m.id + ' ' + m.name) + '</div>').join('');
    }
    if (!adds.length && !dels.length) {
        html += '<div class="hint">追加・削除はありません。台帳はjinjerと同期できています。</div>';
    }
    const tables = document.getElementById('mail-ledger-tables');
    tables.innerHTML = html;
    tables.querySelectorAll('.ledger-add-check, .ledger-del-check')
        .forEach(cb => cb.addEventListener('change', mailLedgerUpdateApplyButton));
    document.getElementById('mail-ledger-result').style.display = 'block';
    mailLedgerUpdateApplyButton();
}

const mailLedgerApplyBtn = document.getElementById('mail-ledger-apply-btn');
if (mailLedgerApplyBtn) {
    mailLedgerApplyBtn.addEventListener('click', async () => {
        const status = document.getElementById('mail-ledger-apply-status');
        mailShowError([]);
        const addIds = mailLedgerSelected('ledger-add-check');
        const delIds = mailLedgerSelected('ledger-del-check');
        if (!addIds.length && !delIds.length) return;
        const msg = '台帳を更新します。\n追加 ' + addIds.length + '人 / 削除 ' + delIds.length + '人\n'
            + '実行前に台帳のバックアップを自動作成します。よろしいですか？';
        if (!window.confirm(msg)) return;
        const fd = new FormData();
        fd.append('address_book', document.getElementById('mail-address-book').value);
        fd.append('add_ids', JSON.stringify(addIds));
        fd.append('delete_ids', JSON.stringify(delIds));
        mailLedgerApplyBtn.disabled = true;
        status.textContent = '台帳を更新中…（バックアップ→Excel書き込み）';
        try {
            const res = await fetch('/mail_ledger_apply', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) { mailShowError(data.errors || ['台帳の更新に失敗しました']); status.textContent = ''; return; }
            const el = document.getElementById('mail-ledger-apply-result');
            el.innerHTML = '✅ 追加 ' + data.added + '人 / 削除 ' + data.deleted + '人 を反映しました'
                + '<br>バックアップ: <code>' + mailEsc(data.backup_path) + '</code>'
                + '<br>ログ: <code>' + mailEsc(data.log_path) + '</code>'
                + '<br><b>もう一度「👀 プレビュー」を押すと、新しい台帳で宛先が突合されます。</b>';
            el.style.display = 'block';
            status.textContent = '完了';
        } catch (e) {
            mailShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            mailLedgerUpdateApplyButton();
        }
    });
}


// =============================================================================
// 健康診断HPMモード（整形済Excel → HPM取込用302列CSV）
// =============================================================================
// 画面に平均血圧の欄は作らない。1回目・2回目の枠しか持たせないことで、
// 「平均を出しておいて」と言われても構造的に出せないようにしておく。

let hhPreview = null;

function hhShowError(messages) {
    const el = document.getElementById('hh-error-area');
    if (!el) return;
    const list = Array.isArray(messages) ? messages : [messages];
    el.innerHTML = list.map(escapeHtml).join('<br>');
    el.style.display = list.length ? 'block' : 'none';
}

function hhClearError() {
    const el = document.getElementById('hh-error-area');
    if (el) { el.innerHTML = ''; el.style.display = 'none'; }
}

function hhIssueHtml(issue) {
    const cls = issue.level === 'error' ? 'hh-issue-error' : 'hh-issue-warning';
    const mark = issue.level === 'error' ? '⛔' : '⚠️';
    return '<div class="hh-issue ' + cls + '">' + mark + ' ' + escapeHtml(issue.message) + '</div>';
}

function hhBpCell(value) {
    return value ? escapeHtml(value) : '<span class="hh-blank">—</span>';
}

function hhInstitutionOptions(master, selected) {
    let html = '<option value="">（選んでください）</option>';
    (master.institutions || []).forEach(inst => {
        const label = inst.name + (inst.hpm_confirmed ? '' : '（HPM未確認・使えません）');
        html += '<option value="' + escapeHtml(inst.name) + '"'
            + (inst.name === selected ? ' selected' : '')
            + (inst.hpm_confirmed ? '' : ' disabled')
            + '>' + escapeHtml(label) + '</option>';
    });
    return html;
}

function hhCourseOptions(master, institutionName, selected) {
    const inst = (master.institutions || []).find(i => i.name === institutionName);
    let html = '<option value="">（選んでください）</option>';
    if (!inst) return html;
    (inst.courses || []).forEach(c => {
        html += '<option value="' + escapeHtml(c.hpm_value) + '"'
            + (c.hpm_value === selected ? ' selected' : '')
            + '>' + escapeHtml(c.display_name) + '</option>';
    });
    return html;
}

function hhJinjerBlock(person, roster) {
    const j = person.jinjer || {};
    if (j.status === 'ok' && j.employee) {
        const e = j.employee;
        return '<div class="hh-box"><div class="hh-box-title">jinjer 社員（自動一致）</div>'
            + '<div style="font-size:13px">' + escapeHtml(e.employee_id) + '　' + escapeHtml(e.name) + '</div>'
            + '<div style="font-size:11px; color:#62707c">' + escapeHtml(e.kana)
            + '／' + escapeHtml(e.birth_date) + '／' + escapeHtml(e.gender) + '</div>'
            + '<input type="hidden" class="hh-emp" data-key="' + escapeHtml(person.key) + '" value="'
            + escapeHtml(e.employee_id) + '"></div>';
    }
    // 候補 → 全員 の順で並べ、人に選んでもらう（同姓同名を自動で決めない）
    const seen = {};
    const list = [];
    (j.candidates || []).forEach(c => { if (!seen[c.employee_id]) { seen[c.employee_id] = 1; list.push(c); } });
    (roster || []).forEach(c => { if (!seen[c.employee_id]) { seen[c.employee_id] = 1; list.push(c); } });
    let options = '<option value="">（社員を選んでください）</option>';
    list.forEach(c => {
        options += '<option value="' + escapeHtml(c.employee_id) + '">'
            + escapeHtml(c.employee_id + '　' + c.name + '（' + (c.birth_date || '生年月日なし')
            + '・' + (c.gender || '性別なし') + '）') + '</option>';
    });
    const reasons = (j.reasons || []).map(r => '<div class="hh-reason">' + escapeHtml(r) + '</div>').join('');
    return '<div class="hh-box"><div class="hh-box-title">jinjer 社員（要選択）</div>'
        + '<select class="hh-select hh-emp" data-key="' + escapeHtml(person.key) + '">' + options + '</select>'
        + reasons + '</div>';
}

function hhRenderPerson(person, master, roster) {
    const bp = person.blood_pressure || { r1: {}, r2: {} };
    const hasError = (person.issues || []).some(i => i.level === 'error');

    let html = '<div class="hh-person' + (hasError ? ' has-error' : '') + '">';
    html += '<div class="hh-person-head">'
        + '<span class="hh-person-name">' + escapeHtml(person.name) + '</span>'
        + '<span style="font-size:12px; color:#62707c">受診日 ' + escapeHtml(person.exam_date)
        + '／受診No. ' + escapeHtml(person.exam_no)
        + '／数値 ' + person.numeric_count + '項目・定性 ' + person.qualitative_count + '項目</span>'
        + '</div>';

    html += '<div class="hh-grid">';
    html += hhJinjerBlock(person, roster);

    html += '<div class="hh-box"><div class="hh-box-title">健診機関・健診種別</div>'
        + '<select class="hh-select hh-inst" data-key="' + escapeHtml(person.key) + '" style="width:100%; margin-bottom:4px">'
        + hhInstitutionOptions(master, '') + '</select>'
        + '<select class="hh-select hh-course" data-key="' + escapeHtml(person.key) + '" style="width:100%">'
        + hhCourseOptions(master, '', '') + '</select></div>';

    html += '<div class="hh-box"><div class="hh-box-title">血圧（原票どおり・平均は作りません）</div>'
        + '<table class="hh-bp"><tr><th></th><th>収縮期</th><th>拡張期</th></tr>'
        + '<tr><th>1回目</th><td>' + hhBpCell(bp.r1 && bp.r1.sys) + '</td><td>' + hhBpCell(bp.r1 && bp.r1.dia) + '</td></tr>'
        + '<tr><th>2回目</th><td>' + hhBpCell(bp.r2 && bp.r2.sys) + '</td><td>' + hhBpCell(bp.r2 && bp.r2.dia) + '</td></tr>'
        + '</table>'
        + (bp.r3_present ? '<div class="hh-reason">3回目以降がありますがHPMには出力しません</div>' : '')
        + '</div>';

    if ((person.qualitative || []).length) {
        const items = person.qualitative.map(q => {
            const where = q.hpm_col !== null && q.hpm_col !== undefined
                ? '→ 列' + q.hpm_col + '（' + escapeHtml(q.col_name) + '）'
                : '<span style="color:#a05a00">' + escapeHtml(q.note || '') + '</span>';
            return '<li>' + escapeHtml(q.item) + '： <b>' + escapeHtml(q.value) + '</b> ' + where + '</li>';
        }).join('');
        html += '<div class="hh-box"><div class="hh-box-title">定性検査</div>'
            + '<ul class="hh-qual">' + items + '</ul></div>';
    }

    // PDFから読んだときだけ、原票そのものを並べて見比べられるようにする
    if (person.page_image_url) {
        const url = escapeHtml(person.page_image_url);
        html += '<div class="hh-box"><div class="hh-box-title">原票（PDF '
            + escapeHtml(person.page) + 'ページ）</div>'
            + '<a href="' + url + '" target="_blank" rel="noopener">'
            + '<img class="hh-thumb" src="' + url + '" loading="lazy" alt="原票"></a>'
            + '<div class="hh-reason">クリックで原寸表示。血圧の1回目・2回目と '
            + '<code>(-)</code> をこの画像で確かめてください。</div></div>';
    }
    html += '</div>';

    if ((person.unmapped_items || []).length) {
        const names = person.unmapped_items.map(u => u.category + '/' + u.item).join('、');
        html += '<div class="hh-issue hh-issue-warning">⚠️ 変換マスタに無いため出力しない項目: '
            + escapeHtml(names) + '</div>';
    }

    // エラー・警告はそのまま出し、AIの「読めなかった」メモは畳んでおく
    const issues = person.issues || [];
    issues.filter(i => i.level !== 'info').forEach(i => { html += hhIssueHtml(i); });
    const notes = issues.filter(i => i.level === 'info');
    if (notes.length) {
        html += '<details class="hh-needs-check"><summary>AIが読めなかった項目のメモ（'
            + notes.length + '件）</summary><ul class="hh-qual">'
            + notes.map(i => '<li>' + escapeHtml(i.message) + '</li>').join('')
            + '</ul></details>';
    }
    html += '</div>';
    return html;
}

function hhRenderPreview(data) {
    hhPreview = data;
    document.getElementById('hh-cnt-persons').textContent = data.counts.persons;
    document.getElementById('hh-cnt-errors').textContent = data.counts.errors;
    document.getElementById('hh-cnt-warnings').textContent = data.counts.warnings;

    const issuesEl = document.getElementById('hh-workbook-issues');
    issuesEl.innerHTML = (data.workbook_issues || []).map(hhIssueHtml).join('');

    let html = '';
    if ((data.master.institutions || []).length && data.persons.length > 1) {
        html += '<div class="hh-box" style="margin-bottom:9px; background:#fafcfe">'
            + '<div class="hh-box-title">全員に同じ健診機関・種別を設定する</div>'
            + '<select class="hh-select" id="hh-bulk-inst" style="min-width:240px">'
            + hhInstitutionOptions(data.master, '') + '</select> '
            + '<select class="hh-select" id="hh-bulk-course" style="min-width:220px">'
            + hhCourseOptions(data.master, '', '') + '</select> '
            + '<button type="button" class="btn btn-secondary" id="hh-bulk-apply">全員に適用</button>'
            + '</div>';
    }
    html += data.persons.map(p => hhRenderPerson(p, data.master, data.roster)).join('');
    document.getElementById('hh-persons').innerHTML = html;

    document.getElementById('hh-genpyo-confirmed').checked = false;
    document.getElementById('hh-generate-result').style.display = 'none';
    document.getElementById('hh-console').style.display = 'none';
    document.getElementById('hh-preview-area').style.display = 'block';
    hhUpdateGenerateButton();
}

function hhCollectSelections() {
    if (!hhPreview) return [];
    return hhPreview.persons.map(p => {
        const pick = (cls) => {
            const el = document.querySelector('.' + cls + '[data-key="' + p.key + '"]');
            return el ? (el.value || '') : '';
        };
        return {
            key: p.key,
            employee_id: pick('hh-emp'),
            institution: pick('hh-inst'),
            course: pick('hh-course'),
        };
    });
}

function hhUpdateGenerateButton() {
    const btn = document.getElementById('hh-generate-btn');
    if (!btn || !hhPreview) return;
    const confirmed = document.getElementById('hh-genpyo-confirmed').checked;
    const picks = hhCollectSelections();
    const allPicked = picks.length > 0
        && picks.every(p => p.employee_id && p.institution && p.course);
    const ready = hhPreview.counts.errors === 0 && allPicked && confirmed;
    btn.disabled = !ready;
    btn.title = ready ? ''
        : (hhPreview.counts.errors > 0
            ? 'エラーが残っているため作成できません'
            : (!allPicked ? '社員・健診機関・健診種別を全員分選んでください'
                : '「原票どおり」のチェックを入れてください'));
}

function hhRenderGenerateResult(data) {
    const el = document.getElementById('hh-generate-result');
    const warnings = (data.warnings || []).map(hhIssueHtml).join('');
    el.innerHTML = '<div class="alert alert-success" style="margin:0 0 6px">'
        + '✅ ' + data.row_count + '名分・' + data.column_count + '列のCSVを作成しました'
        + '<br>保存先: <code>' + escapeHtml(data.output_path) + '</code>'
        + '<br>書き出したファイルを読み直して、行数・列数・全セルの一致を確認済みです。'
        + '<br><b>このCSVをExcelで開いて保存し直さないでください</b>'
        + '（受診番号の先頭ゼロが落ち、行が消えることがあります）。'
        + '<br>HPMへの取り込み・チェック・更新は、これまでどおり手作業でお願いします。'
        + '</div>' + warnings;
    el.style.display = 'block';

    const consoleEl = document.getElementById('hh-console');
    consoleEl.textContent = (data.console || []).join('\n');
    consoleEl.style.display = (data.console || []).length ? 'block' : 'none';
}

async function hhPreviewExcel(file, status) {
    const form = new FormData();
    form.append('health_excel', file);
    form.append('master_path', document.getElementById('hh-master-path').value || '');

    status.textContent = '読み込み中…';
    const res = await fetch('/health_hpm_preview', { method: 'POST', body: form });
    const data = await res.json();
    if (!data.success) {
        hhShowError(data.errors || ['読み込みに失敗しました']);
        document.getElementById('hh-preview-area').style.display = 'none';
        status.textContent = '';
        return;
    }
    hhRenderPreview(data);
    status.textContent = data.counts.persons + '名を読み込みました';
}


// PDFは1人10〜20秒かかるので、SSEで進捗を出しながら読む。
async function hhPreviewPdf(file, status) {
    if (file.size > 45 * 1024 * 1024) {
        hhShowError('PDFが大きすぎます（上限50MB）。ページを分けてから読み込んでください');
        return;
    }
    const form = new FormData();
    form.append('health_pdf', file);
    form.append('master_path', document.getElementById('hh-master-path').value || '');

    status.textContent = 'PDFを送信しています…';
    const res = await fetch('/health_hpm_pdf_preview', { method: 'POST', body: form });

    // 読み取りが始まる前に落ちた場合（拡張子・マスタ・サイズ超過）はJSONで返る
    const contentType = res.headers.get('Content-Type') || '';
    if (!res.ok || contentType.includes('application/json')) {
        let errors = ['読み取りを開始できませんでした'];
        try { errors = (await res.json()).errors || errors; } catch { }
        hhShowError(errors);
        document.getElementById('hh-preview-area').style.display = 'none';
        status.textContent = '';
        return;
    }
    await consumeSSEResponse(res, hhProcessSSEPart);
}


function hhProcessSSEPart(part) {
    const { eventType, data } = parseSSEPart(part);
    if (!eventType || !data) return;
    const status = document.getElementById('hh-status');

    if (eventType === 'progress') {
        status.textContent = data.message || '読み取り中…';
    } else if (eventType === 'done') {
        hhRenderPreview(data);
        status.textContent = data.counts.persons
            + '名を読み取りました（AIの読み取りです。原票と見比べてください）';
    } else if (eventType === 'error') {
        hhShowError(data.message || '読み取りに失敗しました');
        document.getElementById('hh-preview-area').style.display = 'none';
        status.textContent = '';
    }
}


function setupHealthHpmMode() {
    const previewBtn = document.getElementById('hh-preview-btn');
    const generateBtn = document.getElementById('hh-generate-btn');
    if (!previewBtn || !generateBtn) return;

    const personsEl = document.getElementById('hh-persons');

    // 機関を変えたら、その機関の健診種別だけに選択肢を差し替える
    personsEl.addEventListener('change', (e) => {
        if (e.target.classList.contains('hh-inst')) {
            const key = e.target.dataset.key;
            const courseEl = document.querySelector('.hh-course[data-key="' + key + '"]');
            if (courseEl) courseEl.innerHTML = hhCourseOptions(hhPreview.master, e.target.value, '');
        }
        if (e.target.id === 'hh-bulk-inst') {
            const bulkCourse = document.getElementById('hh-bulk-course');
            if (bulkCourse) bulkCourse.innerHTML = hhCourseOptions(hhPreview.master, e.target.value, '');
        }
        hhUpdateGenerateButton();
    });

    personsEl.addEventListener('click', (e) => {
        if (e.target.id !== 'hh-bulk-apply') return;
        const inst = document.getElementById('hh-bulk-inst').value;
        const course = document.getElementById('hh-bulk-course').value;
        if (!inst) { alert('健診機関を選んでください'); return; }
        document.querySelectorAll('.hh-inst').forEach(sel => {
            sel.value = inst;
            const courseEl = document.querySelector('.hh-course[data-key="' + sel.dataset.key + '"]');
            if (courseEl) {
                courseEl.innerHTML = hhCourseOptions(hhPreview.master, inst, course);
                courseEl.value = course;
            }
        });
        hhUpdateGenerateButton();
    });

    document.getElementById('hh-genpyo-confirmed')
        .addEventListener('change', hhUpdateGenerateButton);

    previewBtn.addEventListener('click', async () => {
        const input = document.getElementById('hh-file-input');
        const status = document.getElementById('hh-status');
        hhClearError();
        const file = input.files && input.files[0];
        if (!file) {
            hhShowError('健診結果のPDF、または整形済みExcel（.xlsx）を選んでください');
            return;
        }

        previewBtn.disabled = true;
        try {
            if (/\.pdf$/i.test(file.name)) {
                await hhPreviewPdf(file, status);
            } else {
                await hhPreviewExcel(file, status);
            }
        } catch (e) {
            hhShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            previewBtn.disabled = false;
        }
    });

    generateBtn.addEventListener('click', async () => {
        const status = document.getElementById('hh-generate-status');
        hhClearError();
        if (!confirm('HPM取込用CSVを作成します。\n\n'
            + '・血圧は1回目と2回目を原票どおり別々に出します（平均は作りません）\n'
            + '・(-) は陰性として出力します\n'
            + '・HPMへの取り込み・チェック・更新は手作業で行ってください\n\n'
            + 'よろしいですか？')) return;

        generateBtn.disabled = true;
        status.textContent = '作成中…';
        try {
            const res = await fetch('/health_hpm_generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    session_id: hhPreview.session_id,
                    genpyo_confirmed: document.getElementById('hh-genpyo-confirmed').checked,
                    persons: hhCollectSelections(),
                    output_filename: document.getElementById('hh-output-filename').value || '',
                }),
            });
            const data = await res.json();
            if (!data.success) {
                hhShowError(data.errors || ['作成に失敗しました']);
                const consoleEl = document.getElementById('hh-console');
                consoleEl.textContent = (data.console || []).join('\n');
                consoleEl.style.display = (data.console || []).length ? 'block' : 'none';
                status.textContent = '';
                return;
            }
            hhRenderGenerateResult(data);
            status.textContent = '完了';
        } catch (e) {
            hhShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            hhUpdateGenerateButton();
        }
    });
}

setupHealthHpmMode();

// =============================================================================
// 社労士モード — 前田事務所へ渡す給与CSV（60列・cp932）
// =============================================================================
function sharoushiShowError(msgs, canForce) {
    const el = document.getElementById('sharoushi-error-area');
    if (!el) return;
    const list = Array.isArray(msgs) ? msgs : [msgs];
    let html = list.map(m => '<div>' + escapeHtml(m) + '</div>').join('');
    if (canForce) {
        // 未知項目が出たとき。中身を確かめたうえで、承知のうえ出力できる逃げ道を出す。
        html += '<div style="margin-top:8px">'
            + '<button type="button" class="btn" id="sharoushi-force-btn">'
            + '内容を確認したので、このまま出力する</button></div>';
    }
    el.innerHTML = html;
    el.style.display = list.length ? 'block' : 'none';
}

function sharoushiRenderResult(data) {
    document.getElementById('sharoushi-cnt-rows').textContent = data.rows;
    document.getElementById('sharoushi-cnt-excluded').textContent = (data.excluded || []).length;
    document.getElementById('sharoushi-cnt-ledger').textContent =
        (data.ledger_applied || []).length;
    const warnCount = (data.unknown || []).length + (data.unmapped_systems || []).length
        + (data.multi_statement || []).length;
    document.getElementById('sharoushi-cnt-warn').textContent = warnCount;

    const dlUrl = (fn) => '/sharoushi_download/' + encodeURIComponent(data.ym) + '/'
        + encodeURIComponent(fn);
    const bikoN = (data.biko_rows || []).length;
    document.getElementById('sharoushi-file').innerHTML =
        '<table class="keiri-md-table"><tr><th>ファイル</th><th>内容</th><th></th></tr>'
        + '<tr><td>' + escapeHtml(data.filename) + '</td><td>' + data.rows + '人 / '
        + data.columns + '列</td><td><a class="btn btn-sm" href="' + dlUrl(data.filename)
        + '">ダウンロード</a></td></tr>'
        + '<tr><td>' + escapeHtml(data.biko_filename || '') + '</td><td>イレギュラー発生分の理由 '
        + bikoN + '件</td><td><a class="btn btn-sm" href="' + dlUrl(data.biko_filename)
        + '">ダウンロード</a></td></tr>'
        + '</table>'
        + '<div class="hint" style="margin-top:4px">保存先: ' + escapeHtml(data.out_dir) + '</div>';

    // 要確認（未知項目・想定外の給与体系・複数明細）
    const warn = document.getElementById('sharoushi-warn-area');
    let wh = '';
    if ((data.unknown || []).length) {
        wh += '<div class="alert alert-warning"><b>マッピングに無い支給・控除項目に金額があります'
            + '（' + data.unknown.length + '件）</b>'
            + '<div class="hint">jinjer 側で項目が移設された可能性があります。'
            + '列マッピングCSVに追記してください。</div>'
            + '<table class="keiri-md-table"><tr><th>社員番号</th><th>氏名</th>'
            + '<th>項目ID</th><th>項目名</th><th>金額</th></tr>'
            + data.unknown.map(u => '<tr><td>' + escapeHtml(u['社員番号']) + '</td><td>'
                + escapeHtml(u['氏名']) + '</td><td>' + escapeHtml(u.source_key) + '</td><td>'
                + escapeHtml(u.label) + '</td><td>' + escapeHtml(u['金額']) + '</td></tr>').join('')
            + '</table></div>';
    }
    if ((data.unmapped_systems || []).length) {
        wh += '<div class="alert alert-warning">勤怠列のマッピングが無い給与体系があります: '
            + escapeHtml(data.unmapped_systems.join('、'))
            + '<div class="hint">その体系の人は勤怠列が空のまま出ます。</div></div>';
    }
    if ((data.multi_statement || []).length) {
        wh += '<div class="alert alert-warning">同じ月に給与明細が複数ある人: '
            + escapeHtml(data.multi_statement.join('、'))
            + '<div class="hint">基本給が入っている明細を採用しました。</div></div>';
    }
    if ((data.biko_pending || []).length) {
        wh += '<div class="alert alert-warning">イレギュラー発生分のうち <b>'
            + data.biko_pending.length + '件</b> の理由が未入力です'
            + '<div class="hint">下の「イレギュラー発生分の理由」で入れてから渡してください。'
            + '未入力のままでも備考CSVは出ますが、理由欄が空になります。</div></div>';
    }
    warn.innerHTML = wh;
    warn.style.display = wh ? 'block' : 'none';
    sharoushiRenderBiko(data);

    // 内訳（給与体系別・対象外・台帳の反映）
    let dh = '<div class="hint">給与体系: '
        + Object.entries(data.systems || {}).map(([k, v]) => escapeHtml(k) + ' ' + v + '人').join(' / ')
        + '</div>';
    if ((data.ledger_applied || []).length) {
        dh += '<div style="margin-top:6px"><b>追加支給台帳を反映しました</b>'
            + '<table class="keiri-md-table"><tr><th>社員番号</th><th>氏名</th>'
            + '<th>項目</th><th>金額</th><th>メモ</th></tr>'
            + data.ledger_applied.map(a => '<tr><td>' + escapeHtml(a['社員番号']) + '</td><td>'
                + escapeHtml(a['氏名']) + '</td><td>' + escapeHtml(a['項目']) + '</td><td>'
                + escapeHtml(a['金額']) + '</td><td>' + escapeHtml(a['メモ'] || '') + '</td></tr>').join('')
            + '</table></div>';
    }
    if ((data.excluded || []).length) {
        dh += '<div class="hint" style="margin-top:6px">対象外: '
            + data.excluded.map(e => escapeHtml(e['社員番号']) + '（' + escapeHtml(e['理由']) + '）')
                .join(' / ') + '</div>';
    }
    dh += '<div class="hint" style="margin-top:6px">列マッピング: '
        + escapeHtml(data.mapping_path) + '（' + data.mapping_rows + '行）<br>追加支給台帳: '
        + escapeHtml(data.ledger_path || '（未設定）') + '</div>';
    document.getElementById('sharoushi-detail').innerHTML = dh;

    document.getElementById('sharoushi-preview').textContent = (data.preview || []).join('\n');
    document.getElementById('sharoushi-result-area').style.display = 'block';
}

function sharoushiRenderBiko(data) {
    const area = document.getElementById('sharoushi-biko-area');
    const rows = data.biko_rows || [];
    if (!area) return;
    if (!rows.length) { area.style.display = 'none'; return; }
    document.getElementById('sharoushi-biko-path').textContent =
        '台帳: ' + (data.biko_ledger_path || '（未設定）');
    let html = '<table class="keiri-md-table"><tr><th>社員番号</th><th>氏名</th>'
        + '<th>項目</th><th style="text-align:right">金額</th><th>理由</th></tr>';
    rows.forEach((r, i) => {
        const need = !String(r['理由'] || '').trim();
        html += '<tr' + (need ? ' style="background:#FFF6F6"' : '') + '>'
            + '<td data-biko-emp="' + escapeHtml(r['社員番号']) + '">' + escapeHtml(r['社員番号']) + '</td>'
            + '<td data-biko-name="' + escapeHtml(r['氏名']) + '">' + escapeHtml(r['氏名']) + '</td>'
            + '<td data-biko-item="' + escapeHtml(r['項目']) + '">' + escapeHtml(r['項目']) + '</td>'
            + '<td style="text-align:right">' + escapeHtml(r['金額']) + '</td>'
            + '<td><input type="text" class="biko-reason" data-idx="' + i
            + '" style="width:100%" placeholder="' + (need ? '理由を入れてください' : '')
            + '" value="' + escapeHtml(r['理由'] || '') + '"></td></tr>';
    });
    html += '</table>';
    document.getElementById('sharoushi-biko-rows').innerHTML = html;
    area.style.display = 'block';
}

function sharoushiCollectBiko() {
    const entries = [];
    document.querySelectorAll('#sharoushi-biko-rows tr').forEach(tr => {
        const input = tr.querySelector('.biko-reason');
        if (!input) return;
        entries.push({
            '社員番号': tr.querySelector('[data-biko-emp]').dataset.bikoEmp,
            '氏名': tr.querySelector('[data-biko-name]').dataset.bikoName,
            '項目': tr.querySelector('[data-biko-item]').dataset.bikoItem,
            '理由': input.value.trim(),
        });
    });
    return entries;
}

const sharoushiRunBtn = document.getElementById('sharoushi-run-btn');
if (sharoushiRunBtn) {
    const sharoushiRun = async (allowUnknown) => {
        const status = document.getElementById('sharoushi-status');
        const month = (document.getElementById('sharoushi-month').value || '').trim();
        sharoushiShowError([]);
        if (!/^\d{4}-\d{2}$/.test(month)) {
            sharoushiShowError('支給月は YYYY-MM 形式で入力してください（例: 2026-08）');
            return;
        }
        const fd = new FormData();
        fd.append('month', month);
        fd.append('mapping_csv', document.getElementById('sharoushi-mapping-csv').value);
        fd.append('ledger_csv', document.getElementById('sharoushi-ledger-csv').value);
        fd.append('refresh', document.getElementById('sharoushi-refresh').checked ? '1' : '0');
        if (allowUnknown) fd.append('allow_unknown', '1');

        sharoushiRunBtn.disabled = true;
        status.textContent = '生成中…（給与明細の取得に数分かかることがあります）';
        document.getElementById('sharoushi-result-area').style.display = 'none';
        try {
            const res = await fetch('/sharoushi_run', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) {
                sharoushiShowError(data.errors || ['生成に失敗しました'], data.can_force);
                const forceBtn = document.getElementById('sharoushi-force-btn');
                if (forceBtn) forceBtn.addEventListener('click', () => sharoushiRun(true));
                status.textContent = '';
                return;
            }
            sharoushiRenderResult(data);
            status.textContent = '完了（' + data.month + '）';
        } catch (e) {
            sharoushiShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            sharoushiRunBtn.disabled = false;
        }
    };
    sharoushiRunBtn.addEventListener('click', () => sharoushiRun(false));

    const bikoSaveBtn = document.getElementById('sharoushi-biko-save-btn');
    if (bikoSaveBtn) {
        bikoSaveBtn.addEventListener('click', async () => {
            const status = document.getElementById('sharoushi-biko-status');
            const month = (document.getElementById('sharoushi-month').value || '').trim();
            const entries = sharoushiCollectBiko();
            sharoushiShowError([]);
            if (!entries.length) { sharoushiShowError('保存する行がありません'); return; }
            bikoSaveBtn.disabled = true;
            status.textContent = '保存中…';
            try {
                const res = await fetch('/sharoushi_biko_save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ month, entries }),
                });
                const data = await res.json();
                if (!data.success) {
                    sharoushiShowError(data.errors || ['保存に失敗しました']);
                    status.textContent = '';
                    return;
                }
                status.textContent = '保存しました（理由あり ' + data.saved + '/' + data.total + '）。作り直しています…';
                await sharoushiRun(false);       // 台帳を反映した備考CSVを作り直す
            } catch (e) {
                sharoushiShowError('通信に失敗しました: ' + e);
                status.textContent = '';
            } finally {
                bikoSaveBtn.disabled = false;
            }
        });
    }
}

// =============================================================================
// 標準報酬チェック — 定時決定の検算と保険料突合
// =============================================================================
function shahoShowError(msgs) {
    const el = document.getElementById('shaho-error-area');
    if (!el) return;
    const list = Array.isArray(msgs) ? msgs : [msgs];
    el.innerHTML = list.map(m => '<div>' + escapeHtml(m) + '</div>').join('');
    el.style.display = list.length ? 'block' : 'none';
}

const shahoRunBtn = document.getElementById('shaho-run-btn');
if (shahoRunBtn) {
    shahoRunBtn.addEventListener('click', async () => {
        const status = document.getElementById('shaho-status');
        shahoShowError([]);
        const fd = new FormData();
        fd.append('year', document.getElementById('shaho-year').value.trim());
        fd.append('check_month', document.getElementById('shaho-check-month').value.trim());
        shahoRunBtn.disabled = true;
        status.textContent = '検算中…（30秒ほどかかります）';
        document.getElementById('shaho-result-area').style.display = 'none';
        try {
            const res = await fetch('/shaho_run', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) {
                shahoShowError(data.errors || ['実行に失敗しました']);
                status.textContent = '';
                return;
            }
            document.getElementById('shaho-cnt-n').textContent = data.n;
            document.getElementById('shaho-cnt-review').textContent = data.review_n;
            const openMonths = Object.entries(data.open_months || {});
            document.getElementById('shaho-open-months').innerHTML = openMonths.length
                ? '<div class="alert alert-warning"><b>給与が未確定の月があります: '
                    + escapeHtml(openMonths.map(([m, n]) => m + '（' + n + '名）').join('、'))
                    + '</b><div class="hint">未確定の月は報酬額も登録標準報酬月額もまだ動きます。'
                    + 'その月を含む人は情報不足に落としています。給与を確定してから実行し直してください。'
                    + '</div></div>'
                : '';
            document.getElementById('shaho-statuses').innerHTML =
                '<table class="keiri-md-table"><tr><th>判定</th><th>人数</th></tr>'
                + data.statuses.map(s => '<tr' + (s.review ? ' style="background:#FFF6F6"' : '')
                    + '><td>' + (s.review ? '⚠ ' : '') + escapeHtml(s.label) + '</td><td>'
                    + s.count + '名</td></tr>').join('')
                + '</table>';
            const dl = (fn) => '/shaho_download/' + data.year + '/' + encodeURIComponent(fn);
            document.getElementById('shaho-files').innerHTML =
                '<a class="btn btn-sm" href="' + dl(data.xlsx) + '">📊 Excelをダウンロード</a> '
                + '<a class="btn btn-sm" href="' + dl(data.json) + '">JSON</a>'
                + '<div class="hint" style="margin-top:4px">'
                + escapeHtml(data.year + '年4〜6月支給 → 9月適用予定 / 突合月 ' + data.check_month)
                + '<br>保存先: ' + escapeHtml(data.out_dir)
                + '（個人情報を含むため共有フォルダへ置かないでください）</div>';
            document.getElementById('shaho-result-area').style.display = 'block';
            status.textContent = '完了';
        } catch (e) {
            shahoShowError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            shahoRunBtn.disabled = false;
        }
    });
}


// =============================================================================
// 標報投入 — 社労士の保険料一覧表PDF → jinjer報酬月額
//   このモードだけ jinjer へ書き込む。読み込み → 突合 → ①内容確定 → ②投入 の順。
//   126名だと投入に1時間近くかかるので、進捗はSSEではなくポーリングで追う
//   （ブラウザを閉じてもサーバ側の書き込みは最後まで走る）。
// =============================================================================
const shahoImport = { sessionId: '', planHash: '', targetYm: '', rows: [], timer: null };

function shahoImportError(msgs) {
    const el = document.getElementById('shaho-import-error-area');
    if (!el) return;
    const list = Array.isArray(msgs) ? msgs : [msgs];
    el.innerHTML = list.map(m => '<div>' + escapeHtml(m) + '</div>').join('');
    el.style.display = list.length ? 'block' : 'none';
}

function shahoImportYen(v) {
    return (v === null || v === undefined || v === '') ? '—' : Number(v).toLocaleString();
}

function shahoImportPicked() {
    return Array.from(document.querySelectorAll('.shaho-import-pick:checked'))
        .map(el => ({ emp: el.dataset.emp, forced: el.dataset.force === '1' }));
}

function shahoImportSyncButtons() {
    const n = shahoImportPicked().length;
    const dry = document.getElementById('shaho-import-dryrun-btn');
    if (dry) {
        dry.disabled = n === 0;
        dry.textContent = n ? '① 投入内容を確定（' + n + '名・まだ書きません）'
                            : '① 投入内容を確定（投入する人を選んでください）';
    }
    // 選び直したら確定済みの内容は無効にする（見たものと違うものを書かせない）
    const confirmArea = document.getElementById('shaho-import-confirm-area');
    if (confirmArea) confirmArea.style.display = 'none';
}

function shahoImportRenderRows(data) {
    const forceable = data.rows.filter(r => r.needs_force).length;
    document.getElementById('shaho-import-head').innerHTML = '<div class="hint">'
        + '<b>' + escapeHtml(data.target_ym) + ' 分</b>の保険料一覧表'
        + '（' + escapeHtml(data.pay_ym) + ' の給与から控除）／'
        + escapeHtml(data.office) + '<br>'
        + 'チェックサム照合: ' + escapeHtml((data.checksum || []).join('、') || 'なし')
        + ' → 一致。読み落としはありません。</div>';

    const cnt = (fn) => data.rows.filter(fn).length;
    const reviewStatuses = ['CALC_MISMATCH', 'PDF_INCONSISTENT', 'NOT_IN_JINJER', 'RETIRED'];
    document.getElementById('shaho-import-cnt-n').textContent = data.rows.length;
    document.getElementById('shaho-import-cnt-target').textContent =
        cnt(r => r.status === 'AUTO_OK' || r.status === 'NO_CALC');
    document.getElementById('shaho-import-cnt-review').textContent =
        cnt(r => reviewStatuses.includes(r.status));
    document.getElementById('shaho-import-cnt-nochange').textContent =
        cnt(r => r.status === 'NO_CHANGE');

    const notes = (data.notes || []).map(n => '<div>' + escapeHtml(n) + '</div>').join('');
    document.getElementById('shaho-import-notes').innerHTML = notes
        ? '<div class="alert alert-warning">' + notes + '</div>' : '';

    let html = '';
    if (forceable) {
        // 「承知のうえ投入」は判断が要るので、その権限を持つ人だけが開けられる
        html += '<label style="display:block; margin-bottom:6px">'
            + '<input type="checkbox" id="shaho-import-allow-force"'
            + (data.can_force ? '' : ' disabled') + '> '
            + '要確認の' + forceable + '名も選べるようにする（承知のうえ投入）'
            + (data.can_force ? '' : '<br><span class="hint">'
                + escapeHtml(data.force_reason || '') + '</span>')
            + '</label>';
    }
    html += '<table class="keiri-md-table"><tr><th></th><th>社員番号</th><th>氏名</th>'
        + '<th>判定</th><th>社労士PDF<br>健保／厚年</th><th>jinjer登録<br>健保／厚年</th>'
        + '<th>当方の計算<br>健保／厚年</th><th>操作</th><th>改訂理由・備考</th></tr>';
    data.rows.forEach(r => {
        const review = reviewStatuses.includes(r.status);
        const box = r.selectable
            ? '<input type="checkbox" class="shaho-import-pick" data-emp="' + escapeHtml(r.emp)
              + '" data-force="' + (r.needs_force ? '1' : '0') + '"'
              + (r.default_selected ? ' checked' : '')
              + (r.needs_force ? ' disabled' : '') + '>'
            : '';
        const op = r.operation === 'PATCH' ? '更新' : (r.operation === 'POST' ? '新規' : '—');
        html += '<tr' + (review ? ' style="background:#FFF6F6"' : '') + '>'
            + '<td>' + box + '</td>'
            + '<td>' + escapeHtml(r.emp) + '</td>'
            + '<td>' + escapeHtml(r.name) + '</td>'
            + '<td>' + (review ? '⚠ ' : '') + escapeHtml(r.status_ja) + '</td>'
            + '<td>' + shahoImportYen(r.pdf_kenpo) + '／' + shahoImportYen(r.pdf_konen) + '</td>'
            + '<td>' + shahoImportYen(r.cur_kenpo) + '／' + shahoImportYen(r.cur_konen)
            + (r.cur_ym ? '<br><span class="hint">' + escapeHtml(r.cur_ym) + '</span>' : '') + '</td>'
            + '<td>' + shahoImportYen(r.calc_kenpo) + '／' + shahoImportYen(r.calc_konen) + '</td>'
            + '<td>' + escapeHtml(op) + '</td>'
            + '<td><span class="hint">' + escapeHtml(r.reason || '') + '</span>'
            + (r.notes || []).map(n => '<div class="hint">・' + escapeHtml(n) + '</div>').join('')
            + '</td></tr>';
    });
    html += '</table>';
    document.getElementById('shaho-import-rows').innerHTML = html;

    const allow = document.getElementById('shaho-import-allow-force');
    if (allow) {
        allow.addEventListener('change', () => {
            document.querySelectorAll('.shaho-import-pick[data-force="1"]').forEach(el => {
                el.disabled = !allow.checked;
                if (!allow.checked) el.checked = false;
            });
            shahoImportSyncButtons();
        });
    }
    document.querySelectorAll('.shaho-import-pick')
        .forEach(el => el.addEventListener('change', shahoImportSyncButtons));

    const selectable = data.rows.filter(r => r.selectable).length;
    document.getElementById('shaho-import-exec-area').style.display = 'block';
    const writerEl = document.getElementById('shaho-import-writer');
    if (!data.can_write) {
        writerEl.innerHTML = '<b>' + escapeHtml(data.write_reason) + '</b>';
    } else if (!selectable) {
        // 全員すでに登録済み等で投入対象が0名。ボタンが押せない理由を必ず書く
        const nochange = data.rows.filter(r => r.status === 'NO_CHANGE').length;
        writerEl.innerHTML = '<b>投入する人はいません。</b>'
            + (nochange === data.rows.length
                ? '通知の値は' + nochange + '名全員すでにjinjerに入っています（やることなし）。'
                : nochange + '名は登録済み、残りは要確認のため投入できません。'
                  + '上の表で理由を確認してください。');
    } else {
        writerEl.innerHTML = escapeHtml(data.write_reason) + '／1件あたり約'
            + escapeHtml(String(data.write_interval)) + '秒かかります。'
            + '実行中は経費インポートなど他のAPI投入を走らせないでください。';
    }
    document.getElementById('shaho-import-dryrun-btn').disabled =
        !data.can_write || !selectable;
    document.getElementById('shaho-import-confirm-area').style.display = 'none';
    document.getElementById('shaho-import-progress-area').style.display = 'none';
    if (data.can_write) shahoImportSyncButtons();
}

const shahoImportPreviewBtn = document.getElementById('shaho-import-preview-btn');
if (shahoImportPreviewBtn) {
    shahoImportPreviewBtn.addEventListener('click', async () => {
        const status = document.getElementById('shaho-import-status');
        const fileInput = document.getElementById('shaho-import-pdf');
        shahoImportError([]);
        if (!fileInput.files.length) {
            shahoImportError('保険料一覧表のPDF、または関東ITSの決定通知書CSVを選んでください');
            return;
        }
        if (fileInput.files[0].name.toLowerCase().endsWith('.csv')
            && !document.getElementById('shaho-import-expected-ym').value.trim()) {
            shahoImportError('関東ITSのCSVには適用年月が入っていません。'
                + '「対象年月」に適用する年月（定時決定なら 2026-09）を入力してください');
            return;
        }
        const fd = new FormData();
        fd.append('hoken_pdf', fileInput.files[0]);
        fd.append('expected_ym', document.getElementById('shaho-import-expected-ym').value.trim());
        shahoImportPreviewBtn.disabled = true;
        status.textContent = 'ファイルを読み、jinjer の登録値と突き合わせています…';
        document.getElementById('shaho-import-result-area').style.display = 'none';
        try {
            const res = await fetch('/shaho_import_preview', { method: 'POST', body: fd });
            const data = await res.json();
            if (!data.success) {
                shahoImportError(data.errors || ['読み取りに失敗しました']);
                status.textContent = '';
                return;
            }
            shahoImport.sessionId = data.session_id;
            shahoImport.planHash = data.plan_hash;
            shahoImport.targetYm = data.target_ym;
            shahoImport.rows = data.rows;
            shahoImportRenderRows(data);
            document.getElementById('shaho-import-result-area').style.display = 'block';
            status.textContent = '突合が終わりました（この時点では何も書いていません）';
        } catch (e) {
            shahoImportError('通信に失敗しました: ' + e);
            status.textContent = '';
        } finally {
            shahoImportPreviewBtn.disabled = false;
        }
    });
}

const shahoImportDryRunBtn = document.getElementById('shaho-import-dryrun-btn');
if (shahoImportDryRunBtn) {
    shahoImportDryRunBtn.addEventListener('click', async () => {
        const status = document.getElementById('shaho-import-dryrun-status');
        shahoImportError([]);
        status.textContent = '確認中…';
        try {
            const res = await fetch('/shaho_import_execute', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    session_id: shahoImport.sessionId, plan_hash: shahoImport.planHash,
                    selected: shahoImportPicked(), dry_run: true,
                }),
            });
            const data = await res.json();
            if (!data.success) {
                shahoImportError(data.errors || ['確認に失敗しました']);
                status.textContent = '';
                return;
            }
            document.getElementById('shaho-import-dryrun-detail').innerHTML =
                '<b>' + data.count + '名</b> に書き込みます（所要 約' + data.eta_minutes + '分）。'
                + (data.forced.length
                    ? '<br>承知のうえ投入: ' + escapeHtml(data.forced.join('、')) : '')
                + '<table class="keiri-md-table" style="margin-top:6px">'
                + '<tr><th>社員番号</th><th>氏名</th><th>操作</th><th>健保</th><th>厚年</th></tr>'
                + data.results.map(r => '<tr><td>' + escapeHtml(r.emp) + '</td><td>'
                    + escapeHtml(r.name) + '</td><td>'
                    + escapeHtml(r.operation === 'PATCH' ? '更新' : '新規') + '</td><td>'
                    + shahoImportYen(r.before_kenpo) + ' → <b>' + shahoImportYen(r.after_kenpo)
                    + '</b></td><td>' + shahoImportYen(r.before_konen) + ' → <b>'
                    + shahoImportYen(r.after_konen) + '</b></td></tr>').join('')
                + '</table>';
            document.getElementById('shaho-import-confirm-area').style.display = 'block';
            document.getElementById('shaho-import-confirm-ym').value = '';
            document.getElementById('shaho-import-confirm-check').checked = false;
            document.getElementById('shaho-import-run-btn').disabled = true;
            status.textContent = '内容を確認してください';
        } catch (e) {
            shahoImportError('通信に失敗しました: ' + e);
            status.textContent = '';
        }
    });
}

function shahoImportSyncRunBtn() {
    const ymOk = document.getElementById('shaho-import-confirm-ym').value.trim()
        === shahoImport.targetYm;
    const checked = document.getElementById('shaho-import-confirm-check').checked;
    document.getElementById('shaho-import-run-btn').disabled = !(ymOk && checked);
}
['shaho-import-confirm-ym', 'shaho-import-confirm-check'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', shahoImportSyncRunBtn);
    el.addEventListener('change', shahoImportSyncRunBtn);
});

function shahoImportRenderProgress(p) {
    const area = document.getElementById('shaho-import-progress-area');
    if (!area) return;
    area.style.display = 'block';
    const counts = p.counts || {};
    let html = '<div class="section-title" style="font-size:13px">投入の進み具合</div>';
    if (p.state === 'running') {
        html += '<div><b>' + (p.done || 0) + ' / ' + (p.total || 0) + '名</b> 完了'
            + (p.message ? '（' + escapeHtml(p.message) + '）' : '')
            + '</div><div class="hint">この画面を閉じても投入は最後まで続きます。'
            + '戻ってきたらもう一度このモードを開いてください。</div>';
    } else if (p.state === 'done') {
        const ng = counts.verify_ng || 0;
        html += '<div class="alert ' + (ng || counts.failed ? 'alert-warning' : '')
            + '"><b>投入が終わりました</b>：成功 ' + (counts.ok || 0) + '名'
            + '／失敗 ' + (counts.failed || 0) + '名'
            + '／スキップ ' + (counts.skipped || 0) + '名'
            + '／書込後の照合NG ' + ng + '名</div>';
        if (p.ledger) html += '<div class="hint">実行台帳: ' + escapeHtml(p.ledger) + '</div>';
        if (p.backup) html += '<div class="hint">投入前バックアップ: ' + escapeHtml(p.backup) + '</div>';
        html += '<div class="hint">続きをやり直すときは、同じファイルでもう一度「データを突合する」を'
            + '押してください。投入済みの人は自動で「書込不要」に落ちます。</div>';
    } else if (p.state === 'error') {
        html += '<div class="alert alert-error"><b>投入が止まりました</b>：'
            + escapeHtml(p.message || '') + '</div>'
            + '<div class="hint">同じファイルでもう一度突合すると、どこまで入ったかが分かります。</div>';
    }
    (p.errors || []).forEach(e => {
        html += '<div class="alert alert-warning">' + escapeHtml(e) + '</div>';
    });
    const entries = (p.entries || []).slice().reverse();
    if (entries.length) {
        html += '<table class="keiri-md-table" style="margin-top:6px">'
            + '<tr><th>社員番号</th><th>氏名</th><th>操作</th><th>結果</th><th>確認</th><th>メモ</th></tr>'
            + entries.map(e => '<tr><td>' + escapeHtml(e.emp) + '</td><td>' + escapeHtml(e.name)
                + '</td><td>' + escapeHtml(e.operation || '') + '</td><td>'
                + escapeHtml(e.result || '') + '</td><td>'
                + escapeHtml((p.verified || {})[e.emp] || '') + '</td><td><span class="hint">'
                + escapeHtml(e.message || '') + '</span></td></tr>').join('')
            + '</table>';
    }
    area.innerHTML = html;
}

async function shahoImportPoll() {
    if (!shahoImport.sessionId) return;
    try {
        const res = await fetch('/shaho_import_status?session_id='
            + encodeURIComponent(shahoImport.sessionId));
        const p = await res.json();
        if (p.state && p.state !== 'none') shahoImportRenderProgress(p);
        if (p.state === 'done' || p.state === 'error') {
            clearInterval(shahoImport.timer);
            shahoImport.timer = null;
            document.getElementById('shaho-import-run-status').textContent =
                p.state === 'done' ? '完了' : '停止';
        }
    } catch (e) {
        // 通信が一時的に切れても投入自体は走っている。次の周期で拾い直す
    }
}

const shahoImportRunBtn = document.getElementById('shaho-import-run-btn');
if (shahoImportRunBtn) {
    shahoImportRunBtn.addEventListener('click', async () => {
        const status = document.getElementById('shaho-import-run-status');
        shahoImportError([]);
        shahoImportRunBtn.disabled = true;
        status.textContent = '投入を開始しています…';
        try {
            const res = await fetch('/shaho_import_execute', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    session_id: shahoImport.sessionId, plan_hash: shahoImport.planHash,
                    selected: shahoImportPicked(), dry_run: false,
                    confirm_ym: document.getElementById('shaho-import-confirm-ym').value.trim(),
                }),
            });
            const data = await res.json();
            if (!data.success) {
                shahoImportError(data.errors || ['投入を開始できませんでした']);
                status.textContent = '';
                shahoImportRunBtn.disabled = false;
                return;
            }
            status.textContent = '投入中（約' + data.eta_minutes + '分）…';
            document.getElementById('shaho-import-dryrun-btn').disabled = true;
            if (shahoImport.timer) clearInterval(shahoImport.timer);
            shahoImport.timer = setInterval(shahoImportPoll, 4000);
            shahoImportPoll();
        } catch (e) {
            shahoImportError('通信に失敗しました: ' + e);
            status.textContent = '';
            shahoImportRunBtn.disabled = false;
        }
    });
}

// =============================================================================
// 請求書モード — 共有PDF → freee 売上取引CSV
// =============================================================================
const invoiceState = { month: '', rows: [], signature: '', preview: null };

function invoiceShowError(messages) {
    const area = document.getElementById('invoice-error-area');
    if (!area) return;
    const list = Array.isArray(messages) ? messages : (messages ? [messages] : []);
    area.innerHTML = list.map(message => '<div>' + escapeHtml(String(message)) + '</div>').join('');
    area.style.display = list.length ? '' : 'none';
}

function invoiceValidateRow(row) {
    const main = row._row_type === 'main';
    let required = ['勘定科目', '税区分', '金額', '税計算区分', '税額', '部門', '従業員'];
    if (main) required = ['収支区分', '管理番号', '発生日', '支払期日', '取引先'].concat(required);
    const errors = [];
    required.forEach(key => {
        if (String(row[key] == null ? '' : row[key]).trim() === '') errors.push(key + 'が未入力');
    });
    if (main && row['管理番号'] && !/^\d+$/.test(String(row['管理番号']).trim())) {
        errors.push('管理番号は数字で入力');
    }
    ['発生日', '支払期日'].forEach(key => {
        if (main && row[key] && !/^\d{4}-\d{2}-\d{2}$/.test(String(row[key]).replaceAll('/', '-'))) {
            errors.push(key + 'は正しい日付で入力');
        }
    });
    const amount = Number(String(row['金額'] == null ? '' : row['金額']).replaceAll(',', ''));
    const tax = Number(String(row['税額'] == null ? '' : row['税額']).replaceAll(',', ''));
    if (row['金額'] !== '' && (!Number.isInteger(amount) || amount <= 0)) errors.push('金額は1円以上の整数で入力');
    if (row['税額'] !== '' && (!Number.isInteger(tax) || tax < 0)) errors.push('税額は0円以上の整数で入力');
    if (Number.isFinite(amount) && Number.isFinite(tax) && tax > amount) errors.push('税額が金額を超えています');
    return errors;
}

function invoiceRevalidate() {
    let errorRows = 0;
    invoiceState.rows.forEach((row, index) => {
        row._errors = invoiceValidateRow(row);
        if (row._errors.length) errorRows += 1;
        const tr = document.querySelector('tr[data-invoice-row="' + index + '"]');
        if (tr) {
            tr.style.background = row._errors.length ? '#fff1f1' : '';
            const cell = tr.querySelector('.invoice-row-errors');
            if (cell) {
                const warnings = row._warnings || [];
                cell.innerHTML = row._errors.map(e => '<div style="color:#b00020">' + escapeHtml(e) + '</div>').join('')
                    + warnings.map(w => '<div style="color:#8a5b00">' + escapeHtml(w) + '</div>').join('');
            }
        }
    });
    const count = document.getElementById('invoice-count-errors');
    if (count) count.textContent = String(errorRows);
    const exportBtn = document.getElementById('invoice-export-btn');
    if (exportBtn) exportBtn.disabled = errorRows > 0 || !invoiceState.rows.length;
    return errorRows;
}

function invoiceInput(index, field, type, disabled) {
    const row = invoiceState.rows[index];
    const value = row[field] == null ? '' : row[field];
    const safeType = type || 'text';
    return '<input class="invoice-cell-input" data-index="' + index + '" data-field="'
        + escapeHtml(field) + '" type="' + safeType + '" value="' + escapeHtml(String(value)) + '"'
        + (disabled ? ' disabled' : '') + ' style="min-width:'
        + (field === '取引先' || field === '部門' ? '180px' : '105px') + '; padding:3px 5px">';
}

function invoiceRenderRows() {
    const target = document.getElementById('invoice-preview-table');
    if (!target) return;
    let html = '<table class="keiri-md-table"><tr>'
        + '<th>区分</th><th>従業員</th><th>管理番号</th><th>発生日</th><th>支払期日</th>'
        + '<th>取引先</th><th>勘定科目</th><th>金額</th><th>税額</th><th>部門</th><th>確認</th><th>元PDF</th><th>行操作</th></tr>';
    invoiceState.rows.forEach((row, index) => {
        const commute = row._row_type === 'commute';
        const sourceNames = (row._sources || []).map(path => String(path).split(/[\\/]/).pop());
        html += '<tr data-invoice-row="' + index + '"'
            // 手で足した行・分割した行は元PDFと突き合わせる必要があるので色で分かるようにする。
            // 固定列の背景も一緒に変える必要があるので、インラインstyleではなくクラスで持つ。
            + (row._manual_added ? ' class="invoice-manual-row"' : '') + '>'
            + '<td>' + (commute ? '交通費' : row._manual_added ? '本体(手入力)' : '本体') + '</td>'
            + '<td>' + invoiceInput(index, '従業員') + '</td>'
            + '<td>' + invoiceInput(index, '管理番号', 'text', commute) + '</td>'
            + '<td>' + invoiceInput(index, '発生日', 'date', commute) + '</td>'
            + '<td>' + invoiceInput(index, '支払期日', 'date', commute) + '</td>'
            + '<td>' + invoiceInput(index, '取引先', 'text', commute) + '</td>'
            + '<td>' + invoiceInput(index, '勘定科目') + '</td>'
            + '<td>' + invoiceInput(index, '金額', 'number') + '</td>'
            + '<td>' + invoiceInput(index, '税額', 'number') + '</td>'
            + '<td>' + invoiceInput(index, '部門') + '</td>'
            + '<td class="invoice-row-errors" style="min-width:170px"></td>'
            + '<td title="' + escapeHtml((row._sources || []).join('\n')) + '" style="min-width:160px">'
            + sourceNames.map(name => escapeHtml(name)).join('<br>') + '</td>'
            + '<td style="white-space:nowrap">'
            + (!commute ? '<button type="button" class="btn invoice-row-action" data-action="duplicate" data-index="' + index
                + '" style="padding:2px 6px; font-size:11px">複製</button>' : '')
            + (row._manual_added ? '<button type="button" class="btn invoice-row-action" data-action="delete" data-index="' + index
                + '" style="padding:2px 6px; font-size:11px; margin-left:4px">削除</button>' : '')
            + '</td></tr>';
    });
    html += '</table>';
    target.innerHTML = html;
    target.querySelectorAll('.invoice-cell-input').forEach(input => {
        input.addEventListener('input', () => {
            const index = Number(input.dataset.index);
            invoiceState.rows[index][input.dataset.field] = input.value;
            invoiceRevalidate();
            document.getElementById('invoice-download-link').style.display = 'none';
            document.getElementById('invoice-log-link').style.display = 'none';
        });
    });
    target.querySelectorAll('.invoice-row-action').forEach(button => {
        button.addEventListener('click', () => {
            const index = Number(button.dataset.index);
            if (button.dataset.action === 'duplicate') {
                const source = invoiceState.rows[index];
                const copy = Object.assign({}, source, {
                    '従業員': '',
                    '管理番号': '',
                    '備考': '',
                    _group_id: source._group_id + '-split-' + Date.now(),
                    _sources: (source._sources || []).slice(),
                    _warnings: ['合算請求書を手動で分割した行です。金額と税額の合計を元PDFと照合してください。'],
                    _errors: [],
                    _manual_added: true,
                });
                invoiceState.rows.splice(index + 1, 0, copy);
            } else if (button.dataset.action === 'delete' && invoiceState.rows[index]._manual_added) {
                invoiceState.rows.splice(index, 1);
            }
            document.getElementById('invoice-count-rows').textContent = String(invoiceState.rows.length);
            document.getElementById('invoice-count-commute').textContent = String(
                invoiceState.rows.filter(row => row._row_type === 'commute').length);
            document.getElementById('invoice-download-link').style.display = 'none';
            document.getElementById('invoice-log-link').style.display = 'none';
            invoiceRenderRows();
        });
    });
    invoiceRevalidate();
}

// スポット案件など、共有フォルダのPDFからは拾えない請求分を手で足すための空行。
// 発生日・支払期日・勘定科目・税区分だけ既定を入れておき、あとは人が埋める。
function invoiceAddBlankRow() {
    const sample = invoiceState.rows.find(row => row._row_type === 'main') || {};
    invoiceState.rows.push({
        '収支区分': '収入', '管理番号': '', '発生日': sample['発生日'] || '',
        '支払期日': sample['支払期日'] || '', '取引先': '', '勘定科目': '売上高',
        '税区分': '課税売上10%', '金額': '', '税計算区分': '内税', '税額': '',
        '備考': '', '品目': '', '部門': '', 'メモタグ（複数指定可、カンマ区切り）': '',
        '従業員': '',
        _row_type: 'main', _group_id: 'manual-' + Date.now(), _sources: [],
        _warnings: ['手で追加した行です。元の請求書と金額・税額を照合してください。'],
        _errors: [], _manual_added: true,
    });
    document.getElementById('invoice-count-rows').textContent = String(invoiceState.rows.length);
    document.getElementById('invoice-download-link').style.display = 'none';
    document.getElementById('invoice-log-link').style.display = 'none';
    invoiceRenderRows();
}

function invoiceRenderNotices(data) {
    const target = document.getElementById('invoice-notices');
    if (!target) return;
    let html = '';
    if (data.ignored && data.ignored.length) {
        html += '<div class="alert alert-warning"><b>修正版を優先して除外: '
            + data.ignored.length + '件</b><details><summary>一覧</summary>'
            + data.ignored.map(item => '<div>' + escapeHtml(item.file) + '（'
                + escapeHtml(item.reason) + '）</div>').join('') + '</details></div>';
    }
    if (data.missing_roots && data.missing_roots.length) {
        html += '<div class="alert alert-warning">見つからない対象フォルダ: '
            + data.missing_roots.length + '件</div>';
    }
    const failures = (data.scan_errors || []).concat(data.parse_errors || []);
    if (failures.length) {
        html += '<div class="alert alert-error"><b>読み取れなかったPDF/フォルダがあります</b>'
            + failures.map(message => '<div>' + escapeHtml(message) + '</div>').join('') + '</div>';
    }
    html += '<div class="hint">社員番号の参照元: ' + escapeHtml(data.sales_book || '見つかりません') + '</div>';
    target.innerHTML = html;
}

const invoiceMonthInput = document.getElementById('invoice-month');
if (invoiceMonthInput) {
    const previous = new Date();
    previous.setMonth(previous.getMonth() - 1);
    invoiceMonthInput.value = previous.getFullYear() + '-'
        + String(previous.getMonth() + 1).padStart(2, '0');
}

const invoicePreviewBtn = document.getElementById('invoice-preview-btn');
if (invoicePreviewBtn) {
    invoicePreviewBtn.addEventListener('click', async () => {
        const month = (document.getElementById('invoice-month').value || '').trim();
        const status = document.getElementById('invoice-status');
        invoiceShowError([]);
        if (!/^\d{4}-\d{2}$/.test(month)) {
            invoiceShowError('請求対象月を選択してください');
            return;
        }
        invoicePreviewBtn.disabled = true;
        status.textContent = '28フォルダから請求書PDFを探して読み取っています…';
        document.getElementById('invoice-result-area').style.display = 'none';
        document.getElementById('invoice-download-link').style.display = 'none';
        document.getElementById('invoice-log-link').style.display = 'none';
        try {
            const form = new FormData();
            form.append('month', month);
            const response = await fetch('/invoice_preview', { method: 'POST', body: form });
            const data = await response.json();
            if (!data.success) {
                invoiceShowError(data.errors || ['請求書を読み取れませんでした']);
                status.textContent = '';
                return;
            }
            invoiceState.month = month;
            invoiceState.rows = data.rows || [];
            invoiceState.signature = data.signature || '';
            invoiceState.preview = data;
            document.getElementById('invoice-count-files').textContent = String((data.selected_files || []).length);
            document.getElementById('invoice-count-rows').textContent = String(invoiceState.rows.length);
            document.getElementById('invoice-count-commute').textContent = String(
                invoiceState.rows.filter(row => row._row_type === 'commute').length);
            invoiceRenderNotices(data);
            invoiceRenderRows();
            document.getElementById('invoice-result-area').style.display = 'block';
            status.textContent = invoiceState.rows.length
                ? '読み取り完了。赤い行を修正してからCSVを作成してください。'
                : '対象の請求書PDFが見つかりませんでした。';
        } catch (error) {
            invoiceShowError('通信に失敗しました: ' + error);
            status.textContent = '';
        } finally {
            invoicePreviewBtn.disabled = false;
        }
    });
}

const invoiceAddRowBtn = document.getElementById('invoice-add-row-btn');
if (invoiceAddRowBtn) {
    invoiceAddRowBtn.addEventListener('click', invoiceAddBlankRow);
}

const invoiceExportBtn = document.getElementById('invoice-export-btn');
if (invoiceExportBtn) {
    invoiceExportBtn.addEventListener('click', async () => {
        if (invoiceRevalidate() > 0) {
            invoiceShowError('未入力・不正な行があるためCSVを作成できません');
            return;
        }
        const status = document.getElementById('invoice-export-status');
        invoiceShowError([]);
        invoiceExportBtn.disabled = true;
        status.textContent = 'CSVを作成しています…';
        try {
            const response = await fetch('/invoice_export', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    month: invoiceState.month,
                    rows: invoiceState.rows,
                    signature: invoiceState.signature,
                }),
            });
            const data = await response.json();
            if (!data.success) {
                invoiceShowError(data.errors || ['CSVを作成できませんでした']);
                status.textContent = '';
                return;
            }
            const download = document.getElementById('invoice-download-link');
            const log = document.getElementById('invoice-log-link');
            download.href = data.csv_url;
            download.textContent = data.csv_name + ' をダウンロード';
            download.style.display = '';
            log.href = data.log_url;
            log.style.display = '';
            status.textContent = data.row_count + '行のCSVを作成しました。';
        } catch (error) {
            invoiceShowError('通信に失敗しました: ' + error);
            status.textContent = '';
        } finally {
            invoiceExportBtn.disabled = false;
            invoiceRevalidate();
        }
    });
}
