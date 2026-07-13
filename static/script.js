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

// =============================================================================
// モード切替UI
// =============================================================================
function getCurrentMode() {
    const checked = document.querySelector('input[name="mode"]:checked');
    return checked ? checked.value : 'match';
}

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

    const isSchedule = mode === 'csv_export';
    const isExpense = mode === 'expense';
    const isMatch = !isSchedule && !isExpense;

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

    // ①フォームの実行ボタンはスケジュールモードのみ（スケジュールCSV作成）
    if (runBtn) {
        if (isSchedule) { runBtn.style.display = ''; runBtn.textContent = 'スケジュールCSVを作成'; }
        else runBtn.style.display = 'none';
    }

    // 勤怠チェックの導線（⚡一括＋手順2/手順3）は勤怠チェックモードのみ
    if (batchCompareCard) batchCompareCard.style.display = isMatch ? '' : 'none';
    if (monthlyCompareCard) monthlyCompareCard.style.display = isMatch ? '' : 'none';
    if (monthlyExportCard) monthlyExportCard.style.display = isMatch ? '' : 'none';

    // 経費チェックモード: 経費カードのみ表示
    if (expenseCard) expenseCard.style.display = isExpense ? '' : 'none';
}

document.querySelectorAll('input[name="mode"]').forEach(radio => {
    radio.addEventListener('change', () => applyModeUI(getCurrentMode()));
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


async function consumeSSEResponse(response) {
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
            processSSEPart(part);
        }
    }
}


function processSSEPart(part) {
    const lines = part.split('\n');
    let eventType = null;
    let data = null;
    for (const line of lines) {
        if (line.startsWith('event: ')) eventType = line.slice(7).trim();
        if (line.startsWith('data: ')) {
            try { data = JSON.parse(line.slice(6)); } catch { }
        }
    }
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
    progressBar.style.width = progressStep + '%';
    progressMessage.textContent = '処理を開始しています...';
    progressArea.style.display = 'block';
    errorArea.style.display = 'none';
    resultArea.style.display = 'none';
    if (csvExportArea) csvExportArea.style.display = 'none';
    runBtn.disabled = true;
    runBtn.textContent = '処理中...';
}

function stopProcessing(hideProgress = true) {
    runBtn.disabled = false;
    // モードに応じたボタンラベルに戻す
    runBtn.textContent = (getCurrentMode() === 'csv_export') ? 'スケジュールCSVを作成' : 'チェック実行';
    if (hideProgress) progressArea.style.display = 'none';
}

function showError(msg) {
    errorArea.innerHTML = msg;
    errorArea.style.display = 'block';
    resultArea.style.display = 'none';
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

    resultArea.style.display = 'block';
    resultArea.scrollIntoView({ behavior: 'smooth' });
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
                <th style="width:50%">氏名</th>
                <th style="width:30%">シフト件数</th>
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
        tr.innerHTML = `<td colspan="3">⚠️ 画像から従業員を抽出できませんでした。下の「+ 従業員を追加」から手動で追加してください（ただしシフトは画像から取得できないため、スケジュールCSVは作成できません）。</td>`;
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

    // 詳細行（日別シフト編集）を先に作る
    const detail = document.createElement('tr');
    detail.className = 'legend-employee-detail';
    detail.dataset.empIdx = empIdx;
    detail.style.display = 'none';
    const detailTd = document.createElement('td');
    detailTd.colSpan = 3;
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

    const { csv_files, missing_ids, merges, new_template_filename, new_template_count } = data;

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

    csvExportArea.style.display = 'block';
    csvExportArea.scrollIntoView({ behavior: 'smooth' });
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
