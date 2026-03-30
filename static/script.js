'use strict';

// ドラッグ&ドロップ + ファイル選択
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
        const files = e.dataTransfer.files;
        // DataTransferでinputのfilesを置き換え
        const dt = new DataTransfer();
        for (const f of files) dt.items.add(f);
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

// フォーム送信
const form = document.getElementById('upload-form');
const runBtn = document.getElementById('run-btn');
const progressArea = document.getElementById('progress-area');
const progressBar = document.getElementById('progress-bar');
const progressMessage = document.getElementById('progress-message');
const errorArea = document.getElementById('error-area');
const resultArea = document.getElementById('result-area');

let progressStep = 0;

form.addEventListener('submit', async (e) => {
    e.preventDefault();
    startProcessing();

    const formData = new FormData(form);
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
    } catch (err) {
        showError('通信エラーが発生しました: ' + err.message);
        stopProcessing();
    }
});

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
        progressStep = Math.min(progressStep + 15, 85);
        progressBar.style.width = progressStep + '%';
        progressMessage.textContent = data.message || '処理中...';
    } else if (eventType === 'done') {
        progressBar.style.width = '100%';
        progressMessage.textContent = '完了！';
        setTimeout(() => {
            progressArea.style.display = 'none';
            showResult(data);
        }, 400);
        stopProcessing(false);
    } else if (eventType === 'error') {
        showError(data.message || 'エラーが発生しました');
        stopProcessing();
    }
}

function startProcessing() {
    progressStep = 5;
    progressBar.style.width = progressStep + '%';
    progressMessage.textContent = '処理を開始しています...';
    progressArea.style.display = 'block';
    errorArea.style.display = 'none';
    resultArea.style.display = 'none';
    runBtn.disabled = true;
    runBtn.textContent = '処理中...';
}

function stopProcessing(hideProgress = true) {
    runBtn.disabled = false;
    runBtn.textContent = 'チェック実行';
    if (hideProgress) progressArea.style.display = 'none';
}

function showError(msg) {
    errorArea.innerHTML = msg;
    errorArea.style.display = 'block';
    resultArea.style.display = 'none';
}

function showResult(data) {
    const { summary, table, excel_filename, unsubmitted } = data;

    document.getElementById('cnt-ok').textContent = summary.ok;
    document.getElementById('cnt-ng').textContent = summary.ng;
    document.getElementById('cnt-caution').textContent = summary.caution;
    document.getElementById('cnt-missing').textContent = summary.missing;

    const allOkMsg = document.getElementById('all-ok-msg');
    const ngTableArea = document.getElementById('ng-table-area');

    if (summary.ng === 0 && summary.caution === 0) {
        allOkMsg.style.display = 'block';
        ngTableArea.style.display = 'none';
    } else {
        allOkMsg.style.display = 'none';
        ngTableArea.style.display = 'block';
        renderTable(table);
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
            row.judgment, row.detail
        ];
        cells.forEach((val, idx) => {
            const td = document.createElement('td');
            td.textContent = val !== null && val !== undefined ? val : '';
            if (idx === 8) {
                td.className = judgeClass[row.judgment] || '';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    }
}
