const API = 'api/';
let currentFile = null;
let currentSheet = 0;
let sheetNames = [];
let tableData = [];
let saveTimeout = null;

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// DOM refs
const uploadSection = $('#upload-section');
const editorSection = $('#editor-section');
const fileInput = $('#file-input');
const dropZone = $('#drop-zone');
const fileName = $('#file-name');
const saveStatus = $('#save-status');
const sheetTabs = $('#sheet-tabs');
const tableHead = $('#table-head');
const tableBody = $('#table-body');
const loading = $('#loading-overlay');
const toast = $('#toast');

// Upload handlers
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) uploadFile(e.target.files[0]);
});

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) uploadFile(e.dataTransfer.files[0]);
});

// Buttons
$('#btn-download').addEventListener('click', downloadFile);
$('#btn-close').addEventListener('click', closeEditor);
$('#btn-add-row').addEventListener('click', addRow);
$('#btn-add-col').addEventListener('click', addColumn);

async function uploadFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(ext)) {
        showToast('Format file tidak didukung', 'error');
        return;
    }

    showLoading(true);
    const form = new FormData();
    form.append('excel_file', file);

    try {
        const res = await fetch(API + 'upload.php', { method: 'POST', body: form });
        const data = await res.json();

        if (!data.success) throw new Error(data.message || 'Upload gagal');

        currentFile = data.filename;
        sheetNames = data.sheets;
        currentSheet = 0;

        fileName.textContent = file.name;
        renderSheetTabs();
        await loadSheet(0);

        uploadSection.style.display = 'none';
        editorSection.style.display = 'block';
        showToast('File berhasil dimuat!', 'success');
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        showLoading(false);
    }
}

function renderSheetTabs() {
    sheetTabs.innerHTML = '';
    sheetNames.forEach((name, i) => {
        const tab = document.createElement('div');
        tab.className = 'sheet-tab' + (i === currentSheet ? ' active' : '');
        tab.textContent = name;
        tab.addEventListener('click', () => switchSheet(i));
        sheetTabs.appendChild(tab);
    });
}

async function switchSheet(index) {
    currentSheet = index;
    renderSheetTabs();
    await loadSheet(index);
}

async function loadSheet(index) {
    showLoading(true);
    try {
        const res = await fetch(`${API}read.php?file=${encodeURIComponent(currentFile)}&sheet=${index}`);
        const data = await res.json();

        if (!data.success) throw new Error(data.message || 'Gagal membaca sheet');

        tableData = data.data;
        renderTable(tableData);
        setSaveStatus('saved');
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        showLoading(false);
    }
}

function renderTable(data) {
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    if (!data || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="100" style="padding:24px;text-align:center;color:#94a3b8;">Sheet kosong</td></tr>';
        return;
    }

    // Find max columns
    let maxCols = 0;
    data.forEach(row => { if (row.length > maxCols) maxCols = row.length; });

    // Header row: column letters
    const headRow = document.createElement('tr');
    const cornerTh = document.createElement('th');
    cornerTh.className = 'row-num';
    cornerTh.textContent = '#';
    headRow.appendChild(cornerTh);

    for (let c = 0; c < maxCols; c++) {
        const th = document.createElement('th');
        th.textContent = colLetter(c);
        headRow.appendChild(th);
    }

    // Extra th for delete column
    const delTh = document.createElement('th');
    delTh.className = 'row-num';
    delTh.style.minWidth = '40px';
    delTh.style.width = '40px';
    headRow.appendChild(delTh);
    tableHead.appendChild(headRow);

    // Body rows
    data.forEach((row, rowIdx) => {
        const tr = document.createElement('tr');

        // Row number
        const numTd = document.createElement('td');
        numTd.className = 'row-num-cell';
        numTd.textContent = rowIdx + 1;
        tr.appendChild(numTd);

        for (let colIdx = 0; colIdx < maxCols; colIdx++) {
            const td = document.createElement('td');
            const val = colIdx < row.length ? (row[colIdx] ?? '') : '';

            const display = document.createElement('div');
            display.className = 'cell-display';
            display.textContent = val;
            display.addEventListener('dblclick', () => startEdit(td, rowIdx, colIdx, val));
            td.appendChild(display);
            tr.appendChild(td);
        }

        // Delete button cell
        const delTd = document.createElement('td');
        delTd.className = 'row-num-cell';
        const delBtn = document.createElement('button');
        delBtn.className = 'delete-row-btn';
        delBtn.textContent = '\u2715';
        delBtn.title = 'Hapus baris';
        delBtn.addEventListener('click', () => deleteRow(rowIdx));
        delTd.appendChild(delBtn);
        tr.appendChild(delTd);

        tableBody.appendChild(tr);
    });
}

function startEdit(td, rowIdx, colIdx, currentVal) {
    if (td.querySelector('.cell-input')) return;

    const input = document.createElement('textarea');
    input.className = 'cell-input';
    input.value = currentVal;
    input.rows = 1;

    td.innerHTML = '';
    td.appendChild(input);
    input.focus();
    input.select();

    // Auto-resize
    const resize = () => {
        input.style.height = 'auto';
        input.style.height = input.scrollHeight + 'px';
    };
    input.addEventListener('input', resize);
    resize();

    const finishEdit = () => {
        const newVal = input.value;
        const display = document.createElement('div');
        display.className = 'cell-display';
        display.textContent = newVal;
        display.addEventListener('dblclick', () => startEdit(td, rowIdx, colIdx, newVal));

        td.innerHTML = '';
        td.appendChild(display);

        if (newVal !== String(currentVal)) {
            // Ensure tableData has enough rows/cols
            while (tableData.length <= rowIdx) tableData.push([]);
            while (tableData[rowIdx].length <= colIdx) tableData[rowIdx].push('');
            tableData[rowIdx][colIdx] = newVal;
            scheduleSave(rowIdx, colIdx, newVal);
        }
    };

    input.addEventListener('blur', finishEdit);
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            input.blur();
        }
        if (e.key === 'Escape') {
            input.value = currentVal;
            input.blur();
        }
        if (e.key === 'Tab') {
            e.preventDefault();
            input.blur();
            // Move to next cell
            const nextCol = e.shiftKey ? colIdx - 1 : colIdx + 1;
            const cells = td.parentElement.querySelectorAll('td:not(.row-num-cell)');
            if (nextCol >= 0 && nextCol < cells.length - 1) {
                const nextDisplay = cells[nextCol].querySelector('.cell-display');
                if (nextDisplay) nextDisplay.dispatchEvent(new Event('dblclick'));
            }
        }
    });
}

function scheduleSave(row, col, value) {
    setSaveStatus('saving');
    clearTimeout(saveTimeout);
    saveTimeout = setTimeout(() => saveCell(row, col, value), 500);
}

async function saveCell(row, col, value) {
    try {
        const res = await fetch(API + 'save.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                file: currentFile,
                sheet: currentSheet,
                row: row,
                col: col,
                value: value
            })
        });
        const data = await res.json();
        if (!data.success) throw new Error(data.message);
        setSaveStatus('saved');
    } catch (err) {
        setSaveStatus('error');
        showToast('Gagal menyimpan: ' + err.message, 'error');
    }
}

async function addRow() {
    const newRow = [];
    if (tableData.length > 0) {
        const maxCols = Math.max(...tableData.map(r => r.length));
        for (let i = 0; i < maxCols; i++) newRow.push('');
    } else {
        newRow.push('');
    }
    tableData.push(newRow);
    renderTable(tableData);

    setSaveStatus('saving');
    try {
        const res = await fetch(API + 'add_row.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                file: currentFile,
                sheet: currentSheet
            })
        });
        const data = await res.json();
        if (!data.success) throw new Error(data.message);
        setSaveStatus('saved');
        showToast('Baris ditambahkan', 'success');
    } catch (err) {
        setSaveStatus('error');
        showToast('Gagal: ' + err.message, 'error');
    }
}

async function addColumn() {
    tableData.forEach(row => row.push(''));
    renderTable(tableData);

    setSaveStatus('saving');
    try {
        const res = await fetch(API + 'add_col.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                file: currentFile,
                sheet: currentSheet
            })
        });
        const data = await res.json();
        if (!data.success) throw new Error(data.message);
        setSaveStatus('saved');
        showToast('Kolom ditambahkan', 'success');
    } catch (err) {
        setSaveStatus('error');
        showToast('Gagal: ' + err.message, 'error');
    }
}

async function deleteRow(rowIdx) {
    if (!confirm(`Hapus baris ${rowIdx + 1}?`)) return;

    tableData.splice(rowIdx, 1);
    renderTable(tableData);

    setSaveStatus('saving');
    try {
        const res = await fetch(API + 'delete_row.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                file: currentFile,
                sheet: currentSheet,
                row: rowIdx
            })
        });
        const data = await res.json();
        if (!data.success) throw new Error(data.message);
        setSaveStatus('saved');
        showToast('Baris dihapus', 'success');
    } catch (err) {
        setSaveStatus('error');
        showToast('Gagal: ' + err.message, 'error');
    }
}

function downloadFile() {
    if (!currentFile) return;
    window.location.href = `${API}download.php?file=${encodeURIComponent(currentFile)}`;
}

function closeEditor() {
    if (!confirm('Tutup editor? Pastikan sudah download file yang diedit.')) return;
    uploadSection.style.display = 'block';
    editorSection.style.display = 'none';
    currentFile = null;
    currentSheet = 0;
    sheetNames = [];
    tableData = [];
    fileInput.value = '';
}

// Helpers
function colLetter(n) {
    let s = '';
    while (n >= 0) {
        s = String.fromCharCode(65 + (n % 26)) + s;
        n = Math.floor(n / 26) - 1;
    }
    return s;
}

function setSaveStatus(status) {
    saveStatus.className = 'save-status';
    switch (status) {
        case 'saving':
            saveStatus.textContent = 'Menyimpan...';
            saveStatus.classList.add('saving');
            break;
        case 'saved':
            saveStatus.textContent = 'Tersimpan';
            break;
        case 'error':
            saveStatus.textContent = 'Error';
            saveStatus.classList.add('error');
            break;
    }
}

function showLoading(show) {
    loading.style.display = show ? 'flex' : 'none';
}

function showToast(message, type = '') {
    toast.textContent = message;
    toast.className = 'toast show ' + type;
    setTimeout(() => { toast.className = 'toast'; }, 3000);
}
