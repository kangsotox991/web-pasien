<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Web Editor</title>
    <link rel="stylesheet" href="assets/css/style.css">
</head>
<body>
    <div class="container">
        <header>
            <h1>Excel Web Editor</h1>
            <p class="subtitle">Upload, edit, dan download file Excel langsung dari browser</p>
        </header>

        <div id="upload-section" class="card">
            <form id="upload-form" enctype="multipart/form-data">
                <div class="upload-area" id="drop-zone">
                    <div class="upload-icon">&#128196;</div>
                    <p>Drag & drop file Excel di sini</p>
                    <p class="upload-hint">atau</p>
                    <label class="btn btn-primary" for="file-input">Pilih File</label>
                    <input type="file" id="file-input" name="excel_file" accept=".xlsx,.xls,.csv" hidden>
                    <p class="upload-hint">Format: .xlsx, .xls, .csv</p>
                </div>
            </form>
        </div>

        <div id="editor-section" class="card" style="display:none;">
            <div class="editor-toolbar">
                <div class="toolbar-left">
                    <span id="file-name" class="file-label"></span>
                    <span id="save-status" class="save-status">Tersimpan</span>
                </div>
                <div class="toolbar-right">
                    <button class="btn btn-secondary" id="btn-add-row" title="Tambah Baris">+ Baris</button>
                    <button class="btn btn-secondary" id="btn-add-col" title="Tambah Kolom">+ Kolom</button>
                    <button class="btn btn-primary" id="btn-download">Download Excel</button>
                    <button class="btn btn-danger" id="btn-close">Tutup</button>
                </div>
            </div>

            <div class="sheet-tabs" id="sheet-tabs"></div>

            <div class="table-wrapper">
                <table id="excel-table">
                    <thead id="table-head"></thead>
                    <tbody id="table-body"></tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="loading-overlay" class="loading-overlay" style="display:none;">
        <div class="spinner"></div>
        <p>Memproses...</p>
    </div>

    <div id="toast" class="toast"></div>

    <script src="assets/js/app.js"></script>
</body>
</html>
