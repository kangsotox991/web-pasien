# Excel Logbook Editor

Aplikasi desktop Python untuk membaca, mengedit, dan menyimpan file Excel (Logbook Harian).

## Fitur

- **Buka File Excel** — Load file .xlsx via dialog
- **Tampilan Tabel** — Data Excel ditampilkan dalam tabel dengan navigasi sheet tabs
- **Edit Cell** — Double-click cell untuk edit, perubahan disimpan ke workbook
- **Import Data Pasien** — Paste data pasien, otomatis rename sheet per tanggal (misal: "1 Maret", "4 Maret", dst)
- **Rename Sheet** — Rename sheet secara manual
- **Simpan / Simpan Sebagai** — Simpan perubahan ke file Excel

## Cara Pakai

### Install

```bash
pip install openpyxl
```

### Jalankan

```bash
python app.py
```

### Import Data Pasien

1. Buka file Excel template (Logbook Harian)
2. Klik menu **Tools → Import Data Pasien**
3. Paste data pasien dengan format:
   ```
   01/03/2026 Ny. K No. Reg 1026844 dx: infus, injeksi, ekg.
   01/03/2026 Sdr. A No. Reg 1027099 dx: injeksi, observasi.
   04/03/2026 Tn. S No. Reg 167396 dx: infus, injeksi, lab, ekg.
   ```
4. Klik **Import & Rename**
5. Sheet otomatis di-rename: Sheet1 → "1 Maret", Sheet2 → "4 Maret", dst

### Keyboard Shortcuts

- **Ctrl+O** — Buka file
- **Ctrl+S** — Simpan
- **Ctrl+Shift+S** — Simpan Sebagai
- **Double-click** — Edit cell
- **Enter** — Simpan edit cell
- **Escape** — Batal edit cell

## Requirements

- Python 3.8+
- openpyxl
- tkinter (sudah include di Python standar)
