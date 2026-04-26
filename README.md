# Excel Web Editor

Web application untuk membaca, mengedit, dan menyimpan file Excel langsung dari browser.

## Fitur

- **Upload Excel** - Drag & drop atau pilih file (.xlsx, .xls, .csv)
- **Tampilan Web** - Data Excel ditampilkan dalam tabel web yang bersih
- **Edit Inline** - Double-click cell untuk mengedit, perubahan langsung tersimpan ke file Excel
- **Multi-Sheet** - Support multiple sheet dengan tab navigation
- **Tambah/Hapus** - Tambah baris, tambah kolom, hapus baris
- **Download** - Download file Excel yang sudah diedit
- **Auto-save** - Setiap perubahan otomatis tersimpan ke file Excel

## Tech Stack

- **Backend**: PHP + [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet)
- **Frontend**: HTML + CSS + Vanilla JavaScript
- **Hosting**: Compatible dengan InfinityFree dan shared hosting PHP lainnya

## Instalasi

### Shared Hosting (InfinityFree)

1. Upload semua file ke root hosting (public_html)
2. Pastikan folder `uploads/` writable (chmod 755)
3. Pastikan folder `vendor/` ter-upload lengkap
4. Akses melalui domain

### Local Development

```bash
# Install dependencies
composer install

# Jalankan PHP built-in server
php -S localhost:8000

# Buka browser di http://localhost:8000
```

## Struktur File

```
web-pasien/
├── index.php           # Halaman utama
├── composer.json       # PHP dependencies
├── .htaccess          # Apache config
├── api/
│   ├── upload.php     # API upload file
│   ├── read.php       # API baca sheet
│   ├── save.php       # API simpan cell
│   ├── download.php   # API download file
│   ├── add_row.php    # API tambah baris
│   ├── add_col.php    # API tambah kolom
│   └── delete_row.php # API hapus baris
├── assets/
│   ├── css/style.css  # Stylesheet
│   └── js/app.js      # JavaScript
├── uploads/           # Folder file Excel (auto-created)
└── vendor/            # PHP dependencies (auto-generated)
```

## Keyboard Shortcuts

- **Double-click** - Edit cell
- **Enter** - Simpan dan keluar dari edit
- **Escape** - Batal edit
- **Tab** - Pindah ke cell berikutnya
- **Shift+Tab** - Pindah ke cell sebelumnya
