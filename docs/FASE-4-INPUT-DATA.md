# Fase 4 — Input dan Validasi Data

Status: selesai untuk validasi folder input.

## Hasil

- User cukup memilih satu folder data.
- Sistem otomatis mencari satu file Excel.
- Periode dibaca dari kolom Excel.
- PDF dipetakan berdasarkan periode dan pola nama.
- PDF yang hilang ditampilkan kepada user.
- Periode tanpa PDF tidak ikut diproses dan tidak dianggap lengkap.
- Periode yang tersedia tetap dapat diaudit.

## Verifikasi

- Validasi folder data utama menemukan satu Excel dan 8 periode tersedia.
- Dry-run parser tetap berjalan menggunakan periode yang tervalidasi.
- Modul Python dan GUI berhasil dikompilasi/import.

## Catatan

File masih dibaca dari folder sumber yang dipilih user. Penyalinan ke folder `uploads` aplikasi akan diputuskan saat alur upload dan packaging `.exe` diselesaikan.
