# Fase 2 — Fondasi Aplikasi Lokal

Status: selesai untuk fondasi penyimpanan lokal.

## Hasil

- Lokasi database tidak lagi mengikuti folder `.exe`.
- Folder aplikasi dibuat otomatis di `%LOCALAPPDATA%\\MappingKeuangan`.
- Folder `uploads`, `exports`, dan `corrected-pdf` disiapkan otomatis.
- Aplikasi tidak perlu menulis ke `C:\\`, `Program Files`, atau folder instalasi.
- Parser audit yang sudah ada tidak diubah.

## Verifikasi fase

- Import modul berhasil.
- `ensure_app_dirs()` berhasil membuat folder aplikasi pada lokasi writable.
- `ensure_db()` memakai `database.sqlite3` di folder aplikasi lokal.
- Lokasi lama `map_keuangan_history.sqlite3` di root project tidak dipakai oleh GUI baru.

## Catatan

Riwayat database lama belum dimigrasikan otomatis. Migrasi riwayat akan diputuskan saat fase riwayat/export supaya format database final tidak berubah dua kali.
