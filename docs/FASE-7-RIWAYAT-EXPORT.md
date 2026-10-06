# Fase 7 — Riwayat Koreksi dan Export

Status: selesai untuk audit trail koreksi dan export yang sudah tersedia.

## Hasil

- Tabel `corrections` ditambahkan ke SQLite.
- Setiap koreksi dikaitkan dengan run audit dan username.
- Koreksi Excel menyimpan nilai lama, nilai baru, dan status formula yang dilewati.
- Koreksi PDF menyimpan periode, akun, jenis koreksi, output file, dan status.
- Export Excel/CSV hasil audit tetap tersedia.
- PDF asli tidak disimpan ulang ke database; hanya path output koreksi yang dicatat.

## Verifikasi

- Database dapat membuat tabel `corrections` tanpa menghapus tabel lama.
- Parser dan dry-run tetap menghasilkan ringkasan yang sama.
- Modul GUI berhasil dikompilasi dan di-import.
