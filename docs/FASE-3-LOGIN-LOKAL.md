# Fase 3 — Login Lokal dan Manajemen User

Status: selesai untuk autentikasi lokal dasar.

## Hasil

- Tabel `users` ditambahkan ke SQLite aplikasi.
- Password disimpan dengan PBKDF2-HMAC-SHA256 dan salt acak; password plaintext tidak disimpan.
- User pertama diarahkan membuat akun admin saat aplikasi pertama dibuka.
- Login menolak username/password yang salah atau user nonaktif.
- Role `admin` dan `user` tersedia.
- Admin dapat menambah user dan mengaktifkan/nonaktifkan user lain.
- Admin yang sedang login tidak dapat menonaktifkan dirinya sendiri dari daftar user.

## Verifikasi

- `py_compile` untuk modul auth, path, parser, dan GUI: lulus.
- Self-check autentikasi: membuat admin, login benar, login salah, dan user nonaktif: lulus.
- Import GUI: lulus.

## Batasan fase ini

- Tampilan masih berada di GUI lokal yang ada; halaman Next.js akan menjadi pekerjaan fase UI berikutnya.
- Belum ada fitur ganti password dari layar user.
- Belum ada migrasi akun dari database lama.
