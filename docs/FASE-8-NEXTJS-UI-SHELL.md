# Fase 8 — Next.js UI Shell Lokal

Status: selesai untuk prototype UI lokal.

## Hasil

- App Next.js lokal dibuat di folder `web`.
- Layout mengikuti referensi dashboard: sidebar navy, konten putih, aksen biru, kartu ringkasan, filter kategori, dan rail periode.
- Kategori `Balance Sheet`, `Profit & Loss`, dan `Cash Flow` dapat difilter.
- Data yang tampil masih fixture UI; koneksi ke engine Python dilakukan pada fase integrasi berikutnya.
- Build production Next.js berhasil.
- Preview berjalan di `http://127.0.0.1:3000`.

## Verifikasi

- `npm install`: berhasil.
- `npm run build`: berhasil.
- Filter `Profit & Loss` diuji di browser dan hanya menampilkan baris kategori tersebut.
