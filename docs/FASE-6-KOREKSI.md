# Fase 6 — Preview dan Koreksi Excel

Status: selesai untuk koreksi Excel yang aman; koreksi PDF langsung menunggu mapping koordinat.

## Hasil

- Sebelum koreksi, GUI menampilkan preview baris yang akan diubah.
- Nilai acuan koreksi berasal dari nilai PDF hasil audit saat ini.
- Cell formula dilewati dan dilaporkan.
- Backup Excel memakai nama bertimestamp dan tidak menimpa backup lama.
- Workbook disimpan ke file sementara terlebih dahulu, lalu dipindahkan secara atomik ke file tujuan.
- PDF asli belum diubah.

## Batasan yang disengaja

Koreksi angka langsung pada layout PDF belum diaktifkan. Parser perlu menyimpan halaman dan koordinat angka agar overlay/redaction tidak salah posisi. Tahap berikutnya harus menambahkan mapping koordinat dan menghasilkan file `_corrected.pdf` tanpa menyentuh PDF asli.

## Verifikasi

- Compile modul lulus.
- Dry-run parser tetap menghasilkan ringkasan yang sama.
- Koreksi Excel diuji pada salinan sementara, bukan workbook asli.
