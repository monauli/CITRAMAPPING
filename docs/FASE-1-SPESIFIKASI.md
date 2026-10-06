# Fase 1 — Spesifikasi Aplikasi Mapping Keuangan

Status: draft untuk ditinjau sebelum masuk fase implementasi.

## Tujuan

Membangun aplikasi Windows lokal untuk membandingkan workbook Excel rekap keuangan dengan PDF laporan bulanan, menampilkan selisih secara rapi, dan membuat salinan PDF yang dikoreksi berdasarkan Excel.

## Bentuk aplikasi

- Produk akhir berupa `.exe`.
- User cukup double-click `.exe`; aplikasi menjalankan server lokal dan membuka browser otomatis.
- User tidak perlu memasang Python, Node.js, atau menjalankan terminal.
- Aplikasi tidak membutuhkan koneksi internet atau hak administrator.
- Tidak menggunakan Vercel, MongoDB, Supabase, atau database cloud.

## Login lokal

- Aplikasi mempunyai halaman login.
- Password disimpan sebagai hash, bukan teks biasa.
- Minimal tersedia role `admin` dan `user`.
- Admin dapat mengelola user; user biasa hanya menggunakan proses audit.
- Data login disimpan di SQLite lokal.

## Lokasi data aplikasi

Semua data yang perlu ditulis aplikasi disimpan di folder user yang writable:

```text
%LOCALAPPDATA%\\MappingKeuangan\\
├── database.sqlite3
├── uploads\\
├── exports\\
└── corrected-pdf\\
```

Aplikasi tidak menyimpan database, backup, export, atau riwayat di folder `.exe`, `C:\`, atau `Program Files`.

## Input data

User memilih satu folder data yang berisi:

- satu file `.xlsx` rekap keuangan;
- PDF laporan bulanan dengan pola nama yang sesuai;
- periode yang terbaca dari kolom tanggal di Excel.

Validasi wajib:

- Excel tidak ditemukan → tampilkan pesan yang jelas.
- Lebih dari satu Excel → minta user memilih atau memperbaiki folder.
- PDF periode tertentu tidak ditemukan → tampilkan periode yang hilang; proses tidak boleh diam-diam dianggap lengkap.
- File rusak atau format tidak didukung → tampilkan pesan yang jelas dan jangan menghapus file asli.

## Sumber kebenaran

Istilah “COA” digunakan untuk daftar akun/chart of accounts pada Excel.

- Excel menjadi sumber daftar akun, urutan akun, kategori, dan angka acuan koreksi.
- PDF menjadi sumber pembanding/dokumen laporan.
- Akun yang hanya ada di PDF tidak ditampilkan di tabel hasil utama.
- Akun yang ada di Excel tetapi tidak ditemukan di PDF tetap ditampilkan dengan status `HANYA_EXCEL`.
- Hasil audit utama mengikuti struktur Excel, bukan urutan PDF.

## Kategori tampilan

Kategori mengikuti sheet Excel:

1. `Balance Sheet`
2. `Profit & Loss`
3. `Cash Flow`

Navigasi hasil:

```text
Pilih bulan → pilih kategori → lihat tabel akun
```

Setiap tabel menampilkan:

- nama akun;
- nilai Excel;
- nilai PDF;
- selisih;
- status.

Baris total dan subtotal mengikuti penandaan Excel dan ditampilkan lebih menonjol.

## Status audit

Status utama yang ditampilkan:

- `COCOK` — nilai sama dalam batas toleransi.
- `BEDA` — akun ada di kedua sumber tetapi nilainya berbeda.
- `HANYA_EXCEL` — akun ada di Excel tetapi tidak ditemukan di PDF.

`HANYA_PDF` tidak masuk tabel utama, tetapi boleh tersedia sebagai informasi tambahan untuk pemeriksaan teknis.

Toleransi awal tetap mengikuti aplikasi sekarang: selisih absolut maksimal `1.00` dianggap `COCOK`.

## Koreksi Excel dan PDF

- Excel menjadi nilai acuan saat membuat koreksi.
- Koreksi PDF menghasilkan file baru dengan suffix `_corrected.pdf`.
- PDF asli tidak pernah ditimpa.
- Sebelum mengubah Excel, aplikasi membuat backup otomatis.
- Koreksi tidak boleh langsung menerapkan semua `BEDA` tanpa preview dan konfirmasi user.
- Cell Excel yang berisi formula tidak ditimpa otomatis; harus ditandai untuk pemeriksaan manual.
- Setiap koreksi dicatat di riwayat: user, waktu, sumber, periode, kategori, akun, nilai lama, dan nilai baru.

## Riwayat dan export

Aplikasi menyimpan:

- riwayat proses audit;
- ringkasan jumlah `COCOK`, `BEDA`, dan `HANYA_EXCEL`;
- detail hasil per periode dan kategori;
- riwayat koreksi.

Hasil dapat diekspor ke Excel atau CSV. File export masuk ke folder `exports` atau lokasi yang dipilih user.

## Batasan scope fase ini

Tidak termasuk:

- akses dari internet atau multi-komputer melalui link;
- sinkronisasi cloud;
- integrasi ERP/accounting system;
- OCR untuk PDF scan yang tidak memiliki text layer;
- perubahan isi PDF asli;
- otomatis memperbaiki semua selisih tanpa persetujuan user.

## Kriteria penerimaan fase berikutnya

Implementasi dianggap mengikuti Fase 1 jika:

1. User cukup menjalankan `.exe` dan melihat halaman login.
2. Aplikasi tetap dapat membuat database saat `.exe` berada di folder read-only.
3. User dapat memilih satu folder data dan mendapat validasi Excel/PDF yang jelas.
4. Hasil dapat dibuka per bulan dan per tiga kategori Excel.
5. Tabel utama hanya mengikuti akun dari Excel.
6. PDF asli tetap utuh setelah proses koreksi.
7. Nilai koreksi berasal dari Excel dan dapat ditinjau sebelum diterapkan.
8. Aplikasi dapat berjalan tanpa Python, Node.js, internet, atau akses administrator.

