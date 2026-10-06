# Keputusan Arsitektur — Next.js Lokal

Status: dipilih.

## Keputusan

Aplikasi akan memakai Next.js sebagai UI lokal yang dibuka melalui browser, lalu dibungkus menjadi `.exe` launcher pada tahap packaging.

## Arsitektur

```text
MappingKeuangan.exe
        ↓
Next.js lokal di localhost
        ↓
Python engine lokal untuk Excel/PDF
        ↓
SQLite + folder %LOCALAPPDATA%\\MappingKeuangan
```

## Alasan

- Next.js lebih fleksibel untuk membuat UI seperti referensi dashboard Pinterest.
- Sidebar, tabel, filter, modal, login, dan status lebih mudah dirapikan.
- Parser Python yang sudah menghasilkan audit valid tetap digunakan.
- Tidak perlu MongoDB, Supabase, Vercel, atau koneksi internet.
- User tetap cukup double-click `.exe`.

## Arah visual

- Sidebar navy gelap dengan menu aktif biru.
- Background utama biru sangat muda atau putih.
- Panel konten putih dengan border tipis dan radius sedang.
- Tabel data ringan, dengan baris terpilih berwarna biru muda.
- Status menggunakan warna hijau, merah, dan kuning secara konsisten.
- Konten utama tetap fokus pada data, bukan ilustrasi dekoratif.

## Ruang lingkup migrasi

- UI Tkinter lama digantikan bertahap oleh Next.js.
- `map_keuangan.py`, parser PDF, parser Excel, dan aturan matching tetap menjadi referensi logic.
- Login dan database lokal dipindahkan ke service yang bisa dipakai UI Next.js.
- `.exe` final baru dibuat setelah UI Next.js dan engine lokal teruji.

## Tidak berubah

- Excel tetap sumber COA dan nilai koreksi.
- Akun yang hanya ada di PDF tidak ditampilkan di hasil utama.
- PDF asli tidak ditimpa.
- Data tetap lokal.
