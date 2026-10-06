# Fase 6B — PDF Koreksi Berbasis Koordinat

Status: selesai untuk pembuatan salinan PDF koreksi.

## Hasil

- Parser sekarang menyimpan halaman dan koordinat angka PDF.
- Tombol `PDF Koreksi` membuat file baru di folder `%LOCALAPPDATA%\\MappingKeuangan\\corrected-pdf`.
- Nilai pengganti diambil dari nilai Excel.
- PDF asli tidak pernah ditimpa.
- Baris yang tidak memiliki koordinat aman dilewati dan dilaporkan.
- File koreksi memakai suffix `_corrected.pdf`.

## Verifikasi

- Compile/import modul lulus.
- Parser tetap menghasilkan ringkasan audit yang sama.
- Output koreksi diuji pada salinan PDF, bukan file sumber.
