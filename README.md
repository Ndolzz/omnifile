# OmniFile - Open Almost Anything

<div align="center">

OmniFile
Web app viewer file universal yang berjalan sepenuhnya di browser
Buka hampir semua jenis file

[GitHub Pages](https://ndolzz.github.io/omnifile/)
[License MIT](https://opensource.org/licenses/MIT)

</div>


## Fitur

Dokumen: PDF, DOCX, TXT, Markdown
Spreadsheet: XLSX, CSV, ODS dengan preview tabel interaktif
Gambar: JPG, PNG, WEBP, GIF, SVG, TIFF, AVIF
Audio: MP3, WAV, OGG, FLAC, M4A, AAC, Opus
Video: MP4, WEBM, MKV, MOV, AVI
Arsip: ZIP, 7Z, TAR, GZ
Code: Semua format kode

## Privasi

100 Local - Semua pemrosesan di browser
Tidak ada upload - File tidak meninggalkan perangkat
Open Source - Kode terbuka

## Cepat Mulai

Online
[https://ndolzz.github.io/omnifile/](https://ndolzz.github.io/omnifile/)

Local

    git clone https://github.com/Ndolzz/omnifile.git
    cd omnifile
    npm install
    npm run dev

Buka http://localhost:5173 di browser

## Struktur Project

omnifile
public
  icons
  manifest.json
src
  App.jsx
  index.css
  main.jsx
.github
  workflows
    build.yml
    deploy.yml
index.html
package.json
tailwind.config.js
vite.config.js
README.md

## Tech Stack

Frontend: React 18 + Vite 5
Styling: Tailwind CSS 3
Icons: Lucide React
Docx: Mammoth.js
Spreadsheet: SheetJS (xlsx)
CSV: Papa Parse
Build: Vite + GitHub Actions

## Fitur UI

Dark/Light Mode - Toggle tema dengan mudah
Drag and Drop - Seret file langsung ke area upload
Multiple Files - Buka banyak file sekaligus
Responsive - Tampilan optimal di desktop dan mobile
File Information Panel - Lihat detail file yang dibuka

## Troubleshooting

File tidak bisa dibuka
Pastikan format file didukung
Coba refresh browser
Pastikan file tidak rusak

Deploy gagal
Pastikan base di vite.config.js sudah benar
Pastikan workflow GitHub Actions sudah enabled
Cek logs di tab Actions

Build error

    rm -rf node_modules package-lock.json
    npm install
    npm run build

## Lisensi

MIT License - Bebas digunakan, dimodifikasi, dan didistribusikan

## Kontribusi

Kontribusi selalu diterima
Issue - untuk melaporkan bug atau request fitur
Pull Request - untuk kontribusi kode

## Terima Kasih

React - Library UI
Vite - Build tool
Tailwind CSS - CSS framework
Lucide - Icons
Mammoth.js - DOCX to HTML
SheetJS - Spreadsheet parsing
Papa Parse - CSV parsing


Dibuat oleh [Ndolzz](https://github.com/Ndolzz)
Open almost anything, anywhere