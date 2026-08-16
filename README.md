# OmniFile — Open Almost Anything

Web app viewer file universal (dokumen, gambar, audio, video, spreadsheet, kode) yang berjalan sepenuhnya di browser. Dibuat dengan React + Vite + Tailwind.

## Cara deploy ke GitHub Pages

### 1. Siapkan repo
```bash
# di folder omnifile ini
git init
git add .
git commit -m "Initial commit: OmniFile"
```
Buat repo baru di GitHub (misal `omnifile`), lalu:
```bash
git branch -M main
git remote add origin https://github.com/USERNAME/omnifile.git
git push -u origin main
```

### 2. Sesuaikan `base` path
Buka `vite.config.js`, ganti:
```js
base: "/omnifile/",
```
dengan `/nama-repo-kamu/` (harus sama persis dengan nama repo GitHub, termasuk garis miring di depan & belakang).

Kalau nanti repo ini dipakai sebagai **user/organization page** (nama repo `USERNAME.github.io`), ganti `base` menjadi `"/"`.

### 3. Aktifkan GitHub Pages via GitHub Actions
Workflow deploy otomatis sudah disediakan di `.github/workflows/deploy.yml` — akan build & deploy setiap kamu push ke branch `main`.

Di GitHub:
1. Buka repo → **Settings → Pages**
2. Di bagian **Build and deployment → Source**, pilih **GitHub Actions**
3. Push ke `main` (atau jalankan workflow manual lewat tab **Actions → Deploy to GitHub Pages → Run workflow**)
4. Tunggu 1–2 menit, situs akan tersedia di:
   `https://USERNAME.github.io/omnifile/`

### 4. Coba lokal dulu sebelum deploy (opsional tapi disarankan)
```bash
npm install
npm run dev
```
Buka `http://localhost:5173`. Untuk cek versi production:
```bash
npm run build
npm run preview
```

### Alternatif: deploy manual pakai `gh-pages` (tanpa GitHub Actions)
```bash
npm install
npm run build
npx gh-pages -d dist
```
Lalu di **Settings → Pages**, pilih source branch `gh-pages`.

## Catatan
- File `public/icons/icon-192.png` dan `icon-512.png` masih placeholder sederhana — ganti dengan logo asli kalau ada.
- `manifest.json` sudah cukup untuk "Add to Home Screen" di Android/Chrome, tapi belum ada `service-worker.js` untuk mode offline penuh — bisa ditambahkan pakai plugin `vite-plugin-pwa` kalau dibutuhkan.
- Format ODT/RTF/PPT/PPTX/ODP dan arsip ZIP/7Z/TAR belum punya viewer (belum ada library ringan yang mendukungnya di browser) — file jenis ini akan tampil sebagai "File Type Detected" dengan tombol download.
