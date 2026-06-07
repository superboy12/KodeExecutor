# KodeExecutor — Code to DOCX Generator 📄💻

Aplikasi web sederhana, modern, dan estetis untuk mengonversi kode sumber (*source code*) Anda menjadi file dokumen Word (`.docx`) secara instan dan rapi. 100% diproses di browser Anda secara aman dan tanpa batasan karakter!

## Fitur Utama
* **Instan & Aman:** Pemrosesan secara lokal di browser, kode Anda tidak pernah dikirim ke server luar.
* **Format Monospace Rapi:** Menggunakan font Courier New yang ideal untuk kode di Microsoft Word.
* **Dark Mode UI:** Tampilan modern, bersih, dan memanjakan mata dengan nuansa editor teks.
* **Tanpa Limit:** Bebas memproses baris kode sebanyak apa pun.

---

## Panduan Penggunaan (Untuk Pengguna)

1. Buka website **KodeExecutor** di browser.
2. Salin (*copy*) kode sumber lengkap yang ingin Anda konversi.
3. Tempel (*paste*) kode tersebut ke dalam area editor yang disediakan.
4. Klik tombol **"Generate DOCX"**.
5. File `.docx` akan otomatis terunduh ke perangkat Anda dengan format nama `SourceCode.docx`.

---

## Cara Menjalankan Secara Lokal (Development)

Jika Anda ingin menjalankan aplikasi ini di komputer Anda sendiri:

### Prasyarat
Pastikan Anda sudah menginstal **Node.js** di komputer Anda.

### Langkah-Langkah:
1. Clone repositori ini:
   ```bash
   git clone https://github.com/superboy12/KodeExecutor.git
   cd KodeExecutor
   ```
2. Instal semua dependensi:
   ```bash
   npm install
   ```
3. Jalankan server pengembangan:
   ```bash
   npm run dev
   ```
4. Buka browser dan akses **`http://localhost:5173/`**.

---

## Cara Deploy ke GitHub Pages (Hosting Gratis)

Aplikasi ini sudah dikonfigurasi untuk dideploy ke GitHub Pages dengan mudah.

1. Instal package `gh-pages` di proyek Anda (jika belum):
   ```bash
   npm install gh-pages --save-dev
   ```
2. Tambahkan script deploy di `package.json`:
   ```json
   "scripts": {
     "predeploy": "npm run build",
     "deploy": "gh-pages -d dist"
   }
   ```
3. Jalankan perintah deploy:
   ```bash
   npm run deploy
   ```
   *Aplikasi Anda akan otomatis di-build dan di-upload ke branch `gh-pages` di GitHub!*
