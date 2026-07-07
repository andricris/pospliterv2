# Katalog Project Tools

Halaman statis untuk menampilkan semua project/tools dalam satu katalog.

## Isi

- `index.html` — Halaman utama (Tools Hub)
- `index.css` — Styles halaman utama
- `index.js` — JavaScript halaman utama (data tools + interaksi)
- `404.html` — Custom 404 page
- `404.css` — Styles 404 page
- `404.js` — JavaScript 404 page
- `assets/` — Favicon & OG image
- `netlify.toml` — Konfigurasi deploy Netlify

## Jalankan lokal

Buka `index.html` langsung di browser, atau pakai server statis sederhana.

## Upload ke GitHub

1. Buat repository baru.
2. Upload semua file di folder ini ke root repository.
3. Aktifkan GitHub Pages dari branch `main` dan folder `/root` jika ingin dipakai sebagai website statis.

## Deploy ke Netlify

1. Login ke Netlify.
2. Pilih **Add new site**.
3. Hubungkan repository GitHub Anda.
4. Build command kosong.
5. Publish directory: `.`

