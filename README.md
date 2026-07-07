# 📦 PO Splitter

Split Purchase Order (PO) jadi beberapa PO berdasarkan persentase, dengan pembulatan ke kelipatan Master Box. Format Excel asli (warna, font, border, merged cells, lebar kolom) **100% terjaga**.

---

## Tech Stack

| Layer | Technology | Version |
|---|---|---|
| **Frontend** | React + Vite + Tailwind CSS | React 19, Vite 5 |
| **Backend** | Python Flask + openpyxl | Flask 3.1, openpyxl 3.1 |
| **Excel Processing** | openpyxl (pure Python) | Tidak butuh MS Excel / WPS |
| **Deployment** | Vercel (Serverless Functions) | Python 3.12 runtime |

---

## Cara Kerja Sistem

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER (Browser)                          │
│                                                                 │
│  1. Upload .xlsx                                                │
│  2. Pilih kolom (HNA, Master Box, PO Sisa, PO 1)               │
│  3. Atur jumlah PO & persentase                                 │
│  4. Klik Generate → download file hasil                         │
└──────────────────────────┬──────────────────────────────────────┘
                           │ HTTP (multipart/form-data)
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                    BACKEND (Flask + openpyxl)                    │
│                                                                 │
│  /api/preview                                                   │
│    ├── Baca sheet "PP" dari Excel                               │
│    ├── Cari header row (baris yang ada "po sisa")               │
│    ├── Ambil SEMUA header kolom → kirim ke frontend (dropdown)  │
│    └── Ambil SEMUA nilai kolom per produk → kirim ke frontend   │
│                                                                 │
│  /api/process                                                   │
│    ├── Buka file Excel dengan openpyxl                          │
│    ├── Unprotect sheet (jika di-protect)                        │
│    ├── Copy format dari kolom A header → kolom baru             │
│    ├── Tulis formula Excel asli (A1 notation):                  │
│    │   ├── PO 1..N-1: =MROUND((Sumber)*%, MasterBox)           │
│    │   ├── PO terakhir: =Sumber - SUM(PO 1..N-1)               │
│    │   └── TOTAL: =SUM(PO 1..N)                                │
│    ├── Tulis summary: =SUMPRODUCT(PO, HNA)                     │
│    └── Simpan & return file .xlsx                               │
└─────────────────────────────────────────────────────────────────┘
```

### Kenapa Format 100% Terjaga?

Karena kita **tidak membuat file baru dari nol**. Cara kerjanya:

1. **Copy file asli** → buka copy-nya dengan openpyxl
2. **Tambah kolom baru** di kolom AT, AU, AV, dst (setelah kolom terakhir data)
3. **Copy format** (font, fill, border, alignment, number format) dari kolom A header ke kolom baru
4. **Tulis formula Excel asli** (bukan hitungan Python) → formula di-recalculate saat file dibuka di Excel/WPS
5. **Simpan** → semua formatting yang sudah ada (warna, hidden columns, merged cells, conditional formatting) tetap utuh karena kita hanya menambah, tidak mengubah yang sudah ada

---

## Struktur Project

```
po-splitter/
├── server.py                ← Backend Flask (lokal)
├── api/
│   └── index.py             ← Backend Flask (Vercel serverless)
├── src/
│   ├── App.jsx              ← Frontend React (komponen utama)
│   ├── main.jsx             ← Entry point React
│   └── index.css            ← Tailwind CSS + custom styles
├── dist/                    ← Build output (sudah di-build)
│   ├── index.html
│   └── assets/
├── index.html               ← HTML template untuk Vite
├── package.json             ← Dependencies frontend
├── vite.config.js           ← Vite configuration
├── tailwind.config.js       ← Tailwind CSS configuration
├── postcss.config.js        ← PostCSS configuration
├── vercel.json              ← Vercel deployment config
├── requirements.txt         ← Python dependencies
└── .env                     ← Environment variables (lokal)
```

---

## Langkah-Langkah Setup Lokal

### Prerequisites

- **Python 3.10+** (tested on 3.14)
- **Node.js 18+** (tested on 22)
- **npm** atau **yarn**

### 1. Clone Repository

```bash
git clone https://github.com/YOUR_USERNAME/po-splitter.git
cd po-splitter
```

### 2. Install Python Dependencies

```bash
pip install flask flask-cors openpyxl
```

### 3. Install Node Dependencies

```bash
npm install
```

### 4. Start Backend (Terminal 1)

```bash
python server.py
```

Output:
```
PO Splitter Backend starting on http://localhost:5000
Using: openpyxl (pure Python) — no MS Excel/WPS required
```

### 5. Start Frontend (Terminal 2)

```bash
npm run dev
```

Output:
```
VITE v5.4.21  ready in 300ms
➜  Local:   http://localhost:5173/
```

### 6. Buka Browser

Akses **http://localhost:5173**

---

## Langkah-Langkah Penggunaan

### Step 1: Upload File Excel

1. Klik area upload di halaman utama
2. Pilih file `.xlsx` yang punya sheet **"PP"**
3. Sistem otomatis membaca semua kolom dan header

### Step 2: Konfigurasi

#### Pilih Kolom

| Dropdown | Default | Fungsi |
|---|---|---|
| **Sumber Data Split** | `AS — po sisa` | Kolom yang berisi qty yang mau di-split |
| **Kurangi Kolom PO 1** | ☐ NONAKTIF | Jika aktif: Qty = Sumber - PO 1 |
| **Kolom PO 1** | `AR — po 1 0,75` | Kolom yang berisi qty PO pertama |
| **Kolom HNA (Harga)** | `F — HNA UM` | Kolom harga satuan (untuk hitung total rupiah) |
| **Kolom Master Box** | `P — MASTER BOX` | Kolom kelipatan pembulatan |

#### Atur Jumlah PO & Persentase

1. Pilih jumlah PO: **2, 3, 4, 5, atau 6**
2. Atur persentase per PO (total harus **100%**)
3. Preview tabel langsung update otomatis

#### Dua Mode Split

**Mode 1: Split PO Sisa (default)**
- Sumber = kolom PO Sisa (sudah dikurangi PO 1 di Excel)
- Formula: `=MROUND(AS7 * 0.25, P7)`

**Mode 2: Split Purch Plan + PO 1**
- Aktifkan checkbox "Kurangi Kolom PO 1"
- Sumber = kolom Purch Plan, dikurangi PO 1
- Formula: `=MROUND((W7 - AR7) * 0.25, P7)`

### Step 3: Generate & Download

1. Klik **🚀 Generate & Download Excel**
2. File `.xlsx` otomatis terdownload
3. Buka di Excel atau WPS → formula otomatis di-recalculate

---

## API Endpoints

| Endpoint | Method | Fungsi |
|---|---|---|
| `/api/health` | GET | Health check |
| `/api/preview` | POST | Upload Excel → return headers + data produk (JSON) |
| `/api/process` | POST | Upload Excel + config → return file hasil (.xlsx) |

### Parameter `/api/process`

| Parameter | Type | Default | Contoh |
|---|---|---|---|
| `file` | file | required | (multipart file) |
| `numPO` | int | 4 | `4` |
| `percentages` | JSON array | required | `[25,25,25,25]` |
| `sourceCol` | string | `AS` | `AS` |
| `po1Col` | string | `AR` | `AR` |
| `enablePO1` | bool | `false` | `true` |
| `hnaCol` | string | `F` | `F` |
| `mbCol` | string | `P` | `P` |

### Contoh curl

```bash
curl -X POST http://localhost:5000/api/process \
  -F "file=@template.xlsx" \
  -F "numPO=4" \
  -F "percentages=[25,25,25,25]" \
  -F "sourceCol=AS" \
  -F "enablePO1=false" \
  -F "hnaCol=F" \
  -F "mbCol=P" \
  -o output.xlsx
```

---

## Deploy ke Vercel

### Struktur Deploy

```
po-splitter-deploy/
├── api/index.py         ← Flask API (serverless)
├── dist/                ← Frontend build
├── vercel.json          ← Routing config
├── requirements.txt     ← Python deps
└── package.json         ← (minimal, no build script)
```

### Cara Deploy

**Option 1: Upload Folder**
1. Buka https://vercel.com/new
2. Pilih **Upload**
3. Drag folder `po-splitter-deploy`
4. Klik **Deploy**

**Option 2: GitHub**
1. Push folder ke GitHub
2. Buka https://vercel.com/new → **Import Git Repository**
3. Pilih repo → **Deploy**

### vercel.json

```json
{
  "outputDirectory": "dist",
  "buildCommand": "",
  "rewrites": [
    { "source": "/api/(.*)", "destination": "/api/index.py" }
  ]
}
```

- `outputDirectory: "dist"` → serve frontend dari folder `dist/`
- `buildCommand: ""` → skip npm build (sudah di-build)
- `rewrites` → route `/api/*` ke Python serverless function

### Environment Variables

| Variable | Lokal (.env) | Vercel |
|---|---|---|
| `VITE_API_URL` | `http://localhost:5000` | (kosong = relative) |

---

## Formula Excel yang Dihasilkan

Asumsi: 4 PO, persentase 25% masing-masing, source=AS, MB=P

| Kolom | Header | Formula |
|---|---|---|
| AT | PO Sisa 1 (25%) | `=MROUND(AS7*0.25, P7)` |
| AU | PO Sisa 2 (25%) | `=MROUND(AS7*0.25, P7)` |
| AV | PO Sisa 3 (25%) | `=MROUND(AS7*0.25, P7)` |
| AW | PO Sisa 4 (25%) | `=AS7-(AT7+AU7+AV7)` |
| AX | TOTAL PO Sisa | `=AT7+AU7+AV7+AW7` |

Summary row (baris di atas header):
```
=SUMPRODUCT(AT7:AT114, F7:F114)   ← Total Rupiah PO 1
=SUMPRODUCT(AU7:AU114, F7:F114)   ← Total Rupiah PO 2
```

### MROUND vs Math.ceil

| Nilai | MB | MROUND (Excel) | ceil (lama) |
|---|---|---|---|
| 144 | 100 | **100** | 200 |
| 150 | 100 | **200** | 200 |
| 250 | 100 | **300** | 300 |

`MROUND` bulatkan ke kelipatan **terdekat**, `ceil` bulatkan ke **atas**. Preview di frontend menggunakan `Math.round` yang sama dengan `MROUND`.

---

## Fitur

- ✅ Upload Excel → preview data langsung di browser
- ✅ Pilih kolom sumber (PO Sisa / Purch Plan / kolom manapun)
- ✅ Toggle kurangi PO 1 (untuk mode Purch Plan)
- ✅ Pilih kolom HNA dan Master Box
- ✅ Atur jumlah PO (2-6) dan persentase per PO
- ✅ Preview tabel update otomatis saat persentase diubah
- ✅ Format Excel 100% terjaga (warna, font, border, merged cells)
- ✅ Formula Excel asli (MROUND, SUM, SUMPRODUCT)
- ✅ Tidak butuh MS Excel / WPS terinstall
- ✅ Deploy ke Vercel (serverless, gratis)

---

## Troubleshooting

| Masalah | Solusi |
|---|---|
| Sheet "PP" tidak ditemukan | Pastikan file Excel punya sheet dengan nama mengandung "PP" |
| Header "po sisa" tidak ditemukan | Pastikan ada cell di header row yang berisi teks "po sisa" |
| Preview kosong | Pastikan kolom A berisi nomor urut (1, 2, 3, ...) atau formula `=A7+1` |
| Format hilang di Vercel | Pastikan `api/index.py` menggunakan manual copy (bukan `copy()`) |
| Formula tidak hitung | Buka file di Excel/WPS → Enable Editing → formula auto-recalculate |

---

## License

MIT
