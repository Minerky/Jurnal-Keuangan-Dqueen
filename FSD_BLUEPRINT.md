# 📋 Functional Specification Document (FSD)
## Aplikasi Jurnal Keuangan DQueen Florist
**Versi:** 1.0 | **Tanggal:** November 2026 | **Dibuat oleh:** Rezky

---

## 1. RINGKASAN EKSEKUTIF

### Latar Belakang
DQueen Florist adalah bisnis florist yang sebelumnya mencatat pengeluaran operasional secara manual di **Microsoft Excel**. Proses ini tidak efisien: data tersebar di banyak file, sulit dicari, tidak ada rekapitulasi otomatis, dan berisiko hilang.

### Solusi
Aplikasi web internal **Jurnal Keuangan DQueen Florist** — sistem pencatatan pengeluaran harian yang berjalan langsung di browser, tanpa koneksi internet, tanpa biaya server cloud, dan tanpa dependensi eksternal.

### Nilai Bisnis

| Sebelum (Excel) | Sesudah (Aplikasi) |
|---|---|
| Input manual, rawan salah ketik | Form terstruktur + validasi otomatis |
| Rekapitulasi dihitung manual | Subtotal per kategori real-time |
| File Excel tersebar | Satu database terpusat |
| Tidak bisa dicari cepat | Pencarian instan multi-kolom |
| Export manual | Export Excel 1 klik |
| Tidak ada riwayat bulan | Navigasi periode per bulan |

---

## 2. ALUR BISNIS (BUSINESS FLOW)

### 2.1 Alur Utama Harian

```
SETIAP HARI OPERASIONAL
        │
        ▼
┌───────────────────┐
│  Staf buka browser│
│  → localhost:3000 │
└────────┬──────────┘
         │
         ▼
┌───────────────────┐
│  Pilih Bulan/     │
│  Periode di       │
│  dropdown atas    │
└────────┬──────────┘
         │
         ▼
┌───────────────────────────────────────────────────┐
│              ADA TRANSAKSI BARU?                  │
└──────────────┬───────────────────────┬────────────┘
              YA                      TIDAK
               │                       │
               ▼                       ▼
   ┌───────────────────┐    ┌────────────────────┐
   │ Klik "Catat       │    │ Lihat rekapitulasi │
   │ Pengeluaran"      │    │ & laporan bulan ini│
   └─────────┬─────────┘    └────────────────────┘
             │
             ▼
   ┌─────────────────────┐
   │ PILIH KATEGORI:     │
   │ • Bunga             │
   │ • Keperluan Florist │
   │ • Kendaraan Florist │
   │ • Biaya Lainnya     │
   │ • Dels              │
   └─────────┬───────────┘
             │
      ┌──────┼──────┐
      ▼      ▼      ▼
  [Bunga] [Kend] [Lainnya]
  [Florist]      [Dels]
      │      │      │
      ▼      ▼      ▼
  Form    Form    Form
  Multi   BBM/    1 item
  Item    Plat    keterangan
      │      │      │
      └──────┼──────┘
             │
             ▼
   ┌───────────────────┐
   │ Sistem cek duplikat│
   │ (tanggal+toko sama)│
   └─────────┬─────────┘
             │
   ┌─────────┴──────────┐
  Duplikat?           Tidak
   │                    │
   ▼                    ▼
Tampil             ┌──────────┐
Warning            │ SIMPAN   │
(bisa tetap        └────┬─────┘
 simpan)                │
                        ▼
                 Database SQLite
                 diperbarui
                        │
                        ▼
                 Tabel & Summary
                 refresh otomatis
                        │
                        ▼
               Toast notifikasi
               "Transaksi tersimpan"
```

### 2.2 Alur Export Laporan

```
Owner / Admin
     │
     ▼
Pilih Bulan di dropdown
     │
     ▼
Klik "Export Excel (.xlsx)"
     │
     ▼
Server generate file:
 ├── Ringkasan per kategori
 ├── Subtotal per kategori
 ├── Grand total
 └── Detail seluruh transaksi
     │
     ▼
File otomatis terunduh ke komputer
(nama: Jurnal_DQueen_YYYY-MM.xlsx)
```

### 2.3 Alur Migrasi Data Lama

```
Data Lama (Excel/JSON)
     │
     ▼
Jalankan: node migrate.js
     │
     ▼
Script baca file sumber
     │
     ▼
Parse & bersihkan format angka
(handles: "1.000.000", "Rp 500rb", dll)
     │
     ▼
Masuk ke database SQLite
     │
     ▼
Semua data historis tersedia
di aplikasi
```

---

## 3. FITUR-FITUR APLIKASI

### 3.1 Dashboard Utama

```
┌─────────────────────────────────────────────────────────────────┐
│ 🌿 DQueen Florist  /  Atelier Jurnal Keuangan    [SQLite ●]     │
├─────────────────────────────────────────────────────────────────┤
│ Periode: [◄] [November 2026 (45 transaksi) ▼] [►]              │
│           [📊 Grafik Tren] [+ Catat Pengeluaran] [↓ Export xlsx]│
├──────────────┬──────────┬──────────┬──────────┬──────────┬──────┤
│ TOTAL BEBAN  │  Bunga   │ Kep.     │ Kendaraan│ Biaya    │ Dels │
│ OPERASIONAL  │          │ Florist  │ Florist  │ Lainnya  │      │
│              │          │          │          │          │      │
│ Rp 12.500.000│Rp 5.2jt │Rp 3.1jt  │Rp 2.4jt  │Rp 1.1jt  │Rp 0 │
│ [klik=semua] │[klik=   │[klik=    │[klik=    │[klik=    │[klik]│
│              │ filter] │ filter]  │ filter]  │ filter]  │      │
├──────────────┴──────────┴──────────┴──────────┴──────────┴──────┤
│ [GRAFIK TREN - tampil/sembunyi]                                  │
│  Komposisi Beban  │  Tren Bulanan (klik batang = buka bulan)    │
│  ████░░░░░        │  ▓▓▓▓▓▓▓ ▓▓▓▓▓▓ ▓▓▓▓▓ ▓▓▓▓▓▓▓▓▓▓           │
├─────────────────────────────────────────────────────────────────┤
│ [Semua(45)] [Bunga] [Kep.Florist] [Kendaraan] [Lainnya] [Dels] │
│                                     [🔍 Cari transaksi (/)...] │
├────┬──────────┬──────────┬────────────┬──────────┬──────┬───────┤
│ #  │ Tanggal  │ Kategori │ Toko/Sup.  │ Uraian   │ Qty  │ Total │
├────┼──────────┼──────────┼────────────┼──────────┼──────┼───────┤
│  1 │01/11/26  │ [Bunga]  │ Pasar Rawa │ Mawar    │  50  │130.000│
│  2 │01/11/26  │ [Florist]│ Toko ABC   │ Kawat    │   2  │ 25.000│
│  3 │02/11/26  │ [Kendara]│ SPBU Shell │ BBM      │   -  │ 85.000│
├────┴──────────┴──────────┴────────────┴──────────┴──────┴───────┤
│ Menampilkan 1-45 dari 45 transaksi · Subtotal: Rp 12.500.000    │
│ Tampilkan: [50 ▼]           [«] [‹] Hal 1/1 [›] [»]            │
└─────────────────────────────────────────────────────────────────┘
```

### 3.2 Form Catat Transaksi

**Mode Bunga / Keperluan Florist (Multi-Item dari 1 Nota):**
```
┌──────────────────────────────────────────────────┐
│ Catat Pengeluaran Baru                      [×]  │
│ Isi rincian transaksi belanja florist            │
├──────────────────────────────────────────────────┤
│ [Bunga] [Kep.Florist] [Kend.Florist] [Lainnya]  │
│         [Dels]                                   │
├──────────────────────────────────────────────────┤
│ Tanggal Transaksi     │ Toko / Supplier          │
│ [01/11/2026        ]  │ [Pasar Rawabelong...  ]  │
│                                                  │
│ ┌──────────────────────────────────────────────┐ │
│ │ Jenis Bunga *    Qty  Satuan  H.Sat   Total  │ │
│ ├──────────────────────────────────────────────┤ │
│ │ [Mawar Merah  ] [50] [ikat] [2600] [130.000]│ │
│ │ [Lily Putih   ] [30] [ikat] [4000] [120.000]│ │
│ │ [+ Tambah Jenis Bunga]      Total: Rp 250.000│ │
│ └──────────────────────────────────────────────┘ │
│                                                  │
│ [Ctrl+Enter] Simpan                              │
│         [Batal] [Simpan & Catat Lagi] [Simpan ✓]│
└──────────────────────────────────────────────────┘
```

**Mode Kendaraan Florist:**
```
│ Nomor Plat Mobil       │ Rincian Operasional      │
│ [B 4160 SSV         ]  │ [BBM / Bengkel / Oli   ] │
│                                                    │
│ Total Pengeluaran (Rp)                            │
│ [                 85.000]                          │
```

**Mode Biaya Lainnya / Dels (1 Item):**
```
│ Keterangan Pengeluaran                            │
│ [Uang makan, listrik, ATK...                   ] │
│ Qty    Satuan    Harga Satuan    Total            │
│ [1 ]  [     ]   [          ]   [         0     ] │
```

### 3.3 Grafik & Analitik

```
┌─────────────────────────────────────────────────────┐
│ KOMPOSISI BEBAN PERIODE    │  TREN PENGELUARAN       │
│ November 2026              │  BULANAN                │
│                            │                         │
│ ██████▓▓▓▓░░░░▒▒▒▒         │  ▓                      │
│ (stacked bar)              │  ▓  ▓                   │
│                            │  ▓  ▓    ▓              │
│ ● Bunga      Rp 5.200.000  │  ▓  ▓  ▓ ▓  ▓          │
│ ● Kep.Flo.   Rp 3.100.000  │  ▓  ▓  ▓ ▓  ▓  ▓      │
│ ● Kendaraan  Rp 2.400.000  │ Sep Okt Nov Des Jan Feb │
│ ● Lainnya    Rp 1.100.000  │  (klik batang = buka   │
│ ● Dels       Rp       0    │   bulan tsb)            │
└─────────────────────────────────────────────────────┘
```

---

## 4. KATEGORI TRANSAKSI

### Definisi Kategori

| Kategori | Warna | Jenis Pengeluaran | Input Khusus |
|---|---|---|---|
| **Bunga** | 🌸 Rose | Pembelian bunga segar, rangkaian, bahan bunga | Multi-item per nota |
| **Keperluan Florist** | 🟫 Terracotta | Kawat, foam, pita, selotip, pot, kemasan | Multi-item per nota |
| **Kendaraan Florist** | 🌲 Forest Green | BBM, servis, parkir, tol, oli, ban | Nomor plat + rincian |
| **Biaya Lainnya** | 🔵 Slate | Listrik, air, uang makan, ATK, sewa | 1 item + keterangan |
| **Dels** | 💜 Wisteria | Kategori khusus Dels | 1 item + keterangan |

---

## 5. SPESIFIKASI TEKNIS

### 5.1 Arsitektur Sistem

```
┌──────────────────────────────────────────────────┐
│                  BROWSER                         │
│  ┌────────────────────────────────────────────┐  │
│  │          public/index.html                 │  │
│  │  HTML + CSS (Boutique Atelier Design)      │  │
│  │  Vanilla JavaScript (tanpa framework)      │  │
│  └───────────────────┬────────────────────────┘  │
└──────────────────────│───────────────────────────┘
                       │  HTTP (localhost:3000)
                       │  REST API (JSON)
┌──────────────────────│───────────────────────────┐
│                 NODE.JS SERVER                   │
│  ┌────────────────────────────────────────────┐  │
│  │           server.js                        │  │
│  │  node:http · node:fs · node:zlib           │  │
│  │  node:sqlite (built-in, Node ≥ 22.5)       │  │
│  └───────────────────┬────────────────────────┘  │
│                      │  SQL                      │
│  ┌───────────────────▼────────────────────────┐  │
│  │           jurnal.db (SQLite)               │  │
│  │  Tabel: transaksi                          │  │
│  │  Index: tanggal, kategori                  │  │
│  └────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────┘
```

### 5.2 API Endpoints

| Method | Endpoint | Fungsi |
|---|---|---|
| `GET` | `/api/months` | Daftar bulan tersedia + jumlah transaksi |
| `GET` | `/api/transactions` | Ambil transaksi (filter: bulan, kategori, search, sort) |
| `POST` | `/api/transactions` | Tambah 1 transaksi |
| `POST` | `/api/transactions/batch` | Tambah multi-item dari 1 nota |
| `PUT` | `/api/transactions/:id` | Update transaksi |
| `DELETE` | `/api/transactions/:id` | Hapus transaksi |
| `GET` | `/api/suggestions` | Autocomplete (toko, barang, plat) |
| `GET` | `/api/export?month=YYYY-MM` | Download file `.xlsx` |

### 5.3 Struktur Database

```sql
CREATE TABLE transaksi (
  id           INTEGER PRIMARY KEY AUTOINCREMENT,
  tanggal      TEXT    NOT NULL,          -- Format: YYYY-MM-DD
  kategori     TEXT    NOT NULL,          -- Enum 5 kategori
  nama_toko    TEXT    DEFAULT '',        -- Nama supplier/toko
  jenis_barang TEXT    DEFAULT '',        -- Nama barang/uraian
  qty          REAL    DEFAULT 1,
  satuan       TEXT    DEFAULT '',        -- ikat, pcs, liter, dll
  harga_satuan INTEGER DEFAULT 0,
  jumlah       INTEGER NOT NULL,          -- Total = qty × harga
  detail       TEXT    DEFAULT '',        -- Rincian operasional
  nomor_plat   TEXT    DEFAULT '',        -- Khusus Kendaraan
  no_hp        TEXT    DEFAULT '',
  keterangan   TEXT    DEFAULT '',
  sumber_sheet TEXT    DEFAULT '',        -- 'Web Input' / 'Web Multi-Item'
  created_at   DATETIME DEFAULT CURRENT_TIMESTAMP
);
```

### 5.4 Stack Teknologi

| Komponen | Teknologi | Alasan |
|---|---|---|
| Runtime | Node.js ≥ 22.5 | Built-in SQLite, tanpa npm install |
| Database | SQLite (node:sqlite) | File lokal, zero config, offline |
| Frontend | HTML + CSS + Vanilla JS | Tanpa build tool, langsung jalan |
| Export | OpenXML (.xlsx) built-in | Tanpa library pihak ketiga |
| Font | Playfair Display + Outfit | Google Fonts CDN |

---

## 6. DESAIN VISUAL (DESIGN SYSTEM)

### 6.1 Palet Warna

| Token | Hex | Penggunaan |
|---|---|---|
| Background | `#FAF8F5` | Canvas utama (Warm Ivory Linen) |
| Surface | `#FFFFFF` | Card & modal |
| Brand Primary | `#1B382B` | Tombol utama, header, aksen (Deep Forest Green) |
| Accent Gold | `#B8935B` | Hover scrollbar, border aktif (Antique Gold) |
| Text Primary | `#19261E` | Teks utama |
| Text Muted | `#4B5B52` | Label, placeholder |
| Danger | `#B33939` | Tombol hapus, error |

### 6.2 Warna per Kategori

| Kategori | Warna Badge | Latar |
|---|---|---|
| Bunga | `#9E2A4B` Vintage Rose | `#FDF0F4` |
| Keperluan Florist | `#8F5326` Terracotta | `#FDF5EE` |
| Kendaraan Florist | `#1D6348` Forest Pine | `#EDF7F2` |
| Biaya Lainnya | `#485B6C` Antique Slate | `#F0F4F8` |
| Dels | `#684984` Wisteria Iris | `#F6F1FA` |

### 6.3 Tipografi

- **Display/Header:** Playfair Display (serif, elegan)
- **Body & UI:** Outfit (sans-serif, modern bersih)
- **Angka:** `font-variant-numeric: tabular-nums` (kolom rata)

---

## 7. FITUR UX DETAIL

### 7.1 Keyboard Shortcuts

| Shortcut | Aksi |
|---|---|
| `/` | Fokus ke kotak pencarian |
| `Alt + N` | Buka modal catat transaksi |
| `Ctrl + Enter` | Simpan & tutup form |
| `Esc` | Tutup modal |

### 7.2 Proteksi Data

- **Duplikat detection:** Sistem otomatis deteksi transaksi dengan tanggal + toko + nominal sama, tampilkan peringatan kuning
- **Validasi form:** Tanggal & kategori wajib, jumlah harus > 0
- **SQL injection safe:** Semua query pakai parameterized statement
- **Sort column whitelist:** Kolom sort divalidasi server-side

### 7.3 Autocomplete

Sistem menyimpan dan menyarankan otomatis:
- Nama toko/supplier (dari histori transaksi)
- Nama barang (dari histori transaksi)
- Nomor plat kendaraan (hardcoded 5 plat)

### 7.4 Pagination

- Default: 50 baris per halaman
- Pilihan: 25 / 50 / 100 / Semua
- Navigasi: Pertama, Prev, Nomor Halaman, Next, Terakhir
- Info: "Menampilkan X-Y dari Z transaksi · Subtotal: Rp..."

---

## 8. CARA INSTALASI & PENGGUNAAN

### Prasyarat
- **Node.js versi 22.5 atau lebih baru** (untuk built-in SQLite)
- Windows / Mac / Linux

### Langkah Instalasi

**1. Download/Clone project**
```
Folder: jurnal/
├── server.js
├── migrate.js
├── package.json
├── start.bat
└── public/
    └── index.html
```

**2. Jalankan server**

Cara 1 — Klik `start.bat` (Windows)

Cara 2 — Terminal:
```bash
node server.js
```

**3. Buka browser**
```
http://localhost:3000
```

**4. (Opsional) Import data lama**
```bash
node migrate.js
```

### Cara Backup Data
- File database ada di `jurnal.db`
- Copy file `jurnal.db` ke tempat aman secara berkala
- Atau gunakan tombol Export Excel untuk backup per bulan

---

## 9. RENCANA PENGEMBANGAN LANJUTAN (ROADMAP)

| Prioritas | Fitur | Keterangan |
|---|---|---|
| 🔴 High | Multi-user login | Pisah akses staf & owner |
| 🔴 High | Backup otomatis | Schedule harian ke folder tertentu |
| 🟡 Medium | Laporan laba-rugi | Pemasukan vs Pengeluaran |
| 🟡 Medium | Notifikasi pengeluaran besar | Alert kalau melebihi budget |
| 🟢 Low | Versi mobile (PWA) | Akses dari HP |
| 🟢 Low | Print langsung | Cetak laporan dari browser |
| 🟢 Low | Multi-cabang | Support lebih dari 1 toko |

---

## 10. PERTANYAAN YANG SERING DIAJUKAN (FAQ)

**Q: Apakah data saya aman?**
A: Data tersimpan di file `jurnal.db` di komputer lokal. Tidak ada data yang dikirim ke internet.

**Q: Kalau komputer mati, data hilang?**
A: Tidak. Data permanen di file SQLite. Selama file `jurnal.db` ada, data aman.

**Q: Bisa dipakai lebih dari 1 orang sekaligus?**
A: Bisa, selama dalam jaringan yang sama (LAN). Tapi versi ini belum ada sistem login — semua pengguna punya akses penuh.

**Q: Bagaimana export data ke Excel?**
A: Klik tombol "Export Excel (.xlsx)" di kanan atas. File akan langsung terunduh dengan nama `Jurnal_DQueen_YYYY-MM.xlsx`.

**Q: Apakah bisa diakses dari HP?**
A: Bisa, dari HP yang terhubung jaringan WiFi yang sama. Akses melalui `http://[IP-komputer]:3000`.

---

*Dokumen ini disiapkan untuk presentasi kepada owner DQueen Florist.*
*Dibuat oleh: Rezky — November 2026*
