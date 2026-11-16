# 📒 Jurnal Keuangan DQueen Florist

Aplikasi web jurnal keuangan internal untuk **DQueen Florist** — mencatat, memfilter, dan merekap transaksi harian langsung dari browser, tanpa database eksternal.

## Fitur

- Catat transaksi pemasukan & pengeluaran harian
- Filter per kategori: Bunga, Keperluan Florist, Kendaraan Florist, Biaya Lainnya, Dels
- Rekapitulasi otomatis (ringkasan per kategori)
- Pagination tabel transaksi
- Export & backup data
- Penyimpanan lokal SQLite — tidak butuh internet atau server cloud

## Kategori Transaksi

| Kategori | Keterangan |
|---|---|
| Bunga | Pembelian bunga & bahan rangkaian |
| Keperluan Florist | Perlengkapan dan supplies florist |
| Kendaraan Florist | Biaya operasional kendaraan |
| Biaya Lainnya | Pengeluaran umum lainnya |
| Dels | Kategori khusus Dels |

## Teknologi

- **Backend:** Node.js (built-in `node:http`, `node:sqlite`)
- **Frontend:** HTML + CSS + Vanilla JS (tanpa framework)
- **Database:** SQLite lokal (`jurnal.db`)
- **Tanpa dependensi npm** — langsung jalan dengan Node.js ≥ 22.5

## Cara Pakai

### Jalankan server

```bash
node server.js
```

Atau klik `start.bat` (Windows).

Buka browser ke: `http://localhost:3000`

### Import data lama (migrasi dari Excel)

```bash
node migrate.js
```

## Struktur File

```
jurnal/
├── server.js        # Server + API + skema database
├── migrate.js       # Script migrasi data dari Excel/JSON
├── package.json
├── start.bat        # Shortcut run Windows
└── public/
    └── index.html   # UI aplikasi
```

## Catatan

- File `jurnal.db` tidak di-push ke GitHub (berisi data sensitif)
- Jalankan `node migrate.js` sekali saat pertama setup untuk import data lama
