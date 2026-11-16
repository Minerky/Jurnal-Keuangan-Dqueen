# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users
Pemilik dan staf operasional DQueen Florist pada laptop/komputer kasir lokal.

## Product Purpose
Mencatat dan mengelola seluruh pengeluaran operasional florist (bunga segar, perlengkapan florist, kendaraan florist BBM/bengkel, biaya operasional lainnya, dan dels) secara cepat, akurat, dan dapat diekspor langsung ke laporan Excel (.xlsx) rapi untuk akuntansi bulanan.

## Positioning
Sistem pencatatan kas keluar mandiri berkecepatan tinggi tanpa dependensi internet, tanpa biaya langganan cloud, dan tanpa instalasi dependensi luar (zero external npm).

## Operating Context
Laptop/komputer lokal kasir florist, digunakan saat belanja harian, pengecekan nota supplier bunga (Akasa, Salsabila, dll.), dan pelaporan keuangan akhir bulan.

## Capabilities and Constraints
- Local Node.js native server dengan SQLite (`node:sqlite`).
- Nol dependensi runtime eksternal npm.
- Form modal terfokus dengan keyboard shortcuts (`Alt+N`, `Ctrl+Enter`, `Esc`, `/`).
- Format kategori baku (5): Bunga, Keperluan Florist, Kendaraan Florist, Biaya Lainnya, Dels.
- Ekspor spreadsheet native XML (`.xlsx`) berformat rapi dengan rekapitulasi subtotal.

## Brand Commitments
DQueen Florist — Nuansa butik bunga atelier elegan, tenang, dan profesional.

## Evidence on Hand
- Database lokal riil: `jurnal/jurnal.db` (2.278 transaksi historis 2026 tervalidasi).
- Template ekspor Excel teruji: `.xlsx` dengan format angka akuntansi `#,##0`.

## Product Principles
1. Kecepatan input nomor satu (keyboard-first data entry).
2. Nol dependensi eksternal untuk keandalan jangka panjang tanpa takut update break.
3. Keterbacaan angka tinggi dan nyaman dipandang mata berjam-jam (Boutique Atelier theme).
