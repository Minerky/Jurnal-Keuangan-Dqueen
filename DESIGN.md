# DESIGN.md — Boutique Atelier World

## Aesthetic Identity
- **Dunia Visual:** *Boutique Atelier Florist* — perpaduan hangat kertas linen/ivory atelier, hijau botani tua (*deep forest evergreen*), dan aksen emas champagne kuno.
- **Filosofi:** Elegan, tenang, dan memanjakan mata (*tactile & soothing*). Mengganti tampilan dashboard SaaS abu-abu dingin dengan estetika studio bunga mewah yang tidak melelahkan mata.

## Colors
- **Canvas / Background:** `#FAF8F5` (Warm Ivory Linen)
- **Subtle Surface:** `#F4EFEA` (Warm Stone Paper)
- **Card Surface:** `#FFFFFF` (Porcelain Warm White)
- **Borders:** `#E8E2D8` (Linen Hairline), `#D6CCC0` (Strong Tactile Border)
- **Text Primary:** `#19261E` (Deep Botanical Evergreen Charcoal)
- **Text Muted:** `#4B5B52` (Muted Moss Gray, >= 5.5:1 ratio)
- **Text Subtle:** `#5E6E65` (Soft Foliage Gray, >= 4.6:1 ratio)
- **Brand Primary:** `#1B382B` (Deep Atelier Forest Green)
- **Brand Hover:** `#132A20` (Midnight Pine)
- **Accent Gold:** `#B8935B` (Antique Champagne Brass)
- **Category Indicators:**
  - `Bunga`: `#9E2A4B` (Vintage Rose) & `#FDF0F4` (Blush Tint)
  - `Keperluan Florist`: `#8F5326` (Terracotta Clay) & `#FDF5EE` (Clay Tint)
  - `Kendaraan Florist`: `#1D6348` (Forest Pine) & `#EDF7F2` (Pine Tint)
  - `Biaya Lainnya`: `#485B6C` (Antique Slate) & `#F0F4F8` (Slate Tint)
  - `Dels`: `#684984` (Wisteria Iris) & `#F6F1FA` (Wisteria Tint)

## Typography
- **Display / Headers:** `Playfair Display`, `Georgia`, serif.
- **Body & Controls:** `Outfit`, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif.
- **Numerals:** `font-variant-numeric: tabular-nums` (posisi digit angka sejajar vertikal).
- **Scale:** Minimum font size 12px.

## Components
- **Top Bar:** 56px sticky header dengan brand crest bulat DQ, status database SQLite, dan tombol backup/export.
- **Ledger Summary Cards:** Grid 6 kolom sel rekapitulasi, berfungsi sebagai filter interaktif instan.
- **Modal Dialog:** Native `<dialog>` dengan backdrop blur 3px, form grid terfokus, tombol simpan ganda ("Simpan & Tutup" / "Simpan & Catat Lagi").
- **Tabel Transaksi:** Baris zebra lembut `#FAF8F5`, header berperekat, sorotan animasi `row-flash`, dan pagination bar di bagian bawah.
- **Filter Pills:** Segmentasi kategori bulat 20px berstatus aktif dengan warna hijau botani tua `#1B382B`.

## Elevation & Depth
- 1px hairline border `#E8E2D8` dengan bayangan lembut `0 2px 8px rgba(25, 38, 30, 0.04)`.
- Modal elevation: `0 20px 48px rgba(25, 38, 30, 0.18)` dengan backdrop `#19261E` 45% opacity.
