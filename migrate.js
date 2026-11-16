const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const { DatabaseSync } = require('node:sqlite');

const xlsxPath = "C:\\Users\\Harrys\\OneDrive\\Documents\\Jurnal 2026.xlsx";
const dbPath = path.join(__dirname, "jurnal.db");
const tempDir = path.join(process.env.TEMP, "jurnal_migrate_temp");

console.log("=== MEMULAI MIGRASI DATA JURNAL 2026 ===");
console.log("File Excel:", xlsxPath);
console.log("Database target:", dbPath);

// 1. Ekstrak ZIP dari XLSX
if (fs.existsSync(tempDir)) {
  fs.rmSync(tempDir, { recursive: true, force: true });
}
execSync(`powershell -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::ExtractToDirectory('${xlsxPath}', '${tempDir}')"`);

// 2. Baca Shared Strings
const ssXml = fs.readFileSync(path.join(tempDir, "xl", "sharedStrings.xml"), "utf-8");
const sst = [];
const siRegex = /<si>(.*?)<\/si>/gs;
let match;
while ((match = siRegex.exec(ssXml)) !== null) {
  let text = "";
  const tMatches = match[1].matchAll(/<t[^>]*>(.*?)<\/t>/gs);
  for (const tm of tMatches) text += tm[1];
  sst.push(text);
}
console.log(`Loaded ${sst.length} shared strings.`);

// Helper pembersih angka & teks uang
function cleanMoney(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  let str = String(val).trim();
  // Tangani typo khusus hasil audit
  if (str === '1,056,00') return 1056000;
  if (str === '110,00') return 110000;
  if (str === '43, 000' || str === '43,000') return 43000;
  if (str === '1.020.000') return 1020000;
  const digits = str.replace(/\D/g, '');
  return digits ? parseInt(digits, 10) : 0;
}

function parseQty(val) {
  if (!val) return 1;
  const str = String(val).replace(',', '.').replace(/[^\d.]/g, '');
  const num = parseFloat(str);
  return isNaN(num) || num <= 0 ? 1 : num;
}

function normalizeDate(rawDate, defaultYearMonth) {
  if (!rawDate) return null;
  const str = String(rawDate).trim().replace(/`/g, '').replace(/'/g, '');
  const [defY, defM] = defaultYearMonth.split('-');
  const targetMonthNum = parseInt(defM, 10);

  if (/^\d{5}$/.test(str)) {
    const serial = parseInt(str, 10);
    const utcDays = serial - 25569;
    const date = new Date(utcDays * 86400 * 1000);
    const d = String(date.getUTCDate()).padStart(2, '0');
    return `${defY}-${defM}-${d}`;
  }

  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
  if (m) {
    const n1 = parseInt(m[1], 10);
    const n2 = parseInt(m[2], 10);
    let day, month;
    if (n1 === targetMonthNum && n2 !== targetMonthNum) {
      month = n1;
      day = n2;
    } else if (n2 === targetMonthNum && n1 !== targetMonthNum) {
      day = n1;
      month = n2;
    } else if (n1 > 12) {
      day = n1;
      month = n2;
    } else {
      day = n1;
      month = n2;
    }
    const dStr = String(day).padStart(2, '0');
    const mStr = String(month).padStart(2, '0');
    return `${defY}-${mStr}-${dStr}`;
  }

  m = str.match(/^(\d{1,2})\s+([a-zA-Z]+)(?:\s+(\d{4}))?$/);
  if (m) {
    const d = String(m[1]).padStart(2, '0');
    return `${defY}-${defM}-${d}`;
  }
  return null;
}

// Helper parsing worksheet XML dengan penanganan self-closing tags
function parseSheetXml(file) {
  const xml = fs.readFileSync(path.join(tempDir, "xl", "worksheets", file), 'utf-8');
  const rows = [];
  const rowRegex = /<row r="(\d+)"[^>]*>(.*?)<\/row>/gs;
  let rMatch;
  while ((rMatch = rowRegex.exec(xml)) !== null) {
    const rowNum = parseInt(rMatch[1], 10);
    const content = rMatch[2];
    const cells = {};
    const cRegex = /<c ([^>\/]+)(?:\/>|>(.*?)<\/c>)/gs;
    let cMatch;
    while ((cMatch = cRegex.exec(content)) !== null) {
      const tagAttrs = cMatch[1];
      const innerStr = cMatch[2] || '';
      const rM = tagAttrs.match(/r="([A-Z]+)\d+"/);
      if (!rM) continue;
      const col = rM[1];
      const tM = tagAttrs.match(/t="([^"]+)"/);
      const type = tM ? tM[1] : '';
      let val = '';
      const vM = innerStr.match(/<v>(.*?)<\/v>/);
      if (vM) {
        val = vM[1];
      } else {
        const isM = innerStr.match(/<is><t[^>]*>(.*?)<\/t><\/is>/);
        if (isM) val = isM[1];
      }
      if (type === 's' && val !== '') {
        const idx = parseInt(val, 10);
        val = sst[idx] !== undefined ? sst[idx] : val;
      }
      cells[col] = val !== undefined ? val : '';
    }
    rows.push({ rowNum, cells });
  }
  return rows;
}

// 3. Inisialisasi Database SQLite
const db = new DatabaseSync(dbPath);
db.exec(`
  CREATE TABLE IF NOT EXISTS transaksi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tanggal TEXT NOT NULL,
    kategori TEXT NOT NULL,
    nama_toko TEXT DEFAULT '',
    jenis_barang TEXT DEFAULT '',
    qty REAL DEFAULT 1,
    satuan TEXT DEFAULT '',
    harga_satuan INTEGER DEFAULT 0,
    jumlah INTEGER NOT NULL,
    detail TEXT DEFAULT '',
    nomor_plat TEXT DEFAULT '',
    no_hp TEXT DEFAULT '',
    keterangan TEXT DEFAULT '',
    sumber_sheet TEXT DEFAULT '',
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE INDEX IF NOT EXISTS idx_transaksi_tanggal ON transaksi(tanggal);
  CREATE INDEX IF NOT EXISTS idx_transaksi_kategori ON transaksi(kategori);
`);

// Bersihkan data lama jika migrasi ulang
db.exec("DELETE FROM transaksi;");

const insertStmt = db.prepare(`
  INSERT INTO transaksi (
    tanggal, kategori, nama_toko, jenis_barang, qty, satuan,
    harga_satuan, jumlah, detail, nomor_plat, no_hp, keterangan, sumber_sheet
  ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);

let totalInserted = 0;

function addTx(tx) {
  if (!tx.jumlah || tx.jumlah <= 0) return;
  insertStmt.run(
    tx.tanggal || '2026-01-01',
    tx.kategori || 'Biaya Lainnya',
    tx.nama_toko || '',
    tx.jenis_barang || '',
    tx.qty || 1,
    tx.satuan || '',
    tx.harga_satuan || 0,
    tx.jumlah,
    tx.detail || '',
    tx.nomor_plat || '',
    tx.no_hp || '',
    tx.keterangan || '',
    tx.sumber_sheet || ''
  );
  totalInserted++;
}

// 4. Migrasi Sheet Format Baru (Juli 2026, Agustus 2026, September 2026, Oktober 2026, November 2026)
const newFormatSheets = [
  {
    name: "Juli 2026",
    file: "sheet4.xml",
    startRow: 3,
    endRow: 134,
    dateCol: "A",
    defaultYM: "2026-07",
    bunga: { shop: "C", item: "B", ikat: "D", gabung: "E", price: "F", total: "G" },
    florist: { shop: "H", item: "I", dus: "J", lusin: "K", pak: "", pcs: "L", roll: "M", price: "N", total: "O" },
    vehicle: { detail: "P", plate: "Q", total: "R" },
    other: { desc: "S", qty: "T", price: "U", total: "V" },
    dels: { item: "W", pcs: "X", dus: "Y", price: "Z", total: "AA" },
    shopRef: { shop: "AB", phone: "AC" }
  },
  {
    name: "Agustus 2026",
    file: "sheet5.xml",
    startRow: 4,
    endRow: 388,
    dateCol: "B",
    defaultYM: "2026-08",
    bunga: { shop: "C", item: "D", ikat: "E", gabung: "F", price: "G", total: "H" },
    florist: { shop: "I", item: "J", dus: "K", lusin: "L", pak: "M", pcs: "N", roll: "O", price: "P", total: "Q" },
    vehicle: { detail: "R", plate: "S", total: "T" },
    other: { desc: "U", qty: "V", price: "W", total: "X" },
    dels: { item: "Y", pcs: "Z", dus: "AA", price: "AB", total: "AC" },
    shopRef: { shop: "AD", phone: "AE" }
  },
  {
    name: "September 2026",
    file: "sheet6.xml",
    startRow: 6,
    endRow: 241,
    dateCol: "C",
    defaultYM: "2026-09",
    bunga: { shop: "D", item: "E", ikat: "F", gabung: "G", price: "H", total: "I" },
    florist: { shop: "J", item: "K", dus: "L", lusin: "M", pak: "N", pcs: "O", roll: "P", price: "Q", total: "R" },
    vehicle: { detail: "S", plate: "T", total: "U" },
    other: { desc: "V", qty: "W", price: "X", total: "Y" },
    dels: { item: "Z", pcs: "AA", dus: "AB", price: "AC", total: "AD" },
    shopRef: { shop: "AE", phone: "AF" }
  },
  {
    name: "Oktober 2026",
    file: "sheet7.xml",
    startRow: 6,
    endRow: 264,
    dateCol: "C",
    defaultYM: "2026-10",
    bunga: { shop: "D", item: "E", ikat: "F", gabung: "G", price: "H", total: "I" },
    florist: { shop: "J", item: "K", dus: "L", lusin: "M", pak: "N", pcs: "O", roll: "P", price: "Q", total: "R" },
    vehicle: { detail: "S", plate: "T", total: "U" },
    other: { desc: "V", qty: "W", price: "X", total: "Y" },
    dels: { item: "Z", pcs: "AA", dus: "AB", price: "AC", total: "AD" },
    shopRef: { shop: "AE", phone: "AF" }
  },
  {
    name: "November 2026",
    file: "sheet8.xml",
    startRow: 5,
    endRow: 271,
    dateCol: "B",
    defaultYM: "2026-11",
    bunga: { shop: "C", item: "D", ikat: "E", gabung: "F", price: "G", total: "H" },
    florist: { shop: "I", item: "J", dus: "K", lusin: "L", pak: "M", pcs: "N", roll: "O", price: "P", total: "Q" },
    vehicle: { detail: "R", plate: "S", total: "T" },
    other: { desc: "U", qty: "V", price: "W", total: "X" },
    dels: { item: "Y", pcs: "Z", dus: "AA", price: "AB", total: "AC" },
    shopRef: { shop: "AD", phone: "AE" }
  }
];

db.exec("BEGIN TRANSACTION;");

for (const cfg of newFormatSheets) {
  const rows = parseSheetXml(cfg.file);
  let curDate = cfg.defaultYM + "-01";
  let countSheet = 0;

  for (const r of rows) {
    if (r.rowNum < cfg.startRow || r.rowNum > cfg.endRow) continue;

    // Update current date if row has date
    const dVal = r.cells[cfg.dateCol];
    if (dVal && !dVal.includes("Jumlah")) {
      const norm = normalizeDate(dVal, cfg.defaultYM);
      if (norm) curDate = norm;
    }

    // Reference shop info on the row if any
    const rowShop = r.cells[cfg.shopRef.shop] || '';
    const rowPhone = r.cells[cfg.shopRef.phone] || '';

    // 1. Kategori Bunga
    const bungaTot = cleanMoney(r.cells[cfg.bunga.total]);
    if (bungaTot > 0) {
      const ikat = parseQty(r.cells[cfg.bunga.ikat]);
      const gabung = parseQty(r.cells[cfg.bunga.gabung]);
      let qty = 1;
      let sat = '';
      if (r.cells[cfg.bunga.ikat]) { qty = ikat; sat = 'ikat'; }
      else if (r.cells[cfg.bunga.gabung]) { qty = gabung; sat = 'gabung'; }

      addTx({
        tanggal: curDate,
        kategori: 'Bunga',
        nama_toko: r.cells[cfg.bunga.shop] || rowShop,
        jenis_barang: r.cells[cfg.bunga.item] || '',
        qty,
        satuan: sat,
        harga_satuan: cleanMoney(r.cells[cfg.bunga.price]),
        jumlah: bungaTot,
        no_hp: rowPhone,
        sumber_sheet: cfg.name
      });
      countSheet++;
    }

    // 2. Kategori Keperluan Florist
    const floristTot = cleanMoney(r.cells[cfg.florist.total]);
    if (floristTot > 0) {
      let qty = 1;
      let sat = '';
      if (r.cells[cfg.florist.pcs]) { qty = parseQty(r.cells[cfg.florist.pcs]); sat = 'pcs'; }
      else if (r.cells[cfg.florist.roll]) { qty = parseQty(r.cells[cfg.florist.roll]); sat = 'roll'; }
      else if (r.cells[cfg.florist.dus]) { qty = parseQty(r.cells[cfg.florist.dus]); sat = 'dus'; }
      else if (r.cells[cfg.florist.lusin]) { qty = parseQty(r.cells[cfg.florist.lusin]); sat = 'lusin'; }
      else if (cfg.florist.pak && r.cells[cfg.florist.pak]) { qty = parseQty(r.cells[cfg.florist.pak]); sat = 'pak'; }

      addTx({
        tanggal: curDate,
        kategori: 'Keperluan Florist',
        nama_toko: r.cells[cfg.florist.shop] || rowShop,
        jenis_barang: r.cells[cfg.florist.item] || '',
        qty,
        satuan: sat,
        harga_satuan: cleanMoney(r.cells[cfg.florist.price]),
        jumlah: floristTot,
        no_hp: rowPhone,
        sumber_sheet: cfg.name
      });
      countSheet++;
    }

    // 3. Kategori Kendaraan Florist
    const vehicleTot = cleanMoney(r.cells[cfg.vehicle.total]);
    if (vehicleTot > 0) {
      addTx({
        tanggal: curDate,
        kategori: 'Kendaraan Florist',
        detail: r.cells[cfg.vehicle.detail] || 'BBM',
        nomor_plat: r.cells[cfg.vehicle.plate] || '',
        jumlah: vehicleTot,
        sumber_sheet: cfg.name
      });
      countSheet++;
    }

    // 4. Kategori Biaya Lainnya
    const otherTot = cleanMoney(r.cells[cfg.other.total]);
    if (otherTot > 0) {
      addTx({
        tanggal: curDate,
        kategori: 'Biaya Lainnya',
        keterangan: r.cells[cfg.other.desc] || '',
        qty: parseQty(r.cells[cfg.other.qty]),
        harga_satuan: cleanMoney(r.cells[cfg.other.price]),
        jumlah: otherTot,
        sumber_sheet: cfg.name
      });
      countSheet++;
    }

    // 5. Kategori Dels
    const delsTot = cleanMoney(r.cells[cfg.dels.total]);
    if (delsTot > 0) {
      let qty = 1;
      let sat = '';
      if (r.cells[cfg.dels.pcs]) { qty = parseQty(r.cells[cfg.dels.pcs]); sat = 'pcs'; }
      else if (r.cells[cfg.dels.dus]) { qty = parseQty(r.cells[cfg.dels.dus]); sat = 'dus'; }

      addTx({
        tanggal: curDate,
        kategori: 'Dels',
        jenis_barang: r.cells[cfg.dels.item] || '',
        qty,
        satuan: sat,
        harga_satuan: cleanMoney(r.cells[cfg.dels.price]),
        jumlah: delsTot,
        sumber_sheet: cfg.name
      });
      countSheet++;
    }
  }
  console.log(`Sheet '${cfg.name}': dimasukkan ${countSheet} transaksi.`);
}

// 5. Migrasi Sheet Format Lama: Juni (sheet2.xml) & Juli Awal (sheet3.xml)
console.log("\nMigrasi sheet format lama (Juni & Juli 1-16)...");

function migrateOldFormat(file, sheetName, defaultYM, startRow, endRow) {
  const rows = parseSheetXml(file);
  let curDate = defaultYM + "-01";
  let count = 0;
  for (const r of rows) {
    if (r.rowNum < startRow || r.rowNum > endRow) continue;
    const dVal = r.cells["A"];
    if (dVal && !dVal.includes("JUMLAH") && !dVal.includes("TANGGAL")) {
      const norm = normalizeDate(dVal, defaultYM);
      if (norm) curDate = norm;
    }
    const pengeluaran = cleanMoney(r.cells["C"]);
    const desc = (r.cells["B"] || '').trim();
    const reff = (r.cells["E"] || '').trim();
    if (pengeluaran > 0 && desc && !desc.toLowerCase().includes("total")) {
      let kat = 'Biaya Lainnya';
      let plat = '';
      let detailVal = reff;

      const lowerDesc = desc.toLowerCase();
      const lowerReff = reff.toLowerCase();

      if (
        lowerDesc.includes('bensin') || lowerDesc.includes('bbm') ||
        lowerDesc.includes('tambal ban') ||
        (lowerDesc.includes('service') && (lowerReff.includes('yamaha') || /pwx|ssv|pyx|pmo/i.test(lowerDesc)))
      ) {
        kat = 'Kendaraan Florist';
        plat = reff;
        detailVal = desc;
      }

      addTx({
        tanggal: curDate,
        kategori: kat,
        keterangan: desc,
        jumlah: pengeluaran,
        detail: detailVal,
        nomor_plat: plat,
        sumber_sheet: sheetName
      });
      count++;
    }
  }
  console.log(`Sheet '${sheetName}': dimasukkan ${count} transaksi.`);
}

migrateOldFormat("sheet2.xml", "juni", "2026-06", 5, 471);
migrateOldFormat("sheet3.xml", "juli", "2026-07", 5, 498);

db.exec("COMMIT;");

// Bersihkan folder temp
fs.rmSync(tempDir, { recursive: true, force: true });

console.log(`\n=== MIGRASI SELESAI ===`);
console.log(`Total transaksi tersimpan di SQLite: ${totalInserted}`);

// Tampilkan rekap per bulan dari DB
const summaryQuery = db.prepare(`
  SELECT 
    strftime('%Y-%m', tanggal) as bulan,
    COUNT(*) as total_transaksi,
    SUM(jumlah) as grand_total
  FROM transaksi
  GROUP BY bulan
  ORDER BY bulan
`);
const monthlyReport = summaryQuery.all();
console.table(monthlyReport);
