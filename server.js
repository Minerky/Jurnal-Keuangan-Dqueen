const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const { DatabaseSync } = require('node:sqlite');

const PORT = process.env.PORT || 3000;
const dbPath = path.join(__dirname, 'jurnal.db');
const db = new DatabaseSync(dbPath);

// Pastikan skema tabel sudah siap
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

// Helper pembersih angka (fleksibel titik dan koma)
function parseMoney(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  const str = String(val).trim();
  // Tangani typo spesifik jika ada
  if (str === '1,056,00') return 1056000;
  if (str === '110,00') return 110000;
  // Hapus semua karakter kecuali digit angka
  const digits = str.replace(/\D/g, '');
  return digits ? parseInt(digits, 10) : 0;
}

function parseQty(val) {
  if (typeof val === 'number') return val;
  if (!val) return 1;
  const str = String(val).replace(',', '.').replace(/[^\d.]/g, '');
  const num = parseFloat(str);
  return isNaN(num) || num <= 0 ? 1 : num;
}

// Format mata uang rupiah
function formatRupiah(num) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(num || 0);
}

// Helper membaca body JSON
function getRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (e) {
        reject(e);
      }
    });
  });
}

// === Native OpenXML (.xlsx) Builder (No external npm packages) ===
const crcTable = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
  crcTable[i] = c >>> 0;
}
function crc32(buf) {
  let crc = 0 ^ (-1);
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ crcTable[(crc ^ buf[i]) & 0xff];
  return (crc ^ (-1)) >>> 0;
}

function createZip(files) {
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;

  for (const file of files) {
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf-8');
    const compressed = zlib.deflateRawSync(data);
    const useCompressed = compressed.length < data.length;
    const body = useCompressed ? compressed : data;
    const method = useCompressed ? 8 : 0;
    const crc = crc32(data);
    const nameBuf = Buffer.from(file.name, 'utf-8');

    const lh = Buffer.alloc(30 + nameBuf.length);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(method, 8);
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(body.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    lh.writeUInt16LE(0, 28);
    nameBuf.copy(lh, 30);
    localHeaders.push(lh, body);

    const ch = Buffer.alloc(46 + nameBuf.length);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);
    ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0, 8);
    ch.writeUInt16LE(method, 10);
    ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(body.length, 20);
    ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt16LE(0, 30);
    ch.writeUInt16LE(0, 32);
    ch.writeUInt16LE(0, 34);
    ch.writeUInt16LE(0, 36);
    ch.writeUInt32LE(0, 38);
    ch.writeUInt32LE(offset, 42);
    nameBuf.copy(ch, 46);
    centralHeaders.push(ch);

    offset += lh.length + body.length;
  }

  const centralOffset = offset;
  let centralSize = 0;
  for (const ch of centralHeaders) centralSize += ch.length;

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralSize, 12);
  eocd.writeUInt32LE(centralOffset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localHeaders, ...centralHeaders, eocd]);
}

function xmlEscape(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function generateXlsx(periodLabel, rows, subtotals, grandTotal) {
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Jurnal Pengeluaran" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0"/>
  </numFmts>
  <fonts count="4">
    <font><name val="Calibri"/><sz val="11"/></font>
    <font><b/><name val="Calibri"/><sz val="11"/></font>
    <font><b/><name val="Calibri"/><sz val="15"/><color rgb="FF0F172A"/></font>
    <font><b/><name val="Calibri"/><sz val="11"/><color rgb="FFFFFFFF"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2E8F0"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
    </border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FF0F172A"/></top>
      <bottom style="double"><color rgb="FF0F172A"/></bottom>
    </border>
  </borders>
  <cellXfs count="16">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="1" fillId="3" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="164" fontId="3" fillId="2" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="1" applyNumberFormat="1" applyFont="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="2" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="1" fillId="3" borderId="2" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right"/></xf>
  </cellXfs>
</styleSheet>`;

  const categories = [
    'Bunga',
    'Keperluan Florist',
    'Kendaraan Florist',
    'Biaya Lainnya',
    'Dels'
  ];

  let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="6" customWidth="1"/>
    <col min="2" max="2" width="14" customWidth="1"/>
    <col min="3" max="3" width="20" customWidth="1"/>
    <col min="4" max="4" width="28" customWidth="1"/>
    <col min="5" max="5" width="36" customWidth="1"/>
    <col min="6" max="6" width="9" customWidth="1"/>
    <col min="7" max="7" width="10" customWidth="1"/>
    <col min="8" max="8" width="18" customWidth="1"/>
    <col min="9" max="9" width="20" customWidth="1"/>
    <col min="10" max="10" width="24" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1" ht="28" customHeight="1">
      <c r="A1" s="1" t="inlineStr"><is><t>DQUEEN FLORIST — LAPORAN JURNAL PENGELUARAN</t></is></c>
    </row>
    <row r="2" ht="18" customHeight="1">
      <c r="A2" s="2" t="inlineStr"><is><t>Periode: ${xmlEscape(periodLabel)}   |   Tanggal Unduh: ${new Date().toLocaleDateString('id-ID')}   |   Total Data: ${rows.length} Transaksi</t></is></c>
    </row>
    <row r="3"></row>
    <row r="4" ht="22" customHeight="1">
      <c r="B4" s="3" t="inlineStr"><is><t>Ringkasan Kategori</t></is></c>
      <c r="C4" s="3" t="inlineStr"><is><t>Subtotal (Rp)</t></is></c>
    </row>`;

  let summaryRow = 5;
  for (const cat of categories) {
    sheetXml += `
    <row r="${summaryRow}" ht="19" customHeight="1">
      <c r="B${summaryRow}" s="4" t="inlineStr"><is><t>${xmlEscape(cat)}</t></is></c>
      <c r="C${summaryRow}" s="5"><v>${subtotals[cat] || 0}</v></c>
    </row>`;
    summaryRow++;
  }

  // Row Total Ringkasan
  sheetXml += `
    <row r="${summaryRow}" ht="22" customHeight="1">
      <c r="B${summaryRow}" s="6" t="inlineStr"><is><t>TOTAL PENGELUARAN</t></is></c>
      <c r="C${summaryRow}" s="7"><v>${grandTotal}</v></c>
    </row>`;

  const tableHeaderRow = summaryRow + 2;
  sheetXml += `
    <row r="${summaryRow + 1}"></row>
    <row r="${tableHeaderRow}" ht="26" customHeight="1">
      <c r="A${tableHeaderRow}" s="8" t="inlineStr"><is><t>No</t></is></c>
      <c r="B${tableHeaderRow}" s="8" t="inlineStr"><is><t>Tanggal</t></is></c>
      <c r="C${tableHeaderRow}" s="8" t="inlineStr"><is><t>Kategori</t></is></c>
      <c r="D${tableHeaderRow}" s="8" t="inlineStr"><is><t>Toko / Supplier</t></is></c>
      <c r="E${tableHeaderRow}" s="8" t="inlineStr"><is><t>Nama Barang / Uraian</t></is></c>
      <c r="F${tableHeaderRow}" s="8" t="inlineStr"><is><t>Qty</t></is></c>
      <c r="G${tableHeaderRow}" s="8" t="inlineStr"><is><t>Satuan</t></is></c>
      <c r="H${tableHeaderRow}" s="8" t="inlineStr"><is><t>Harga Satuan (Rp)</t></is></c>
      <c r="I${tableHeaderRow}" s="8" t="inlineStr"><is><t>Jumlah Total (Rp)</t></is></c>
      <c r="J${tableHeaderRow}" s="8" t="inlineStr"><is><t>Plat / Keterangan</t></is></c>
    </row>`;

  let rNum = tableHeaderRow + 1;
  rows.forEach((r, idx) => {
    const itemDesc = r.jenis_barang || r.keterangan || (r.kategori === 'Kendaraan Florist' ? (r.detail || 'BBM') : '');
    const ref = r.nomor_plat || r.detail || '';

    sheetXml += `
    <row r="${rNum}" ht="19" customHeight="1">
      <c r="A${rNum}" s="10"><v>${idx + 1}</v></c>
      <c r="B${rNum}" s="10" t="inlineStr"><is><t>${xmlEscape(r.tanggal)}</t></is></c>
      <c r="C${rNum}" s="10" t="inlineStr"><is><t>${xmlEscape(r.kategori)}</t></is></c>
      <c r="D${rNum}" s="9" t="inlineStr"><is><t>${xmlEscape(r.nama_toko || '')}</t></is></c>
      <c r="E${rNum}" s="9" t="inlineStr"><is><t>${xmlEscape(itemDesc)}</t></is></c>
      <c r="F${rNum}" s="11"><v>${r.qty || 1}</v></c>
      <c r="G${rNum}" s="10" t="inlineStr"><is><t>${xmlEscape(r.satuan || '')}</t></is></c>
      <c r="H${rNum}" s="12"><v>${r.harga_satuan || 0}</v></c>
      <c r="I${rNum}" s="13"><v>${r.jumlah || 0}</v></c>
      <c r="J${rNum}" s="9" t="inlineStr"><is><t>${xmlEscape(ref)}</t></is></c>
    </row>`;
    rNum++;
  });

  sheetXml += `
    <row r="${rNum}" ht="24" customHeight="1">
      <c r="A${rNum}" s="14"></c>
      <c r="B${rNum}" s="14"></c>
      <c r="C${rNum}" s="14"></c>
      <c r="D${rNum}" s="14"></c>
      <c r="E${rNum}" s="14"></c>
      <c r="F${rNum}" s="14"></c>
      <c r="G${rNum}" s="14"></c>
      <c r="H${rNum}" s="14" t="inlineStr"><is><t>TOTAL PENGELUARAN:</t></is></c>
      <c r="I${rNum}" s="15"><v>${grandTotal}</v></c>
      <c r="J${rNum}" s="14"></c>
    </row>
  </sheetData>
</worksheet>`;

  const zipFiles = [
    { name: '[Content_Types].xml', data: contentTypesXml },
    { name: '_rels/.rels', data: relsXml },
    { name: 'xl/workbook.xml', data: workbookXml },
    { name: 'xl/_rels/workbook.xml.rels', data: wbRelsXml },
    { name: 'xl/styles.xml', data: stylesXml },
    { name: 'xl/worksheets/sheet1.xml', data: sheetXml }
  ];

  return createZip(zipFiles);
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname;
  const method = req.method;

  // Set CORS & JSON headers helper
  const sendJson = (statusCode, data) => {
    res.writeHead(statusCode, {
      'Content-Type': 'application/json; charset=utf-8',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(JSON.stringify(data));
  };

  try {
    // 1. API: Daftar Bulan Tersedia
    if (pathname === '/api/months' && method === 'GET') {
      const rows = db.prepare(`
        SELECT 
          strftime('%Y-%m', tanggal) as bulan,
          COUNT(*) as total_transaksi,
          SUM(jumlah) as grand_total
        FROM transaksi
        GROUP BY bulan
        ORDER BY bulan DESC
      `).all();
      return sendJson(200, rows);
    }

    // 1b. API: Rekomendasi / Autocomplete (Toko, Barang, Plat)
    if (pathname === '/api/suggestions' && method === 'GET') {
      const shops = db.prepare(`
        SELECT DISTINCT TRIM(nama_toko) as nama
        FROM transaksi
        WHERE nama_toko != '' AND nama_toko IS NOT NULL
        ORDER BY nama ASC
      `).all().map(r => r.nama);

      const plates = [
        'B 4160 SSV', 'B 3320 PMO', 'B 3225 PYX', 'B 6818 PWX', 'B 4160 SSK'
      ];

      const items = db.prepare(`
        SELECT DISTINCT TRIM(jenis_barang) as item
        FROM transaksi
        WHERE jenis_barang != '' AND jenis_barang IS NOT NULL
        ORDER BY item ASC
      `).all().map(r => r.item);

      return sendJson(200, {
        shops: Array.from(new Set(shops.filter(Boolean))),
        plates,
        items: Array.from(new Set(items.filter(Boolean)))
      });
    }

    // 2. API: Daftar Transaksi & Subtotal (Filter, Sortir, Pencarian)
    if (pathname === '/api/transactions' && method === 'GET') {
      const month = parsedUrl.searchParams.get('month') || '';
      const kategori = parsedUrl.searchParams.get('kategori') || '';
      const search = parsedUrl.searchParams.get('q') || '';
      const startDate = parsedUrl.searchParams.get('startDate') || '';
      const endDate = parsedUrl.searchParams.get('endDate') || '';
      const sortBy = parsedUrl.searchParams.get('sortBy') || 'tanggal';
      const sortOrder = parsedUrl.searchParams.get('sortOrder') === 'asc' ? 'ASC' : 'DESC';

      let query = "SELECT * FROM transaksi WHERE 1=1";
      const params = [];

      if (startDate && endDate) {
        query += " AND tanggal BETWEEN ? AND ?";
        params.push(startDate, endDate);
      } else if (month && month !== 'ALL') {
        query += " AND strftime('%Y-%m', tanggal) = ?";
        params.push(month);
      }

      if (kategori && kategori !== 'Semua') {
        query += " AND kategori = ?";
        params.push(kategori);
      }
      if (search) {
        query += " AND (nama_toko LIKE ? OR jenis_barang LIKE ? OR keterangan LIKE ? OR nomor_plat LIKE ?)";
        const term = `%${search}%`;
        params.push(term, term, term, term);
      }

      // Validasi kolom sortir agar aman dari SQL injection
      const allowedSortCols = {
        'tanggal': 'tanggal',
        'jumlah': 'jumlah',
        'nama_toko': 'nama_toko',
        'jenis_barang': 'jenis_barang',
        'kategori': 'kategori'
      };
      const sortCol = allowedSortCols[sortBy] || 'tanggal';
      query += ` ORDER BY ${sortCol} ${sortOrder}, id ${sortOrder}`;

      const rows = db.prepare(query).all(...params);

      // Hitung subtotal untuk periode terpilih + filter pencarian
      let sumQuery = "SELECT kategori, SUM(jumlah) as subtotal, COUNT(*) as count FROM transaksi WHERE 1=1";
      const sumParams = [];
      if (startDate && endDate) {
        sumQuery += " AND tanggal BETWEEN ? AND ?";
        sumParams.push(startDate, endDate);
      } else if (month && month !== 'ALL') {
        sumQuery += " AND strftime('%Y-%m', tanggal) = ?";
        sumParams.push(month);
      }
      if (search) {
        sumQuery += " AND (nama_toko LIKE ? OR jenis_barang LIKE ? OR keterangan LIKE ? OR nomor_plat LIKE ?)";
        const term = `%${search}%`;
        sumParams.push(term, term, term, term);
      }
      sumQuery += " GROUP BY kategori";
      const subtotals = db.prepare(sumQuery).all(...sumParams);

      let grandTotal = 0;
      const subtotalMap = {
        'Bunga': 0,
        'Keperluan Florist': 0,
        'Kendaraan Florist': 0,
        'Biaya Lainnya': 0,
        'Dels': 0
      };

      for (const s of subtotals) {
        grandTotal += s.subtotal;
        if (subtotalMap[s.kategori] !== undefined) {
          subtotalMap[s.kategori] = s.subtotal;
        }
      }

      const filteredTotal = rows.reduce((acc, r) => acc + (r.jumlah || 0), 0);

      return sendJson(200, {
        items: rows,
        subtotals: subtotalMap,
        grandTotal,
        filteredTotal,
        totalItems: rows.length
      });
    }

    // 3. API: Tambah Transaksi Baru
    if (pathname === '/api/transactions' && method === 'POST') {
      const data = await getRequestBody(req);

      if (!data.tanggal || !data.kategori) {
        return sendJson(400, { error: 'Tanggal dan kategori wajib diisi.' });
      }

      const jumlah = parseMoney(data.jumlah);
      if (!jumlah || jumlah <= 0) {
        return sendJson(400, { error: 'Jumlah uang harus lebih besar dari 0.' });
      }

      const hargaSatuan = parseMoney(data.harga_satuan);
      const qty = parseQty(data.qty);

      const stmt = db.prepare(`
        INSERT INTO transaksi (
          tanggal, kategori, nama_toko, jenis_barang, qty, satuan,
          harga_satuan, jumlah, detail, nomor_plat, no_hp, keterangan, sumber_sheet
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Web Input')
      `);

      stmt.run(
        data.tanggal,
        data.kategori,
        data.nama_toko || '',
        data.jenis_barang || '',
        qty,
        data.satuan || '',
        hargaSatuan,
        jumlah,
        data.detail || '',
        data.nomor_plat || '',
        data.no_hp || '',
        data.keterangan || ''
      );

      return sendJson(201, { success: true });
    }

    // 3b. API: Tambah Transaksi Batch (Entri Multi-Item 1 Nota)
    if (pathname === '/api/transactions/batch' && method === 'POST') {
      const data = await getRequestBody(req);
      if (!data.tanggal || !data.kategori || !Array.isArray(data.items) || !data.items.length) {
        return sendJson(400, { error: 'Data nota atau daftar item tidak valid.' });
      }

      const sumber = (data.items.length === 1) ? 'Web Input' : 'Web Multi-Item';
      const insertStmt = db.prepare(`
        INSERT INTO transaksi (
          tanggal, kategori, nama_toko, jenis_barang, qty, satuan,
          harga_satuan, jumlah, detail, nomor_plat, no_hp, keterangan, sumber_sheet
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);

      let count = 0;
      db.exec('BEGIN');
      try {
        for (const it of data.items) {
          const jumlah = parseMoney(it.jumlah);
          if (jumlah > 0) {
            insertStmt.run(
              data.tanggal,
              data.kategori,
              (data.nama_toko || '').trim(),
              (it.jenis_barang || '').trim(),
              parseQty(it.qty),
              (it.satuan || '').trim(),
              parseMoney(it.harga_satuan),
              jumlah,
              (it.detail || '').trim(),
              (data.nomor_plat || '').trim(),
              '',
              '',
              sumber
            );
            count++;
          }
        }
        db.exec('COMMIT');
      } catch (e) {
        db.exec('ROLLBACK');
        throw e;
      }

      return sendJson(201, { success: true, count });
    }

    // 4. API: Update Transaksi
    if (pathname.startsWith('/api/transactions/') && method === 'PUT') {
      const id = parseInt(pathname.split('/').pop(), 10);
      const data = await getRequestBody(req);

      const jumlah = parseMoney(data.jumlah);
      const hargaSatuan = parseMoney(data.harga_satuan);
      const qty = parseQty(data.qty);

      const stmt = db.prepare(`
        UPDATE transaksi SET
          tanggal = ?,
          kategori = ?,
          nama_toko = ?,
          jenis_barang = ?,
          qty = ?,
          satuan = ?,
          harga_satuan = ?,
          jumlah = ?,
          detail = ?,
          nomor_plat = ?,
          no_hp = ?,
          keterangan = ?
        WHERE id = ?
      `);

      stmt.run(
        data.tanggal,
        data.kategori,
        data.nama_toko || '',
        data.jenis_barang || '',
        qty,
        data.satuan || '',
        hargaSatuan,
        jumlah,
        data.detail || '',
        data.nomor_plat || '',
        data.no_hp || '',
        data.keterangan || '',
        id
      );

      return sendJson(200, { success: true });
    }

    // 5. API: Hapus Transaksi
    if (pathname.startsWith('/api/transactions/') && method === 'DELETE') {
      const id = parseInt(pathname.split('/').pop(), 10);
      db.prepare("DELETE FROM transaksi WHERE id = ?").run(id);
      return sendJson(200, { success: true });
    }

    // 6. API: Export Excel Terformat (.xlsx native tanpa library pihak ketiga)
    if (pathname === '/api/export' && method === 'GET') {
      const month = parsedUrl.searchParams.get('month') || '';
      let query = "SELECT * FROM transaksi WHERE 1=1";
      const params = [];
      if (month && month !== 'ALL') {
        query += " AND strftime('%Y-%m', tanggal) = ?";
        params.push(month);
      }
      query += " ORDER BY tanggal ASC, id ASC";

      const rows = db.prepare(query).all(...params);

      const monthNames = {
        '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
        '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
        '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
      };

      let periodLabel = 'Semua Periode';
      if (month && month !== 'ALL' && month.includes('-')) {
        const [y, m] = month.split('-');
        periodLabel = `${monthNames[m] || m} ${y}`;
      }

      const subtotals = {
        'Bunga': 0,
        'Keperluan Florist': 0,
        'Kendaraan Florist': 0,
        'Biaya Lainnya': 0,
        'Dels': 0
      };
      let grandTotal = 0;
      for (const r of rows) {
        const amt = Number(r.jumlah) || 0;
        if (subtotals[r.kategori] !== undefined) {
          subtotals[r.kategori] += amt;
        }
        grandTotal += amt;
      }

      const xlsxBuffer = generateXlsx(periodLabel, rows, subtotals, grandTotal);
      const filename = (month && month !== 'ALL') ? `Jurnal_DQueen_${month}.xlsx` : 'Jurnal_DQueen_Semua_Periode.xlsx';

      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`
      });
      return res.end(xlsxBuffer);
    }

    // 7. Static Files (public/index.html)
    let filePath = path.join(__dirname, 'public', pathname === '/' ? 'index.html' : pathname);
    if (!fs.existsSync(filePath)) {
      filePath = path.join(__dirname, 'public', 'index.html');
    }

    const ext = path.extname(filePath);
    const mimeTypes = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'application/javascript; charset=utf-8',
      '.css': 'text/css; charset=utf-8',
      '.json': 'application/json',
      '.png': 'image/png',
      '.jpg': 'image/jpeg'
    };

    const contentType = mimeTypes[ext] || 'text/plain';
    const content = fs.readFileSync(filePath);
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(content);

  } catch (err) {
    console.error("Server Error:", err);
    sendJson(500, { error: err.message });
  }
});

server.listen(PORT, () => {
  console.log(`===============================================`);
  console.log(`  Jurnal Keuangan DQueen Florist Aktif!`);
  console.log(`  Akses di browser: http://localhost:${PORT}`);
  console.log(`===============================================`);
});
