const { setBorder, mergeAndStyle, getColumnLetter, getNamaPendekKapal } = require('./excelHelpers');
const ProduksiPenumpangModel = require('../../models/ProduksiPenumpang');
const ProduksiKendaraanModel = require('../../models/ProduksiKendaraan');
const path = require('path');

const GOL_ROW_MAP = [
  { row: 51, nomor: 1, tipe: null },
  { row: 53, nomor: 2, tipe: null },
  { row: 55, nomor: 3, tipe: null },
  { row: 57, nomor: 4, tipe: "penumpang" },
  { row: 59, nomor: 4, tipe: "barang" },
  { row: 61, nomor: 5, tipe: "penumpang" },
  { row: 63, nomor: 5, tipe: "barang" },
  { row: 65, nomor: 6, tipe: "penumpang" },
  { row: 67, nomor: 6, tipe: "barang" },
  { row: 69, nomor: 7, tipe: null },
  { row: 71, nomor: 8, tipe: null },
  { row: 73, nomor: 9, tipe: null },
];

async function generateSheetPerKapal(
  workbook,
  kapal,
  rute,
  produksiRute,
  tarifKendaraan,
  tarifPenumpang,
  tanggalDari,
  tanggalSampai,
  suratDokumen,
  kapalIndex,
  dataPerRute
) {
  const namaKapalPendek = getNamaPendekKapal(kapal.nama_kapal);
  const pelabuhanAsal = rute.pelabuhan_asal.nama_pelabuhan;
  const pelabuhanTujuan = rute.pelabuhan_tujuan.nama_pelabuhan;
  
  const labelAsalShort = pelabuhanAsal.substring(0, 3);
  const labelTujuanShort = pelabuhanTujuan.substring(0, 3);

  // Cari kapalAsdpTemplate dari dataPerRute untuk rute ini
  const ruteData = dataPerRute.find(rd => rd.rute.rute_id === rute.rute_id);
  const kapalAsdpTemplateList = ruteData ? ruteData.kapalAsdpTemplate : [];

  // Filter produksi untuk kapal ini
  const produksiKapal = produksiRute.filter(p => p.kapal_id === kapal.kapal_id);
  
  console.log(`Generating sheets for ${namaKapalPendek} on route ${rute.nama_rute}:`);
  console.log(`  Produksi count: ${produksiKapal.length}`);

  // Pisahkan produksi berdasarkan arah
  const produksiAsalTujuan = produksiKapal.filter(p => {
    const asalProd = String(p.nama_pelabuhan_asal || "").toUpperCase();
    const tujuanProd = String(p.nama_pelabuhan_tujuan || "").toUpperCase();
    return asalProd.includes(pelabuhanAsal.toUpperCase()) && 
           tujuanProd.includes(pelabuhanTujuan.toUpperCase());
  });

  const produksiTujuanAsal = produksiKapal.filter(p => {
    const asalProd = String(p.nama_pelabuhan_asal || "").toUpperCase();
    const tujuanProd = String(p.nama_pelabuhan_tujuan || "").toUpperCase();
    return asalProd.includes(pelabuhanTujuan.toUpperCase()) && 
           tujuanProd.includes(pelabuhanAsal.toUpperCase());
  });

  console.log(`  Asal→Tujuan: ${produksiAsalTujuan.length}, Tujuan→Asal: ${produksiTujuanAsal.length}`);

  // Format bulan
  const bulanNama = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
  ];
  const bulan = bulanNama[tanggalDari.getMonth()];
  const tahun = tanggalDari.getFullYear();
  const jarak = rute.jarak || 0;

  // Preload data penumpang dan kendaraan untuk semua produksi
  const penumpangPerProduksi = {};
  const kendaraanPerProduksi = {};
  
  for (const p of produksiKapal) {
    try {
      penumpangPerProduksi[p.produksi_id] = await ProduksiPenumpangModel.getByProduksi(p.produksi_id);
      kendaraanPerProduksi[p.produksi_id] = await ProduksiKendaraanModel.getByProduksi(p.produksi_id);
    } catch (e) {
      penumpangPerProduksi[p.produksi_id] = [];
      kendaraanPerProduksi[p.produksi_id] = [];
    }
  }

  // Helper function to ensure unique sheet name
  const ensureUniqueSheetName = (baseName) => {
    let sheetName = baseName.substring(0, 31);
    let counter = 1;
    
    // Check if sheet already exists
    while (workbook.getWorksheet(sheetName)) {
      const suffix = `_${counter}`;
      const maxLength = 31 - suffix.length;
      sheetName = baseName.substring(0, maxLength) + suffix;
      counter++;
    }
    
    return sheetName;
  };

  // Sheet 1: Kapal-Asal-Tujuan (Rodhita-LEMBAR-JANGKAR)
  const sheetName1 = ensureUniqueSheetName(`${namaKapalPendek}-${pelabuhanAsal.toUpperCase()}-${pelabuhanTujuan.toUpperCase()}`);
  await createDetailSheet(
    workbook,
    sheetName1,
    kapal,
    rute,
    produksiAsalTujuan,
    penumpangPerProduksi,
    kendaraanPerProduksi,
    tarifKendaraan,
    tarifPenumpang,
    tanggalDari,
    tanggalSampai,
    suratDokumen,
    bulan,
    tahun,
    "asal_tujuan",
    null,
    null,
    jarak,
    kapalIndex,
    dataPerRute,
    kapalAsdpTemplateList
  );
  
  // Sheet 2: Kapal-Tujuan-Asal (Rodhita-JANGKAR-LEMBAR)
  const sheetName2 = ensureUniqueSheetName(`${namaKapalPendek}-${pelabuhanTujuan.toUpperCase()}-${pelabuhanAsal.toUpperCase()}`);
  await createDetailSheet(
    workbook,
    sheetName2,
    kapal,
    rute,
    produksiTujuanAsal,
    penumpangPerProduksi,
    kendaraanPerProduksi,
    tarifKendaraan,
    tarifPenumpang,
    tanggalDari,
    tanggalSampai,
    suratDokumen,
    bulan,
    tahun,
    "tujuan_asal",
    sheetName1,
    null,
    jarak,
    kapalIndex,
    dataPerRute,
    kapalAsdpTemplateList
  );
  
  // Sheet 3: Kapal-LBR-Tujuan (Rodhita-LEM-JAN) - GABUNGAN
  const sheetName3 = ensureUniqueSheetName(`${namaKapalPendek}-${labelAsalShort.toUpperCase()}-${labelTujuanShort.toUpperCase()}`);
  await createDetailSheet(
    workbook,
    sheetName3,
    kapal,
    rute,
    produksiKapal, // Semua produksi (gabungan)
    penumpangPerProduksi,
    kendaraanPerProduksi,
    tarifKendaraan,
    tarifPenumpang,
    tanggalDari,
    tanggalSampai,
    suratDokumen,
    bulan,
    tahun,
    "gabungan1",
    sheetName1,
    sheetName2,
    jarak,
    kapalIndex,
    dataPerRute,
    kapalAsdpTemplateList
  );

  return { sheetName1, sheetName2, sheetName3 };
}

module.exports = { generateSheetPerKapal };

async function createDetailSheet(
  workbook,
  sheetName,
  kapal,
  rute,
  produksiData,
  penumpangPerProduksi,
  kendaraanPerProduksi,
  tarifKendaraan,
  tarifPenumpang,
  tanggalDari,
  tanggalSampai,
  suratDokumen,
  bulan,
  tahun,
  arahRute,
  namaSheetLbr,
  namaSheetPdg,
  jarakRute,
  kapalIndex,
  dataPerRute,
  kapalAsdpTemplate
) {
  const ws = workbook.addWorksheet(sheetName);
  ws.properties.defaultFont = { name: "Calibri", size: 11 };
  ws.properties.tabColor = { argb: "FF00B050" };

  const pelabuhanAsal = rute.pelabuhan_asal.nama_pelabuhan;
  const pelabuhanTujuan = rute.pelabuhan_tujuan.nama_pelabuhan;
  const namaKapalPendek = getNamaPendekKapal(kapal.nama_kapal);
  const namaKapalDisplay = `KMP. ${namaKapalPendek.toUpperCase()}`;
  const gt = kapal.gt || "0";
  const kapasitasPnp = kapal.kapasitas_penumpang || "0";
  const kapasitasKnd = kapal.kapasitas_kendaraan || "0";

  const isGabungan1 = arahRute === "gabungan1";
  const is1Arah = arahRute === "asal_tujuan" || arahRute === "tujuan_asal";

  // Helper functions
  const b = (cell, s = "thin") => {
    cell.border = {
      top: { style: s },
      left: { style: s },
      bottom: { style: s },
      right: { style: s },
    };
  };

  const bOuter = (r1, c1, r2, c2, s = "medium") => {
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(r, c);
        const cur = cell.border || {};
        cell.border = {
          top: r === r1 ? { style: s } : cur.top || {},
          bottom: r === r2 ? { style: s } : cur.bottom || {},
          left: c === c1 ? { style: s } : cur.left || {},
          right: c === c2 ? { style: s } : cur.right || {},
        };
      }
  };

  const sc = (addr, val, opts = {}) => {
    const c = ws.getCell(addr);
    if (val !== undefined) c.value = val;
    if (opts.bold)
      c.font = { name: "Calibri", size: opts.size || 11, bold: true };
    if (opts.align) c.alignment = opts.align;
    if (opts.border) b(c, opts.border);
    if (opts.numFmt) c.numFmt = opts.numFmt;
    return c;
  };

  const mc = (range, val, opts = {}) => {
    try {
      ws.mergeCells(range);
    } catch (e) {}
    return sc(range.split(":")[0], val, opts);
  };

  // Lintasan dinamis berdasarkan rute
  const lintasan =
    arahRute === "asal_tujuan"
      ? `${pelabuhanAsal.toUpperCase()} - ${pelabuhanTujuan.toUpperCase()} / ${jarakRute} NM`
      : arahRute === "tujuan_asal"
        ? `${pelabuhanTujuan.toUpperCase()} - ${pelabuhanAsal.toUpperCase()} / ${jarakRute} NM`
        : `${pelabuhanAsal.toUpperCase()} - ${pelabuhanTujuan.toUpperCase()} / ${jarakRute} NM`;

  // ========== COLUMN WIDTHS ==========
  [4, 6, 25, 15, 12, 12, 7, 7, 18, 18, 10, 10].forEach(
    (w, i) => (ws.getColumn(i + 1).width = w),
  );

  // ========== HEADER SECTION (Rows 2-12) ==========
  bOuter(2, 2, 5, 3, "thin");
  
  // Logo
  try {
    const logoPath = path.join(__dirname, "../../public/images/logo_asdp.png");
    const logoId = workbook.addImage({
      filename: logoPath,
      extension: "png",
    });
    ws.addImage(logoId, {
      tl: { col: 2.5, row: 1.2 },
      ext: { width: 70, height: 70 },
      editAs: "absolute",
    });
  } catch (e) {
    console.warn("Failed to add logo:", e.message);
  }

  mc("D3:I4", "LAPORAN PRODUKSI\nDAN PENDAPATAN KAPAL BULANAN ", {
    bold: true,
    size: 14,
    align: { horizontal: "center", vertical: "middle", wrapText: true },
  });

  bOuter(2, 4, 5, 9, "thin");

  sc("J2", "No Dokumen");
  sc("K2", `: ${suratDokumen.no_dokumen || ""}`);
  sc("J3", "Revisi");
  sc("K3", `: ${suratDokumen.revisi || ""}`);
  sc("J4", "Berlaku Efektif");
  sc("K4", ":");
  sc("J5", "Halaman");
  sc("K5", `: ${suratDokumen.halaman || "1 dari 1"}`);
  bOuter(2, 10, 5, 12, "thin");

  sc("B7", "CABANG", { bold: true, size: 14 });
  sc("D7", `: ${pelabuhanAsal.toUpperCase()}`, { bold: true, size: 14 });
  sc("B8", "USAHA", { bold: true, size: 14 });
  sc("D8", ": BISNIS PENYEBERANGAN", { bold: true, size: 14 });
  sc("B9", "BULAN", { bold: true, size: 14 });
  sc("D9", `: ${bulan.toUpperCase()} ${tahun}`, { bold: true, size: 14 });
  sc("J9", "BULANAN", {
    name: "Times New Roman",
    bold: true,
    border: "thin",
    align: { horizontal: "center", vertical: "middle" },
  });
  sc("B10", "LINTASAN / JARAK", { bold: true, size: 14 });
  sc("D10", `: ${lintasan}`, { bold: true, size: 14 });
  sc("B11", "NAMA KAPAL / GRT", { bold: true, size: 14 });
  const gtFormatted = gt && gt !== "0" ? `(${Math.round(parseFloat(gt))})` : "";
  sc("D11", `: ${namaKapalDisplay.toUpperCase()} ${gtFormatted}`, { bold: true, size: 14 });
  sc("B12", "KAPASITAS ANGKUT", { bold: true, size: 14 });
  sc("D12", `: a. Penumpang: ${kapasitasPnp} org, -Kendaraan: ${kapasitasKnd} unit (campuran)`, { bold: true, size: 11 });

  // ========== ROW 14-15: Kategori header ==========
  ["D", "E", "F", "G", "H", "I", "J", "K", "L"].forEach((col, i) => {
    const labels = ["Eks", "Bis.I", "Bis.II", "Eko.", "Gol IV", "Gol V", "Gol VI", "Gol VII", "Gol VIII"];
    const cell14 = sc(`${col}14`, labels[i], {
      border: "thin",
      align: { horizontal: "center", vertical: "middle" },
    });
    cell14.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "ffccffff" },
    };
    cell14.font = { name: "Calibri", size: 11 };
    sc(`${col}15`, "", { border: "thin" });
  });

  const blueHeaderFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "ffccffff" },
  };
  const blueHeaderFont = {
    name: "Calibri",
    size: 11,
    bold: true,
  };

  const yellowFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };

  // Yellow fills untuk data cells
  ["F20", "F21", "F22"].forEach(addr => (ws.getCell(addr).fill = yellowFill));
  ["D38", "D39", "D40", "D41"].forEach(addr => (ws.getCell(addr).fill = yellowFill));
  for (let r = 38; r <= 41; r++) {
    for (let c = 6; c <= 8; c++) ws.getCell(r, c).fill = yellowFill;
    for (let c = 10; c <= 12; c++) ws.getCell(r, c).fill = yellowFill;
  }
  ["F48", "G48", "H48", "J48", "K48", "L48"].forEach(addr => (ws.getCell(addr).fill = yellowFill));
  for (let r = 51; r <= 75; r++) {
    ws.getCell(r, 4).fill = yellowFill;
    for (let c = 6; c <= 8; c++) ws.getCell(r, c).fill = yellowFill;
    for (let c = 10; c <= 12; c++) ws.getCell(r, c).fill = yellowFill;
  }
  for (let r = 102; r <= 104; r++) {
    ws.getCell(r, 4).fill = yellowFill;
    for (let c = 6; c <= 8; c++) ws.getCell(r, c).fill = yellowFill;
    for (let c = 10; c <= 12; c++) ws.getCell(r, c).fill = yellowFill;
  }

  // ========== TABLE HEADERS (Rows 17-19) ==========
  const h17b18 = mc("B17:B18", "NO", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h17b18.fill = blueHeaderFill;
  h17b18.font = blueHeaderFont;

  const h17c18 = mc("C17:C18", "JENIS TI-T", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h17c18.fill = blueHeaderFill;
  h17c18.font = blueHeaderFont;

  const h17d18 = mc("D17:D18", "TARIF (Rp.)", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h17d18.fill = blueHeaderFill;
  h17d18.font = blueHeaderFont;

  const h17prod = mc("E17:H17", "PRODUKSI", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h17prod.fill = blueHeaderFill;
  h17prod.font = blueHeaderFont;

  const h17pend = mc("I17:L17", "PENDAPATAN", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h17pend.fill = blueHeaderFill;
  h17pend.font = blueHeaderFont;

  const h18renc = sc("E18", "RENC.", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18renc.fill = blueHeaderFill;
  h18renc.font = blueHeaderFont;

  const h18real = sc("F18", "REAL", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18real.fill = blueHeaderFill;
  h18real.font = blueHeaderFont;

  const h18pct1 = mc("G18:H18", "(%)", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18pct1.fill = blueHeaderFill;
  h18pct1.font = blueHeaderFont;

  const h18rencana = sc("I18", "RENCANA", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18rencana.fill = blueHeaderFill;
  h18rencana.font = blueHeaderFont;

  const h18realisasi = sc("J18", "REALISASI", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18realisasi.fill = blueHeaderFill;
  h18realisasi.font = blueHeaderFont;

  const h18pct2 = mc("K18:L18", "(%)", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  h18pct2.fill = blueHeaderFill;
  h18pct2.font = blueHeaderFont;

  [
    ["B19", "1"],
    ["C19", "2"],
    ["D19", "3"],
    ["E19", "4"],
    ["F19", "5"],
    ["I19", "7"],
    ["J19", "8"],
  ].forEach(([a, v]) => {
    const c19 = sc(a, v, {
      border: "medium",
      align: { horizontal: "center" },
    });
    c19.fill = blueHeaderFill;
    c19.font = blueHeaderFont;
  });
  const g19 = mc("G19:H19", "6", {
    border: "medium",
    align: { horizontal: "center" },
  });
  g19.fill = blueHeaderFill;
  g19.font = blueHeaderFont;
  const k19 = mc("K19:L19", "9", {
    border: "medium",
    align: { horizontal: "center" },
  });
  k19.fill = blueHeaderFill;
  k19.font = blueHeaderFont;

  // ========== HELPER FUNCTIONS FOR DATA ==========
  const hitungPenumpangDetail = (produksiFiltArr) => {
    const result = {
      dewasa: { normal: 0, diskon: 0, tarifNormal: 0, tarifDiskon: 0 },
      anak: { normal: 0, diskon: 0, tarifNormal: 0, tarifDiskon: 0 },
    };
    
    produksiFiltArr.forEach((p) => {
      const pArr = penumpangPerProduksi[p.produksi_id] || [];
      pArr.forEach((k) => {
        const nama = String(k.nama_kategori || "").toLowerCase();
        const jumlah = k.jumlah || 0;
        const tarif = k.tarif || 0;
        const isCustom = k.is_tarif_custom === true || k.is_tarif_custom === 1;
        
        if (nama.includes("dewasa")) {
          if (isCustom) {
            result.dewasa.diskon += jumlah;
            if (jumlah > 0 && result.dewasa.tarifDiskon === 0) {
              result.dewasa.tarifDiskon = tarif;
            }
          } else {
            result.dewasa.normal += jumlah;
            if (jumlah > 0 && result.dewasa.tarifNormal === 0) {
              result.dewasa.tarifNormal = tarif;
            }
          }
        } else if (nama.includes("bayi") || nama.includes("anak")) {
          if (isCustom) {
            result.anak.diskon += jumlah;
            if (jumlah > 0 && result.anak.tarifDiskon === 0) {
              result.anak.tarifDiskon = tarif;
            }
          } else {
            result.anak.normal += jumlah;
            if (jumlah > 0 && result.anak.tarifNormal === 0) {
              result.anak.tarifNormal = tarif;
            }
          }
        }
      });
    });
    
    return result;
  };

  const hitungKendaraanDetail = (produksiFiltArr) => {
    const result = {};
    
    produksiFiltArr.forEach((p) => {
      const kndArr = kendaraanPerProduksi[p.produksi_id] || [];
      kndArr.forEach((k) => {
        const key = `${k.nomor_golongan}_${(k.tipe_muatan || "").toLowerCase()}`;
        const jumlah = k.jumlah || 0;
        const tarif = k.tarif || 0;
        const isCustom = k.is_tarif_custom === true || k.is_tarif_custom === 1;
        
        if (!result[key]) {
          result[key] = {
            normal: 0,
            diskon: 0,
            tarifNormal: 0,
            tarifDiskon: 0,
          };
        }
        
        if (isCustom) {
          result[key].diskon += jumlah;
          if (jumlah > 0 && result[key].tarifDiskon === 0) {
            result[key].tarifDiskon = tarif;
          }
        } else {
          result[key].normal += jumlah;
          if (jumlah > 0 && result[key].tarifNormal === 0) {
            result[key].tarifNormal = tarif;
          }
        }
      });
    });
    
    return result;
  };

  const getKndKey = (g) => {
    if (!g.tipe) return `${g.nomor}_`;
    return `${g.nomor}_${g.tipe}`;
  };

  // Hitung data
  const pnpDetail = hitungPenumpangDetail(produksiData);
  const kndDetail = hitungKendaraanDetail(produksiData);

  const totalTripAktual = produksiData.length;
  const tanggalSet = new Set(
    produksiData.map((p) => new Date(p.tanggal_produksi).toDateString()),
  );
  const hariOpsAktual = tanggalSet.size;

  // Tarif penumpang dari database
  const cariTarifPenumpang = (tarifArr, kategori) => {
    const found = tarifArr.find((t) => {
      const namaKategori = String(t.kategori?.nama_kategori || "").toLowerCase();
      return namaKategori.includes(kategori.toLowerCase());
    });
    return found ? found.tarif || 0 : 0;
  };

  const tarifDewasaNormal = pnpDetail.dewasa.tarifNormal || cariTarifPenumpang(tarifPenumpang, "dewasa");
  const tarifDewasaDiskon = pnpDetail.dewasa.tarifDiskon || null;
  const tarifAnakNormal = pnpDetail.anak.tarifNormal || cariTarifPenumpang(tarifPenumpang, "bayi");
  const tarifAnakDiskon = pnpDetail.anak.tarifDiskon || null;

  // Tarif kendaraan
  const cariTarif = (tarifArr, nomorGol, tipe) => {
    const found = tarifArr.find((t) => {
      const nMatch = parseInt(t.golongan?.nomor_golongan || t.nomor_golongan || 0) === nomorGol;
      if (!tipe) return nMatch && !t.golongan?.tipe_muatan && !t.tipe_muatan;
      const tMatch = String(t.golongan?.tipe_muatan || t.tipe_muatan || "").toLowerCase().includes(tipe);
      return nMatch && tMatch;
    });
    return found ? found.tarif || 0 : 0;
  };

  const tarifPerRow = {};
  GOL_ROW_MAP.forEach((g) => {
    tarifPerRow[g.row] = cariTarif(tarifKendaraan, g.nomor, g.tipe);
  });

  // Formula helpers
  const rencE = (row) => {
    if (!is1Arah || row === 20) return "";
    const namaSheetGab = `${namaKapalPendek}-${pelabuhanAsal.substring(0, 3)}-${pelabuhanTujuan.substring(0, 3)}`.substring(0, 31);
    return {
      formula: `IF(OR('${namaSheetGab}'!E${row}=0,'${namaSheetGab}'!E${row}=""),"",'${namaSheetGab}'!E${row}/2)`,
    };
  };

  const realFGab1 = (row) => {
    if (!isGabungan1 || !namaSheetLbr || !namaSheetPdg) return null;
    return {
      formula: `'${namaSheetLbr}'!F${row}+'${namaSheetPdg}'!F${row}`,
    };
  };

  const realJFormula = (row) => {
    if (isGabungan1 && namaSheetLbr && namaSheetPdg)
      return `'${namaSheetLbr}'!J${row}+'${namaSheetPdg}'!J${row}`;
    return `F${row}*D${row}`;
  };

  // ========== DATA ROWS FUNCTION ==========
  const brsStd = (row, no, label, tarif, eVal, fVal, iFormula, jFormula) => {
    if (no !== null)
      sc(`B${row}`, no, { align: { horizontal: "center", vertical: "middle" } });
    else b(ws.getCell(`B${row}`));
    sc(`C${row}`, label, { align: { vertical: "middle" } });

    // D tarif
    const dc = ws.getCell(`D${row}`);
    if (tarif || tarif === 0) {
      dc.value = Math.round(tarif);
      dc.numFmt = "#,##0";
      dc.alignment = { horizontal: "right", vertical: "middle" };
    }
    b(dc);

    // E rencana
    const ec = ws.getCell(`E${row}`);
    if (eVal !== "" && eVal !== null && eVal !== undefined)
      ec.value = eVal;
    ec.numFmt = "#,##0";
    ec.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
    b(ec);

    // F realisasi
    const fc = ws.getCell(`F${row}`);
    if (fVal !== null && fVal !== undefined && fVal !== "") {
      fc.value = fVal;
      fc.numFmt = "#,##0";
    }
    b(fc);

    // G:H = %
    try {
      ws.mergeCells(`G${row}:H${row}`);
    } catch (e2) {}
    const gc = ws.getCell(`G${row}`);
    gc.value = {
      formula: `IF(OR(E${row}=0,F${row}=0),"",F${row}/E${row})`,
    };
    gc.numFmt = "0.00%";
    b(gc);

    // I rencana pendapatan
    const ic = ws.getCell(`I${row}`);
    if (iFormula) {
      ic.value = { formula: iFormula };
      ic.numFmt = "#,##0";
    }
    ic.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
    b(ic);

    // J realisasi pendapatan
    const jc = ws.getCell(`J${row}`);
    if (jFormula) {
      jc.value = { formula: jFormula };
      jc.numFmt = "#,##0";
    }
    b(jc);

    // K:L = %
    try {
      ws.mergeCells(`K${row}:L${row}`);
    } catch (e2) {}
    const kc = ws.getCell(`K${row}`);
    if (iFormula && jFormula) {
      kc.value = { formula: `IF(I${row}=0,"",J${row}/I${row})` };
      kc.numFmt = "0.00%";
    }
    b(kc);
  };

  const brsHdr = (row, no, label, eRef) => {
    if (no)
      sc(`B${row}`, no, {
        border: "thin",
        align: { bold: true, horizontal: "center", vertical: "middle" },
      });
    else b(ws.getCell(`B${row}`));
    sc(`C${row}`, label, {
      bold: true,
      border: "thin",
      align: { vertical: "middle" },
    });
    b(ws.getCell(`D${row}`));
    const ec = ws.getCell(`E${row}`);
    if (eRef) {
      ec.value = eRef;
      ec.numFmt = "#,##0";
    }
    ec.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
    b(ec);
    b(ws.getCell(`F${row}`));
    try {
      ws.mergeCells(`G${row}:H${row}`);
    } catch (e2) {}
    b(ws.getCell(`G${row}`));
    ws.getCell(`I${row}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
    b(ws.getCell(`I${row}`));
    b(ws.getCell(`J${row}`));
    try {
      ws.mergeCells(`K${row}:L${row}`);
    } catch (e2) {}
    b(ws.getCell(`K${row}`));
  };

  // ========== DATA SECTION (Rows 20-104) ==========
  
  // Rows 20-22: Hari Operasi & Trip
  // Hitung row di sheet Kinerja KAPAL bagian bawah (detail per arah)
  // Struktur bagian bawah dimulai dari row 40
  // Untuk setiap rute:
  //   - 1 baris nama rute
  //   - 1 baris header tabel
  //   - 1 baris tanggal
  //   - 1 baris "total trip/tidak ops/hari ops"
  //   - 3 baris per kapal (asal→tujuan, tujuan→asal, total)
  //   - 1 baris spasi antar kapal
  //   - 1 baris spasi antar rute
  
  // Hitung offset untuk rute ini
  let kinerjaDetailRow = 40; // Baris awal bagian detail
  
  // Loop semua rute sebelum rute ini untuk menghitung offset
  const ruteIndex = dataPerRute.findIndex(rd => rd.rute.rute_id === rute.rute_id);
  for (let i = 0; i < ruteIndex; i++) {
    const prevRute = dataPerRute[i];
    kinerjaDetailRow += 1; // Nama rute
    kinerjaDetailRow += 1; // Header tabel
    kinerjaDetailRow += 1; // Tanggal
    kinerjaDetailRow += 1; // "total trip/tidak ops/hari ops"
    kinerjaDetailRow += prevRute.kapalAsdpTemplate.length * 4; // 3 baris data + 1 spasi per kapal
    kinerjaDetailRow += 1; // Spasi antar rute
  }
  
  // Sekarang kinerjaDetailRow menunjuk ke baris nama rute saat ini
  kinerjaDetailRow += 1; // Skip nama rute
  kinerjaDetailRow += 1; // Skip header tabel
  kinerjaDetailRow += 1; // Skip tanggal
  kinerjaDetailRow += 1; // Baris "total trip/tidak ops/hari ops"
  
  // Hitung posisi kapal ini dalam rute
  const kapalIndexInRute = kapalAsdpTemplate.findIndex(k => k.kapal_id === kapal.kapal_id);
  const kinerjaRowAsalTujuan = kinerjaDetailRow + (kapalIndexInRute * 4); // 3 baris data + 1 spasi per kapal sebelumnya
  const kinerjaRowTujuanAsal = kinerjaRowAsalTujuan + 1;
  const kinerjaRowGabungan = kinerjaRowAsalTujuan + 2;

  let f20val, f21val, f22val;
  if (arahRute === "asal_tujuan") {
    f20val = 0;
    f21val = { formula: `'Kinerja KAPAL'!AL${kinerjaRowAsalTujuan}` };
    f22val = { formula: `'Kinerja KAPAL'!AI${kinerjaRowAsalTujuan}` };
  } else if (arahRute === "tujuan_asal") {
    f20val = { formula: `'${namaSheetLbr}'!F20` };
    f21val = { formula: `'Kinerja KAPAL'!AL${kinerjaRowTujuanAsal}` };
    f22val = { formula: `'Kinerja KAPAL'!AI${kinerjaRowTujuanAsal}` };
  } else if (arahRute === "gabungan1") {
    f20val = { formula: `'${namaSheetLbr}'!F20` };
    f21val = { formula: `'Kinerja KAPAL'!AL${kinerjaRowGabungan}` };
    f22val = { formula: `'${namaSheetLbr}'!F22+'${namaSheetPdg}'!F22` };
  } else {
    f21val = hariOpsAktual;
    f22val = totalTripAktual;
  }

  // Baris 20-23: Tanpa bold
  sc("B20", "1", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("C20", "HARI SIAP OPERASI", { border: "thin", align: { vertical: "middle" } });
  b(ws.getCell("D20"));
  const e20 = ws.getCell("E20");
  e20.value = rencE(20);
  e20.numFmt = "#,##0";
  e20.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e20);
  ws.getCell("F20").value = f20val;
  ws.getCell("F20").numFmt = "#,##0";
  b(ws.getCell("F20"));
  try { ws.mergeCells("G20:H20"); } catch (e2) {}
  ws.getCell("G20").value = { formula: 'IF(OR(E20=0,F20=0,E20="",F20=""),"",F20/E20)' };
  ws.getCell("G20").numFmt = "0.00%";
  b(ws.getCell("G20"));
  ws.getCell("I20").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(ws.getCell("I20"));
  b(ws.getCell("J20"));
  try { ws.mergeCells("K20:L20"); } catch (e2) {}
  b(ws.getCell("K20"));

  sc("B21", "2", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("C21", "HARI OPERASI", { border: "thin", align: { vertical: "middle" } });
  b(ws.getCell("D21"));
  const e21 = ws.getCell("E21");
  e21.value = rencE(21);
  e21.numFmt = "#,##0";
  e21.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e21);
  ws.getCell("F21").value = f21val;
  ws.getCell("F21").numFmt = "#,##0";
  b(ws.getCell("F21"));
  try { ws.mergeCells("G21:H21"); } catch (e2) {}
  ws.getCell("G21").value = { formula: 'IF(OR(E21=0,F21=0,E21="",F21=""),"",F21/E21)' };
  ws.getCell("G21").numFmt = "0.00%";
  b(ws.getCell("G21"));
  ws.getCell("I21").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(ws.getCell("I21"));
  b(ws.getCell("J21"));
  try { ws.mergeCells("K21:L21"); } catch (e2) {}
  b(ws.getCell("K21"));

  sc("B22", "3", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("C22", "TRIP", { border: "thin", align: { vertical: "middle" } });
  b(ws.getCell("D22"));
  const e22 = ws.getCell("E22");
  e22.value = rencE(22);
  e22.numFmt = "#,##0";
  e22.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e22);
  ws.getCell("F22").value = f22val;
  ws.getCell("F22").numFmt = "#,##0";
  b(ws.getCell("F22"));
  try { ws.mergeCells("G22:H22"); } catch (e2) {}
  ws.getCell("G22").value = { formula: 'IF(OR(E22=0,F22=0,E22="",F22=""),"",F22/E22)' };
  ws.getCell("G22").numFmt = "0.00%";
  b(ws.getCell("G22"));
  ws.getCell("I22").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(ws.getCell("I22"));
  b(ws.getCell("J22"));
  try { ws.mergeCells("K22:L22"); } catch (e2) {}
  b(ws.getCell("K22"));

  sc("B23", "4", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("C23", "PRODUKSI :", { border: "thin", align: { vertical: "middle" } });
  b(ws.getCell("D23"));
  const e23 = ws.getCell("E23");
  e23.value = rencE(23);
  e23.numFmt = "#,##0";
  e23.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e23);
  b(ws.getCell("F23"));
  try { ws.mergeCells("G23:H23"); } catch (e2) {}
  b(ws.getCell("G23"));
  ws.getCell("I23").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(ws.getCell("I23"));
  b(ws.getCell("J23"));
  try { ws.mergeCells("K23:L23"); } catch (e2) {}
  b(ws.getCell("K23"));
  brsHdr(24, "a.", "PENUMPANG", rencE(24));

  brsStd(25, "", "1) Eksekutif", null, rencE(25), null, null, null);
  brsStd(26, "", "- Dewasa", null, rencE(26), null, null, null);
  brsStd(27, "", "- Anak", null, rencE(27), null, null, null);
  brsStd(28, "", "- Bayi", null, rencE(28), null, null, null);

  brsStd(29, "", "2) Bisnis I", null, rencE(29), null, null, null);
  brsStd(30, "", "- Dewasa", null, rencE(30), null, null, null);
  brsStd(31, "", "- Anak", null, rencE(31), null, null, null);
  brsStd(32, "", "- Bayi", null, rencE(32), null, null, null);

  brsStd(33, "", "3) Bisnis II", null, rencE(33), null, null, null);
  brsStd(34, "", "- Dewasa", null, rencE(34), null, null, null);
  brsStd(35, "", "- Anak", null, rencE(35), null, null, null);
  brsStd(36, "", "- Bayi", null, rencE(36), null, null, null);

  brsStd(37, "", "4) Ekonomi", null, rencE(37), null, null, null);
  
  // Data penumpang dengan pemisahan normal dan diskon
  const fVal38 = isGabungan1
    ? { formula: `+'${namaSheetLbr}'!F38+'${namaSheetPdg}'!F38` }
    : pnpDetail.dewasa.normal;
    
  const fVal39 = isGabungan1
    ? { formula: `+'${namaSheetLbr}'!F39+'${namaSheetPdg}'!F39` }
    : pnpDetail.dewasa.diskon;
    
  const fVal40 = isGabungan1
    ? { formula: `+'${namaSheetLbr}'!F40+'${namaSheetPdg}'!F40` }
    : pnpDetail.anak.normal;
    
  const fVal41 = isGabungan1
    ? { formula: `+'${namaSheetLbr}'!F41+'${namaSheetPdg}'!F41` }
    : pnpDetail.anak.diskon;

  brsStd(
    38,
    "",
    "- Dewasa",
    tarifDewasaNormal,
    rencE(38),
    fVal38,
    `E38*D38`,
    isGabungan1 ? `'${namaSheetLbr}'!J38+'${namaSheetPdg}'!J38` : `F38*D38`,
  );
  
  brsStd(
    39,
    "",
    "- Dewasa (Diskon)",
    tarifDewasaDiskon,
    rencE(39),
    fVal39,
    tarifDewasaDiskon ? `E39*D39` : "",
    isGabungan1 ? `'${namaSheetLbr}'!J39+'${namaSheetPdg}'!J39` : tarifDewasaDiskon ? `F39*D39` : "",
  );
  
  brsStd(
    40,
    "",
    "- Anak",
    tarifAnakNormal,
    rencE(40),
    fVal40,
    `E40*D40`,
    isGabungan1 ? `'${namaSheetLbr}'!J40+'${namaSheetPdg}'!J40` : `F40*D40`,
  );
  
  brsStd(
    41,
    "",
    "- Anak (Diskon)",
    tarifAnakDiskon,
    rencE(41),
    fVal41,
    tarifAnakDiskon ? `E41*D41` : "",
    isGabungan1 ? `'${namaSheetLbr}'!J41+'${namaSheetPdg}'!J41` : tarifAnakDiskon ? `F41*D41` : "",
  );

  sc("C42", "- Bayi");
  b(ws.getCell("D42"));
  b(ws.getCell("E42"));
  b(ws.getCell("F42"));
  try { ws.mergeCells("G42:H42"); } catch (e2) {}
  b(ws.getCell("G42"));
  b(ws.getCell("I42"));
  b(ws.getCell("J42"));
  try { ws.mergeCells("K42:L42"); } catch (e2) {}
  b(ws.getCell("K42"));

  brsStd(43, "", "5) Suplesi", null, rencE(43), null, null, null);
  brsStd(44, "", "- Eksekutif", null, rencE(44), null, null, null);
  brsStd(45, "", "- Bisnis I", null, rencE(45), null, null, null);
  brsStd(46, "", "- Bisnis II Dewasa", null, rencE(46), null, null, null);
  brsStd(47, "", "- Bisnis II Anak", null, rencE(47), null, null, null);

  // Jumlah penumpang (row 48)
  sc("C48", "Jumlah (4.a.1 s/d 4.a.5)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D48"));
  const e48val = rencE(48);
  if (isGabungan1) {
    ws.getCell("E48").value = { formula: "SUM(E38:E47)" };
  } else if (e48val) {
    ws.getCell("E48").value = e48val;
  }
  ws.getCell("E48").numFmt = "#,##0";
  b(ws.getCell("E48"));
  ws.getCell("F48").value = { formula: "SUM(F38:F47)" };
  ws.getCell("F48").numFmt = "#,##0";
  b(ws.getCell("F48"));
  try { ws.mergeCells("G48:H48"); } catch (e2) {}
  ws.getCell("G48").value = { formula: 'IF(E48=0,"",F48/E48)' };
  ws.getCell("G48").numFmt = "0.00%";
  b(ws.getCell("G48"));
  ws.getCell("I48").value = { formula: "SUM(I38:I47)" };
  ws.getCell("I48").numFmt = "#,##0";
  b(ws.getCell("I48"));
  ws.getCell("J48").value = { formula: "SUM(J38:J47)" };
  ws.getCell("J48").numFmt = "#,##0";
  b(ws.getCell("J48"));
  try { ws.mergeCells("K48:L48"); } catch (e2) {}
  ws.getCell("K48").value = { formula: 'IF(I48=0,"",J48/I48)' };
  ws.getCell("K48").numFmt = "0.00%";
  b(ws.getCell("K48"));

  ["B", "C", "D", "E", "F", "I", "J"].forEach((col) => b(ws.getCell(`${col}49`)));
  try { ws.mergeCells("G49:H49"); } catch (e2) {}
  b(ws.getCell("G49"));
  try { ws.mergeCells("K49:L49"); } catch (e2) {}
  b(ws.getCell("K49"));

  // Kendaraan (row 50)
  brsHdr(50, "b.", "KENDARAAN", rencE(50));

  // Kendaraan per golongan (rows 51-74)
  const golLabel = [
    "- Golongan I",
    "- Golongan II",
    "- Golongan III",
    "- Golongan IV",
    "- Golongan IV Pick Up",
    "- Golongan V Bus",
    "- Golongan V Truk",
    "- Golongan VI Bus",
    "- Golongan VI Truk",
    "- Golongan VII",
    "- Golongan VIII",
    "- Golongan IX",
  ];

  let currentRow = 51;
  GOL_ROW_MAP.forEach((g, i) => {
    const tarif = tarifPerRow[g.row] || 0;
    const key = getKndKey(g);
    const detail = kndDetail[key] || { normal: 0, diskon: 0, tarifNormal: 0, tarifDiskon: 0 };
    
    // Baris normal
    const fValNormal = isGabungan1
      ? { formula: `'${namaSheetLbr}'!F${currentRow}+'${namaSheetPdg}'!F${currentRow}` }
      : detail.normal;
    const tarifNormal = detail.tarifNormal || tarif;
    const iFNormal = `E${currentRow}*D${currentRow}`;
    const jFNormal = isGabungan1 
      ? `'${namaSheetLbr}'!J${currentRow}+'${namaSheetPdg}'!J${currentRow}` 
      : `F${currentRow}*D${currentRow}`;
    
    brsStd(currentRow, "", golLabel[i], tarifNormal, rencE(currentRow), fValNormal, iFNormal, jFNormal);
    currentRow++;
    
    // Baris diskon
    const fValDiskon = isGabungan1
      ? { formula: `'${namaSheetLbr}'!F${currentRow}+'${namaSheetPdg}'!F${currentRow}` }
      : detail.diskon;
    const tarifDiskon = detail.tarifDiskon || null;
    const iFDiskon = tarifDiskon ? `E${currentRow}*D${currentRow}` : "";
    const jFDiskon = isGabungan1 
      ? `'${namaSheetLbr}'!J${currentRow}+'${namaSheetPdg}'!J${currentRow}` 
      : tarifDiskon ? `F${currentRow}*D${currentRow}` : "";
    
    brsStd(currentRow, "", golLabel[i] + " (Diskon)", tarifDiskon, rencE(currentRow), fValDiskon, iFDiskon, jFDiskon);
    currentRow++;
  });

  // Jumlah kendaraan (row 75)
  sc("C75", "Jumlah (4.b)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D75"));
  const e75val = rencE(75);
  if (isGabungan1) {
    ws.getCell("E75").value = { formula: "SUM(E51:E74)" };
  } else if (e75val) {
    ws.getCell("E75").value = e75val;
  }
  ws.getCell("E75").numFmt = "#,##0";
  b(ws.getCell("E75"));
  ws.getCell("F75").value = { formula: "SUM(F51:F74)" };
  ws.getCell("F75").numFmt = "#,##0";
  b(ws.getCell("F75"));
  try { ws.mergeCells("G75:H75"); } catch (e2) {}
  ws.getCell("G75").value = { formula: 'IF(E75=0,"",F75/E75)' };
  ws.getCell("G75").numFmt = "0.00%";
  b(ws.getCell("G75"));
  ws.getCell("I75").value = { formula: "SUM(I51:I74)" };
  ws.getCell("I75").numFmt = "#,##0";
  b(ws.getCell("I75"));
  ws.getCell("J75").value = { formula: "SUM(J51:J74)" };
  ws.getCell("J75").numFmt = "#,##0";
  b(ws.getCell("J75"));
  try { ws.mergeCells("K75:L75"); } catch (e2) {}
  ws.getCell("K75").value = { formula: 'IF(I75=0,"",J75/I75)' };
  ws.getCell("K75").numFmt = "0.00%";
  b(ws.getCell("K75"));

  // Load Factor (rows 76-79)
  brsHdr(76, "c.", "LOAD FACTOR", null);
  brsStd(77, "", "- Penumpang (%)", null, rencE(77), null, null, null);
  brsStd(78, "", "- Kendaraan (%)", null, rencE(78), null, null, null);
  sc("C79", "Jumlah (4.c)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D79"));
  const e79val = rencE(79);
  if (e79val) ws.getCell("E79").value = e79val;
  b(ws.getCell("E79"));
  b(ws.getCell("F79"));
  try { ws.mergeCells("G79:H79"); } catch (e2) {}
  b(ws.getCell("G79"));
  b(ws.getCell("I79"));
  b(ws.getCell("J79"));
  try { ws.mergeCells("K79:L79"); } catch (e2) {}
  b(ws.getCell("K79"));

  // Barang (rows 80-84)
  brsHdr(80, "d.", "BARANG", null);
  brsStd(81, "", "1) Ton/M Kubik", null, rencE(81), null, null, null);
  brsStd(82, "", "2) Brg curah/M Kubik", null, rencE(82), null, null, null);
  brsStd(83, "", "3) Brg tentengan/kg", null, rencE(83), null, null, null);
  sc("C84", "Jumlah (4.d)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D84"));
  const e84val = rencE(84);
  if (e84val) ws.getCell("E84").value = e84val;
  b(ws.getCell("E84"));
  b(ws.getCell("F84"));
  try { ws.mergeCells("G84:H84"); } catch (e2) {}
  b(ws.getCell("G84"));
  b(ws.getCell("I84"));
  b(ws.getCell("J84"));
  try { ws.mergeCells("K84:L84"); } catch (e2) {}
  b(ws.getCell("K84"));

  // Pendapatan lain (rows 85-100)
  brsHdr(85, "e.", "PENDAPATAN LAIN-2", null);
  [
    [86, "1) Hewan"],
    [87, "    a). Kambing & sejenisnya"],
    [88, "    b). Sapi & sejenisnya"],
    [89, "2) Charter"],
    [90, "3) Gayor"],
    [91, "4) Angkutan Pos"],
    [92, "5) Angkutan Khusus"],
    [93, "6) -nd. Sbg muatan"],
    [94, "   - Golongan IIa"],
    [95, "   - Golongan III"],
    [96, "   - Golongan IV"],
    [97, "   - Golongan V"],
    [98, "   - Golongan VIa"],
    [99, "   - Golongan VIb"],
    [100, "7) Lain-lain"],
  ].forEach(([row, lbl]) => {
    ["B", "C", "D", "F", "I", "J"].forEach((col) => b(ws.getCell(`${col}${row}`)));
    ws.getCell(`C${row}`).value = lbl;
    const ec = ws.getCell(`E${row}`);
    const eRef = rencE(row);
    if (eRef) {
      ec.value = eRef;
      ec.numFmt = "#,##0";
    }
    b(ec);
    try { ws.mergeCells(`G${row}:H${row}`); } catch (e2) {}
    b(ws.getCell(`G${row}`));
    try { ws.mergeCells(`K${row}:L${row}`); } catch (e2) {}
    b(ws.getCell(`K${row}`));
  });

  // Remove borders between NO and JENIS TI-T columns
  const ranges = [
    [20, 23],
    [25, 48],
    [51, 75],
    [77, 79],
    [81, 84],
    [86, 100],
  ];

  ranges.forEach(([start, end]) => {
    for (let r = start; r <= end; r++) {
      for (let c = 2; c <= 3; c++) {
        ws.getCell(r, c).border = {};
      }
      ws.getCell(r, 2).border = { right: { style: "thin" } };
    }
  });

  b(ws.getCell("B101"));
  sc("C101", "Jumlah (4.e)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D101"));
  const e101val = rencE(101);
  if (e101val) ws.getCell("E101").value = e101val;
  ws.getCell("E101").numFmt = "#,##0";
  b(ws.getCell("E101"));
  b(ws.getCell("F101"));
  try { ws.mergeCells("G101:H101"); } catch (e2) {}
  b(ws.getCell("G101"));
  b(ws.getCell("I101"));
  b(ws.getCell("J101"));
  try { ws.mergeCells("K101:L101"); } catch (e2) {}
  b(ws.getCell("K101"));

  // Jumlah Total (row 102)
  mc("B102:C102", "JUMLAH TOTAL (4.a s/d 4.e)", {
    bold: true,
    border: "medium",
    align: { horizontal: "center", vertical: "middle" },
  });
  b(ws.getCell("D102"));
  b(ws.getCell("E102"));
  b(ws.getCell("F102"));
  try { ws.mergeCells("G102:H102"); } catch (e2) {}
  b(ws.getCell("G102"));
  ws.getCell("I102").value = { formula: "I48+I75" };
  ws.getCell("I102").numFmt = "#,##0";
  b(ws.getCell("I102"));
  ws.getCell("J102").value = { formula: "J48+J75" };
  ws.getCell("J102").numFmt = "#,##0";
  b(ws.getCell("J102"));
  try { ws.mergeCells("K102:L102"); } catch (e2) {}
  ws.getCell("K102").value = { formula: 'IF(I102=0,"",J102/I102)' };
  ws.getCell("K102").numFmt = "0.00%";
  b(ws.getCell("K102"));

  // Reduksi (row 103)
  sc("B103", "5", {
    border: "thin",
    align: { horizontal: "center", vertical: "middle" },
  });
  sc("C103", "Reduksi Pend.Penyebergn", { bold: true });
  ws.getCell("D103").value = 0;
  ws.getCell("D103").numFmt = "#,##0";
  b(ws.getCell("D103"));
  ws.getCell("E103").value = { formula: "E48+E75" };
  ws.getCell("E103").numFmt = "#,##0";
  b(ws.getCell("E103"));
  ws.getCell("F103").value = { formula: "F48+F75" };
  ws.getCell("F103").numFmt = "#,##0";
  b(ws.getCell("F103"));
  try { ws.mergeCells("G103:H103"); } catch (e2) {}
  ws.getCell("G103").value = { formula: 'IF(E103=0,"",F103/E103)' };
  ws.getCell("G103").numFmt = "0.00%";
  b(ws.getCell("G103"));
  ws.getCell("I103").value = { formula: "E103*D103" };
  ws.getCell("I103").numFmt = "#,##0";
  b(ws.getCell("I103"));
  ws.getCell("J103").value = { formula: "F103*D103" };
  ws.getCell("J103").numFmt = "#,##0";
  b(ws.getCell("J103"));
  try { ws.mergeCells("K103:L103"); } catch (e2) {}
  ws.getCell("K103").value = { formula: 'IF(I103=0,"",J103/I103)' };
  ws.getCell("K103").numFmt = "0.00%";
  b(ws.getCell("K103"));

  // Total Akhir (row 104)
  b(ws.getCell("B104"));
  sc("C104", "Jumlah (3-4)", {
    bold: true,
    border: "thin",
    align: { horizontal: "right", vertical: "middle" },
  });
  b(ws.getCell("D104"));
  b(ws.getCell("E104"));
  b(ws.getCell("F104"));
  try { ws.mergeCells("G104:H104"); } catch (e2) {}
  b(ws.getCell("G104"));
  ws.getCell("I104").value = { formula: "I102-I103" };
  ws.getCell("I104").numFmt = "#,##0";
  b(ws.getCell("I104"));
  ws.getCell("J104").value = { formula: "J102-J103" };
  ws.getCell("J104").numFmt = "#,##0";
  b(ws.getCell("J104"));
  try { ws.mergeCells("K104:L104"); } catch (e2) {}
  ws.getCell("K104").value = { formula: 'IF(I104=0,"",J104/I104)' };
  ws.getCell("K104").numFmt = "0.00%";
  b(ws.getCell("K104"));

  // ========== HIDE BARIS DISKON YANG TIDAK ADA DATANYA ==========
  const barisDiskonPenumpang = [39, 41];
  const barisDiskonKendaraan = [];
  for (let rowNum = 52; rowNum <= 74; rowNum += 2) {
    barisDiskonKendaraan.push(rowNum);
  }
  const semuaBarisDiskon = [...barisDiskonPenumpang, ...barisDiskonKendaraan];
  
  if (isGabungan1) {
    // Untuk gabungan, hide berdasarkan data detail
    semuaBarisDiskon.forEach((rowNum) => {
      let adaData = false;
      
      if (rowNum === 39) {
        adaData = pnpDetail.dewasa.diskon > 0;
      } else if (rowNum === 41) {
        adaData = pnpDetail.anak.diskon > 0;
      } else {
        const golIndex = Math.floor((rowNum - 52) / 2);
        if (golIndex >= 0 && golIndex < GOL_ROW_MAP.length) {
          const g = GOL_ROW_MAP[golIndex];
          const key = getKndKey(g);
          const detail = kndDetail[key];
          adaData = detail && detail.diskon > 0;
        }
      }
      
      if (!adaData) {
        ws.getRow(rowNum).hidden = true;
      }
    });
  } else {
    // Untuk sheet per arah, hide berdasarkan nilai cell
    semuaBarisDiskon.forEach((rowNum) => {
      const cellValue = ws.getCell(`F${rowNum}`).value;
      if (!cellValue || cellValue === 0 || cellValue === "" || cellValue === null) {
        ws.getRow(rowNum).hidden = true;
      }
    });
  }

  // Border luar tabel
  bOuter(20, 2, 108, 12, "medium");

  // Gray fill untuk kolom rencana
  for (let r = 20; r <= 104; r++) {
    ws.getCell(`E${r}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
    ws.getCell(`I${r}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9D9D9" },
    };
  }

  // Bold semua baris Jumlah
  [48, 75, 79, 84, 101, 102, 103, 104].forEach((r) => {
    ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"].forEach((col) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.font = { ...cell.font, name: "Calibri", size: 11, bold: true };
    });
  });

  // ========== KETERANGAN (Rows 105-108) ==========
  mc("B105:K105", "-KETERANGAN", { bold: true });
  sc("B106", "1. Docking");
  sc("D106", ":");
  sc("B107", "2. Rusak");
  sc("D107", ":");
  sc("B108", "3. Lain - lain", { vertical: "middle" });
  sc("D108", ":");
  try { ws.mergeCells("D108:K108"); } catch (e2) {}
  ws.getCell("D108").alignment = {
    vertical: "middle",
    horizontal: "left",
    wrapText: true,
  };
  ws.getRow(108).height = 30;

  // ========== TANDA TANGAN (Rows 110-118) ==========
  const tglSur = new Date(tanggalSampai);
  tglSur.setDate(tglSur.getDate() + 1);
  const bulanNama2 = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
  ];
  const ttdStr = `${pelabuhanAsal}, ${String(tglSur.getDate()).padStart(2, "0")} ${bulanNama2[tglSur.getMonth()]} ${tglSur.getFullYear()}`;
  mc("I110:K110", ttdStr, { align: { horizontal: "center", vertical: "middle" } });
  mc("C111:E111", "Mengetahui :", { align: { horizontal: "center" } });
  mc("C112:E112", "GENERAL MANAGER", { bold: true, align: { horizontal: "center" } });
  mc("I112:K112", "MANAGER USAHA", { bold: true, align: { horizontal: "center" } });
  mc("C117:E117", suratDokumen.general_manager, { bold: true, align: { horizontal: "center" } });
  ws.getCell("C117").font = { ...ws.getCell("C117").font, underline: true };
  mc("I117:K117", suratDokumen.manager_usaha, { bold: true, align: { horizontal: "center" } });
  ws.getCell("I117").font = { ...ws.getCell("I117").font, underline: true };
  mc("C118:E118", "");
  mc("I118:K118", "");

  // ========== PAGE SETUP ==========
  ws.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: "17:19",
  };
  ws.views = [{ showGridLines: true }];
}

module.exports = { generateSheetPerKapal };
