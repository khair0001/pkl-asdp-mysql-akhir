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

async function generateSheetGabunganKapalSemuaLintasan(
  workbook,
  namaKapal,
  ruteList,
  sheetNamesLBRKapal,
  suratDokumen,
  tanggalDari,
  tanggalSampai
) {
  const sheetName = `GAB ${namaKapal}`;
  let finalSheetName = sheetName;
  if (finalSheetName.length > 31) {
    finalSheetName = finalSheetName.substring(0, 31);
  }
  
  // Ensure unique sheet name
  let counter = 1;
  while (workbook.getWorksheet(finalSheetName)) {
    const suffix = `_${counter}`;
    const maxLength = 31 - suffix.length;
    finalSheetName = sheetName.substring(0, maxLength) + suffix;
    counter++;
  }
  
  console.log(`Generating sheet: ${finalSheetName}`);

  const ws = workbook.addWorksheet(finalSheetName);
  ws.properties.defaultFont = { name: "Calibri", size: 11 };
  ws.properties.tabColor = { argb: "FF00B050" }; // Warna ungu untuk gabungan kapal semua lintasan

  const bulanNama = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const bulan = bulanNama[tanggalDari.getMonth()];
  const tahun = tanggalDari.getFullYear();

  const namaKapalDisplay = `KMP. ${namaKapal}`;
  const lintasan = `GABUNGAN SEMUA LINTASAN`;

  const b = (cell, s = "thin") => {
    cell.border = { top: { style: s }, left: { style: s }, bottom: { style: s }, right: { style: s } };
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
    if (opts.bold) c.font = { name: "Calibri", size: opts.size || 11, bold: true };
    if (opts.align) c.alignment = opts.align;
    if (opts.border) b(c, opts.border);
    if (opts.numFmt) c.numFmt = opts.numFmt;
    return c;
  };

  const mc = (range, val, opts = {}) => {
    try { ws.mergeCells(range); } catch (e) {}
    return sc(range.split(":")[0], val, opts);
  };

  const sumF = (row) => sheetNamesLBRKapal.map(n => `'${n}'!F${row}`).join('+');
  const sumE = (row) => sheetNamesLBRKapal.map(n => `'${n}'!E${row}`).join('+');
  const sumI = (row) => sheetNamesLBRKapal.map(n => `'${n}'!I${row}`).join('+');
  const sumJ = (row) => sheetNamesLBRKapal.map(n => `'${n}'!J${row}`).join('+');

  [4, 6, 25, 15, 12, 12, 7, 7, 18, 18, 10, 10].forEach((w, i) => (ws.getColumn(i + 1).width = w));

  bOuter(2, 2, 5, 3, "thin");

  try {
    const logoPath = path.join(__dirname, "../../public/images/logo_asdp.png");
    const logoId = workbook.addImage({ filename: logoPath, extension: "png" });
    ws.addImage(logoId, { tl: { col: 2.5, row: 1.2 }, ext: { width: 70, height: 70 }, editAs: "absolute" });
  } catch (e) {
    console.warn("Failed to add logo:", e.message);
  }

  mc("D3:I4", "LAPORAN PRODUKSI\nDAN PENDAPATAN KAPAL BULANAN ", {
    bold: true, size: 14, align: { horizontal: "center", vertical: "middle", wrapText: true }
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
  sc("D7", `: GABUNGAN SEMUA LINTASAN`, { bold: true, size: 14 });
  sc("B8", "USAHA", { bold: true, size: 14 });
  sc("D8", ": BISNIS PENYEBERANGAN", { bold: true, size: 14 });
  sc("B9", "BULAN", { bold: true, size: 14 });
  sc("D9", `: ${bulan.toUpperCase()} ${tahun}`, { bold: true, size: 14 });
  sc("J9", "BULANAN", { name: "Times New Roman", bold: true, border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("B10", "LINTASAN / JARAK", { bold: true, size: 14 });
  sc("D10", `: ${lintasan}`, { bold: true, size: 14 });
  sc("B11", "NAMA KAPAL / GRT", { bold: true, size: 14 });
  sc("D11", `: ${namaKapalDisplay.toUpperCase()}`, { bold: true, size: 14 });
  sc("B12", "KAPASITAS ANGKUT", { bold: true, size: 14 });
  sc("D12", `: a. Penumpang:  org, -Kendaraan:  unit (campuran)`, { bold: true, size: 11 });

  // Kategori header (rows 14-15)
  ["D", "E", "F", "G", "H", "I", "J", "K", "L"].forEach((col, i) => {
    const labels = ["Eks", "Bis.I", "Bis.II", "Eko.", "Gol IV", "Gol V", "Gol VI", "Gol VII", "Gol VIII"];
    const cell14 = sc(`${col}14`, labels[i], { border: "thin", align: { horizontal: "center", vertical: "middle" } });
    cell14.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "ffccffff" } };
    cell14.font = { name: "Calibri", size: 11 };
    sc(`${col}15`, "", { border: "thin" });
  });

  const blueHeaderFill = { type: "pattern", pattern: "solid", fgColor: { argb: "ffccffff" } };
  const blueHeaderFont = { name: "Calibri", size: 11, bold: true };
  const yellowFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };

  // Yellow fills
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

  // Table headers (rows 17-19)
  const h17b18 = mc("B17:B18", "NO", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h17b18.fill = blueHeaderFill;
  h17b18.font = blueHeaderFont;

  const h17c18 = mc("C17:C18", "JENIS TI-T", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h17c18.fill = blueHeaderFill;
  h17c18.font = blueHeaderFont;

  const h17d18 = mc("D17:D18", "TARIF (Rp.)", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h17d18.fill = blueHeaderFill;
  h17d18.font = blueHeaderFont;

  const h17prod = mc("E17:H17", "PRODUKSI", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h17prod.fill = blueHeaderFill;
  h17prod.font = blueHeaderFont;

  const h17pend = mc("I17:L17", "PENDAPATAN", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h17pend.fill = blueHeaderFill;
  h17pend.font = blueHeaderFont;

  const h18renc = sc("E18", "RENC.", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18renc.fill = blueHeaderFill;
  h18renc.font = blueHeaderFont;

  const h18real = sc("F18", "REAL", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18real.fill = blueHeaderFill;
  h18real.font = blueHeaderFont;

  const h18pct1 = mc("G18:H18", "(%)", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18pct1.fill = blueHeaderFill;
  h18pct1.font = blueHeaderFont;

  const h18rencana = sc("I18", "RENCANA", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18rencana.fill = blueHeaderFill;
  h18rencana.font = blueHeaderFont;

  const h18realisasi = sc("J18", "REALISASI", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18realisasi.fill = blueHeaderFill;
  h18realisasi.font = blueHeaderFont;

  const h18pct2 = mc("K18:L18", "(%)", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
  h18pct2.fill = blueHeaderFill;
  h18pct2.font = blueHeaderFont;

  [["B19", "1"], ["C19", "2"], ["D19", "3"], ["E19", "4"], ["F19", "5"], ["I19", "7"], ["J19", "8"]].forEach(([a, v]) => {
    const c19 = sc(a, v, { border: "medium", align: { horizontal: "center" } });
    c19.fill = blueHeaderFill;
    c19.font = blueHeaderFont;
  });
  const g19 = mc("G19:H19", "6", { border: "medium", align: { horizontal: "center" } });
  g19.fill = blueHeaderFill;
  g19.font = blueHeaderFont;
  const k19 = mc("K19:L19", "9", { border: "medium", align: { horizontal: "center" } });
  k19.fill = blueHeaderFill;
  k19.font = blueHeaderFont;

  // Helper functions for data rows
  const brsStd = (row, no, label, tarif, eVal, fVal, iFormula, jFormula) => {
    if (no !== null) sc(`B${row}`, no, { align: { horizontal: "center", vertical: "middle" } });
    else b(ws.getCell(`B${row}`));
    sc(`C${row}`, label, { align: { vertical: "middle" } });

    const dc = ws.getCell(`D${row}`);
    if (tarif || tarif === 0) {
      dc.value = Math.round(tarif);
      dc.numFmt = "#,##0";
      dc.alignment = { horizontal: "right", vertical: "middle" };
    }
    b(dc);

    const ec = ws.getCell(`E${row}`);
    if (eVal !== "" && eVal !== null && eVal !== undefined) ec.value = eVal;
    ec.numFmt = "#,##0";
    ec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    b(ec);

    const fc = ws.getCell(`F${row}`);
    if (fVal !== null && fVal !== undefined && fVal !== "") {
      fc.value = fVal;
      fc.numFmt = "#,##0";
    }
    b(fc);

    try { ws.mergeCells(`G${row}:H${row}`); } catch (e2) {}
    const gc = ws.getCell(`G${row}`);
    gc.value = { formula: `IF(OR(E${row}=0,F${row}=0),"",F${row}/E${row})` };
    gc.numFmt = "0.00%";
    b(gc);

    const ic = ws.getCell(`I${row}`);
    if (iFormula) {
      ic.value = { formula: iFormula };
      ic.numFmt = "#,##0";
    }
    ic.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    b(ic);

    const jc = ws.getCell(`J${row}`);
    if (jFormula) {
      jc.value = { formula: jFormula };
      jc.numFmt = "#,##0";
    }
    b(jc);

    try { ws.mergeCells(`K${row}:L${row}`); } catch (e2) {}
    const kc = ws.getCell(`K${row}`);
    if (iFormula && jFormula) {
      kc.value = { formula: `IF(I${row}=0,"",J${row}/I${row})` };
      kc.numFmt = "0.00%";
    }
    b(kc);
  };

  const brsHdr = (row, no, label, eRef) => {
    if (no) sc(`B${row}`, no, { border: "thin", align: { bold: true, horizontal: "center", vertical: "middle" } });
    else b(ws.getCell(`B${row}`));
    sc(`C${row}`, label, { bold: true, border: "thin", align: { vertical: "middle" } });
    b(ws.getCell(`D${row}`));
    const ec = ws.getCell(`E${row}`);
    if (eRef) {
      ec.value = eRef;
      ec.numFmt = "#,##0";
    }
    ec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    b(ec);
    b(ws.getCell(`F${row}`));
    try { ws.mergeCells(`G${row}:H${row}`); } catch (e2) {}
    b(ws.getCell(`G${row}`));
    ws.getCell(`I${row}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    b(ws.getCell(`I${row}`));
    b(ws.getCell(`J${row}`));
    try { ws.mergeCells(`K${row}:L${row}`); } catch (e2) {}
    b(ws.getCell(`K${row}`));
  };

  // Data rows (20-104)
  // Baris 20-22: Tanpa bold
  sc("B20", "1", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
  sc("C20", "HARI SIAP OPERASI", { border: "thin", align: { vertical: "middle" } });
  b(ws.getCell("D20"));
  const e20 = ws.getCell("E20");
  e20.value = ""; // Kosongkan E20
  e20.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e20);
  b(ws.getCell("F20")); // F20 juga dikosongkan
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
  e21.value = { formula: sumE(21) };
  e21.numFmt = "#,##0";
  e21.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e21);
  ws.getCell("F21").value = { formula: sumF(21) };
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
  e22.value = { formula: sumE(22) };
  e22.numFmt = "#,##0";
  e22.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  b(e22);
  ws.getCell("F22").value = { formula: sumF(22) };
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
  brsHdr(24, "a.", "PENUMPANG", { formula: sumE(24) });

  brsStd(25, "", "1) Eksekutif", null, null, null, null, null);
  brsStd(26, "", "- Dewasa", null, null, null, null, null);
  brsStd(27, "", "- Anak", null, null, null, null, null);
  brsStd(28, "", "- Bayi", null, null, null, null, null);

  brsStd(29, "", "2) Bisnis I", null, null, null, null, null);
  brsStd(30, "", "- Dewasa", null, null, null, null, null);
  brsStd(31, "", "- Anak", null, null, null, null, null);
  brsStd(32, "", "- Bayi", null, null, null, null, null);

  brsStd(33, "", "3) Bisnis II", null, null, null, null, null);
  brsStd(34, "", "- Dewasa", null, null, null, null, null);
  brsStd(35, "", "- Anak", null, null, null, null, null);
  brsStd(36, "", "- Bayi", null, null, null, null, null);

  brsStd(37, "", "4) Ekonomi", null, { formula: sumE(37) }, null, null, null);
  
  // Tarif penumpang - kosongkan karena tiap rute beda tarif
  brsStd(38, "", "- Dewasa", null, { formula: sumE(38) }, { formula: sumF(38) }, sumI(38), sumJ(38));
  brsStd(39, "", "- Dewasa (Diskon)", null, { formula: sumE(39) }, { formula: sumF(39) }, sumI(39), sumJ(39));
  brsStd(40, "", "- Anak", null, { formula: sumE(40) }, { formula: sumF(40) }, sumI(40), sumJ(40));
  brsStd(41, "", "- Anak (Diskon)", null, { formula: sumE(41) }, { formula: sumF(41) }, sumI(41), sumJ(41));

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

  brsStd(43, "", "5) Suplesi", null, null, null, null, null);
  brsStd(44, "", "- Eksekutif", null, null, null, null, null);
  brsStd(45, "", "- Bisnis I", null, null, null, null, null);
  brsStd(46, "", "- Bisnis II Dewasa", null, null, null, null, null);
  brsStd(47, "", "- Bisnis II Anak", null, null, null, null, null);

  // Jumlah penumpang (row 48)
  sc("C48", "Jumlah (4.a.1 s/d 4.a.5)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
  b(ws.getCell("D48"));
  ws.getCell("E48").value = { formula: "SUM(E38:E47)" };
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
  brsHdr(50, "b.", "KENDARAAN", { formula: sumE(50) });

  const golLabel = [
    "- Golongan I", "- Golongan II", "- Golongan III", "- Golongan IV", "- Golongan IV Pick Up",
    "- Golongan V Bus", "- Golongan V Truk", "- Golongan VI Bus", "- Golongan VI Truk",
    "- Golongan VII", "- Golongan VIII", "- Golongan IX"
  ];

  let currentRow = 51;
  GOL_ROW_MAP.forEach((g, i) => {
    // Tarif kendaraan - kosongkan karena tiap rute beda tarif
    brsStd(currentRow, "", golLabel[i], null, { formula: sumE(currentRow) }, { formula: sumF(currentRow) }, sumI(currentRow), sumJ(currentRow));
    currentRow++;
    
    // Tarif diskon
    brsStd(currentRow, "", golLabel[i] + " (Diskon)", null, { formula: sumE(currentRow) }, { formula: sumF(currentRow) }, sumI(currentRow), sumJ(currentRow));
    currentRow++;
  });

  // Jumlah kendaraan (row 75)
  sc("C75", "Jumlah (4.b)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
  b(ws.getCell("D75"));
  ws.getCell("E75").value = { formula: "SUM(E51:E74)" };
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
  brsStd(77, "", "- Penumpang (%)", null, null, null, null, null);
  brsStd(78, "", "- Kendaraan (%)", null, null, null, null, null);
  sc("C79", "Jumlah (4.c)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
  b(ws.getCell("D79"));
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
  brsStd(81, "", "1) Ton/M Kubik", null, null, null, null, null);
  brsStd(82, "", "2) Brg curah/M Kubik", null, null, null, null, null);
  brsStd(83, "", "3) Brg tentengan/kg", null, null, null, null, null);
  sc("C84", "Jumlah (4.d)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
  b(ws.getCell("D84"));
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
    [86, "1) Hewan"], [87, "    a). Kambing & sejenisnya"], [88, "    b). Sapi & sejenisnya"],
    [89, "2) Charter"], [90, "3) Gayor"], [91, "4) Angkutan Pos"], [92, "5) Angkutan Khusus"],
    [93, "6) -nd. Sbg muatan"], [94, "    - Golongan IIa"], [95, "    - Golongan III"],
    [96, "   - Golongan IV"], [97, "    - Golongan V"], [98, "    - Golongan VIa"],
    [99, "   - Golongan VIb"], [100, "7) Lain-lain"]
  ].forEach(([row, lbl]) => {
    ["B", "C", "D", "F", "I", "J"].forEach((col) => b(ws.getCell(`${col}${row}`)));
    ws.getCell(`C${row}`).value = lbl;
    const ec = ws.getCell(`E${row}`);
    ec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    b(ec);
    try { ws.mergeCells(`G${row}:H${row}`); } catch (e2) {}
    b(ws.getCell(`G${row}`));
    try { ws.mergeCells(`K${row}:L${row}`); } catch (e2) {}
    b(ws.getCell(`K${row}`));
  });

  // Remove borders between NO and JENIS TI-T columns
  const ranges = [[20, 23], [25, 48], [51, 75], [77, 79], [81, 84], [86, 100]];
  ranges.forEach(([start, end]) => {
    for (let r = start; r <= end; r++) {
      for (let c = 2; c <= 3; c++) ws.getCell(r, c).border = {};
      ws.getCell(r, 2).border = { right: { style: "thin" } };
    }
  });

  b(ws.getCell("B101"));
  sc("C101", "Jumlah (4.e)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
  b(ws.getCell("D101"));
  b(ws.getCell("E101"));
  b(ws.getCell("F101"));
  try { ws.mergeCells("G101:H101"); } catch (e2) {}
  b(ws.getCell("G101"));
  b(ws.getCell("I101"));
  b(ws.getCell("J101"));
  try { ws.mergeCells("K101:L101"); } catch (e2) {}
  b(ws.getCell("K101"));

  // Jumlah Total (row 102)
  mc("B102:C102", "JUMLAH TOTAL (4.a s/d 4.e)", { bold: true, border: "medium", align: { horizontal: "center", vertical: "middle" } });
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
  sc("B103", "5", { border: "thin", align: { horizontal: "center", vertical: "middle" } });
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
  ws.getCell("I103").value = { formula: sumI(103) };
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
  sc("C104", "Jumlah (3-4)", { bold: true, border: "thin", align: { horizontal: "right", vertical: "middle" } });
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

  // Border luar tabel
  bOuter(20, 2, 108, 12, "medium");

  // Gray fill untuk kolom rencana
  for (let r = 20; r <= 104; r++) {
    ws.getCell(`E${r}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
    ws.getCell(`I${r}`).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  }

  // Bold semua baris Jumlah
  [48, 75, 79, 84, 101, 102, 103, 104].forEach((r) => {
    ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"].forEach((col) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.font = { ...cell.font, name: "Calibri", size: 11, bold: true };
    });
  });

  // Keterangan (rows 105-108)
  mc("B105:K105", "-KETERANGAN", { bold: true });
  sc("B106", "1. Docking");
  sc("D106", ":");
  sc("B107", "2. Rusak");
  sc("D107", ":");
  sc("B108", "3. Lain - lain", { vertical: "middle" });
  sc("D108", ":");
  try { ws.mergeCells("D108:K108"); } catch (e2) {}
  ws.getCell("D108").alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  ws.getRow(108).height = 30;

  // Tanda tangan (rows 110-118)
  const tglSur = new Date(tanggalSampai);
  tglSur.setDate(tglSur.getDate() + 1);
  const bulanNama2 = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const ttdStr = `Lembar, ${String(tglSur.getDate()).padStart(2, "0")} ${bulanNama2[tglSur.getMonth()]} ${tglSur.getFullYear()}`;
  mc("I110:K110", ttdStr, { bold: true, align: { horizontal: "center", vertical: "middle" } });
  mc("C111:E111", "Mengetahui :", { align: { horizontal: "center" } });
  mc("C112:E112", "GENERAL MANAGER", { bold: true, align: { horizontal: "center" } });
  mc("I112:K112", "MANAGER USAHA", { bold: true, align: { horizontal: "center" } });
  mc("C117:E117", suratDokumen.general_manager, { bold: true, align: { horizontal: "center" } });
  ws.getCell("C117").font = { ...ws.getCell("C117").font, underline: true };
  mc("I117:K117", suratDokumen.manager_usaha, { bold: true, align: { horizontal: "center" } });
  ws.getCell("I117").font = { ...ws.getCell("I117").font, underline: true };
  mc("C118:E118", "");
  mc("I118:K118", "");

  ws.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "17:19" };
  ws.views = [{ showGridLines: true }];

  return finalSheetName;
}

module.exports = { generateSheetGabunganKapalSemuaLintasan };
