const { setBorder, mergeAndStyle } = require('./excelHelpers');

async function generateSheetKinerjaKapal(workbook, dataPerRute, bulan, tahun, suratDokumen) {
  const worksheet = workbook.addWorksheet("Kinerja KAPAL");
  worksheet.properties.defaultFont = { name: "Calibri", size: 11 };
  worksheet.properties.tabColor = { argb: "FF00B050" };

  const bulanTahun = `${bulan}-${tahun.toString().slice(-2)}`;

  // Header
  worksheet.getCell("A1").value = "LAPORAN KINERJA OPERASI KAPAL";
  worksheet.getCell("A1").font = { name: "Calibri", size: 16, bold: true };
  
  // Ambil pelabuhan asal dari rute pertama (asumsi semua rute dari cabang yang sama)
  const pelabuhanAsal = dataPerRute[0]?.pelabuhanAsal || "LEMBAR";
  worksheet.getCell("A2").value = `CABANG ${pelabuhanAsal.toUpperCase()}`;
  worksheet.getCell("A2").font = { name: "Calibri", size: 16, bold: true };
  
  worksheet.getCell("A3").value = `BULAN ${bulan.toUpperCase()} ${tahun}`;
  worksheet.getCell("A3").font = { name: "Calibri", size: 16, bold: true };

  // Set row heights
  for (let i = 5; i <= 15; i++) {
    worksheet.getRow(i).height = 20;
  }

  // Header tabel
  mergeAndStyle(worksheet, "A5:C5", "BULAN", {
    font: { name: "Calibri", size: 11, bold: true },
    alignment: { horizontal: "right", vertical: "middle" },
    border: "thin",
  });
  
  mergeAndStyle(worksheet, "D5:AH5", bulanTahun, {
    font: { name: "Calibri", size: 11, bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });
  
  mergeAndStyle(worksheet, "AI5:AJ5", "TRIP", {
    font: { name: "Calibri", size: 11, bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });

  mergeAndStyle(worksheet, "A6:C6", "NAMA KAPAL", {
    font: { name: "Calibri", size: 11, bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: "thin",
  });

  // Tanggal 1-31
  for (let i = 1; i <= 31; i++) {
    const colIndex = 3 + i;
    const cell = worksheet.getCell(6, colIndex);
    cell.value = i;
    cell.font = { name: "Calibri", size: 11, bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    setBorder(cell, "thin");
    worksheet.getColumn(colIndex).width = 10;
  }

  worksheet.getCell("AI6").value = "RENC";
  worksheet.getCell("AI6").font = { bold: true };
  worksheet.getCell("AI6").alignment = { horizontal: "center", vertical: "middle" };
  setBorder(worksheet.getCell("AI6"), "thin");

  worksheet.getCell("AJ6").value = "REAL";
  worksheet.getCell("AJ6").font = { bold: true };
  worksheet.getCell("AJ6").alignment = { horizontal: "center", vertical: "middle" };
  setBorder(worksheet.getCell("AJ6"), "thin");

  let currentRow = 7;

  // A. Kapal Komersil
  mergeAndStyle(worksheet, `A${currentRow}:A${currentRow}`, "A.", {
    font: { bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });
  mergeAndStyle(worksheet, `B${currentRow}:C${currentRow}`, "Kapal Komersil", {
    font: { bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: "thin",
  });
  setBorder(worksheet.getCell("AI7"), "thin");
  setBorder(worksheet.getCell("AJ7"), "thin");
  currentRow++;

  // Loop setiap rute
  for (const ruteData of dataPerRute) {
    const { rute, kapalAsdpTemplate, tripPerKapalPerTanggal } = ruteData;
    
    // Header lintas
    mergeAndStyle(worksheet, `A${currentRow}:C${currentRow}`, 
      `LINTAS ${rute.nama_rute.toUpperCase()}`, {
      font: { bold: true },
      alignment: { horizontal: "center", vertical: "middle" },
      border: "thin",
    });
    setBorder(worksheet.getCell(`AI${currentRow}`), "thin");
    setBorder(worksheet.getCell(`AJ${currentRow}`), "thin");
    currentRow++;

    // Data kapal per rute
    kapalAsdpTemplate.forEach((kapal, idx) => {
      let namaKapal = kapal.nama_kapal.toUpperCase();
      if (namaKapal.startsWith("KMP.")) {
        namaKapal = namaKapal.substring(4).trim();
      }

      const row = worksheet.getRow(currentRow);
      setBorder(row.getCell(1), "thin");

      row.getCell(2).value = idx + 1;
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      setBorder(row.getCell(2), "thin");

      row.getCell(3).value = namaKapal;
      row.getCell(3).alignment = { horizontal: "left", vertical: "middle" };
      setBorder(row.getCell(3), "thin");

      // Trip data per tanggal
      const tripData = tripPerKapalPerTanggal?.[namaKapal] || {};
      for (let i = 1; i <= 31; i++) {
        const colIndex = 3 + i;
        const cell = row.getCell(colIndex);
        cell.value = tripData[i] || 0;
        cell.font = { name: "Calibri", size: 8 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        setBorder(cell, "thin");

        // Background color
        if (tripData[i] && tripData[i] > 0) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "ff00ffff" },
          };
        } else {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" },
          };
        }
      }

      // RENCANA & REALISASI
      row.getCell(35).value = 0;
      row.getCell(35).font = { name: "Calibri", size: 12 };
      row.getCell(35).alignment = { horizontal: "center", vertical: "middle" };
      setBorder(row.getCell(35), "thin");

      row.getCell(36).value = { formula: `SUM(D${currentRow}:AH${currentRow})` };
      row.getCell(36).font = { name: "Calibri", size: 12 };
      row.getCell(36).alignment = { horizontal: "center", vertical: "middle" };
      setBorder(row.getCell(36), "thin");

      currentRow++;
    });
  }

  // B. Kapal Perintis
  mergeAndStyle(worksheet, `A${currentRow}:A${currentRow}`, "B.", {
    font: { bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });
  mergeAndStyle(worksheet, `B${currentRow}:C${currentRow}`, "Kapal Perintis", {
    font: { bold: true },
    alignment: { horizontal: "left", vertical: "middle" },
    border: "thin",
  });
  for (let i = 4; i <= 36; i++) {
    setBorder(worksheet.getCell(currentRow, i), "thin");
  }
  currentRow++;

  // Baris kosong
  for (let i = 1; i <= 36; i++) {
    setBorder(worksheet.getCell(currentRow, i), "thin");
  }
  currentRow++;

  // JUMLAH KAPAL - hitung unique kapal
  const uniqueKapal = new Set();
  dataPerRute.forEach(rd => {
    rd.kapalAsdpTemplate.forEach(k => {
      let namaKapal = k.nama_kapal.toUpperCase();
      if (namaKapal.startsWith("KMP.")) {
        namaKapal = namaKapal.substring(4).trim();
      }
      uniqueKapal.add(namaKapal);
    });
  });

  mergeAndStyle(worksheet, `A${currentRow}:C${currentRow}`, "JUMLAH KAPAL", {
    font: { name: "Calibri", size: 12, bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });
  mergeAndStyle(worksheet, `D${currentRow}:E${currentRow}`, uniqueKapal.size, {
    font: { name: "Calibri", size: 12, bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });
  mergeAndStyle(worksheet, `F${currentRow}:G${currentRow}`, "UNIT", {
    font: { name: "Calibri", size: 12, bold: true },
    alignment: { horizontal: "center", vertical: "middle" },
    border: "thin",
  });

  // Column widths
  worksheet.getColumn(1).width = 5;
  worksheet.getColumn(2).width = 5;
  worksheet.getColumn(3).width = 25;
  for (let i = 4; i <= 34; i++) {
    worksheet.getColumn(i).width = 4;
  }
  worksheet.getColumn(35).width = 8;
  worksheet.getColumn(36).width = 8;

  worksheet.views = [{ showGridLines: false }];
  worksheet.pageSetup = {
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  return worksheet;
}

module.exports = { generateSheetKinerjaKapal };
