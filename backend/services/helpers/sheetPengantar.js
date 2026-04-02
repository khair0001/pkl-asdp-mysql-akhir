const { getNamaPendekKapal } = require('./excelHelpers');

async function generateSheetPengantar(
  workbook,
  dataPerRute,
  suratDokumen,
  tanggalDari,
  tanggalSampai,
  pelabuhanAsalCabang
) {
  const worksheet = workbook.addWorksheet("Pengantar");
  worksheet.properties.defaultFont = { name: "Frutiger-Normal", size: 11 };
  worksheet.properties.tabColor = { argb: "FF00B050" };

  const setBorder = (cell, style = "thin") => {
    cell.border = {
      top: { style },
      left: { style },
      bottom: { style },
      right: { style },
    };
  };

  const bulanNama = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
  ];
  const bulan = bulanNama[tanggalDari.getMonth()];
  const tahun = tanggalDari.getFullYear();

  // Tanggal surat (1 hari setelah tanggal sampai)
  const tanggalSurat = new Date(tanggalSampai);
  tanggalSurat.setDate(tanggalSurat.getDate() + 1);
  const hariSurat = tanggalSurat.getDate();
  const bulanSurat = bulanNama[tanggalSurat.getMonth()];
  const tahunSurat = tanggalSurat.getFullYear();

  // HEADER SURAT
  worksheet.getCell("H29").value =
    `${pelabuhanAsalCabang}, ${String(hariSurat).padStart(2, "0")} ${bulanSurat} ${tahunSurat}`;
  worksheet.getCell("H29").alignment = {
    horizontal: "left",
    vertical: "top",
  };

  worksheet.getCell("B30").value = "Lampiran";
  worksheet.getCell("C30").value = ":";
  worksheet.getCell("D30").value = "1 (satu) berkas.";

  worksheet.getCell("B31").value = "Perihal";
  worksheet.getCell("C31").value = ":";
  worksheet.mergeCells("D31:F32");
  worksheet.getCell("D31").value =
    `Laporan Bulanan Produksi dan Pendapatan Kapal\nBulan ${bulan} ${tahun}`;
  worksheet.getCell("D31").alignment = { wrapText: true, vertical: "top" };

  worksheet.getCell("G33").value = "Yth.";
  worksheet.getCell("H32").value = "K e p a d a";
  worksheet.getCell("H33").value = "Direktur Komersil";
  worksheet.getCell("H33").font = { bold: true };
  worksheet.getCell("H34").value = "PT. ASDP Indonesia Ferry (Persero)";
  worksheet.getCell("H34").font = { bold: true };
  worksheet.getCell("H35").value = "di";
  worksheet.mergeCells("H36:J36");
  worksheet.getCell("H36").value = "J A K A R T A";
  worksheet.getCell("H36").font = { bold: true };
  worksheet.getCell("H36").alignment = {
    horizontal: "center",
    vertical: "top",
  };

  worksheet.getCell("C39").value = "1.";

  worksheet.mergeCells("D39:M39");
  worksheet.getCell("D39").value =
    `Bersama ini terlampir disampaikan laporan Produksi dan Pendapatan Kapal (Penyeberangan) Bulan ${bulan} ${tahun} sebagai berikut :`;
  worksheet.getCell("D39").alignment = {
    horizontal: "left",
    vertical: "middle",
    wrapText: true,
  };

  worksheet.getCell("D40").value = "a.";
  worksheet.mergeCells("E40:M40");
  worksheet.getCell("E40").value =
    `Laporan Produksi dan Pendapatan Kapal (Penyeberangan) Bulan ${bulan} sebagai berikut :`;
  worksheet.getCell("E40").alignment = {
    wrapText: true,
    vertical: "top",
  };
  worksheet.getCell("E40").font = {
    bold: true,
  };

  // HEADER TABEL
  worksheet.mergeCells("E42:E43");
  worksheet.getCell("E42").value = "No";
  worksheet.getCell("E42").font = { bold: true };
  worksheet.getCell("E42").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("E42"), "thin");

  worksheet.mergeCells("F42:F43");
  worksheet.getCell("F42").value = "Uraian";
  worksheet.getCell("F42").font = { bold: true };
  worksheet.getCell("F42").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("F42"), "thin");

  worksheet.mergeCells("G42:G43");
  worksheet.getCell("G42").value = `RKA ${tahun}`;
  worksheet.getCell("G42").font = { bold: true };
  worksheet.getCell("G42").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("G42"), "thin");

  worksheet.mergeCells("H42:H43");
  worksheet.getCell("H42").value = `REALISASI ${tahun}`;
  worksheet.getCell("H42").font = { bold: true };
  worksheet.getCell("H42").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("H42"), "thin");

  worksheet.mergeCells("I42:M42");
  worksheet.getCell("I42").value = "Deviasi";
  worksheet.getCell("I42").font = { bold: true };
  worksheet.getCell("I42").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("I42"), "thin");

  worksheet.mergeCells("I43:J43");
  worksheet.getCell("I43").value = "( 4 - 3 )";
  worksheet.getCell("I43").font = { bold: true };
  worksheet.getCell("I43").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("I43"), "thin");

  worksheet.mergeCells("K43:M43");
  worksheet.getCell("K43").value = "( 4 / 3 ) %";
  worksheet.getCell("K43").font = { bold: true };
  worksheet.getCell("K43").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("K43"), "thin");

  // Baris 44: Nomor kolom
  worksheet.getCell("E44").value = "1";
  worksheet.getCell("E44").font = { bold: true };
  worksheet.getCell("E44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("E44"), "thin");

  worksheet.getCell("F44").value = "2";
  worksheet.getCell("F44").font = { bold: true };
  worksheet.getCell("F44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("F44"), "thin");

  worksheet.getCell("G44").value = "3";
  worksheet.getCell("G44").font = { bold: true };
  worksheet.getCell("G44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("G44"), "thin");

  worksheet.getCell("H44").value = "4";
  worksheet.getCell("H44").font = { bold: true };
  worksheet.getCell("H44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("H44"), "thin");

  worksheet.mergeCells("I44:J44");
  worksheet.getCell("I44").value = "5";
  worksheet.getCell("I44").font = { bold: true };
  worksheet.getCell("I44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("I44"), "thin");

  worksheet.mergeCells("K44:M44");
  worksheet.getCell("K44").value = "6";
  worksheet.getCell("K44").font = { bold: true };
  worksheet.getCell("K44").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  setBorder(worksheet.getCell("K44"), "thin");

  // Header Produksi
  for (let col = 5; col <= 13; col++) {
    const cell = worksheet.getCell(45, col);
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    setBorder(cell, "thin");
  }
  worksheet.getCell("F45").value = "Produksi";
  worksheet.getCell("F45").font = { bold: true };
  worksheet.getCell("F45").alignment = {
    horizontal: "left",
    vertical: "middle",
  };

  // Fungsi helper untuk menambahkan data kapal
  const addKapalData = (startRow, huruf, namaKapal, namaRute, namaSheetKapalList) => {
    worksheet.getCell(`E${startRow}`).value = huruf;
    worksheet.getCell(`E${startRow}`).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    setBorder(worksheet.getCell(`E${startRow}`), "thin");

    worksheet.getCell(`F${startRow}`).value = `Kmp. ${namaKapal}`;
    worksheet.getCell(`F${startRow}`).font = { bold: true };
    worksheet.getCell(`F${startRow}`).alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    setBorder(worksheet.getCell(`F${startRow}`), "thin");

    ["G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
      setBorder(worksheet.getCell(`${col}${startRow}`), "thin");
    });

    const details = [
      { label: "- Hari Operasi", refRow: 21 },
      { label: "- Trip", refRow: 22 },
      { label: "- Penumpang", refRow: 48 },
      { label: "- Kendaraan", refRow: 75 },
    ];

    details.forEach((detail, idx) => {
      const row = startRow + 1 + idx;

      worksheet.getCell(`E${row}`).value = "";
      setBorder(worksheet.getCell(`E${row}`), "thin");

      worksheet.getCell(`F${row}`).value = detail.label;
      worksheet.getCell(`F${row}`).alignment = {
        horizontal: "left",
        vertical: "middle",
      };
      setBorder(worksheet.getCell(`F${row}`), "thin");

      // Gabungkan formula dari semua sheet untuk kolom E (RKA/Rencana)
      const formulaE = namaSheetKapalList.map(sheet => `'${sheet}'!E${detail.refRow}`).join('+');
      worksheet.getCell(`G${row}`).value = {
        formula: formulaE,
      };
      worksheet.getCell(`G${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`G${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`G${row}`), "thin");

      // Gabungkan formula dari semua sheet untuk kolom F (Realisasi)
      const formulaF = namaSheetKapalList.map(sheet => `'${sheet}'!F${detail.refRow}`).join('+');
      worksheet.getCell(`H${row}`).value = {
        formula: formulaF,
      };
      worksheet.getCell(`H${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`H${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`H${row}`), "thin");

      worksheet.mergeCells(`I${row}:J${row}`);
      worksheet.getCell(`I${row}`).value = { formula: `H${row}-G${row}` };
      worksheet.getCell(`I${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`I${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`I${row}`), "thin");

      worksheet.mergeCells(`K${row}:M${row}`);
      worksheet.getCell(`K${row}`).value = {
        formula: `IF(G${row}=0,"",H${row}/G${row})`,
      };
      worksheet.getCell(`K${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`K${row}`).numFmt = "0.00%";
      setBorder(worksheet.getCell(`K${row}`), "thin");
    });
  };

  // Fungsi helper untuk menambahkan data pendapatan kapal
  const addPendapatanKapal = (startRow, huruf, namaKapal, namaRute, namaSheetKapalList) => {
    worksheet.getCell(`E${startRow}`).value = huruf;
    worksheet.getCell(`E${startRow}`).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    setBorder(worksheet.getCell(`E${startRow}`), "thin");

    worksheet.getCell(`F${startRow}`).value = `Kmp. ${namaKapal}`;
    worksheet.getCell(`F${startRow}`).font = { bold: true };
    worksheet.getCell(`F${startRow}`).alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    setBorder(worksheet.getCell(`F${startRow}`), "thin");

    ["G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
      setBorder(worksheet.getCell(`${col}${startRow}`), "thin");
    });

    const details = [
      { label: "- Penumpang", refRow: 48 },
      { label: "- Kendaraan", refRow: 75 },
    ];

    details.forEach((detail, idx) => {
      const row = startRow + 1 + idx;

      worksheet.getCell(`E${row}`).value = "";
      setBorder(worksheet.getCell(`E${row}`), "thin");

      worksheet.getCell(`F${row}`).value = detail.label;
      worksheet.getCell(`F${row}`).alignment = {
        horizontal: "left",
        vertical: "middle",
      };
      setBorder(worksheet.getCell(`F${row}`), "thin");

      // Gabungkan formula dari semua sheet untuk kolom I (rencana pendapatan)
      const formulaI = namaSheetKapalList.map(sheet => `'${sheet}'!I${detail.refRow}`).join('+');
      worksheet.getCell(`G${row}`).value = {
        formula: formulaI,
      };
      worksheet.getCell(`G${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`G${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`G${row}`), "thin");

      // Gabungkan formula dari semua sheet untuk kolom J (realisasi pendapatan)
      const formulaJ = namaSheetKapalList.map(sheet => `'${sheet}'!J${detail.refRow}`).join('+');
      worksheet.getCell(`H${row}`).value = {
        formula: formulaJ,
      };
      worksheet.getCell(`H${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`H${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`H${row}`), "thin");

      worksheet.mergeCells(`I${row}:J${row}`);
      worksheet.getCell(`I${row}`).value = { formula: `H${row}-G${row}` };
      worksheet.getCell(`I${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`I${row}`).numFmt = "#,##0";
      setBorder(worksheet.getCell(`I${row}`), "thin");

      worksheet.mergeCells(`K${row}:M${row}`);
      worksheet.getCell(`K${row}`).value = {
        formula: `IF(G${row}=0,"",H${row}/G${row})`,
      };
      worksheet.getCell(`K${row}`).alignment = {
        horizontal: "right",
        vertical: "middle",
      };
      worksheet.getCell(`K${row}`).numFmt = "0.00%";
      setBorder(worksheet.getCell(`K${row}`), "thin");
    });

    // Baris Jumlah per kapal
    const jumlahRow = startRow + 3;
    worksheet.getCell(`E${jumlahRow}`).value = "";
    setBorder(worksheet.getCell(`E${jumlahRow}`), "thin");

    worksheet.getCell(`F${jumlahRow}`).value = `Jumlah-${huruf}`;
    worksheet.getCell(`F${jumlahRow}`).font = { bold: true };
    worksheet.getCell(`F${jumlahRow}`).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
    setBorder(worksheet.getCell(`F${jumlahRow}`), "thin");

    worksheet.getCell(`G${jumlahRow}`).value = {
      formula: `SUM(G${startRow + 1}:G${startRow + 2})`,
    };
    worksheet.getCell(`G${jumlahRow}`).font = { bold: true };
    worksheet.getCell(`G${jumlahRow}`).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
    worksheet.getCell(`G${jumlahRow}`).numFmt = "#,##0";
    setBorder(worksheet.getCell(`G${jumlahRow}`), "thin");

    worksheet.getCell(`H${jumlahRow}`).value = {
      formula: `SUM(H${startRow + 1}:H${startRow + 2})`,
    };
    worksheet.getCell(`H${jumlahRow}`).font = { bold: true };
    worksheet.getCell(`H${jumlahRow}`).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
    worksheet.getCell(`H${jumlahRow}`).numFmt = "#,##0";
    setBorder(worksheet.getCell(`H${jumlahRow}`), "thin");

    worksheet.mergeCells(`I${jumlahRow}:J${jumlahRow}`);
    worksheet.getCell(`I${jumlahRow}`).value = {
      formula: `H${jumlahRow}-G${jumlahRow}`,
    };
    worksheet.getCell(`I${jumlahRow}`).font = { bold: true };
    worksheet.getCell(`I${jumlahRow}`).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
    worksheet.getCell(`I${jumlahRow}`).numFmt = "#,##0";
    setBorder(worksheet.getCell(`I${jumlahRow}`), "thin");

    worksheet.mergeCells(`K${jumlahRow}:M${jumlahRow}`);
    worksheet.getCell(`K${jumlahRow}`).value = {
      formula: `IF(G${jumlahRow}=0,"",H${jumlahRow}/G${jumlahRow})`,
    };
    worksheet.getCell(`K${jumlahRow}`).font = { bold: true };
    worksheet.getCell(`K${jumlahRow}`).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
    worksheet.getCell(`K${jumlahRow}`).numFmt = "0.00%";
    setBorder(worksheet.getCell(`K${jumlahRow}`), "thin");
  };

  // Kumpulkan semua kombinasi kapal-rute, dikelompokkan per kapal
  const kapalDataMap = new Map();
  dataPerRute.forEach(rd => {
    rd.kapalAsdpTemplate.forEach(kapal => {
      const namaKapalPendek = getNamaPendekKapal(kapal.nama_kapal);
      const labelAsalShort = rd.pelabuhanAsal.substring(0, 3).toUpperCase();
      const labelTujuanShort = rd.pelabuhanTujuan.substring(0, 3).toUpperCase();
      const namaSheetGabungan = `${namaKapalPendek}-${labelAsalShort}-${labelTujuanShort}`.substring(0, 31);
      
      if (!kapalDataMap.has(namaKapalPendek)) {
        kapalDataMap.set(namaKapalPendek, {
          namaKapal: namaKapalPendek,
          rutes: [],
          sheets: []
        });
      }
      
      const kapalData = kapalDataMap.get(namaKapalPendek);
      kapalData.rutes.push(rd.rute.nama_rute);
      kapalData.sheets.push(namaSheetGabungan);
    });
  });
  
  const semuaKapalRute = Array.from(kapalDataMap.values());

  // Generate huruf untuk semua kombinasi kapal-rute
  const hurufKapal = semuaKapalRute.map((_, idx) => String.fromCharCode(65 + idx));

  // DATA PRODUKSI PER KAPAL PER RUTE
  let currentRow = 46;
  semuaKapalRute.forEach((kr, idx) => {
    const namaRuteGabungan = kr.rutes.join(", ");
    addKapalData(
      currentRow,
      hurufKapal[idx],
      kr.namaKapal,
      namaRuteGabungan,
      kr.sheets
    );
    currentRow += 5;
  });

  // Header Pendapatan
  const headerPendapatanRow = currentRow;
  for (let col = 5; col <= 13; col++) {
    const cell = worksheet.getCell(headerPendapatanRow, col);
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    setBorder(cell, "thin");
  }
  worksheet.getCell(`F${headerPendapatanRow}`).value = "Pendapatan";
  worksheet.getCell(`F${headerPendapatanRow}`).font = { bold: true };
  worksheet.getCell(`F${headerPendapatanRow}`).alignment = {
    horizontal: "left",
    vertical: "middle",
  };

  // DATA PENDAPATAN PER KAPAL PER RUTE
  currentRow = headerPendapatanRow + 1;
  semuaKapalRute.forEach((kr, idx) => {
    const namaRuteGabungan = kr.rutes.join(", ");
    addPendapatanKapal(
      currentRow,
      hurufKapal[idx],
      kr.namaKapal,
      namaRuteGabungan,
      kr.sheets
    );
    currentRow += 4;
  });

  // Baris Jumlah Total
  const jumlahTotalRow = currentRow;

  worksheet.getCell(`E${jumlahTotalRow}`).value = "";
  setBorder(worksheet.getCell(`E${jumlahTotalRow}`), "thin");

  // Generate formula jumlah dinamis (A+B+C+D+...)
  const hurufFormula = hurufKapal.join("+");
  worksheet.getCell(`F${jumlahTotalRow}`).value = `Jumlah (${hurufFormula})`;
  worksheet.getCell(`F${jumlahTotalRow}`).font = { bold: true };
  worksheet.getCell(`F${jumlahTotalRow}`).alignment = {
    horizontal: "left",
    vertical: "middle",
  };
  setBorder(worksheet.getCell(`F${jumlahTotalRow}`), "thin");

  // Generate formula SUM untuk semua kapal
  // Setiap kapal di bagian pendapatan memiliki 4 baris (header + 2 detail + jumlah)
  // Header pendapatan dimulai setelah semua data produksi
  const headerPendapatanRowCalc = 46 + semuaKapalRute.length * 5;
  const rowJumlahKapal = semuaKapalRute.map((_, idx) => {
    // Setiap kapal di pendapatan: startRow = headerPendapatanRow + 1 + (idx * 4)
    // Jumlah row = startRow + 3
    const startRowPendapatan = headerPendapatanRowCalc + 1 + (idx * 4);
    return startRowPendapatan + 3;
  });

  const g74Formula = rowJumlahKapal.map((r) => `G${r}`).join("+");
  const h74Formula = rowJumlahKapal.map((r) => `H${r}`).join("+");

  worksheet.getCell(`G${jumlahTotalRow}`).value = { formula: g74Formula };
  worksheet.getCell(`G${jumlahTotalRow}`).font = { bold: true };
  worksheet.getCell(`G${jumlahTotalRow}`).alignment = {
    horizontal: "right",
    vertical: "middle",
  };
  worksheet.getCell(`G${jumlahTotalRow}`).numFmt = "#,##0";
  setBorder(worksheet.getCell(`G${jumlahTotalRow}`), "thin");

  worksheet.getCell(`H${jumlahTotalRow}`).value = { formula: h74Formula };
  worksheet.getCell(`H${jumlahTotalRow}`).font = { bold: true };
  worksheet.getCell(`H${jumlahTotalRow}`).alignment = {
    horizontal: "right",
    vertical: "middle",
  };
  worksheet.getCell(`H${jumlahTotalRow}`).numFmt = "#,##0";
  setBorder(worksheet.getCell(`H${jumlahTotalRow}`), "thin");

  worksheet.mergeCells(`I${jumlahTotalRow}:J${jumlahTotalRow}`);
  worksheet.getCell(`I${jumlahTotalRow}`).value = {
    formula: `H${jumlahTotalRow}-G${jumlahTotalRow}`,
  };
  worksheet.getCell(`I${jumlahTotalRow}`).font = { bold: true };
  worksheet.getCell(`I${jumlahTotalRow}`).alignment = {
    horizontal: "right",
    vertical: "middle",
  };
  worksheet.getCell(`I${jumlahTotalRow}`).numFmt = "#,##0";
  setBorder(worksheet.getCell(`I${jumlahTotalRow}`), "thin");

  worksheet.mergeCells(`K${jumlahTotalRow}:M${jumlahTotalRow}`);
  worksheet.getCell(`K${jumlahTotalRow}`).value = {
    formula: `IF(G${jumlahTotalRow}=0,"",H${jumlahTotalRow}/G${jumlahTotalRow})`,
  };
  worksheet.getCell(`K${jumlahTotalRow}`).font = { bold: true };
  worksheet.getCell(`K${jumlahTotalRow}`).alignment = {
    horizontal: "right",
    vertical: "middle",
  };
  worksheet.getCell(`K${jumlahTotalRow}`).numFmt = "0.00%";
  setBorder(worksheet.getCell(`K${jumlahTotalRow}`), "thin");

  // Bagian penutup
  const row76 = jumlahTotalRow + 3;
  const row78 = jumlahTotalRow + 5;
  const row83 = jumlahTotalRow + 10;
  const row85 = jumlahTotalRow + 12;

  worksheet.getCell(`D${row76}`).value = "2.";
  worksheet.mergeCells(`E${row76}:M${row76}`);
  worksheet.getCell(`E${row76}`).value =
    "Demikian atas perhatian Direksi di ucapkan terima kasih.";
  worksheet.getCell(`E${row76}`).alignment = {
    wrapText: true,
    vertical: "top",
  };

  worksheet.mergeCells(`H${row78}:J${row78}`);
  worksheet.getCell(`H${row78}`).value = "GENERAL MANAGER";
  worksheet.getCell(`H${row78}`).font = { bold: true };
  worksheet.getCell(`H${row78}`).alignment = {
    horizontal: "center",
    vertical: "middle",
  };

  worksheet.mergeCells(`H${row83}:J${row83}`);
  worksheet.getCell(`H${row83}`).value = suratDokumen.general_manager;
  worksheet.getCell(`H${row83}`).font = { bold: true, underline: true };
  worksheet.getCell(`H${row83}`).alignment = {
    horizontal: "center",
    vertical: "middle",
  };

  worksheet.getCell(`B${row85}`).value = "Tembusan Yth :";
  worksheet.getCell(`B${row85 + 1}`).value = "1. Direktur Utama";
  worksheet.getCell(`B${row85 + 2}`).value = "2. Direktur Keuangan";
  worksheet.getCell(`B${row85 + 3}`).value = "3. Kepala SPI";

  // Column widths
  worksheet.getColumn(1).width = 3;
  worksheet.getColumn(2).width = 15;
  worksheet.getColumn(3).width = 5;
  worksheet.getColumn(4).width = 5;
  worksheet.getColumn(5).width = 5;
  worksheet.getColumn(6).width = 30;
  worksheet.getColumn(7).width = 15;
  worksheet.getColumn(8).width = 15;
  worksheet.getColumn(9).width = 0;
  worksheet.getColumn(10).width = 20;
  worksheet.getColumn(11).width = 0;
  worksheet.getColumn(12).width = 0;
  worksheet.getColumn(13).width = 25;

  worksheet.views = [{ showGridLines: true }];
  worksheet.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  return worksheet;
}

module.exports = { generateSheetPengantar };
