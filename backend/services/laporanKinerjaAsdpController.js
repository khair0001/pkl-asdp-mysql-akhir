const ProduksiModel = require("../models/Produksi");
const ProduksiKendaraanModel = require("../models/ProduksiKendaraan");
const ProduksiPenumpangModel = require("../models/ProduksiPenumpang");
const RuteModel = require("../models/Rute");
const TemplateKapalRuteModel = require("../models/TemplateKapalRute");
const TarifKendaraanModel = require("../models/TarifKendaraan");
const TarifPenumpangModel = require("../models/TarifPenumpang");
const SuratDokumenModel = require("../models/SuratDokumen");
const ExcelJS = require("exceljs");

class LaporanKinerjaAsdpController {
  static async exportKinerjaAsdp(req, res, next) {
    try {
      console.log("=== EXPORT KINERJA ASDP MULTI-RUTE START ===");
      console.log("Query params:", req.query);

      // Helper functions
      const getColumnLetter = (colNumber) => {
        let letter = "";
        while (colNumber > 0) {
          const remainder = (colNumber - 1) % 26;
          letter = String.fromCharCode(65 + remainder) + letter;
          colNumber = Math.floor((colNumber - 1) / 26);
        }
        return letter;
      };

      const setBorder = (cell, style = "thin") => {
        cell.border = {
          top: { style },
          left: { style },
          bottom: { style },
          right: { style },
        };
      };

      const mergeAndStyle = (worksheet, range, value, options = {}) => {
        worksheet.mergeCells(range);
        const cell = worksheet.getCell(range.split(":")[0]);
        cell.value = value;
        if (options.font) cell.font = options.font;
        if (options.alignment) cell.alignment = options.alignment;
        if (options.fill) cell.fill = options.fill;
        if (options.border) setBorder(cell, options.border);
        return cell;
      };

      // Validasi tanggal
      if (!req.query.tanggal_dari || !req.query.tanggal_sampai) {
        return res.status(400).json({
          error: "Pilih periode (tanggal dari dan tanggal sampai) terlebih dahulu.",
        });
      }

      // Validasi 1 bulan
      const tanggalDari = new Date(req.query.tanggal_dari);
      const tanggalSampai = new Date(req.query.tanggal_sampai);

      if (
        tanggalDari.getMonth() !== tanggalSampai.getMonth() ||
        tanggalDari.getFullYear() !== tanggalSampai.getFullYear()
      ) {
        return res.status(400).json({
          error: "Periode harus dalam 1 bulan yang sama untuk export Kinerja ASDP",
        });
      }

      // STEP 1: Query semua rute yang punya kapal ASDP
      const ruteList = await TemplateKapalRuteModel.getRutesWithAsdpKapal();
      
      if (ruteList.length === 0) {
        return res.status(404).json({
          error: "Tidak ada rute dengan kapal ASDP ditemukan",
        });
      }

      console.log(`Ditemukan ${ruteList.length} rute dengan kapal ASDP:`, ruteList.map(r => r.nama_rute));

      // Format bulan untuk header
      const bulanNama = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember",
      ];
      const bulan = bulanNama[tanggalDari.getMonth()];
      const tahun = tanggalDari.getFullYear();
      const bulanTahun = `${bulan}-${tahun.toString().slice(-2)}`;

      // Ambil data surat dokumen
      const suratDokumen = await SuratDokumenModel.getActive();
      console.log("Data surat dokumen:", suratDokumen);

      // Query semua produksi dalam periode
      const filters = {
        tanggal_dari: req.query.tanggal_dari,
        tanggal_sampai: req.query.tanggal_sampai,
      };
      const produksiResult = await ProduksiModel.getAll(filters);
      const produksiList = produksiResult.data || [];
      console.log("Total produksi count (semua):", produksiList.length);

      // STEP 2: Untuk setiap rute, ambil kapal ASDP dan data produksi
      const semuaKapalSemuaRute = []; // Untuk sheet Pengantar
      const dataPerRute = [];

      for (const rute of ruteList) {
        console.log(`\n=== Processing Rute: ${rute.nama_rute} ===`);
        
        // Ambil kapal ASDP di rute ini
        const kapalAsdpTemplate = await TemplateKapalRuteModel.getAsdpKapalByRute(rute.rute_id);
        console.log(`Kapal ASDP di rute ${rute.nama_rute}:`, kapalAsdpTemplate.map(k => k.nama_kapal));

        const pelabuhanAsal = rute.pelabuhan_asal.nama_pelabuhan;
        const pelabuhanTujuan = rute.pelabuhan_tujuan.nama_pelabuhan;

        // Filter produksi untuk rute ini
        const kapalAsdpIds = kapalAsdpTemplate.map(k => k.kapal_id);
        const produksiRute = produksiList.filter(p => {
          if (!kapalAsdpIds.includes(p.kapal_id)) return false;

          const pelabuhanAsalProd = String(p.nama_pelabuhan_asal || "").toUpperCase();
          const pelabuhanTujuanProd = String(p.nama_pelabuhan_tujuan || "").toUpperCase();
          const pelabuhanAsalRute = pelabuhanAsal.toUpperCase();
          const pelabuhanTujuanRute = pelabuhanTujuan.toUpperCase();

          const isAsalTujuan = 
            pelabuhanAsalProd.includes(pelabuhanAsalRute) && 
            pelabuhanTujuanProd.includes(pelabuhanTujuanRute);
          
          const isTujuanAsal = 
            pelabuhanAsalProd.includes(pelabuhanTujuanRute) && 
            pelabuhanTujuanProd.includes(pelabuhanAsalRute);
          
          return isAsalTujuan || isTujuanAsal;
        });

        console.log(`Produksi untuk rute ${rute.nama_rute}:`, produksiRute.length);

        // Hitung trip per kapal per tanggal untuk rute ini
        const tripPerKapalPerTanggal = {};
        
        produksiRute.forEach((p) => {
          let namaKapal = p.nama_kapal.toUpperCase();
          if (namaKapal.startsWith("KMP.")) {
            namaKapal = namaKapal.substring(4).trim();
          }
          const tanggal = new Date(p.tanggal_produksi).getDate();

          if (!tripPerKapalPerTanggal[namaKapal]) {
            tripPerKapalPerTanggal[namaKapal] = {};
          }
          if (!tripPerKapalPerTanggal[namaKapal][tanggal]) {
            tripPerKapalPerTanggal[namaKapal][tanggal] = 0;
          }
          tripPerKapalPerTanggal[namaKapal][tanggal]++;
        });

        // Simpan data rute
        dataPerRute.push({
          rute,
          kapalAsdpTemplate,
          produksiRute,
          pelabuhanAsal,
          pelabuhanTujuan,
          tripPerKapalPerTanggal,
        });

        // Simpan untuk pengantar
        kapalAsdpTemplate.forEach(kapal => {
          const labelAsalShort = pelabuhanAsal.substring(0, 3).toUpperCase();
          const labelTujuanShort = pelabuhanTujuan.substring(0, 3).toUpperCase();
          let namaKapal = kapal.nama_kapal.toUpperCase();
          if (namaKapal.startsWith("KMP.")) {
            namaKapal = namaKapal.substring(4).trim();
          }
          
          semuaKapalSemuaRute.push({
            kapal: kapal,
            namaKapal: namaKapal,
            rute: rute,
            namaSheetGabungan: `${namaKapal}-${labelAsalShort}-${labelTujuanShort}`.substring(0, 31),
          });
        });
      }

      console.log(`\nTotal kombinasi kapal-rute: ${semuaKapalSemuaRute.length}`);

      // Urutkan dataPerRute berdasarkan jumlah kapal (terbanyak dulu)
      dataPerRute.sort((a, b) => b.kapalAsdpTemplate.length - a.kapalAsdpTemplate.length);
      console.log("\nUrutan rute (berdasarkan jumlah kapal):");
      dataPerRute.forEach(rd => {
        console.log(`  ${rd.rute.nama_rute}: ${rd.kapalAsdpTemplate.length} kapal`);
      });

      // Buat workbook
      const workbook = new ExcelJS.Workbook();

      // ========================================
      // SHEET 1: KINERJA KAPAL (GABUNGAN SEMUA RUTE)
      // ========================================
      const worksheet = workbook.addWorksheet("Kinerja KAPAL");
      worksheet.properties.defaultFont = { name: "Calibri", size: 11 };
      worksheet.properties.tabColor = { argb: "FF00B050" };

      // Header
      worksheet.getCell("A1").value = "LAPORAN KINERJA OPERASI KAPAL";
      worksheet.getCell("A1").font = { name: "Calibri", size: 16, bold: true };
      
      // Ambil pelabuhan asal dari rute pertama (asumsi semua rute dari cabang yang sama)
      const pelabuhanAsalCabang = dataPerRute[0]?.pelabuhanAsal || "LEMBAR";
      worksheet.getCell("A2").value = `CABANG LEMBAR`;
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
        worksheet.getColumn(colIndex).width = 4;
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
      setBorder(worksheet.getCell(`AI${currentRow}`), "thin");
      setBorder(worksheet.getCell(`AJ${currentRow}`), "thin");
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
          const tanggalDariDate = tanggalDari.getDate();
          const tanggalSampaiDate = tanggalSampai.getDate();

          for (let i = 1; i <= 31; i++) {
            const colIndex = 3 + i;
            const cell = row.getCell(colIndex);

            // Cek apakah tanggal ini dalam periode
            const dalamPeriode = i >= tanggalDariDate && i <= tanggalSampaiDate;

            if (dalamPeriode) {
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
            } else {
              // Di luar periode: kosongkan
              cell.value = "";
              setBorder(cell, "thin");
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
      currentRow++;
      currentRow++;

      // KETERANGAN WARNA
      worksheet.getCell(`A${currentRow}`).value = "Keterangan :";
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      currentRow++;
      currentRow++;

      worksheet.getCell(`C${currentRow}`).value = "RENC";
      worksheet.getCell(`E${currentRow}`).value = "POLA OPERASI KAPAL";
      currentRow++;

      worksheet.getCell(`C${currentRow}`).value = "REAL";
      worksheet.getCell(`E${currentRow}`).value = "REALISASI OPERASI";
      currentRow++;
      currentRow++;

      const keteranganWarna = [
        { color: "FF00FFFF", label: "KAPAL OPERASI", desc: "kondisi dimana kapal berlayar/operasi dari pelabuhan asal hingga ke pelabuhan tujuan" },
        { color: "FFFFFF00", label: "KAPAL ISTIRAHAT", desc: "kondisi kapal yang sandar/engker karena jadwal operasi" },
        { color: "ff9900cc", label: "KAPAL DOCKING", desc: "Kondisi kapal dinyatakan keluar lintasan untuk menjalani docking hingga kembali ke lintasan" },
        { color: "FFFF0000", label: "KAPAL RUSAK", desc: "Kondisi kapal tidak dapat beroperasi yang diakibatkan oleh kerusakan yang terjadi pada kapal" },
        { color: "FFFFFFFF", label: "CUACA BURUK", desc: "Kondisi kapal tidak dapat dioperasikan karena keadaan alam yang dapat membahayakan keselamatan kapal" },
        { color: "FF808000", label: "KAPAL SCRAP", desc: "Kondisi kapal untuk dihapuskan sebagai asset perusahaan" },
        { color: "FF632523", label: "SURAT/DOKUMENT", desc: "" },
      ];

      keteranganWarna.forEach((ket) => {
        mergeAndStyle(worksheet, `B${currentRow}:C${currentRow}`, "", {
          font: { name: "Calibri", size: 10 },
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: ket.color },
          },
          border: "thin",
        });
        worksheet.getCell(`E${currentRow}`).value = ket.label;
        worksheet.getCell(`I${currentRow}`).value = ket.desc;
        currentRow++;
        currentRow++;
      });

      // TANDA TANGAN
      const ttRow = 19;
      worksheet.getCell(`AI${ttRow}`).value = "GENERAL MANAGER";
      worksheet.getCell(`AI${ttRow}`).alignment = { horizontal: "center", vertical: "middle" };
      worksheet.getCell(`AI${ttRow + 6}`).value = suratDokumen.general_manager;
      worksheet.getCell(`AI${ttRow + 6}`).alignment = { horizontal: "center", vertical: "middle" };
      worksheet.getCell(`AI${ttRow + 6}`).font = { bold: true, underline: true };

      // ========================================
      // BAGIAN BAWAH: DETAIL TRIP PER KAPAL PER ARAH (GROUPED BY RUTE)
      // ========================================
      currentRow = 40;

      // Hitung trip per kapal per tanggal per arah PER RUTE (tidak digabung)
      dataPerRute.forEach(ruteData => {
        const { produksiRute, pelabuhanAsal, pelabuhanTujuan } = ruteData;
        const tripPerKapalPerTanggalPerArah = {};

        produksiRute.forEach((p) => {
          let namaKapal = p.nama_kapal.toUpperCase();
          if (namaKapal.startsWith("KMP.")) {
            namaKapal = namaKapal.substring(4).trim();
          }
          const tanggal = new Date(p.tanggal_produksi).getDate();

          if (!tripPerKapalPerTanggalPerArah[namaKapal]) {
            tripPerKapalPerTanggalPerArah[namaKapal] = {
              asalTujuan: {},
              tujuanAsal: {},
            };
          }

          const pelabuhanAsalProduksi = String(p.nama_pelabuhan_asal || "").toUpperCase();
          const pelabuhanTujuanProduksi = String(p.nama_pelabuhan_tujuan || "").toUpperCase();

          // Cek arah: asal → tujuan (Lembar → Tujuan)
          if (
            pelabuhanAsalProduksi.includes(pelabuhanAsal.toUpperCase()) &&
            pelabuhanTujuanProduksi.includes(pelabuhanTujuan.toUpperCase())
          ) {
            if (!tripPerKapalPerTanggalPerArah[namaKapal].asalTujuan[tanggal]) {
              tripPerKapalPerTanggalPerArah[namaKapal].asalTujuan[tanggal] = 0;
            }
            tripPerKapalPerTanggalPerArah[namaKapal].asalTujuan[tanggal]++;
          } 
          // Cek arah: tujuan → asal (Tujuan → Lembar)
          else if (
            pelabuhanAsalProduksi.includes(pelabuhanTujuan.toUpperCase()) &&
            pelabuhanTujuanProduksi.includes(pelabuhanAsal.toUpperCase())
          ) {
            if (!tripPerKapalPerTanggalPerArah[namaKapal].tujuanAsal[tanggal]) {
              tripPerKapalPerTanggalPerArah[namaKapal].tujuanAsal[tanggal] = 0;
            }
            tripPerKapalPerTanggalPerArah[namaKapal].tujuanAsal[tanggal]++;
          }
        });

        // Simpan ke ruteData untuk digunakan di bagian bawah
        ruteData.tripPerKapalPerTanggalPerArah = tripPerKapalPerTanggalPerArah;
      });

      // Loop per rute untuk menampilkan detail
      dataPerRute.forEach((ruteData, ruteIdx) => {
        const { rute, kapalAsdpTemplate, pelabuhanAsal, pelabuhanTujuan, tripPerKapalPerTanggalPerArah } = ruteData;

        // Header nama rute (hanya di kolom C)
        worksheet.getCell(`C${currentRow}`).value = rute.nama_rute.toUpperCase();
        worksheet.getCell(`C${currentRow}`).font = { name: "Calibri", size: 10, bold: true };
        worksheet.getCell(`C${currentRow}`).alignment = { horizontal: "center", vertical: "middle" };
        currentRow++;

        // Header tabel detail untuk rute ini
        mergeAndStyle(worksheet, `C${currentRow}:C${currentRow + 1}`, "NAMA KAPAL", {
          font: { name: "Calibri", size: 10, bold: true },
          alignment: { horizontal: "center", vertical: "middle" },
        });
        mergeAndStyle(worksheet, `D${currentRow}:AH${currentRow}`, "TANGGAL", {
          font: { name: "Calibri", size: 10, bold: true },
          alignment: { horizontal: "center", vertical: "middle" },
        });
        mergeAndStyle(worksheet, `AI${currentRow}:AJ${currentRow + 1}`, "JUMLAH", {
          font: { name: "Calibri", size: 10, bold: true },
          alignment: { horizontal: "center", vertical: "middle" },
        });

        currentRow++;
        // Tanggal 1-31 untuk rute ini
        const tanggalRowForRute = currentRow; // Simpan baris tanggal untuk formula hari ops
        for (let i = 1; i <= 31; i++) {
          const colIndex = 3 + i;
          const cell = worksheet.getCell(currentRow, colIndex);
          cell.value = i;
          cell.font = { name: "Calibri", size: 10, bold: true };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }

        currentRow++;

        // Header kolom kanan untuk rute ini
        worksheet.getCell(`AI${currentRow}`).value = "total trip";
        worksheet.getCell(`AI${currentRow}`).font = { name: "Calibri", size: 10, bold: true };
        worksheet.getCell(`AK${currentRow}`).value = "tidak ops";
        worksheet.getCell(`AK${currentRow}`).font = { name: "Calibri", size: 10, bold: true };
        worksheet.getCell(`AL${currentRow}`).value = "hari ops";
        worksheet.getCell(`AL${currentRow}`).font = { name: "Calibri", size: 10, bold: true };
        currentRow++;

        // Loop setiap kapal di rute ini
        kapalAsdpTemplate.forEach((kapal) => {
          let namaKapal = kapal.nama_kapal.toUpperCase();
          if (namaKapal.startsWith("KMP.")) {
            namaKapal = namaKapal.substring(4).trim();
          }

          const startRow = currentRow;

          // Merge nama kapal untuk 3 baris
          mergeAndStyle(worksheet, `C${startRow}:C${startRow + 2}`, namaKapal, {
            font: { name: "Calibri", size: 10, bold: true },
            alignment: { horizontal: "center", vertical: "middle" },
          });

          const tripDataPerArah = tripPerKapalPerTanggalPerArah[namaKapal] || {
            asalTujuan: {},
            tujuanAsal: {},
          };

          // Baris 1: Arah Asal → Tujuan (Lembar → Tujuan)
          for (let i = 1; i <= 31; i++) {
            const colIndex = 3 + i;
            const cell = worksheet.getCell(startRow, colIndex);
            const trip = tripDataPerArah.asalTujuan[i] || 0;

            cell.font = { name: "Calibri", size: 10, bold: true };
            cell.value = trip;
            cell.alignment = { horizontal: "center", vertical: "middle" };

            // Background kuning untuk arah asal→tujuan
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFFFF00" },
            };
          }

          // Baris 2: Arah Tujuan → Asal (Tujuan → Lembar)
          for (let i = 1; i <= 31; i++) {
            const colIndex = 3 + i;
            const cell = worksheet.getCell(startRow + 1, colIndex);
            const trip = tripDataPerArah.tujuanAsal[i] || 0;

            cell.font = { name: "Calibri", size: 10, bold: true };
            cell.value = trip;
            cell.alignment = { horizontal: "center", vertical: "middle" };

            // Background biru untuk arah tujuan→asal
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "ff0070c0" },
            };
          }

          // Baris 3: Total (formula)
          for (let i = 1; i <= 31; i++) {
            const colIndex = 3 + i;
            const cell = worksheet.getCell(startRow + 2, colIndex);
            const colLetter = getColumnLetter(colIndex);
            cell.font = { name: "Calibri", size: 10, bold: true };
            cell.value = {
              formula: `${colLetter}${startRow}+${colLetter}${startRow + 1}`,
            };
            cell.alignment = { horizontal: "center", vertical: "middle" };
          }

          // Kolom JUMLAH (AI)
          worksheet.getCell(`AI${startRow}`).value = {
            formula: `SUM(D${startRow}:AH${startRow})`,
          };
          worksheet.getCell(`AI${startRow}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AI${startRow + 1}`).value = {
            formula: `SUM(D${startRow + 1}:AH${startRow + 1})`,
          };
          worksheet.getCell(`AI${startRow + 1}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AI${startRow + 2}`).value = {
            formula: `SUM(D${startRow + 2}:AH${startRow + 2})`,
          };
          worksheet.getCell(`AI${startRow + 2}`).font = { name: "Calibri", size: 10, bold: true };

          // Label arah (AJ)
          const labelAsal = pelabuhanAsal.substring(0, 1).toUpperCase();
          const labelTujuan = pelabuhanTujuan.substring(0, 3).toUpperCase();

          const cellAsal = worksheet.getCell(`AJ${startRow}`);
          cellAsal.value = labelAsal;
          cellAsal.font = { name: "Calibri", size: 10, bold: true };
          cellAsal.alignment = { horizontal: "center", vertical: "middle" };

          const cellTujuan = worksheet.getCell(`AJ${startRow + 1}`);
          cellTujuan.value = labelTujuan.substring(0, 2);
          cellTujuan.font = { name: "Calibri", size: 10, bold: true };
          cellTujuan.alignment = { horizontal: "center", vertical: "middle" };

          const cellGabungan = worksheet.getCell(`AJ${startRow + 2}`);
          cellGabungan.value = `${labelAsal} & ${labelTujuan.substring(0, 2)}`;
          cellGabungan.font = { name: "Calibri", size: 10, bold: true };
          cellGabungan.alignment = { horizontal: "center", vertical: "middle" };

          // Kolom tidak ops (AK)
          worksheet.getCell(`AK${startRow}`).value = {
            formula: `COUNTIF(D${startRow}:AH${startRow},0)`,
          };
          worksheet.getCell(`AK${startRow}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AK${startRow + 1}`).value = {
            formula: `COUNTIF(D${startRow + 1}:AH${startRow + 1},0)`,
          };
          worksheet.getCell(`AK${startRow + 1}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AK${startRow + 2}`).value = {
            formula: `COUNTIF(D${startRow + 2}:AH${startRow + 2},0)`,
          };
          worksheet.getCell(`AK${startRow + 2}`).font = { name: "Calibri", size: 10, bold: true };

          // Kolom hari ops (AL) - gunakan baris tanggal rute ini
          worksheet.getCell(`AL${startRow}`).value = {
            formula: `$AH$${tanggalRowForRute}-AK${startRow}`,
          };
          worksheet.getCell(`AL${startRow}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AL${startRow + 1}`).value = {
            formula: `$AH$${tanggalRowForRute}-AK${startRow + 1}`,
          };
          worksheet.getCell(`AL${startRow + 1}`).font = { name: "Calibri", size: 10, bold: true };
          worksheet.getCell(`AL${startRow + 2}`).value = {
            formula: `$AH$${tanggalRowForRute}-AK${startRow + 2}`,
          };
          worksheet.getCell(`AL${startRow + 2}`).font = { name: "Calibri", size: 10, bold: true };

          currentRow += 3;
          currentRow++; // Spasi antar kapal
        });

        // Spasi antar rute
        currentRow++;
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
      worksheet.getColumn(37).width = 8;
      worksheet.getColumn(38).width = 8;
      worksheet.getColumn(39).width = 8;

      worksheet.views = [{ showGridLines: false }];
      worksheet.pageSetup = {
        paperSize: 9,
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      };

      // ========================================
      // GENERATE SHEET PENGANTAR (POSISI KE-2)
      // ========================================
      const { generateSheetPengantar } = require('./helpers/sheetPengantar');
      console.log("\n=== Generating sheet Pengantar ===");
      await generateSheetPengantar(
        workbook,
        dataPerRute,
        suratDokumen,
        tanggalDari,
        tanggalSampai,
        pelabuhanAsalCabang
      );

      // ========================================
      // GENERATE SHEET PER KAPAL PER RUTE
      // ========================================
      const { generateSheetPerKapal } = require('./helpers/sheetPerKapal');
      const { generateSheetGabunganRute } = require('./helpers/sheetGabunganRute');

      console.log("\n=== Generating sheets per kapal per rute ===");

      // Hitung kinerjaRowIndex untuk setiap kapal
      // Di sheet Kinerja KAPAL, baris dimulai dari row 44 (setelah header)
      // Setiap kapal menempati 4 baris (3 data + 1 spasi)
      let globalKapalIndex = 0;

      for (const ruteData of dataPerRute) {
        const { rute, kapalAsdpTemplate, produksiRute } = ruteData;
        
        // Query tarif untuk rute ini
        const tarifKendaraan = await TarifKendaraanModel.getByRute(rute.rute_id);
        const tarifPenumpang = await TarifPenumpangModel.getByRute(rute.rute_id);

        // Preload data penumpang dan kendaraan untuk produksi rute ini
        const penumpangPerProduksi = {};
        const kendaraanPerProduksi = {};
        
        for (const p of produksiRute) {
          try {
            penumpangPerProduksi[p.produksi_id] = await ProduksiPenumpangModel.getByProduksi(p.produksi_id);
            kendaraanPerProduksi[p.produksi_id] = await ProduksiKendaraanModel.getByProduksi(p.produksi_id);
          } catch (e) {
            penumpangPerProduksi[p.produksi_id] = [];
            kendaraanPerProduksi[p.produksi_id] = [];
          }
        }

        const sheetNamesLBR = [];

        // Generate 3 sheet untuk setiap kapal di rute ini
        for (const kapal of kapalAsdpTemplate) {
          const result = await generateSheetPerKapal(
            workbook,
            kapal,
            rute,
            produksiRute,
            tarifKendaraan,
            tarifPenumpang,
            tanggalDari,
            tanggalSampai,
            suratDokumen,
            globalKapalIndex,
            dataPerRute
          );
          
          sheetNamesLBR.push(result.sheetName3);
          globalKapalIndex++; // Increment untuk kapal berikutnya
        }

        // Simpan nama sheet LBR untuk digunakan nanti
        ruteData.sheetNamesLBR = sheetNamesLBR;

        // Generate sheet gabungan untuk rute ini (hanya jika ada lebih dari 1 kapal)
        if (kapalAsdpTemplate.length > 1) {
          const sheetGabunganName = await generateSheetGabunganRute(
            workbook,
            rute,
            kapalAsdpTemplate,
            sheetNamesLBR,
            suratDokumen,
            tanggalDari,
            tanggalSampai,
            tarifKendaraan,
            tarifPenumpang,
            produksiRute,
            penumpangPerProduksi,
            kendaraanPerProduksi
          );
          // Simpan nama sheet gabungan yang actual
          ruteData.sheetGabunganName = sheetGabunganName;
        } else {
          console.log(`Skip sheet gabungan untuk ${rute.nama_rute} (hanya 1 kapal)`);
        }
      }

      console.log("\n=== Generating gabungan kapal semua lintasan ===");

      // Identifikasi kapal yang ada di lebih dari 1 rute
      const kapalPerRute = {};
      dataPerRute.forEach(rd => {
        rd.kapalAsdpTemplate.forEach((k, idx) => {
          let namaKapal = k.nama_kapal.toUpperCase();
          if (namaKapal.startsWith("KMP.")) {
            namaKapal = namaKapal.substring(4).trim();
          }
          
          if (!kapalPerRute[namaKapal]) {
            kapalPerRute[namaKapal] = [];
          }
          
          kapalPerRute[namaKapal].push({
            rute: rd.rute,
            kapal: k,
            pelabuhanAsal: rd.pelabuhanAsal,
            pelabuhanTujuan: rd.pelabuhanTujuan,
            sheetNameLBR: rd.sheetNamesLBR[idx], // Ambil nama sheet LBR yang actual
          });
        });
      });

      // Generate sheet gabungan untuk kapal yang ada di lebih dari 1 rute
      const { generateSheetGabunganKapalSemuaLintasan } = require('./helpers/sheetGabunganKapalSemuaLintasan');
      const sheetGabunganRuteNames = []; // Untuk sheet gabungan seluruh lintasan

      for (const [namaKapal, ruteList] of Object.entries(kapalPerRute)) {
        if (ruteList.length > 1) {
          console.log(`Generating gabungan semua lintasan untuk ${namaKapal} (${ruteList.length} rute)`);
          
          // Kumpulkan nama sheet LBR untuk kapal ini dari semua rute (gunakan nama actual)
          const sheetNamesLBRKapal = ruteList.map(r => r.sheetNameLBR);

          // Simpan nama sheet yang actual yang dikembalikan oleh fungsi
          const actualSheetName = await generateSheetGabunganKapalSemuaLintasan(
            workbook,
            namaKapal,
            ruteList,
            sheetNamesLBRKapal,
            suratDokumen,
            tanggalDari,
            tanggalSampai
          );
          
          console.log(`Generated sheet: ${actualSheetName}`);
        }
      }

      console.log("\n=== Generating gabungan seluruh lintasan ===");

      // Kumpulkan semua sheet LBR dari semua rute (gunakan nama actual)
      const allSheetNamesLBR = [];
      dataPerRute.forEach(rd => {
        if (rd.sheetNamesLBR && rd.sheetNamesLBR.length > 0) {
          allSheetNamesLBR.push(...rd.sheetNamesLBR);
        }
      });

      // Generate sheet gabungan seluruh lintasan (hanya jika ada lebih dari 1 rute)
      if (dataPerRute.length > 1 && allSheetNamesLBR.length > 0) {
        const { generateSheetGabunganSeluruhLintasan } = require('./helpers/sheetGabunganSeluruhLintasan');
        const actualSheetName = await generateSheetGabunganSeluruhLintasan(
          workbook,
          allSheetNamesLBR,
          suratDokumen,
          tanggalDari,
          tanggalSampai
        );
        console.log(`Generated sheet: ${actualSheetName}`);
      }

      console.log("\n=== All sheets generated ===");

      // Send file
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="Laporan Kinerja ASDP ${bulan} ${tahun}.xlsx"`,
      );

      await workbook.xlsx.write(res);
      res.end();

      console.log("=== EXPORT KINERJA ASDP MULTI-RUTE COMPLETE ===");
    } catch (error) {
      console.error("Error in exportKinerjaAsdp:", error);
      next(error);
    }
  }
}

module.exports = LaporanKinerjaAsdpController;
