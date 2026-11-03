/**
 * JELAJAH MALAYSIA – FINGER MOTION PUZZLE CHALLENGE
 * Backend Controller & Google Sheets Database Interface
 */

// Konfigurasi ID Google Spreadsheet (Tinggalkan kosong untuk menggunakan sheet yang terikat / bound, 
// atau masukkan ID Spreadsheet anda jika menggunakan standalone script)
const SPREADSHEET_ID = ""; 
const SHEET_NAME = "Rekod_Pemain";

/**
 * Fungsi utama untuk memuatkan antaramuka Web App
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    return template.evaluate()
      .setTitle('Jelajah Malaysia – Finger Motion Puzzle Challenge')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Ralat memuatkan aplikasi: ' + err.message + '</h3>');
  }
}

/**
 * Fungsi utiliti untuk menyertakan fail separa (Style, Script, Map) ke dalam Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mendapatkan helaian rekod atau mencipta helaian baru berserta tajuk lajur secara automatik
 */
function getOrCreateSheet() {
  let ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Spreadsheet tidak ditemui. Sila tetapkan SPREADSHEET_ID di dalam Code.gs.");
    }
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Format tajuk lajur (A hingga J)
    const headers = [
      "Timestamp",
      "ID Sesi",
      "Nama Pemain",
      "Kelas",
      "Markah",
      "Tempoh Masa (Saat)",
      "Baki Nyawa",
      "Jawapan Betul",
      "Jawapan Salah",
      "Status Permainan"
    ];
    
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#1e293b");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    // Laraskan kelebaran lajur asas
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 130); // ID Sesi
    sheet.setColumnWidth(3, 180); // Nama
    sheet.setColumnWidth(4, 100); // Kelas
    sheet.setColumnWidth(5, 80);  // Markah
    sheet.setColumnWidth(6, 140); // Tempoh Masa
    sheet.setColumnWidth(7, 90);  // Baki Nyawa
    sheet.setColumnWidth(8, 110); // Jawapan Betul
    sheet.setColumnWidth(9, 110); // Jawapan Salah
    sheet.setColumnWidth(10, 130); // Status
  }
  return sheet;
}

/**
 * Menyimpan rekod skor permainan ke dalam Google Sheets secara atomik & selamat
 * @param {Object} data - Objek keputusan dihantar dari klien
 * @return {Object} Status kejayaan
 */
function simpanRekodPemain(data) {
  // Gunakan LockService untuk mengelakkan pertembungan penulisan data serentak
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000); // Tunggu maksimum 10 saat
  
  if (!hasLock) {
    return {
      success: false,
      message: "Pelayan sedang sibuk mengendalikan permintaan lain. Sila cuba sebentar lagi."
    };
  }

  try {
    // 1. Sanitasi dan Pengesyoran Data Input
    if (!data || !data.nama || !data.kelas || !data.sessionId) {
      throw new Error("Maklumat sesi permainan tidak lengkap atau tidak sah.");
    }

    const nama = String(data.nama).trim().substring(0, 50);
    const kelas = String(data.kelas).trim().substring(0, 20);
    const sessionId = String(data.sessionId).trim();
    const masaSaat = Math.max(0, parseInt(data.masaSaat, 10) || 0);
    const jawapanBetul = Math.min(16, Math.max(0, parseInt(data.jawapanBetul, 10) || 0));
    const jawapanSalah = Math.max(0, parseInt(data.jawapanSalah, 10) || 0);
    const bakiNyawa = Math.min(3, Math.max(0, parseInt(data.bakiNyawa, 10) || 0));
    const statusPermainan = data.statusPermainan === "BERJAYA" ? "BERJAYA" : "GAME OVER";

    // 2. Pengesahan Integriti Pemarkahan di Peringkat Pelayan (Anti-tamper verification)
    // +10 markah setiap negeri betul (maksimum 160)
    let markahKira = jawapanBetul * 10;
    
    // Bonus tamat semua 16 label: +20 markah
    if (jawapanBetul === 16 && statusPermainan === "BERJAYA") {
      markahKira += 20;
      // Bonus masa: +20 markah jika selesai dalam 180 saat atau kurang
      if (masaSaat <= 180) {
        markahKira += 20;
      }
    }
    
    // Pastikan markah berada dalam julat 0 hingga 200
    const markahSah = Math.min(200, Math.max(0, markahKira));

    const sheet = getOrCreateSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    // 3. Semakan Pencegahan Rekod Berganda berasaskan ID Sesi
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]) === sessionId) {
        return {
          success: true,
          message: "Rekod untuk sesi ini telah disimpan sebelum ini."
        };
      }
    }

    // 4. Masukkan rekod ke helaian
    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      sessionId,
      nama,
      kelas,
      markahSah,
      masaSaat,
      bakiNyawa,
      jawapanBetul,
      jawapanSalah,
      statusPermainan
    ]);

    return {
      success: true,
      message: "Tahniah! Rekod markah berjaya disimpan.",
      finalScore: markahSah
    };

  } catch (err) {
    return {
      success: false,
      message: "Ralat semasa menyimpan rekod: " + err.message
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membaca 10 pemain terbaik untuk dipaparkan di Papan Pendahulu (Leaderboard)
 * Kriteria susunan:
 * 1. Markah tertinggi (Descending)
 * 2. Masa terpantas (Ascending)
 * 3. Baki nyawa tertinggi (Descending)
 * @return {Object} Senarai pemain kedudukan 1 hingga 10
 */
function dapatkanPapanPendahulu() {
  try {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();
    
    if (rows.length <= 1) {
      return { success: true, data: [] };
    }

    // Ekstrak baris data tanpa baris tajuk
    const dataRows = rows.slice(1);
    
    const senaraiPemain = dataRows.map(function(r) {
      return {
        timestamp: r[0],
        sessionId: r[1],
        nama: r[2],
        kelas: r[3],
        markah: Number(r[4]) || 0,
        masaSaat: Number(r[5]) || 0,
        bakiNyawa: Number(r[6]) || 0,
        jawapanBetul: Number(r[7]) || 0,
        jawapanSalah: Number(r[8]) || 0,
        status: r[9]
      };
    });

    // Susun data mengikut spesifikasi pertandingan
    senaraiPemain.sort(function(a, b) {
      if (b.markah !== a.markah) {
        return b.markah - a.markah; // 1. Markah tertinggi
      }
      if (a.masaSaat !== b.masaSaat) {
        return a.masaSaat - b.masaSaat; // 2. Masa terpantas
      }
      return b.bakiNyawa - a.bakiNyawa; // 3. Baki nyawa tertinggi
    });

    // Ambil 10 teratas
    const top10 = senaraiPemain.slice(0, 10).map(function(item, index) {
      return {
        kedudukan: index + 1,
        nama: item.nama,
        kelas: item.kelas,
        markah: item.markah,
        masaSaat: item.masaSaat,
        bakiNyawa: item.bakiNyawa,
        status: item.status
      };
    });

    return {
      success: true,
      data: top10
    };

  } catch (err) {
    return {
      success: false,
      message: "Gagal memuat turun data papan pendahulu: " + err.message,
      data: []
    };
  }
}
