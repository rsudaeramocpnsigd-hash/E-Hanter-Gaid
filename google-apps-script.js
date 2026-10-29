/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT - BACKEND API HANDOVER DOKTER IGD
 * ==============================================================================
 * 
 * Target Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1gdqvWMOfRmyBALEymuYrHmu7kWviHXf8CsgGK1EitFA/edit
 * 
 * CARA PEMASANGAN (HANYA 1 MENIT):
 * 1. Buka spreadsheet Google Sheets Anda di browser.
 * 2. Klik menu "Ekstensi" (Extensions) -> "Apps Script".
 * 3. Hapus semua kode default di dalam editor (myFunction), lalu SALIN & TEMPEL
 *    seluruh isi kode di bawah ini.
 * 4. Klik tombol "Simpan" (ikon disket / Ctrl+S).
 * 5. Klik tombol biru "Terapkan" (Deploy) di kanan atas -> pilih "Penerapan baru" (New deployment).
 * 6. Pada ikon gerigi (Select type), pilih "Aplikasi Web" (Web app).
 * 7. Atur konfigurasi:
 *    - Deskripsi: Handover Dokter IGD API
 *    - Jalankan sebagai (Execute as): Saya (email Anda)
 *    - Yang memiliki akses (Who has access): Siapa saja (Anyone) -> WAJIB agar web app bisa kirim data!
 * 8. Klik "Terapkan" (Deploy) dan berikan izin akses akun Google Anda jika diminta.
 * 9. Salin URL Aplikasi Web (contoh: https://script.google.com/macros/s/AKfycb.../exec).
 * 10. Masukkan URL tersebut ke aplikasi Handover Dokter IGD pada tab "Master Data" -> "Integrasi Google Sheets".
 * ==============================================================================
 */

// 1. GET REQUEST: Tarik data dari Google Sheets ke Aplikasi Web
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getAll';

    if (action === 'getDoctors') {
      const userSheet = ss.getSheetByName('USER');
      const dpjpSheet = ss.getSheetByName('USER DPJP');
      return jsonResponse({
        success: true,
        doctors: getSheetData(userSheet),
        dpjps: getSheetData(dpjpSheet)
      });
    }

    // Default: Ambil semua data dari 5 sheets
    const handoverSheet = ss.getSheetByName('Handover');
    const pasienSheet = ss.getSheetByName('Pasien IGD');
    const userSheet = ss.getSheetByName('USER');
    const dpjpSheet = ss.getSheetByName('USER DPJP');
    const ringkasanSheet = ss.getSheetByName('Ringkasan Pasien');

    return jsonResponse({
      success: true,
      timestamp: new Date().toISOString(),
      handover: getSheetData(handoverSheet),
      pasien: getSheetData(pasienSheet),
      doctors: getSheetData(userSheet),
      dpjps: getSheetData(dpjpSheet),
      ringkasan: getSheetData(ringkasanSheet)
    });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// 2. POST REQUEST: Simpan sesi Handover & Data Pasien ke Google Sheets
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'Tidak ada data POST yang diterima.' });
    }

    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = payload.action || 'saveHandover';

    if (action === 'saveHandover') {
      const handover = payload.handover || {};
      const patients = payload.patients || [];

      // A. Simpan / Tambah baris ke sheet 'Handover'
      const handoverSheet = ss.getSheetByName('Handover') || ss.insertSheet('Handover');
      ensureHandoverHeaders(handoverSheet);

      handoverSheet.appendRow([
        handover.id || '',
        handover.date || '',
        handover.time || '',
        handover.shift || '',
        handover.reporter || '',
        handover.receiver || '',
        patients.length || 0,
        handover.notes || ''
      ]);

      // B. Simpan seluruh pasien ke sheet 'Pasien IGD' & 'Ringkasan Pasien'
      const pasienSheet = ss.getSheetByName('Pasien IGD') || ss.insertSheet('Pasien IGD');
      ensurePasienHeaders(pasienSheet);

      const ringkasanSheet = ss.getSheetByName('Ringkasan Pasien') || ss.insertSheet('Ringkasan Pasien');
      ensureRingkasanHeaders(ringkasanSheet);

      patients.forEach(function(p) {
        // Tulis baris pasien
        pasienSheet.appendRow([
          p.id || '',
          p.handoverId || handover.id || '',
          p.rm || '',
          p.name || '',
          p.age || '',
          p.triage || '',
          p.diagnosis || '',
          p.dpjp || '-',
          p.lab || '-',
          p.rad || '-',
          p.condition || '-',
          p.pending || '-',
          p.spri || '-',
          p.status || 'Observasi'
        ]);

        // Tulis baris ringkasan pasien
        const isObs = (p.status === 'Observasi') ? 1 : 0;
        const isPulang = (p.status === 'Boleh Pulang' || p.status === 'Pulang') ? 1 : 0;
        const isRanap = (p.status === 'Rawat Inap') ? 1 : 0;
        const isRujuk = (p.status === 'Rujuk') ? 1 : 0;
        const isMeninggal = (p.status === 'Meninggal') ? 1 : 0;

        ringkasanSheet.appendRow([
          p.id || '',
          handover.date || '',
          p.triage || '',
          isObs,
          isPulang,
          isRanap,
          isRujuk,
          isMeninggal
        ]);
      });

      return jsonResponse({
        success: true,
        message: 'Data Handover dan ' + patients.length + ' pasien berhasil dicatat ke Google Sheets!',
        handoverId: handover.id,
        savedPatientsCount: patients.length,
        timestamp: new Date().toISOString()
      });
    }

    // Aksi: Tambah Dokter Jaga baru ke sheet USER
    if (action === 'addDoctor') {
      const userSheet = ss.getSheetByName('USER');
      if (userSheet) {
        userSheet.appendRow([payload.name, 1]);
        return jsonResponse({ success: true, message: 'Dokter jaga berhasil ditambahkan ke sheet USER.' });
      }
    }

    // Aksi: Tambah DPJP baru ke sheet USER DPJP
    if (action === 'addDpjp') {
      const dpjpSheet = ss.getSheetByName('USER DPJP');
      if (dpjpSheet) {
        dpjpSheet.appendRow([payload.name, 1]);
        return jsonResponse({ success: true, message: 'Dokter DPJP berhasil ditambahkan ke sheet USER DPJP.' });
      }
    }

    return jsonResponse({ success: false, error: 'Aksi "' + action + '" tidak dikenali.' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// 3. HELPER FUNCTIONS
function getSheetData(sheet) {
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  const results = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]).trim();
      if (key) {
        row[key] = values[i][j];
      }
    }
    results.push(row);
  }
  return results;
}

function ensureHandoverHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID_Handover',
      'Tanggal Handover',
      'Waktu Handover',
      'Shift Jaga',
      'Dokter yang Melaporkan',
      'Dokter Penerima',
      'Jumlah Pasien Aktif',
      'Catatan Shift'
    ]);
  }
}

function ensurePasienHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID_Pasien_IGD',
      'ID_Handover',
      'No_Rekam_Medis',
      'Nama_Pasien',
      'Usia_Pasien',
      'Triase',
      'Diagnosis_Kerja',
      'Konsul DPJP',
      'Laboratorium',
      'Radiologi',
      'Kondisi_Terakhir',
      'Instruksi_Pending_Terapi',
      'SPRI',
      'Status_Pasien'
    ]);
  }
}

function ensureRingkasanHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID_Pasien_IGD',
      'Tanggal Handover',
      'Triase',
      'Observasi',
      'Pulang',
      'Rawat Inap',
      'Rujuk',
      'Meninggal'
    ]);
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ==============================================================================
 * 4. FUNGSI SETUP OTOMATIS (Klik 'Jalankan' / 'Run' pada fungsi ini di Apps Script)
 * ==============================================================================
 * Otomatis membuat dan merapikan 5 sheet utama:
 * - Handover
 * - Pasien IGD
 * - Ringkasan Pasien
 * - USER (Dokter Jaga)
 * - USER DPJP (Dokter Spesialis)
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Handover
  const handoverSheet = ss.getSheetByName('Handover') || ss.insertSheet('Handover');
  ensureHandoverHeaders(handoverSheet);
  handoverSheet.getRange(1, 1, 1, 8).setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold');

  // 2. Sheet Pasien IGD
  const pasienSheet = ss.getSheetByName('Pasien IGD') || ss.insertSheet('Pasien IGD');
  ensurePasienHeaders(pasienSheet);
  pasienSheet.getRange(1, 1, 1, 14).setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold');

  // 3. Sheet Ringkasan Pasien
  const ringkasanSheet = ss.getSheetByName('Ringkasan Pasien') || ss.insertSheet('Ringkasan Pasien');
  ensureRingkasanHeaders(ringkasanSheet);
  ringkasanSheet.getRange(1, 1, 1, 8).setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold');

  // 4. Sheet USER (Dokter Jaga IGD)
  const userSheet = ss.getSheetByName('USER') || ss.insertSheet('USER');
  if (userSheet.getLastRow() === 0) {
    userSheet.appendRow(['Nama Dokter', 'Status Aktif']);
    userSheet.getRange(1, 1, 1, 2).setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold');
    const defaultDoctors = [
      ['dr. Dian', 1],
      ['dr. Camelia', 1],
      ['dr. Maria Monika', 1],
      ['dr. Anskariani', 1],
      ['dr. Rounali', 1],
      ['dr. Gabriela', 1],
      ['dr. Ernilinda', 1],
      ['dr. Kandida', 1],
      ['dr. Adelaide', 1]
    ];
    defaultDoctors.forEach(function(r) { userSheet.appendRow(r); });
  }

  // 5. Sheet USER DPJP (Dokter Konsultan Spesialis)
  const dpjpSheet = ss.getSheetByName('USER DPJP') || ss.insertSheet('USER DPJP');
  if (dpjpSheet.getLastRow() === 0) {
    dpjpSheet.appendRow(['Nama Dokter Spesialis DPJP', 'Status Aktif']);
    dpjpSheet.getRange(1, 1, 1, 2).setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold');
    const defaultDpjps = [
      ['dr. Erik, Sp.PD', 1],
      ['dr. Andreas, Sp.PD', 1],
      ['dr. Rastu, Sp.B', 1],
      ['dr. Josef, Sp.B', 1],
      ['dr. Sari, Sp.A', 1],
      ['dr. Polo, Sp.OG', 1],
      ['dr. Yona, Sp.OG', 1],
      ['dr. Asih, Sp.S', 1],
      ['dr. Ngurah, Sp.M', 1],
      ['dr. Andris, Sp.THT', 1],
      ['dr. Roland, Sp.KJ', 1]
    ];
    defaultDpjps.forEach(function(r) { dpjpSheet.appendRow(r); });
  }

  try {
    SpreadsheetApp.getUi().alert('Setup Selesai! Seluruh 5 Sheet untuk E-Hanter Gaid berhasil disiapkan.');
  } catch (e) {
    Logger.log('Setup selesai!');
  }
}
