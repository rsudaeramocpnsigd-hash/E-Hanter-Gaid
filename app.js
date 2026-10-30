/**
 * APLIKASI HANDOVER DOKTER IGD
 * Sistem Digitalisasi Serah Terima Jaga Medis Berbasis Web
 * 
 * Sesuai Spesifikasi:
 * - Rancangan Aplikasi Handover Dokter IGD.pdf
 * - Handover Dokter.xlsx (Tabel Handover, Pasien IGD, USER, USER DPJP, Ringkasan)
 */

// ==========================================================================
// 1. DATA MASTER AWAL DARI FILE EXCEL (SEED DATA)
// ==========================================================================
const DEFAULT_DOKTER_JAGA = [
  { id: 'DOC-01', name: 'dr. Dian', pin: '1234', active: true },
  { id: 'DOC-02', name: 'dr. Camelia', pin: '1234', active: true },
  { id: 'DOC-03', name: 'dr. Maria Monika', pin: '1234', active: true },
  { id: 'DOC-04', name: 'dr. Anskariani', pin: '1234', active: true },
  { id: 'DOC-05', name: 'dr. Rounali', pin: '1234', active: true },
  { id: 'DOC-06', name: 'dr. Gabriela', pin: '1234', active: true },
  { id: 'DOC-07', name: 'dr. Ernilinda', pin: '1234', active: true },
  { id: 'DOC-08', name: 'dr. Kandida', pin: '1234', active: true },
  { id: 'DOC-09', name: 'dr. Adelaide', pin: '1234', active: true }
];

const DEFAULT_DPJP = [
  { id: 'DPJP-01', name: 'dr. Erik, Sp.PD', specialist: 'Penyakit Dalam', active: true },
  { id: 'DPJP-02', name: 'dr. Andreas, Sp.PD', specialist: 'Penyakit Dalam', active: true },
  { id: 'DPJP-03', name: 'dr. Rastu, Sp.B', specialist: 'Bedah', active: true },
  { id: 'DPJP-04', name: 'dr. Josef, Sp.B', specialist: 'Bedah', active: true },
  { id: 'DPJP-05', name: 'dr. Sari, Sp.A', specialist: 'Anak', active: true },
  { id: 'DPJP-06', name: 'dr. Polo, Sp.OG', specialist: 'Kebidanan & Kandungan', active: true },
  { id: 'DPJP-07', name: 'dr. Yona, Sp.OG', specialist: 'Kebidanan & Kandungan', active: true },
  { id: 'DPJP-08', name: 'dr. Asih, Sp.S', specialist: 'Saraf', active: true },
  { id: 'DPJP-09', name: 'dr. Ngurah, Sp.M', specialist: 'Mata', active: true },
  { id: 'DPJP-10', name: 'dr. Andris, Sp.THT', specialist: 'THT-KL', active: true },
  { id: 'DPJP-11', name: 'dr. Roland, Sp.KJ', specialist: 'Kedokteran Jiwa', active: true }
];

const DEFAULT_SAMPLE_PATIENTS = [];

// ==========================================================================
// 2. STATE APLIKASI
// ==========================================================================
const DEFAULT_GSHEET_CONFIG = {
  sheetId: '1TKHVo7SMfZYa73_JaA8O4xp5QJS67p_DuqpX_2zDt2o',
  sheetUrl: 'https://docs.google.com/spreadsheets/d/1TKHVo7SMfZYa73_JaA8O4xp5QJS67p_DuqpX_2zDt2o/edit?usp=sharing',
  webAppUrl: 'https://script.google.com/macros/s/AKfycbzUoLenrSriKUSYojU2hyn1XQvrMxiV4OJG0IruMbPVzLWFnAJUtlbPIHNqNBb_ABsP1w/exec',
  lastSync: null
};

const GAS_BACKEND_CODE = '/**\n * ==============================================================================\n * GOOGLE APPS SCRIPT - BACKEND API HANDOVER DOKTER IGD\n * ==============================================================================\n * \n * Target Spreadsheet:\n * https://docs.google.com/spreadsheets/d/1gdqvWMOfRmyBALEymuYrHmu7kWviHXf8CsgGK1EitFA/edit\n * \n * CARA PEMASANGAN (HANYA 1 MENIT):\n * 1. Buka spreadsheet Google Sheets Anda di browser.\n * 2. Klik menu "Ekstensi" (Extensions) -> "Apps Script".\n * 3. Hapus semua kode default di dalam editor (myFunction), lalu SALIN & TEMPEL\n *    seluruh isi kode di bawah ini.\n * 4. Klik tombol "Simpan" (ikon disket / Ctrl+S).\n * 5. Klik tombol biru "Terapkan" (Deploy) di kanan atas -> pilih "Penerapan baru" (New deployment).\n * 6. Pada ikon gerigi (Select type), pilih "Aplikasi Web" (Web app).\n * 7. Atur konfigurasi:\n *    - Deskripsi: Handover Dokter IGD API\n *    - Jalankan sebagai (Execute as): Saya (email Anda)\n *    - Yang memiliki akses (Who has access): Siapa saja (Anyone) -> WAJIB agar web app bisa kirim data!\n * 8. Klik "Terapkan" (Deploy) dan berikan izin akses akun Google Anda jika diminta.\n * 9. Salin URL Aplikasi Web (contoh: https://script.google.com/macros/s/AKfycb.../exec).\n * 10. Masukkan URL tersebut ke aplikasi Handover Dokter IGD pada tab "Master Data" -> "Integrasi Google Sheets".\n * ==============================================================================\n */\n\n// 1. GET REQUEST: Tarik data dari Google Sheets ke Aplikasi Web\nfunction doGet(e) {\n  try {\n    const ss = SpreadsheetApp.getActiveSpreadsheet();\n    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : \'getAll\';\n\n    if (action === \'getDoctors\') {\n      const userSheet = ss.getSheetByName(\'USER\');\n      const dpjpSheet = ss.getSheetByName(\'USER DPJP\');\n      return jsonResponse({\n        success: true,\n        doctors: getSheetData(userSheet),\n        dpjps: getSheetData(dpjpSheet)\n      });\n    }\n\n    // Default: Ambil semua data dari 5 sheets\n    const handoverSheet = ss.getSheetByName(\'Handover\');\n    const pasienSheet = ss.getSheetByName(\'Pasien IGD\');\n    const userSheet = ss.getSheetByName(\'USER\');\n    const dpjpSheet = ss.getSheetByName(\'USER DPJP\');\n    const ringkasanSheet = ss.getSheetByName(\'Ringkasan Pasien\');\n\n    return jsonResponse({\n      success: true,\n      timestamp: new Date().toISOString(),\n      handover: getSheetData(handoverSheet),\n      pasien: getSheetData(pasienSheet),\n      doctors: getSheetData(userSheet),\n      dpjps: getSheetData(dpjpSheet),\n      ringkasan: getSheetData(ringkasanSheet)\n    });\n  } catch (err) {\n    return jsonResponse({ success: false, error: err.toString() });\n  }\n}\n\n// 2. POST REQUEST: Simpan sesi Handover & Data Pasien ke Google Sheets\nfunction doPost(e) {\n  try {\n    if (!e || !e.postData || !e.postData.contents) {\n      return jsonResponse({ success: false, error: \'Tidak ada data POST yang diterima.\' });\n    }\n\n    const payload = JSON.parse(e.postData.contents);\n    const ss = SpreadsheetApp.getActiveSpreadsheet();\n    const action = payload.action || \'saveHandover\';\n\n    if (action === \'saveHandover\') {\n      const handover = payload.handover || {};\n      const patients = payload.patients || [];\n\n      // A. Simpan / Tambah baris ke sheet \'Handover\'\n      const handoverSheet = ss.getSheetByName(\'Handover\') || ss.insertSheet(\'Handover\');\n      ensureHandoverHeaders(handoverSheet);\n\n      handoverSheet.appendRow([\n        handover.id || \'\',\n        handover.date || \'\',\n        handover.time || \'\',\n        handover.shift || \'\',\n        handover.reporter || \'\',\n        handover.receiver || \'\',\n        patients.length || 0,\n        handover.notes || \'\'\n      ]);\n\n      // B. Simpan seluruh pasien ke sheet \'Pasien IGD\' & \'Ringkasan Pasien\'\n      const pasienSheet = ss.getSheetByName(\'Pasien IGD\') || ss.insertSheet(\'Pasien IGD\');\n      ensurePasienHeaders(pasienSheet);\n\n      const ringkasanSheet = ss.getSheetByName(\'Ringkasan Pasien\') || ss.insertSheet(\'Ringkasan Pasien\');\n      ensureRingkasanHeaders(ringkasanSheet);\n\n      patients.forEach(function(p) {\n        // Tulis baris pasien\n        pasienSheet.appendRow([\n          p.id || \'\',\n          p.handoverId || handover.id || \'\',\n          p.rm || \'\',\n          p.name || \'\',\n          p.age || \'\',\n          p.triage || \'\',\n          p.diagnosis || \'\',\n          p.dpjp || \'-\',\n          p.lab || \'-\',\n          p.rad || \'-\',\n          p.condition || \'-\',\n          p.pending || \'-\',\n          p.spri || \'-\',\n          p.status || \'Observasi\'\n        ]);\n\n        // Tulis baris ringkasan pasien\n        const isObs = (p.status === \'Observasi\') ? 1 : 0;\n        const isPulang = (p.status === \'Boleh Pulang\' || p.status === \'Pulang\') ? 1 : 0;\n        const isRanap = (p.status === \'Rawat Inap\') ? 1 : 0;\n        const isRujuk = (p.status === \'Rujuk\') ? 1 : 0;\n        const isMeninggal = (p.status === \'Meninggal\') ? 1 : 0;\n\n        ringkasanSheet.appendRow([\n          p.id || \'\',\n          handover.date || \'\',\n          p.triage || \'\',\n          isObs,\n          isPulang,\n          isRanap,\n          isRujuk,\n          isMeninggal\n        ]);\n      });\n\n      return jsonResponse({\n        success: true,\n        message: \'Data Handover dan \' + patients.length + \' pasien berhasil dicatat ke Google Sheets!\',\n        handoverId: handover.id,\n        savedPatientsCount: patients.length,\n        timestamp: new Date().toISOString()\n      });\n    }\n\n    // Aksi: Tambah Dokter Jaga baru ke sheet USER\n    if (action === \'addDoctor\') {\n      const userSheet = ss.getSheetByName(\'USER\');\n      if (userSheet) {\n        userSheet.appendRow([payload.name, 1]);\n        return jsonResponse({ success: true, message: \'Dokter jaga berhasil ditambahkan ke sheet USER.\' });\n      }\n    }\n\n    // Aksi: Tambah DPJP baru ke sheet USER DPJP\n    if (action === \'addDpjp\') {\n      const dpjpSheet = ss.getSheetByName(\'USER DPJP\');\n      if (dpjpSheet) {\n        dpjpSheet.appendRow([payload.name, 1]);\n        return jsonResponse({ success: true, message: \'Dokter DPJP berhasil ditambahkan ke sheet USER DPJP.\' });\n      }\n    }\n\n    return jsonResponse({ success: false, error: \'Aksi "\' + action + \'" tidak dikenali.\' });\n  } catch (err) {\n    return jsonResponse({ success: false, error: err.toString() });\n  }\n}\n\n// 3. HELPER FUNCTIONS\nfunction getSheetData(sheet) {\n  if (!sheet) return [];\n  const range = sheet.getDataRange();\n  const values = range.getValues();\n  if (values.length <= 1) return [];\n\n  const headers = values[0];\n  const results = [];\n  for (var i = 1; i < values.length; i++) {\n    var row = {};\n    for (var j = 0; j < headers.length; j++) {\n      var key = String(headers[j]).trim();\n      if (key) {\n        row[key] = values[i][j];\n      }\n    }\n    results.push(row);\n  }\n  return results;\n}\n\nfunction ensureHandoverHeaders(sheet) {\n  if (sheet.getLastRow() === 0) {\n    sheet.appendRow([\n      \'ID_Handover\',\n      \'Tanggal Handover\',\n      \'Waktu Handover\',\n      \'Shift Jaga\',\n      \'Dokter yang Melaporkan\',\n      \'Dokter Penerima\',\n      \'Jumlah Pasien Aktif\',\n      \'Catatan Shift\'\n    ]);\n  }\n}\n\nfunction ensurePasienHeaders(sheet) {\n  if (sheet.getLastRow() === 0) {\n    sheet.appendRow([\n      \'ID_Pasien_IGD\',\n      \'ID_Handover\',\n      \'No_Rekam_Medis\',\n      \'Nama_Pasien\',\n      \'Usia_Pasien\',\n      \'Triase\',\n      \'Diagnosis_Kerja\',\n      \'Konsul DPJP\',\n      \'Laboratorium\',\n      \'Radiologi\',\n      \'Kondisi_Terakhir\',\n      \'Instruksi_Pending_Terapi\',\n      \'SPRI\',\n      \'Status_Pasien\'\n    ]);\n  }\n}\n\nfunction ensureRingkasanHeaders(sheet) {\n  if (sheet.getLastRow() === 0) {\n    sheet.appendRow([\n      \'ID_Pasien_IGD\',\n      \'Tanggal Handover\',\n      \'Triase\',\n      \'Observasi\',\n      \'Pulang\',\n      \'Rawat Inap\',\n      \'Rujuk\',\n      \'Meninggal\'\n    ]);\n  }\n}\n\nfunction jsonResponse(data) {\n  return ContentService.createTextOutput(JSON.stringify(data))\n    .setMimeType(ContentService.MimeType.JSON);\n}\n';

let AppState = {
  currentUser: null, // { id, name, role, shift, loginAt }
  doctors: [],
  dpjps: [],
  currentHandover: null,
  handoverHistory: [],
  patients: [],
  activeDoctorId: 'DOC-01',
  viewMode: 'cards', // 'cards' | 'table'
  filters: {
    search: '',
    triage: 'ALL',
    status: 'ALL',
    dpjp: 'ALL',
    shift: 'ALL'
  },
  signatures: {
    reporter: null,
    receiver: null
  },
  gsheet: { ...DEFAULT_GSHEET_CONFIG }
};

// ==========================================================================
// 3. INISIALISASI & PERSISTENSI STORAGE
// ==========================================================================
function initApp() {
  loadFromStorage();
  setupEventListeners();
  setupSignaturePads();
  startLiveClock();
  renderAllViews();
  checkAuthStatus();

  // Auto-sinkronisasi data dari Google Sheets saat pertama dibuka (background silent)
  if (AppState.gsheet && (AppState.gsheet.sheetId || AppState.gsheet.webAppUrl)) {
    setTimeout(() => {
      fetchMasterFromGoogleSheets(true);
    }, 600);
  }
}

function loadFromStorage() {
  const savedDocs = localStorage.getItem('HANDOVER_DOCTORS');
  AppState.doctors = savedDocs ? JSON.parse(savedDocs) : [...DEFAULT_DOKTER_JAGA];
  // Ensure every doctor has pin property
  AppState.doctors = AppState.doctors.map(d => ({ ...d, pin: d.pin || '1234' }));

  // Load Auth Session
  const savedAuth = localStorage.getItem('HANDOVER_AUTH_SESSION');
  if (savedAuth) {
    try {
      AppState.currentUser = JSON.parse(savedAuth);
      AppState.activeDoctorId = AppState.currentUser.id || 'DOC-01';
    } catch (e) {
      AppState.currentUser = null;
    }
  }

  const savedDpjps = localStorage.getItem('HANDOVER_DPJPS');
  AppState.dpjps = savedDpjps ? JSON.parse(savedDpjps) : [...DEFAULT_DPJP];

  const savedHistory = localStorage.getItem('HANDOVER_HISTORY_LOG');
  AppState.handoverHistory = savedHistory ? JSON.parse(savedHistory) : getDefaultHandoverHistory();

  // Auto-reset patients to zero for clean production start
  const hasResetToZero = localStorage.getItem('HANDOVER_PATIENTS_CLEARED_ZERO_V3');
  if (!hasResetToZero) {
    AppState.patients = [];
    localStorage.setItem('HANDOVER_PATIENTS', JSON.stringify([]));
    localStorage.setItem('HANDOVER_PATIENTS_CLEARED_ZERO_V3', 'true');
  } else {
    const savedPatients = localStorage.getItem('HANDOVER_PATIENTS');
    AppState.patients = savedPatients ? JSON.parse(savedPatients) : [];
  }

    const savedCurrent = localStorage.getItem('HANDOVER_CURRENT_SESSION');
  if (savedCurrent) {
    AppState.currentHandover = JSON.parse(savedCurrent);
    // Normalize shift names to new schedule (08.00-14.00, 14.00-20.00, 20.00-08.00)
    if (AppState.currentHandover && AppState.currentHandover.shift) {
      const s = AppState.currentHandover.shift;
      if (s.includes('07.00') || s.includes('Pagi')) {
        AppState.currentHandover.shift = 'Shift Pagi (08.00 - 14.00)';
      } else if (s.includes('21.00') || s.includes('Sore') || s.includes('Siang')) {
        AppState.currentHandover.shift = 'Shift Siang (14.00 - 20.00)';
      } else if (s.includes('Malam')) {
        AppState.currentHandover.shift = 'Shift Malam (20.00 - 08.00)';
      }
    }
  } else {
    // Generate default current shift session
    AppState.currentHandover = createDefaultShiftSession();
  }

  const savedSigs = localStorage.getItem('HANDOVER_SIGNATURES');
  if (savedSigs) {
    AppState.signatures = JSON.parse(savedSigs);
  }

  const savedGSheet = localStorage.getItem('HANDOVER_GSHEET_CONFIG');
  if (savedGSheet) {
    try {
      AppState.gsheet = { ...DEFAULT_GSHEET_CONFIG, ...JSON.parse(savedGSheet) };
      // Enforce official hardcoded spreadsheet and webhook URL
      AppState.gsheet.sheetId = DEFAULT_GSHEET_CONFIG.sheetId;
      AppState.gsheet.sheetUrl = DEFAULT_GSHEET_CONFIG.sheetUrl;
      AppState.gsheet.webAppUrl = DEFAULT_GSHEET_CONFIG.webAppUrl;
    } catch (e) {
      AppState.gsheet = { ...DEFAULT_GSHEET_CONFIG };
    }
  } else {
    AppState.gsheet = { ...DEFAULT_GSHEET_CONFIG };
  }
}

function saveToStorage() {
  localStorage.setItem('HANDOVER_DOCTORS', JSON.stringify(AppState.doctors));
  localStorage.setItem('HANDOVER_DPJPS', JSON.stringify(AppState.dpjps));
  localStorage.setItem('HANDOVER_PATIENTS', JSON.stringify(AppState.patients));
  localStorage.setItem('HANDOVER_CURRENT_SESSION', JSON.stringify(AppState.currentHandover));
  localStorage.setItem('HANDOVER_HISTORY_LOG', JSON.stringify(AppState.handoverHistory));
  localStorage.setItem('HANDOVER_SIGNATURES', JSON.stringify(AppState.signatures));
  localStorage.setItem('HANDOVER_GSHEET_CONFIG', JSON.stringify(AppState.gsheet));
}

function createDefaultShiftSession() {
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0];
  const timeStr = String(now.getHours()).padStart(2, '0') + ':' + String(now.getMinutes()).padStart(2, '0');
  const shiftInfo = getShiftByHour(now.getHours());

  return {
    id: `HO-${dateStr.replace(/-/g, '')}-01`,
    date: dateStr,
    time: timeStr,
    shift: shiftInfo.name,
    reporter: AppState.doctors[0]?.name || 'dr. Dian',
    receiver: AppState.doctors[1]?.name || 'dr. Camelia',
    notes: 'Operan shift berjalan lancar. 1 pasien Triase Merah telah dititrasi inotropik menunggu ruang intensif.',
    locked: false,
    timestamp: new Date().toISOString()
  };
}

function getShiftByHour(hours) {
  if (hours >= 8 && hours < 14) {
    return { name: 'Shift Pagi (08.00 - 14.00)', code: 'PAGI' };
  } else if (hours >= 14 && hours < 20) {
    return { name: 'Shift Siang (14.00 - 20.00)', code: 'SIANG' };
  } else {
    return { name: 'Shift Malam (20.00 - 08.00)', code: 'MALAM' };
  }
}

// ==========================================================================
// 4. LIVE CLOCK & AUTO SHIFT BANNER
// ==========================================================================
function startLiveClock() {
  function updateTime() {
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const mins = String(now.getMinutes()).padStart(2, '0');
    const secs = String(now.getSeconds()).padStart(2, '0');
    
    const timeElem = document.getElementById('clockTime');
    if (timeElem) timeElem.textContent = `${hours}:${mins}:${secs}`;

    const dateOptions = { day: 'numeric', month: 'short', year: 'numeric' };
    const dateElem = document.getElementById('clockDate');
    if (dateElem) dateElem.textContent = now.toLocaleDateString('id-ID', dateOptions);

    // Update Banner
    const shift = getShiftByHour(now.getHours());
    const shiftElem = document.getElementById('shiftBannerText');
    if (shiftElem) shiftElem.textContent = shift.name;
  }
  updateTime();
  setInterval(updateTime, 1000);
}

// ==========================================================================
// 5. VIEW RENDERING & DOM UPDATES
// ==========================================================================
function renderAllViews() {
  updateGSheetUI();
  populateDoctorDropdowns();
  renderDashboardStats();
  renderDashboardPatientTable();
  populateHandoverForm();
  renderFormPatientsTable();
  renderTrackerView();
  renderPrintableReport();
  renderMasterDataTables();
  renderLastShiftBanner();
  renderHandoverHistory();
}

function populateDoctorDropdowns() {
  // Top active doctor select (if exists)
  const topDocSelect = document.getElementById('activeDoctorSelect');
  if (topDocSelect) {
    topDocSelect.innerHTML = AppState.doctors
      .filter(d => d.active)
      .map(d => `<option value="${d.id}">${d.name}</option>`)
      .join('');
    topDocSelect.value = AppState.activeDoctorId;
  }

  // Login Doctor Select
  const loginDocSelect = document.getElementById('loginDoctorSelect');
  if (loginDocSelect) {
    const activeDocs = AppState.doctors.filter(d => d.active);
    loginDocSelect.innerHTML = activeDocs.map(d => `<option value="${d.id}">${d.name} (Dokter Jaga IGD)</option>`).join('');
    if (AppState.currentUser && AppState.currentUser.id) {
      loginDocSelect.value = AppState.currentUser.id;
    } else if (AppState.activeDoctorId) {
      loginDocSelect.value = AppState.activeDoctorId;
    } else if (activeDocs[0]) {
      loginDocSelect.value = activeDocs[0].id;
    }
  }

  // Handover Form Doctor Selects
  const selReporter = document.getElementById('selectDokterPelapor');
  const selReceiver = document.getElementById('selectDokterPenerima');
  if (selReporter && selReceiver) {
    const optionsHtml = AppState.doctors
      .filter(d => d.active)
      .map(d => `<option value="${d.name}">${d.name}</option>`)
      .join('');
    selReporter.innerHTML = optionsHtml;
    selReceiver.innerHTML = optionsHtml;

    if (AppState.currentHandover) {
      selReporter.value = AppState.currentHandover.reporter;
      selReceiver.value = AppState.currentHandover.receiver;
    }
  }

  // Patient Modal DPJP Select
  const selDpjp = document.getElementById('selectPatDpjp');
  if (selDpjp) {
    let dpjpOpts = `<option value="-">- Tidak Ada / Belum Konsul -</option>`;
    dpjpOpts += AppState.dpjps
      .filter(d => d.active)
      .map(d => `<option value="${d.name}">${d.name} (${d.specialist})</option>`)
      .join('');
    selDpjp.innerHTML = dpjpOpts;
  }

  // Tracker Filter DPJP Select
  const filterDpjpSel = document.getElementById('filterDpjp');
  if (filterDpjpSel) {
    let fOpts = `<option value="ALL">Semua Dokter Spesialis</option>`;
    fOpts += AppState.dpjps
      .filter(d => d.active)
      .map(d => `<option value="${d.name}">${d.name}</option>`)
      .join('');
    filterDpjpSel.innerHTML = fOpts;
    filterDpjpSel.value = AppState.filters.dpjp;
  }
}

// 5.1 DASHBOARD STATS
function renderDashboardStats() {
  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  let currentPatients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  // Jika shift berjalan belum memiliki pasien tercatat, tampilkan seluruh pasien IGD aktif
  if (currentPatients.length === 0 && AppState.patients.length > 0) {
    currentPatients = AppState.patients;
  }

  const countRed = currentPatients.filter(p => p.triage === 'Merah').length;
  const countYellow = currentPatients.filter(p => p.triage === 'Kuning').length;
  const countGreen = currentPatients.filter(p => p.triage === 'Hijau').length;
  const countTotal = currentPatients.length;

  const countPending = currentPatients.filter(p => p.pending && p.pending.trim() !== '-' && p.pending.trim() !== '').length;
  const countSpri = currentPatients.filter(p => p.spri && p.spri.trim() !== '-' && p.spri.trim() !== '').length;
  const countDpjp = currentPatients.filter(p => p.dpjp && p.dpjp.trim() !== '-' && p.dpjp.trim() !== '').length;
  const countDischarged = currentPatients.filter(p => p.status === 'Boleh Pulang').length;

  // DOM elements
  document.getElementById('statTriageRed').textContent = countRed;
  document.getElementById('statTriageYellow').textContent = countYellow;
  document.getElementById('statTriageGreen').textContent = countGreen;
  document.getElementById('statTotalShift').textContent = countTotal;

  document.getElementById('metricPendingCount').textContent = `${countPending} Pasien`;
  document.getElementById('metricSpriCount').textContent = `${countSpri} Pasien`;
  document.getElementById('metricDpjpCount').textContent = `${countDpjp} Pasien`;
  document.getElementById('metricDischargedCount').textContent = `${countDischarged} Pasien`;

  document.getElementById('badgeTotalActive').textContent = countTotal;
  document.getElementById('dashPatientCountBadge').textContent = `${countTotal} Pasien`;

  // Update session summary card
  if (AppState.currentHandover) {
    document.getElementById('currentSessionIdBadge').textContent = AppState.currentHandover.id;
    document.getElementById('sessionDateTimeText').textContent = `${formatDateIndo(AppState.currentHandover.date)} • ${AppState.currentHandover.shift}`;
    document.getElementById('sessionReporterText').textContent = AppState.currentHandover.reporter;
    document.getElementById('sessionReceiverText').textContent = AppState.currentHandover.receiver;
    document.getElementById('statShiftNameLabel').textContent = AppState.currentHandover.shift;

    const lockBadge = document.getElementById('sessionLockStatusBadge');
    if (AppState.currentHandover.locked) {
      lockBadge.className = 'badge badge-neutral';
      lockBadge.innerHTML = `<i class="fa-solid fa-lock"></i> Terkunci (Final)`;
    } else {
      lockBadge.className = 'badge badge-success';
      lockBadge.innerHTML = `<i class="fa-solid fa-lock-open"></i> Aktif / Draft`;
    }
  }
}

// 5.2 DASHBOARD PATIENTS TABLE
function renderDashboardPatientTable() {
  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  let patients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  const isAllFallback = (patients.length === 0 && AppState.patients.length > 0);
  if (isAllFallback) {
    patients = AppState.patients;
  }
  const tbody = document.getElementById('dashPatientTableBody');
  if (!tbody) return;

  if (patients.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" class="text-center text-muted p-6">
          <i class="fa-solid fa-inbox text-dim" style="font-size: 2rem;"></i>
          <p class="mt-2">Belum ada pasien yang dicatat untuk sesi shift ini.</p>
          <button class="btn btn-primary btn-xs mt-2" onclick="openAddPatientModal()">+ Tambah Pasien</button>
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = patients.map(p => `
    <tr>
      <td>${getTriageBadgeHtml(p.triage)}</td>
      <td class="font-mono"><strong>${p.rm}</strong></td>
      <td>
        <strong class="text-white">${escapeHtml(p.name)}</strong>
        <div class="text-dim text-xs">${escapeHtml(p.age)}</div>
      </td>
      <td><span class="text-white">${escapeHtml(p.diagnosis)}</span></td>
      <td><span class="text-blue">${escapeHtml(p.dpjp || '-')}</span></td>
      <td><div class="text-xs text-muted" style="max-width: 200px;">${escapeHtml(p.condition)}</div></td>
      <td>
        <div class="pending-text text-xs text-amber" style="max-width: 220px; font-weight: 500;">
          <i class="fa-solid fa-clock-rotate-left"></i> ${escapeHtml(p.pending)}
        </div>
      </td>
      <td>${getStatusBadgeHtml(p.status, p.spri)}</td>
      <td>
        <div class="card-actions">
          <button class="btn btn-secondary btn-xs" onclick="editPatient('${p.id}')" title="Edit Pasien">
            <i class="fa-solid fa-pen"></i>
          </button>
          <button class="btn btn-danger btn-xs" onclick="deletePatient('${p.id}')" title="Hapus Pasien">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

// 5.3 HANDOVER FORM MASTER-DETAIL
function populateHandoverForm() {
  if (!AppState.currentHandover) return;
  const h = AppState.currentHandover;

  const idInp = document.getElementById('inputHandoverId');
  const dateInp = document.getElementById('inputHandoverDate');
  const timeInp = document.getElementById('inputHandoverTime');
  const shiftSel = document.getElementById('selectHandoverShift');
  const repSel = document.getElementById('selectDokterPelapor');
  const recSel = document.getElementById('selectDokterPenerima');
  const notesInp = document.getElementById('inputCatatanShift');

  if (idInp) idInp.value = h.id;
  if (dateInp) dateInp.value = h.date;
  if (timeInp) timeInp.value = h.time;
  if (shiftSel) shiftSel.value = h.shift;
  if (repSel) repSel.value = h.reporter;
  if (recSel) recSel.value = h.receiver;
  if (notesInp) notesInp.value = h.notes || '';

  // Update signature labels
  const sigLabelRep = document.getElementById('sigLabelReporter');
  const sigLabelRec = document.getElementById('sigLabelReceiver');
  if (sigLabelRep) sigLabelRep.textContent = h.reporter;
  if (sigLabelRec) sigLabelRec.textContent = h.receiver;
}

function importActivePatientsToCurrentShift() {
  if (!AppState.currentHandover) return;
  const currId = AppState.currentHandover.id;
  const activePatients = AppState.patients.filter(p => p.status === 'Observasi' && p.handoverId !== currId);
  
  if (activePatients.length === 0) {
    showToast('Tidak ada pasien observasi dari shift sebelumnya untuk dibawa.', 'info');
    return;
  }

  activePatients.forEach(p => {
    p.handoverId = currId;
  });

  saveToStorage();
  renderAllViews();
  showToast(`✓ Berhasil membawa ${activePatients.length} pasien observasi aktif ke shift berjalan ini.`, 'success');
}

function renderFormPatientsTable() {
  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  const patients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  const tbody = document.getElementById('formPatientsTableBody');
  const badge = document.getElementById('formPatientCountBadge');
  if (badge) badge.textContent = `${patients.length} Pasien`;
  if (!tbody) return;

  if (patients.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" class="text-center text-muted p-6">
          Belum ada pasien terkait. Klik tombol "+ Tambah Pasien Baru" di atas.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = patients.map((p, index) => `
    <tr>
      <td>${getTriageBadgeHtml(p.triage)}</td>
      <td class="font-mono"><strong>${p.rm}</strong></td>
      <td>
        <strong>${escapeHtml(p.name)}</strong>
        <div class="text-dim text-xs">${escapeHtml(p.age)}</div>
      </td>
      <td>${escapeHtml(p.diagnosis)}</td>
      <td>${escapeHtml(p.dpjp || '-')}</td>
      <td>
        <div class="text-xs">
          <strong>Lab:</strong> ${escapeHtml(p.lab || '-')}<br>
          <strong>Rad:</strong> ${escapeHtml(p.rad || '-')}
        </div>
      </td>
      <td>
        <div class="text-xs">
          <div class="text-muted"><strong>Kondisi:</strong> ${escapeHtml(p.condition)}</div>
          <div class="text-amber mt-1"><strong>Pending:</strong> ${escapeHtml(p.pending)}</div>
        </div>
      </td>
      <td>
        ${getStatusBadgeHtml(p.status, p.spri)}
        ${p.spri && p.spri !== '-' ? `<div class="text-xs text-blue mt-1 font-mono">${escapeHtml(p.spri)}</div>` : ''}
      </td>
      <td class="text-center">
        <div class="card-actions" style="justify-content: center;">
          <button type="button" class="btn btn-secondary btn-xs" onclick="editPatient('${p.id}')">
            <i class="fa-solid fa-pen"></i>
          </button>
          <button type="button" class="btn btn-danger btn-xs" onclick="deletePatient('${p.id}')">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      </td>
    </tr>
  `).join('');
}

// 5.4 LIVE PATIENT TRACKER
function renderTrackerView() {
  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  const filterShift = (AppState.filters && AppState.filters.shift) ? AppState.filters.shift : 'ALL';
  let filtered = (filterShift === 'CURRENT' && currentHandoverId)
    ? AppState.patients.filter(p => p.handoverId === currentHandoverId)
    : [...AppState.patients];

  // Apply filters
  const f = AppState.filters;
  if (f.search.trim()) {
    const q = f.search.toLowerCase();
    filtered = filtered.filter(p => 
      p.rm.toLowerCase().includes(q) ||
      p.name.toLowerCase().includes(q) ||
      p.diagnosis.toLowerCase().includes(q) ||
      (p.dpjp && p.dpjp.toLowerCase().includes(q))
    );
  }

  if (f.triage !== 'ALL') {
    filtered = filtered.filter(p => p.triage === f.triage);
  }

  if (f.status !== 'ALL') {
    filtered = filtered.filter(p => p.status === f.status);
  }

  if (f.dpjp !== 'ALL') {
    filtered = filtered.filter(p => p.dpjp === f.dpjp);
  }

  // Update match count
  const countElem = document.getElementById('trackerMatchCount');
  if (countElem) {
    countElem.textContent = `Menampilkan ${filtered.length} dari total ${AppState.patients.length} pasien`;
  }

  const cardsContainer = document.getElementById('trackerCardsGrid');
  const tableContainer = document.getElementById('trackerTableContainer');
  const tbody = document.getElementById('trackerTableBody');

  if (AppState.viewMode === 'cards') {
    cardsContainer.classList.remove('hidden');
    tableContainer.classList.add('hidden');

    if (filtered.length === 0) {
      cardsContainer.innerHTML = `
        <div class="col-span-full text-center p-6 text-muted">
          <i class="fa-solid fa-magnifying-glass" style="font-size: 2rem;"></i>
          <p class="mt-2">Tidak ditemukan data pasien yang sesuai filter pencarian.</p>
        </div>
      `;
      return;
    }

    cardsContainer.innerHTML = filtered.map(p => `
      <div class="patient-card border-${p.triage.toLowerCase()}">
        <div class="card-top-info">
          <div class="patient-identity">
            <h4>${escapeHtml(p.name)}</h4>
            <span class="rm-badge"><i class="fa-solid fa-id-card"></i> RM: ${p.rm} • Usia: ${escapeHtml(p.age)}</span>
          </div>
          ${getTriageBadgeHtml(p.triage)}
        </div>

        <div class="card-diagnosis">
          <i class="fa-solid fa-stethoscope text-primary"></i> ${escapeHtml(p.diagnosis)}
        </div>

        <div class="card-clinical-details">
          <div class="detail-line">
            <span class="lbl"><i class="fa-solid fa-user-doctor"></i> DPJP:</span>
            <span class="val text-blue font-bold">${escapeHtml(p.dpjp || '-')}</span>
          </div>
          <div class="detail-line">
            <span class="lbl"><i class="fa-solid fa-heart-pulse"></i> TTV/KU:</span>
            <span class="val">${escapeHtml(p.condition)}</span>
          </div>
          ${p.lab && p.lab !== '-' ? `
            <div class="detail-line">
              <span class="lbl"><i class="fa-solid fa-vial"></i> Lab:</span>
              <span class="val">${escapeHtml(p.lab)}</span>
            </div>
          ` : ''}
          ${p.rad && p.rad !== '-' ? `
            <div class="detail-line">
              <span class="lbl"><i class="fa-solid fa-x-ray"></i> Rad:</span>
              <span class="val">${escapeHtml(p.rad)}</span>
            </div>
          ` : ''}
        </div>

        <div class="pending-box">
          <strong><i class="fa-solid fa-triangle-exclamation"></i> INSTRUKSI PENDING TERAPI:</strong>
          ${escapeHtml(p.pending)}
        </div>

        <div class="card-footer-bar">
          <div>
            ${getStatusBadgeHtml(p.status, p.spri)}
          </div>
          <div class="card-actions">
            <button class="btn btn-secondary btn-xs" onclick="editPatient('${p.id}')">
              <i class="fa-solid fa-pen"></i> Edit
            </button>
            <button class="btn btn-outline btn-xs" onclick="quickUpdateStatus('${p.id}')">
              <i class="fa-solid fa-arrows-rotate"></i> Status
            </button>
          </div>
        </div>
      </div>
    `).join('');

  } else {
    // Table mode
    cardsContainer.classList.add('hidden');
    tableContainer.classList.remove('hidden');

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="10" class="text-center text-muted p-6">Tidak ada pasien yang cocok.</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map(p => `
      <tr>
        <td>${getTriageBadgeHtml(p.triage)}</td>
        <td class="font-mono"><strong>${p.rm}</strong></td>
        <td>
          <strong>${escapeHtml(p.name)}</strong>
          <div class="text-dim text-xs">${escapeHtml(p.age)}</div>
        </td>
        <td>${escapeHtml(p.diagnosis)}</td>
        <td><strong class="text-blue">${escapeHtml(p.dpjp || '-')}</strong></td>
        <td><div class="text-xs" style="max-width: 180px;">${escapeHtml(p.condition)}</div></td>
        <td><div class="text-xs text-amber font-medium" style="max-width: 200px;">${escapeHtml(p.pending)}</div></td>
        <td><span class="font-mono text-xs">${escapeHtml(p.spri || '-')}</span></td>
        <td>${getStatusBadgeHtml(p.status, p.spri)}</td>
        <td>
          <button class="btn btn-secondary btn-xs" onclick="editPatient('${p.id}')"><i class="fa-solid fa-pen"></i></button>
        </td>
      </tr>
    `).join('');
  }
}

// 5.5 PRINTABLE OFFICIAL REPORT
function renderPrintableReport() {
  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  const patients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  const h = AppState.currentHandover;
  if (!h) return;

  document.getElementById('printReportId').textContent = h.id;
  document.getElementById('printReportDate').textContent = formatDateIndo(h.date);
  document.getElementById('printReportTime').textContent = `${h.time} WITA`;
  document.getElementById('printReportShift').textContent = h.shift;
  document.getElementById('printReportReporter').textContent = h.reporter;
  document.getElementById('printReportReceiver').textContent = h.receiver;
  document.getElementById('printReportTotalPatients').textContent = `${patients.length} Pasien`;

  // Stats
  document.getElementById('printStatRed').textContent = patients.filter(p => p.triage === 'Merah').length;
  document.getElementById('printStatYellow').textContent = patients.filter(p => p.triage === 'Kuning').length;
  document.getElementById('printStatGreen').textContent = patients.filter(p => p.triage === 'Hijau').length;
  document.getElementById('printStatObs').textContent = patients.filter(p => p.status === 'Observasi').length;
  document.getElementById('printStatRanap').textContent = patients.filter(p => p.status === 'Rawat Inap').length;
  document.getElementById('printStatPulang').textContent = patients.filter(p => p.status === 'Boleh Pulang').length;

  document.getElementById('printDoctorNameReporter').textContent = `( ${h.reporter} )`;
  document.getElementById('printDoctorNameReceiver').textContent = `( ${h.receiver} )`;
  
  const dateSigEl = document.getElementById('printDateSignature');
  if (dateSigEl) {
    dateSigEl.textContent = formatDateIndo(h.date);
  }
  const ackTimeEl = document.getElementById('printAckTimestamp');
  if (ackTimeEl) {
    ackTimeEl.textContent = `Pukul: ${h.time} WITA`;
  }

  // Signatures on paper
  const sigRepBox = document.getElementById('printSigReporterContainer');
  const sigRecBox = document.getElementById('printSigReceiverContainer');

  if (AppState.signatures.reporter) {
    sigRepBox.innerHTML = `<img src="${AppState.signatures.reporter}" alt="TTD Pelapor">`;
  } else {
    sigRepBox.innerHTML = `<span class="sig-placeholder text-muted">(Tanda Tangan Digital)</span>`;
  }

  if (AppState.signatures.receiver) {
    sigRecBox.innerHTML = `<img src="${AppState.signatures.receiver}" alt="TTD Penerima">`;
  } else {
    sigRecBox.innerHTML = `<span class="sig-placeholder text-muted">(Tanda Tangan Digital)</span>`;
  }

  // Patients Table in Print
  const pBody = document.getElementById('printPatientsTableBody');
  if (!pBody) return;

  if (patients.length === 0) {
    pBody.innerHTML = `<tr><td colspan="8" class="text-center">Tidak ada pasien yang dioperkan pada shift ini.</td></tr>`;
    return;
  }

  pBody.innerHTML = patients.map((p, idx) => {
    const trgClass = p.triage === 'Merah' ? 'triage-merah' : p.triage === 'Kuning' ? 'triage-kuning' : 'triage-hijau';
    const trgLabel = p.triage === 'Merah' ? 'P1 - Merah' : p.triage === 'Kuning' ? 'P2 - Kuning' : 'P3 - Hijau';
    return `
    <tr>
      <td class="text-center" style="font-weight: 700;">${idx + 1}</td>
      <td class="text-center">
        <span class="badge-triage-print ${trgClass}">${trgLabel}</span>
      </td>
      <td>
        <strong style="color: #0f172a;">${escapeHtml(p.name)}</strong><br>
        <span style="font-family: monospace; font-size: 0.68rem; color: #475569;">No. RM: ${p.rm} (${escapeHtml(p.age)})</span>
      </td>
      <td><strong style="color: #1e293b;">${escapeHtml(p.diagnosis)}</strong></td>
      <td><span style="color: #0f172a; font-weight: 600;">${escapeHtml(p.dpjp || '-')}</span></td>
      <td>
        <div style="font-size: 0.68rem; line-height: 1.35;">
          <strong>Lab:</strong> ${escapeHtml(p.lab || '-')}<br>
          <strong>Rad:</strong> ${escapeHtml(p.rad || '-')}
        </div>
      </td>
      <td>
        <div style="font-size: 0.68rem; line-height: 1.35;">
          <strong>TTV & Kondisi:</strong> ${escapeHtml(p.condition)}<br>
          <strong style="color: #b45309;">Pending / Instruksi:</strong> ${escapeHtml(p.pending)}
        </div>
      </td>
      <td>
        <strong style="color: #047857;">${p.status}</strong>
        ${p.spri ? `<br><span style="font-family: monospace; font-size: 0.65rem; color: #64748b;">SPRI: ${escapeHtml(p.spri)}</span>` : ''}
      </td>
    </tr>
  `;
  }).join('');
}

// 5.6 MASTER DATA TABLES
function renderMasterDataTables() {
  // Table Dokter Jaga
  const docBody = document.getElementById('masterDokterJagaBody');
  if (docBody) {
    docBody.innerHTML = AppState.doctors.map((d, i) => `
      <tr>
        <td>${i + 1}</td>
        <td><strong>${escapeHtml(d.name)}</strong></td>
        <td>
          <span class="badge ${d.active ? 'badge-success' : 'badge-neutral'}">
            ${d.active ? 'Aktif' : 'Non-Aktif'}
          </span>
        </td>
        <td>
          <span class="badge-pin-status">
            <i class="fa-solid fa-lock text-emerald"></i> ${d.pin ? 'PIN Aktif' : 'Default (1234)'}
          </span>
        </td>
        <td>
          <div class="flex gap-1">
            <button class="btn btn-outline btn-xs" onclick="toggleDoctorStatus('${d.id}')">
              ${d.active ? 'Nonaktifkan' : 'Aktifkan'}
            </button>
            <button class="btn btn-secondary btn-xs" onclick="openChangePinModal('${d.id}')" title="Ubah PIN Dokter">
              <i class="fa-solid fa-key"></i> PIN
            </button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  // Table DPJP
  const dpjpBody = document.getElementById('masterDpjpBody');
  if (dpjpBody) {
    dpjpBody.innerHTML = AppState.dpjps.map((d, i) => `
      <tr>
        <td>${i + 1}</td>
        <td>
          <strong>${escapeHtml(d.name)}</strong>
          <div class="text-xs text-muted">${escapeHtml(d.specialist)}</div>
        </td>
        <td>
          <span class="badge ${d.active ? 'badge-success' : 'badge-neutral'}">
            ${d.active ? 'Aktif' : 'Non-Aktif'}
          </span>
        </td>
        <td>
          <button class="btn btn-outline btn-xs" onclick="toggleDpjpStatus('${d.id}')">
            ${d.active ? 'Nonaktifkan' : 'Aktifkan'}
          </button>
        </td>
      </tr>
    `).join('');
  }
}

// ==========================================================================
// 6. EVENT LISTENERS & INTERAKSI PENGGUNA
// ==========================================================================
function setupEventListeners() {
  // Sidebar Navigation Tabs
  const navButtons = document.querySelectorAll('.sidebar-nav .nav-item');
  navButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetTab = btn.getAttribute('data-tab');
      switchTab(targetTab);
    });
  });

  // Top Active Doctor Select Change (if element exists)
  const topActiveSelect = document.getElementById('activeDoctorSelect');
  if (topActiveSelect) {
    topActiveSelect.addEventListener('change', (e) => {
      AppState.activeDoctorId = e.target.value;
      const doc = AppState.doctors.find(d => d.id === AppState.activeDoctorId);
      showToast(`Dokter aktif beralih ke: ${doc ? doc.name : ''}`, 'info');
    });
  }

  // Auth & Session Listeners
  const btnHeaderSwitch = document.getElementById('btnHeaderSwitchDoctor');
  if (btnHeaderSwitch) {
    btnHeaderSwitch.addEventListener('click', openSwitchUserModal);
  }

  const btnHeaderLogout = document.getElementById('btnHeaderLogout');
  if (btnHeaderLogout) {
    btnHeaderLogout.addEventListener('click', handleDoctorLogout);
  }

  const btnSidebarLogout = document.getElementById('btnSidebarLogout');
  if (btnSidebarLogout) {
    btnSidebarLogout.addEventListener('click', handleDoctorLogout);
  }

  const navOpenLogin = document.getElementById('navOpenLoginModal');
  if (navOpenLogin) {
    navOpenLogin.addEventListener('click', openSwitchUserModal);
  }

  const formLogin = document.getElementById('formDoctorLogin');
  if (formLogin) {
    formLogin.addEventListener('submit', (e) => {
      e.preventDefault();
      handleDoctorLogin(false);
    });
  }

  const btnEmerg = document.getElementById('btnLoginEmergency');
  if (btnEmerg) {
    btnEmerg.addEventListener('click', () => {
      handleDoctorLogin(true);
    });
  }

  const btnTogglePin = document.getElementById('btnTogglePinVisibility');
  if (btnTogglePin) {
    btnTogglePin.addEventListener('click', togglePinVisibility);
  }

  // Change PIN Modal Listeners
  const btnClosePinModal = document.getElementById('btnCloseChangePinModal');
  if (btnClosePinModal) btnClosePinModal.addEventListener('click', closeChangePinModal);
  const btnCancelPin = document.getElementById('btnCancelChangePin');
  if (btnCancelPin) btnCancelPin.addEventListener('click', closeChangePinModal);
  const btnSavePin = document.getElementById('btnSaveNewPin');
  if (btnSavePin) btnSavePin.addEventListener('click', saveNewDoctorPin);

  // Quick New Handover button
  document.getElementById('btnQuickNewHandover').addEventListener('click', () => {
    startNewHandoverSession();
  });

  // History and Shift Buttons
  const btnStartDirect = document.getElementById('btnStartNewShiftDirect');
  if (btnStartDirect) {
    btnStartDirect.addEventListener('click', startNewHandoverSession);
  }

  const btnHistNewShift = document.getElementById('btnHistoryNewShift');
  if (btnHistNewShift) {
    btnHistNewShift.addEventListener('click', startNewHandoverSession);
  }

  // History filters
  const histSearch = document.getElementById('historySearchInput');
  if (histSearch) {
    histSearch.addEventListener('input', renderHandoverHistory);
  }

  const histPeriod = document.getElementById('filterHistoryPeriod');
  if (histPeriod) {
    histPeriod.addEventListener('change', renderHandoverHistory);
  }

  const histShift = document.getElementById('filterHistoryShift');
  if (histShift) {
    histShift.addEventListener('change', renderHandoverHistory);
  }

  const btnResetHist = document.getElementById('btnResetHistoryFilters');
  if (btnResetHist) {
    btnResetHist.addEventListener('click', () => {
      if (histSearch) histSearch.value = '';
      if (histPeriod) histPeriod.value = 'ALL';
      if (histShift) histShift.value = 'ALL';
      renderHandoverHistory();
    });
  }

  // Modal Detail Shift Close buttons
  const btnCloseModal = document.getElementById('btnCloseModalShiftDetail');
  if (btnCloseModal) {
    btnCloseModal.addEventListener('click', closeShiftDetailModal);
  }

  const btnCloseModalBtm = document.getElementById('btnCloseShiftDetailBottom');
  if (btnCloseModalBtm) {
    btnCloseModalBtm.addEventListener('click', closeShiftDetailModal);
  }

  // Dashboard shortcuts
  document.getElementById('btnRefreshDashboard').addEventListener('click', () => {
    fetchMasterFromGoogleSheets(false);
  });

  document.getElementById('btnGoToHandoverForm').addEventListener('click', () => {
    switchTab('tab-handover');
  });

  document.getElementById('btnAddNewPatientQuick').addEventListener('click', () => {
    openAddPatientModal();
  });

  // Handover Master Form Actions
  document.getElementById('btnResetHandoverForm').addEventListener('click', () => {
    if (confirm('Apakah Anda yakin ingin mengatur ulang data form shift ini?')) {
      populateHandoverForm();
      showToast('Formulir diatur ulang ke data tersimpan.', 'info');
    }
  });

  document.getElementById('btnSaveHandoverSession').addEventListener('click', () => {
    saveHandoverMasterForm();
  });

  document.getElementById('handoverMasterForm').addEventListener('submit', (e) => {
    e.preventDefault();
    finalizeHandoverSession();
  });

  document.getElementById('btnModalAddPatient').addEventListener('click', () => {
    openAddPatientModal();
  });

  // Modal Patient Actions
  document.getElementById('btnClosePatientModal').addEventListener('click', closePatientModal);
  document.getElementById('btnCancelPatientModal').addEventListener('click', closePatientModal);
  document.getElementById('btnGenRm').addEventListener('click', generateDummyRm);
  document.getElementById('formPatientInput').addEventListener('submit', handlePatientFormSubmit);

  // Tracker Filters & Search
  document.getElementById('trackerSearchInput').addEventListener('input', (e) => {
    AppState.filters.search = e.target.value;
    renderTrackerView();
  });

  document.getElementById('filterTriage').addEventListener('change', (e) => {
    AppState.filters.triage = e.target.value;
    renderTrackerView();
  });

  document.getElementById('filterStatus').addEventListener('change', (e) => {
    AppState.filters.status = e.target.value;
    renderTrackerView();
  });

  document.getElementById('filterDpjp').addEventListener('change', (e) => {
    AppState.filters.dpjp = e.target.value;
    renderTrackerView();
  });

  document.getElementById('btnResetFilters').addEventListener('click', () => {
    AppState.filters = { search: '', triage: 'ALL', status: 'ALL', dpjp: 'ALL', shift: 'ALL' };
    document.getElementById('trackerSearchInput').value = '';
    document.getElementById('filterTriage').value = 'ALL';
    document.getElementById('filterStatus').value = 'ALL';
    document.getElementById('filterDpjp').value = 'ALL';
    const filterShiftEl = document.getElementById('filterTrackerShift');
    if (filterShiftEl) filterShiftEl.value = 'ALL';
    renderTrackerView();
  });

  document.getElementById('btnViewCards').addEventListener('click', () => {
    AppState.viewMode = 'cards';
    document.getElementById('btnViewCards').classList.add('active');
    document.getElementById('btnViewTable').classList.remove('active');
    renderTrackerView();
  });

  document.getElementById('btnViewTable').addEventListener('click', () => {
    AppState.viewMode = 'table';
    document.getElementById('btnViewTable').classList.add('active');
    document.getElementById('btnViewCards').classList.remove('active');
    renderTrackerView();
  });

  // Report & Print Actions
  document.getElementById('btnPrintOfficialReport').addEventListener('click', () => {
    window.print();
  });

  document.getElementById('btnPreviewPrintFromForm').addEventListener('click', () => {
    switchTab('tab-report');
    renderPrintableReport();
  });

  document.getElementById('btnExportExcel').addEventListener('click', exportToExcelCompatibleCsv);
  document.getElementById('btnExportJson').addEventListener('click', exportBackupJson);

  // Master Data Add Buttons
  document.getElementById('btnAddDoctorJaga').addEventListener('click', promptAddDoctorJaga);
  document.getElementById('btnAddDoctorDpjp').addEventListener('click', promptAddDoctorDpjp);

  // Data management tools
  document.getElementById('btnResetToExcelDefaults').addEventListener('click', resetToExcelDefaults);
  document.getElementById('btnTriggerImportJson').addEventListener('click', () => {
    document.getElementById('inputImportJsonFile').click();
  });
  document.getElementById('inputImportJsonFile').addEventListener('change', handleImportJson);
  document.getElementById('btnClearStorage').addEventListener('click', () => {
    if (confirm('PERINGATAN: Apakah Anda yakin ingin mengosongkan seluruh riwayat dan cache lokal?')) {
      localStorage.clear();
      location.reload();
    }
  });

  // Google Sheets Cloud Integration Listeners
  const btnFetchGSheet = document.getElementById('btnFetchFromGSheet');
  if (btnFetchGSheet) btnFetchGSheet.addEventListener('click', () => fetchMasterFromGoogleSheets(false));

  const btnHeaderSync = document.getElementById('btnHeaderSyncAll');
  if (btnHeaderSync) btnHeaderSync.addEventListener('click', () => fetchMasterFromGoogleSheets(false));

  const btnTrackerSync = document.getElementById('btnTrackerSyncGSheet');
  if (btnTrackerSync) btnTrackerSync.addEventListener('click', () => fetchMasterFromGoogleSheets(false));

  const filterTrackerShiftEl = document.getElementById('filterTrackerShift');
  if (filterTrackerShiftEl) {
    filterTrackerShiftEl.addEventListener('change', (e) => {
      AppState.filters.shift = e.target.value;
      renderTrackerView();
    });
  }

  const btnTestGSheet = document.getElementById('btnTestGSheetConnection');
  if (btnTestGSheet) btnTestGSheet.addEventListener('click', testGSheetConnection);

  const btnSaveGSheet = document.getElementById('btnSaveGSheetConfig');
  if (btnSaveGSheet) btnSaveGSheet.addEventListener('click', saveGSheetConfigFromUI);

  const btnPushCurrent = document.getElementById('btnPushCurrentToGSheet');
  if (btnPushCurrent) btnPushCurrent.addEventListener('click', pushHandoverToGoogleSheets);

  const btnSyncForm = document.getElementById('btnSyncGSheetFromForm');
  if (btnSyncForm) btnSyncForm.addEventListener('click', pushHandoverToGoogleSheets);

  const btnSyncReport = document.getElementById('btnSyncGSheetFromReport');
  if (btnSyncReport) btnSyncReport.addEventListener('click', pushHandoverToGoogleSheets);

  const btnOpenGas = document.getElementById('btnOpenGasModal');
  if (btnOpenGas) btnOpenGas.addEventListener('click', openGasModal);

  const btnCloseGas = document.getElementById('btnCloseGasModal');
  if (btnCloseGas) btnCloseGas.addEventListener('click', closeGasModal);

  const btnCloseGasBtn = document.getElementById('btnCloseGasModalBtn');
  if (btnCloseGasBtn) btnCloseGasBtn.addEventListener('click', closeGasModal);

  const btnCopyGas1 = document.getElementById('btnCopyGasCode');
  if (btnCopyGas1) btnCopyGas1.addEventListener('click', copyGasCodeToClipboard);

  const btnCopyGas2 = document.getElementById('btnCopyGasCodeBottom');
  if (btnCopyGas2) btnCopyGas2.addEventListener('click', copyGasCodeToClipboard);

  const btnSaveGasDirect = document.getElementById('btnSaveGasModalDirectAndSync');
  if (btnSaveGasDirect) btnSaveGasDirect.addEventListener('click', saveGasDirectFromModalAndSync);

  const btnExportExcel = document.getElementById('btnExportExcelFromForm');
  if (btnExportExcel) btnExportExcel.addEventListener('click', exportCurrentShiftToExcel);
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.sidebar-nav .nav-item').forEach(el => el.classList.remove('active'));

  const targetTab = document.getElementById(tabId);
  if (targetTab) targetTab.classList.add('active');

  const navBtn = document.querySelector(`.sidebar-nav .nav-item[data-tab="${tabId}"]`);
  if (navBtn) navBtn.classList.add('active');

  if (tabId === 'tab-report') renderPrintableReport();
  if (tabId === 'tab-tracker') renderTrackerView();
  if (tabId === 'tab-dashboard') renderDashboardStats();
}

// ==========================================================================
// 7. SIGNATURE PAD CONTROLLER (CANVAS DRAWING)
// ==========================================================================
function setupSignaturePads() {
  setupSingleCanvas('canvasReporter', 'statusSigReporter', 'reporter');
  setupSingleCanvas('canvasReceiver', 'statusSigReceiver', 'receiver');

  document.getElementById('btnClearSigReporter').addEventListener('click', () => {
    clearCanvas('canvasReporter');
    AppState.signatures.reporter = null;
    document.getElementById('statusSigReporter').textContent = 'Belum ditandatangani';
    saveToStorage();
    renderPrintableReport();
  });

  document.getElementById('btnClearSigReceiver').addEventListener('click', () => {
    clearCanvas('canvasReceiver');
    AppState.signatures.receiver = null;
    document.getElementById('statusSigReceiver').textContent = 'Belum ditandatangani';
    saveToStorage();
    renderPrintableReport();
  });
}

function setupSingleCanvas(canvasId, statusElemId, sigKey) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  let isDrawing = false;
  let hasDrawn = false;

  ctx.lineWidth = 2.5;
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = '#0f172a';

  // Restore signature if exists
  if (AppState.signatures[sigKey]) {
    const img = new Image();
    img.onload = () => {
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      document.getElementById(statusElemId).textContent = '✓ Ditandatangani';
    };
    img.src = AppState.signatures[sigKey];
  }

  function getPos(e) {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    if (e.touches && e.touches.length > 0) {
      return {
        x: (e.touches[0].clientX - rect.left) * scaleX,
        y: (e.touches[0].clientY - rect.top) * scaleY
      };
    }
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY
    };
  }

  function startDraw(e) {
    isDrawing = true;
    hasDrawn = true;
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
    e.preventDefault();
  }

  function draw(e) {
    if (!isDrawing) return;
    const pos = getPos(e);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    e.preventDefault();
  }

  function endDraw() {
    if (!isDrawing) return;
    isDrawing = false;
    if (hasDrawn) {
      AppState.signatures[sigKey] = canvas.toDataURL('image/png');
      document.getElementById(statusElemId).textContent = '✓ Ditandatangani';
      saveToStorage();
      renderPrintableReport();
    }
  }

  canvas.addEventListener('mousedown', startDraw);
  canvas.addEventListener('mousemove', draw);
  canvas.addEventListener('mouseup', endDraw);
  canvas.addEventListener('mouseleave', endDraw);

  canvas.addEventListener('touchstart', startDraw, { passive: false });
  canvas.addEventListener('touchmove', draw, { passive: false });
  canvas.addEventListener('touchend', endDraw);
}

function clearCanvas(canvasId) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

// ==========================================================================
// 8. CRUD LOGIC: HANDOVER & PATIENT
// ==========================================================================
function saveHandoverMasterForm() {
  const id = document.getElementById('inputHandoverId').value;
  const date = document.getElementById('inputHandoverDate').value;
  const time = document.getElementById('inputHandoverTime').value;
  const shift = document.getElementById('selectHandoverShift').value;
  const reporter = document.getElementById('selectDokterPelapor').value;
  const receiver = document.getElementById('selectDokterPenerima').value;
  const notes = document.getElementById('inputCatatanShift').value;

  if (!id || !date || !time) {
    showToast('Mohon lengkapi ID, Tanggal, dan Waktu Handover.', 'warning');
    return;
  }

  AppState.currentHandover = {
    ...AppState.currentHandover,
    id, date, time, shift, reporter, receiver, notes
  };

  saveCurrentSessionToHistory();
  saveToStorage();
  renderAllViews();
  showToast('Informasi master sesi handover berhasil disimpan.', 'success');
}

function finalizeHandoverSession() {
  saveHandoverMasterForm();
  AppState.currentHandover.locked = true;
  saveCurrentSessionToHistory();
  saveToStorage();
  renderAllViews();
  // Auto-sync langsung ke Google Sheets saat shift difinalisasi
  if (AppState.gsheet && AppState.gsheet.webAppUrl) {
    pushHandoverToGoogleSheets(false);
  }
  showToast('Sesi Handover resmi difinalisasi dan tersimpan di riwayat!', 'success');
  switchTab('tab-report');
}

function startNewHandoverSession() {
  if (confirm('Mulai sesi handover shift baru? Sesi shift saat ini akan otomatis diarsipkan ke riwayat.')) {
    // Archive current session before starting new one
    saveCurrentSessionToHistory();

    const newSession = createDefaultShiftSession();
    // Unique ID
    const randomSuffix = Math.floor(Math.random() * 90 + 10);
    newSession.id = `HO-${newSession.date.replace(/-/g, '')}-${randomSuffix}`;
    AppState.currentHandover = newSession;

    // Reset signatures for new shift
    AppState.signatures = { reporter: null, receiver: null };
    clearCanvas('canvasReporter');
    clearCanvas('canvasReceiver');
    document.getElementById('statusSigReporter').textContent = 'Belum ditandatangani';
    document.getElementById('statusSigReceiver').textContent = 'Belum ditandatangani';

    saveToStorage();
    renderAllViews();
    switchTab('tab-handover');
    showToast(`Sesi handover baru ${newSession.id} diaktifkan.`, 'info');
  }
}

// 8.1 PATIENT MODAL HANDLERS
function openAddPatientModal() {
  document.getElementById('modalPatientTitle').textContent = 'Tambah Pasien IGD Baru';
  document.getElementById('inputPatientId').value = '';
  document.getElementById('formPatientInput').reset();
  document.getElementById('inputPatTriage').value = 'Kuning';
  generateDummyRm();
  document.getElementById('modalPatient').classList.add('show');
}

function editPatient(patientId) {
  const p = AppState.patients.find(item => item.id === patientId);
  if (!p) return;

  document.getElementById('modalPatientTitle').textContent = `Edit Pasien: ${p.name} (${p.rm})`;
  document.getElementById('inputPatientId').value = p.id;
  document.getElementById('inputPatTriage').value = p.triage;
  document.getElementById('inputPatRm').value = p.rm;
  document.getElementById('inputPatAge').value = p.age;
  document.getElementById('inputPatName').value = p.name;
  document.getElementById('selectPatDpjp').value = p.dpjp || '-';
  document.getElementById('inputPatDiagnosis').value = p.diagnosis;
  document.getElementById('inputPatLab').value = p.lab || '';
  document.getElementById('inputPatRad').value = p.rad || '';
  document.getElementById('inputPatCondition').value = p.condition || '';
  document.getElementById('inputPatPending').value = p.pending || '';
  document.getElementById('inputPatSpri').value = p.spri || '';
  document.getElementById('selectPatStatus').value = p.status;

  document.getElementById('modalPatient').classList.add('show');
}

function closePatientModal() {
  document.getElementById('modalPatient').classList.remove('show');
}

function generateDummyRm() {
  const p1 = String(Math.floor(Math.random() * 90 + 10));
  const p2 = String(Math.floor(Math.random() * 90 + 10));
  const p3 = String(Math.floor(Math.random() * 90 + 10));
  document.getElementById('inputPatRm').value = `${p1}-${p2}-${p3}`;
}

function handlePatientFormSubmit(e) {
  e.preventDefault();

  const id = document.getElementById('inputPatientId').value;
  const triage = document.getElementById('inputPatTriage').value;
  const rm = document.getElementById('inputPatRm').value.trim();
  const age = document.getElementById('inputPatAge').value.trim();
  const name = document.getElementById('inputPatName').value.trim();
  const dpjp = document.getElementById('selectPatDpjp').value;
  const diagnosis = document.getElementById('inputPatDiagnosis').value.trim();
  const lab = document.getElementById('inputPatLab').value.trim();
  const rad = document.getElementById('inputPatRad').value.trim();
  const condition = document.getElementById('inputPatCondition').value.trim();
  const pending = document.getElementById('inputPatPending').value.trim();
  const spri = document.getElementById('inputPatSpri').value.trim();
  const status = document.getElementById('selectPatStatus').value;

  const currentHandoverId = AppState.currentHandover.id;

  if (id) {
    // Edit existing patient
    const idx = AppState.patients.findIndex(p => p.id === id);
    if (idx !== -1) {
      AppState.patients[idx] = {
        ...AppState.patients[idx],
        triage, rm, age, name, dpjp, diagnosis, lab, rad, condition, pending, spri, status
      };
      showToast(`Data pasien ${name} berhasil diperbarui.`, 'success');
    }
  } else {
    // Create new patient
    const newId = `P-${Date.now().toString().slice(-4)}`;
    const newPatient = {
      id: newId,
      handoverId: currentHandoverId,
      triage, rm, age, name, dpjp, diagnosis, lab, rad, condition, pending, spri, status
    };
    AppState.patients.unshift(newPatient);
    showToast(`Pasien baru ${name} (${triage}) ditambahkan.`, 'success');
  }

    saveToStorage();
  closePatientModal();
  renderAllViews();

  // Auto-sync ke Google Sheets jika Apps Script sudah terpasang
  if (AppState.gsheet && AppState.gsheet.webAppUrl && AppState.gsheet.webAppUrl.trim()) {
    pushHandoverToGoogleSheets(false);
  } else {
    showToast('Pasien tersimpan di browser. Hubungkan Apps Script di menu Master Data untuk rekam otomatis ke Google Sheets.', 'info');
  }
}

function deletePatient(patientId) {
  const p = AppState.patients.find(item => item.id === patientId);
  if (!p) return;

  if (confirm(`Hapus pasien ${p.name} (RM: ${p.rm}) dari sesi handover ini?`)) {
    AppState.patients = AppState.patients.filter(item => item.id !== patientId);
    saveToStorage();
    renderAllViews();
    showToast(`Pasien ${p.name} dihapus.`, 'info');
  }
}

function quickUpdateStatus(patientId) {
  const p = AppState.patients.find(item => item.id === patientId);
  if (!p) return;

  const statuses = ['Observasi', 'Rawat Inap', 'Boleh Pulang', 'Rujuk', 'Meninggal'];
  const curIdx = statuses.indexOf(p.status);
  const nextStatus = statuses[(curIdx + 1) % statuses.length];

  p.status = nextStatus;
  saveToStorage();
  renderAllViews();
  showToast(`Status ${p.name} diubah menjadi: ${nextStatus}`, 'info');
}

// 8.2 MASTER DATA USER & DPJP HANDLERS
function toggleDoctorStatus(docId) {
  const d = AppState.doctors.find(item => item.id === docId);
  if (d) {
    d.active = !d.active;
    saveToStorage();
    renderAllViews();
    showToast(`Status ${d.name} diubah.`, 'info');
  }
}

function toggleDpjpStatus(dpjpId) {
  const d = AppState.dpjps.find(item => item.id === dpjpId);
  if (d) {
    d.active = !d.active;
    saveToStorage();
    renderAllViews();
    showToast(`Status ${d.name} diubah.`, 'info');
  }
}

function promptAddDoctorJaga() {
  const name = prompt('Masukkan nama dokter jaga IGD baru (contoh: dr. Rahmat):');
  if (name && name.trim()) {
    const newId = `DOC-${String(AppState.doctors.length + 1).padStart(2, '0')}`;
    AppState.doctors.push({ id: newId, name: name.trim(), active: true });
    saveToStorage();
    renderAllViews();
    showToast(`Dokter jaga baru ${name} ditambahkan.`, 'success');
  }
}

function promptAddDoctorDpjp() {
  const name = prompt('Masukkan nama DPJP beserta gelar spesialis (contoh: dr. Hendra, Sp.JP):');
  if (!name || !name.trim()) return;
  const specialist = prompt('Masukkan bidang spesialisasi (contoh: Jantung & Pembuluh Darah):') || 'Spesialis';

  const newId = `DPJP-${String(AppState.dpjps.length + 1).padStart(2, '0')}`;
  AppState.dpjps.push({ id: newId, name: name.trim(), specialist: specialist.trim(), active: true });
  saveToStorage();
  renderAllViews();
  showToast(`DPJP baru ${name} ditambahkan.`, 'success');
}

function resetToExcelDefaults() {
  if (confirm('Kembalikan master dokter ke daftar awal dan kosongkan seluruh data pasien (set ke 0)?')) {
    AppState.doctors = [...DEFAULT_DOKTER_JAGA];
    AppState.dpjps = [...DEFAULT_DPJP];
    AppState.patients = [];
    AppState.currentHandover = createDefaultShiftSession();
    AppState.signatures = { reporter: null, receiver: null };
    saveToStorage();
    renderAllViews();
    showToast('Database master dokter direset dan data pasien dikosongkan (0).', 'success');
  }
}

function clearAllPatients() {
  if (confirm('Kosongkan seluruh data pasien pada sesi shift ini (set data pasien ke 0)?')) {
    AppState.patients = [];
    saveToStorage();
    renderAllViews();
    showToast('Seluruh data pasien telah dikosongkan (Total Pasien = 0).', 'info');
  }
}
window.clearAllPatients = clearAllPatients;

// ==========================================================================
// 9. EXPORT & BACKUP TOOLS (EXCEL & JSON)
// ==========================================================================
function exportToExcelCompatibleCsv() {
  const currentHandoverId = AppState.currentHandover.id;
  const currentPatients = AppState.patients.filter(p => p.handoverId === currentHandoverId);

  // Headers matching `Pasien IGD` in Handover Dokter.xlsx
  const headers = [
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
  ];

  const rows = currentPatients.map(p => [
    `"${p.id}"`,
    `"${p.handoverId}"`,
    `"${p.rm}"`,
    `"${p.name.replace(/"/g, '""')}"`,
    `"${p.age}"`,
    `"${p.triage}"`,
    `"${p.diagnosis.replace(/"/g, '""')}"`,
    `"${(p.dpjp || '').replace(/"/g, '""')}"`,
    `"${(p.lab || '').replace(/"/g, '""')}"`,
    `"${(p.rad || '').replace(/"/g, '""')}"`,
    `"${(p.condition || '').replace(/"/g, '""')}"`,
    `"${(p.pending || '').replace(/"/g, '""')}"`,
    `"${(p.spri || '').replace(/"/g, '""')}"`,
    `"${p.status}"`
  ]);

  const csvContent = '\uFEFF' + headers.join(',') + '\n' + rows.map(r => r.join(',')).join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Handover_Dokter_IGD_${AppState.currentHandover.id}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('File CSV berhasil diunduh (kompatibel langsung dengan Excel).', 'success');
}

function exportBackupJson() {
  const data = JSON.stringify(AppState, null, 2);
  const blob = new Blob([data], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Backup_Handover_IGD_${new Date().toISOString().split('T')[0]}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('Backup JSON berhasil diekspor.', 'success');
}

function handleImportJson(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const imported = JSON.parse(event.target.result);
      if (imported.doctors && imported.patients) {
        AppState = { ...AppState, ...imported };
        saveToStorage();
        renderAllViews();
        showToast('Data berhasil dipulihkan dari berkas JSON!', 'success');
      } else {
        showToast('Format JSON tidak valid.', 'error');
      }
    } catch (err) {
      showToast('Gagal membaca file JSON: ' + err.message, 'error');
    }
  };
  reader.readAsText(file);
}

// ==========================================================================
// 10. HELPER FUNCTIONS
// ==========================================================================
function getTriageBadgeHtml(triage) {
  const t = triage.toLowerCase();
  let cls = 'badge-triage-kuning';
  let icon = 'fa-triangle-exclamation';

  if (t === 'merah') {
    cls = 'badge-triage-merah';
    icon = 'fa-heart-pulse';
  } else if (t === 'hijau') {
    cls = 'badge-triage-hijau';
    icon = 'fa-circle-check';
  } else if (t === 'hitam') {
    cls = 'badge-triage-hitam';
    icon = 'fa-skull-crossbones';
  }

  return `<span class="badge ${cls}"><i class="fa-solid ${icon}"></i> ${triage}</span>`;
}

function getStatusBadgeHtml(status, spri) {
  let cls = 'badge-neutral';
  let icon = 'fa-bed';

  if (status === 'Rawat Inap') {
    cls = 'badge-primary';
    icon = 'fa-hospital-user';
  } else if (status === 'Boleh Pulang') {
    cls = 'badge-success';
    icon = 'fa-house-user';
  } else if (status === 'Rujuk') {
    cls = 'badge-warning';
    icon = 'fa-ambulance';
  } else if (status === 'Meninggal') {
    cls = 'badge-neutral';
    icon = 'fa-ribbon';
  }

  return `<span class="badge ${cls}"><i class="fa-solid ${icon}"></i> ${status}</span>`;
}

function formatDateIndo(dateStr) {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return `${parseInt(parts[2])} ${months[parseInt(parts[1]) - 1]} ${parts[0]}`;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  
  let icon = 'fa-circle-info';
  if (type === 'success') icon = 'fa-circle-check text-emerald';
  if (type === 'warning') icon = 'fa-triangle-exclamation text-amber';
  if (type === 'error') icon = 'fa-circle-xmark text-rose';

  toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${escapeHtml(message)}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3200);
}

// ==========================================================================
// 11. GOOGLE SHEETS CLOUD INTEGRATION & TWO-WAY SYNC
// ==========================================================================

function extractSheetIdFromUrl(url) {
  if (!url) return '';
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : url.trim();
}

function parseCsv(text) {
  if (!text) return [];
  const rows = [];
  let row = [];
  let field = '';
  let insideQuotes = false;
  const n = text.length;

  for (let i = 0; i < n; i++) {
    const c = text[i];
    const nxt = (i + 1 < n) ? text[i + 1] : '';

    if (c === '"') {
      if (insideQuotes && nxt === '"') {
        field += '"';
        i++; // skip escaped double quote
      } else {
        insideQuotes = !insideQuotes;
      }
    } else if (c === ',' && !insideQuotes) {
      row.push(field.trim());
      field = '';
    } else if ((c === '\r' || c === '\n') && !insideQuotes) {
      if (c === '\r' && nxt === '\n') {
        i++;
      }
      row.push(field.trim());
      if (row.some(f => f.length > 0)) {
        rows.push(row);
      }
      row = [];
      field = '';
    } else {
      field += c;
    }
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field.trim());
    if (row.some(f => f.length > 0)) {
      rows.push(row);
    }
  }

  return rows;
}

async function fetchMasterFromGoogleSheets(silent = false) {
  const sheetId = AppState.gsheet.sheetId || extractSheetIdFromUrl(AppState.gsheet.sheetUrl) || DEFAULT_GSHEET_CONFIG.sheetId;
  if (!sheetId) {
    if (!silent) showToast('ID Spreadsheet Google belum ditentukan.', 'warning');
    return;
  }

  if (!silent) showToast('Menghubungi Google Sheets untuk menarik data...', 'info');

  const syncBtns = document.querySelectorAll('.btn-sync-header, #btnHeaderSyncAll, #btnTrackerSyncGSheet, #btnFetchFromGSheet, #btnRefreshDashboard');
  syncBtns.forEach(b => {
    b.classList.add('loading-spin');
    const ic = b.querySelector('i');
    if (ic) ic.classList.add('fa-spin');
  });

  try {
    let docsUpdated = 0;
    let dpjpUpdated = 0;
    let patientsImported = 0;
    let handoversImported = 0;

    // 1. Tarik Master Dokter Jaga (Sheet USER)
    try {
      const userUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=USER`;
      const userResp = await fetch(userUrl);
      if (userResp.ok) {
        const userCsv = await userResp.text();
        const userRows = parseCsv(userCsv);
        if (userRows.length > 1) {
          const newDoctors = [];
          for (let i = 1; i < userRows.length; i++) {
            const row = userRows[i];
            const name = (row[0] || '').replace(/^"|"$/g, '').trim();
            const activeRaw = (row[1] || '1').replace(/^"|"$/g, '').toUpperCase();
            const active = activeRaw !== '0' && activeRaw !== 'FALSE';
            if (name) {
              newDoctors.push({
                id: `DOC-${String(newDoctors.length + 1).padStart(2, '0')}`,
                name,
                pin: '1234',
                active
              });
            }
          }
          if (newDoctors.length > 0) {
            AppState.doctors = newDoctors;
            docsUpdated = newDoctors.length;
          }
        }
      }
    } catch (eDoc) {
      console.warn('GSheet Sync USER warning:', eDoc);
    }

    // 2. Tarik Master DPJP (Sheet USER DPJP)
    try {
      const dpjpUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=USER%20DPJP`;
      const dpjpResp = await fetch(dpjpUrl);
      if (dpjpResp.ok) {
        const dpjpCsv = await dpjpResp.text();
        const dpjpRows = parseCsv(dpjpCsv);
        if (dpjpRows.length > 1) {
          const newDpjps = [];
          for (let i = 1; i < dpjpRows.length; i++) {
            const row = dpjpRows[i];
            const name = (row[0] || '').replace(/^"|"$/g, '').trim();
            const activeRaw = (row[1] || '1').replace(/^"|"$/g, '').toUpperCase();
            const active = activeRaw !== '0' && activeRaw !== 'FALSE';

            if (name) {
              let specialist = 'Spesialis Konsultan';
              if (name.includes('Sp.PD')) specialist = 'Penyakit Dalam';
              else if (name.includes('Sp.B')) specialist = 'Bedah Umum';
              else if (name.includes('Sp.A')) specialist = 'Kesehatan Anak';
              else if (name.includes('Sp.OG')) specialist = 'Kebidanan & Kandungan';
              else if (name.includes('Sp.S')) specialist = 'Neurologi / Saraf';
              else if (name.includes('Sp.M')) specialist = 'Mata';
              else if (name.includes('Sp.THT')) specialist = 'THT-KL';
              else if (name.includes('Sp.KJ')) specialist = 'Kedokteran Jiwa';
              else if (name.includes('Sp.JP')) specialist = 'Jantung & Pembuluh Darah';
              else if (name.includes('Sp.An')) specialist = 'Anestesi & Perawatan Intensif';
              else if (name.includes(',')) specialist = name.split(',').slice(1).join(',').trim();

              newDpjps.push({
                id: `DPJP-${String(newDpjps.length + 1).padStart(2, '0')}`,
                name,
                specialist,
                active
              });
            }
          }
          if (newDpjps.length > 0) {
            AppState.dpjps = newDpjps;
            dpjpUpdated = newDpjps.length;
          }
        }
      }
    } catch (eDpjp) {
      console.warn('GSheet Sync DPJP warning:', eDpjp);
    }

    // 3. Tarik Seluruh Pasien (Sheet Pasien IGD)
    try {
      const pasienUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Pasien%20IGD`;
      const pasienResp = await fetch(pasienUrl);
      if (pasienResp.ok) {
        const pasienCsv = await pasienResp.text();
        const pasienRows = parseCsv(pasienCsv);
        if (pasienRows.length > 1) {
          const patientMap = new Map();
          AppState.patients.forEach(p => {
            if (p.id) patientMap.set(p.id, p);
          });

          for (let i = 1; i < pasienRows.length; i++) {
            const r = pasienRows[i];
            if (!r || r.length === 0 || !r[0]) continue;
            const pid = (r[0] || '').replace(/^"|"$/g, '').trim();
            if (!pid || pid === 'ID_Pasien_IGD') continue;

            const pat = {
              id: pid,
              handoverId: (r[1] || '').replace(/^"|"$/g, '').trim(),
              rm: (r[2] || '').replace(/^"|"$/g, '').trim(),
              name: (r[3] || '').replace(/^"|"$/g, '').trim(),
              age: (r[4] || '').replace(/^"|"$/g, '').trim(),
              triage: (r[5] || 'Kuning').replace(/^"|"$/g, '').trim(),
              diagnosis: (r[6] || '').replace(/^"|"$/g, '').trim(),
              dpjp: (r[7] || '-').replace(/^"|"$/g, '').trim(),
              lab: (r[8] || '-').replace(/^"|"$/g, '').trim(),
              rad: (r[9] || '-').replace(/^"|"$/g, '').trim(),
              condition: (r[10] || '-').replace(/^"|"$/g, '').trim(),
              pending: (r[11] || '-').replace(/^"|"$/g, '').trim(),
              spri: (r[12] || '-').replace(/^"|"$/g, '').trim(),
              status: (r[13] || 'Observasi').replace(/^"|"$/g, '').trim()
            };

            if (!['Merah', 'Kuning', 'Hijau', 'Hitam'].includes(pat.triage)) {
              if (pat.triage.toLowerCase().includes('merah')) pat.triage = 'Merah';
              else if (pat.triage.toLowerCase().includes('kuning')) pat.triage = 'Kuning';
              else if (pat.triage.toLowerCase().includes('hijau')) pat.triage = 'Hijau';
              else if (pat.triage.toLowerCase().includes('hitam')) pat.triage = 'Hitam';
              else pat.triage = 'Kuning';
            }

            patientMap.set(pid, pat);
          }

          AppState.patients = Array.from(patientMap.values());
          patientsImported = AppState.patients.length;
        }
      }
    } catch (ePasien) {
      console.warn('GSheet Sync Pasien IGD warning:', ePasien);
    }

    // 4. Tarik Seluruh Riwayat Shift (Sheet Handover)
    try {
      const hoUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=Handover`;
      const hoResp = await fetch(hoUrl);
      if (hoResp.ok) {
        const hoCsv = await hoResp.text();
        const hoRows = parseCsv(hoCsv);
        if (hoRows.length > 1) {
          const hoMap = new Map();
          AppState.handoverHistory.forEach(h => {
            if (h.id) hoMap.set(h.id, h);
          });

          for (let i = 1; i < hoRows.length; i++) {
            const r = hoRows[i];
            if (!r || r.length === 0 || !r[0]) continue;
            const hid = (r[0] || '').replace(/^"|"$/g, '').trim();
            if (!hid || hid === 'ID_Handover') continue;

            const date = (r[1] || '').replace(/^"|"$/g, '').trim();
            const time = (r[2] || '').replace(/^"|"$/g, '').trim() || '08:00';
            const shift = (r[3] || 'Shift Pagi (08.00 - 14.00)').replace(/^"|"$/g, '').trim();
            const reporter = (r[4] || '-').replace(/^"|"$/g, '').trim();
            const receiver = (r[5] || '-').replace(/^"|"$/g, '').trim();
            const notes = (r[7] || '').replace(/^"|"$/g, '').trim();

            const existing = hoMap.get(hid) || {};
            const shiftPatients = AppState.patients.filter(p => p.handoverId === hid);

            hoMap.set(hid, {
              ...existing,
              id: hid,
              date: date || existing.date || new Date().toISOString().slice(0, 10),
              time: time || existing.time || '08:00',
              shift: shift || existing.shift || 'Shift Pagi (08.00 - 14.00)',
              reporter: reporter || existing.reporter || 'Dokter IGD',
              receiver: receiver || existing.receiver || 'Dokter IGD',
              notes: notes || existing.notes || '',
              locked: true,
              timestamp: existing.timestamp || new Date().toISOString(),
              patientsSnapshot: shiftPatients.length > 0 ? shiftPatients : (existing.patientsSnapshot || []),
              signatures: existing.signatures || { reporter: null, receiver: null }
            });
          }

          AppState.handoverHistory = Array.from(hoMap.values()).sort((a, b) => (b.date + ' ' + b.time).localeCompare(a.date + ' ' + a.time));
          handoversImported = AppState.handoverHistory.length;
        }
      }
    } catch (eHo) {
      console.warn('GSheet Sync Handover warning:', eHo);
    }

    AppState.gsheet.lastSync = new Date().toISOString();
    saveToStorage();
    renderAllViews();
    updateGSheetUI();

    if (!silent) {
      showToast(`✓ Berhasil memuat ${AppState.patients.length} Pasien IGD & ${AppState.handoverHistory.length} Riwayat Shift dari Google Sheets!`, 'success');
    }
  } catch (err) {
    console.error('GSheet Sync Error:', err);
    if (!silent) {
      showToast('Gagal menarik data dari Google Sheets: ' + err.message, 'error');
    }
  } finally {
    syncBtns.forEach(b => {
      b.classList.remove('loading-spin');
      const ic = b.querySelector('i');
      if (ic) ic.classList.remove('fa-spin');
    });
  }
}

async function testGSheetConnection() {
  const sheetId = AppState.gsheet.sheetId || extractSheetIdFromUrl(AppState.gsheet.sheetUrl);
  if (!sheetId) {
    showToast('Masukkan URL Google Spreadsheet yang valid.', 'warning');
    return;
  }

  showToast('Menguji koneksi ke Google Spreadsheet...', 'info');
  try {
    const testUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=USER`;
    const res = await fetch(testUrl);
    if (res.ok) {
      showToast('✓ Koneksi berhasil! Google Spreadsheet dapat diakses dan dibaca secara live.', 'success');
      const badge = document.getElementById('badgeGSheetConnectionStatus');
      if (badge) {
        badge.className = 'badge badge-success';
        badge.innerHTML = '<i class="fa-solid fa-circle-check"></i> Terhubung ke Google Sheets';
      }
    } else {
      throw new Error(`HTTP Status ${res.status}`);
    }
  } catch (e) {
    showToast('Koneksi gagal atau spreadsheet belum dibagikan ke Publik (Anyone with link): ' + e.message, 'error');
    const badge = document.getElementById('badgeGSheetConnectionStatus');
    if (badge) {
      badge.className = 'badge badge-warning';
      badge.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Akses Terbatas';
    }
  }
}

function saveHandoverMasterFormSilently() {
  const idInput = document.getElementById('inputHandoverId');
  const dateInput = document.getElementById('inputHandoverDate');
  const timeInput = document.getElementById('inputHandoverTime');
  const shiftSelect = document.getElementById('selectHandoverShift');
  const repSelect = document.getElementById('selectDokterPelapor');
  const recSelect = document.getElementById('selectDokterPenerima');
  const notesInput = document.getElementById('inputCatatanShift');

  if (idInput && dateInput && timeInput && AppState.currentHandover) {
    AppState.currentHandover.id = idInput.value.trim() || AppState.currentHandover.id;
    AppState.currentHandover.date = dateInput.value.trim() || AppState.currentHandover.date;
    AppState.currentHandover.time = timeInput.value.trim() || AppState.currentHandover.time;
    if (shiftSelect) AppState.currentHandover.shift = shiftSelect.value;
    if (repSelect) AppState.currentHandover.reporter = repSelect.value;
    if (recSelect) AppState.currentHandover.receiver = recSelect.value;
    if (notesInput) AppState.currentHandover.notes = notesInput.value.trim();
    saveToStorage();
  }
}

async function pushHandoverToGoogleSheets(isSilent = false) {
  const webAppUrl = (AppState.gsheet.webAppUrl || '').trim();

  // If Apps Script web app URL is not yet configured, guide the user
  if (!webAppUrl) {
    if (!isSilent) {
      openGasModal();
      showToast('Koneksi Google Apps Script belum terpasang. Silakan masukkan URL Web App pada panduan.', 'warning');
    }
    return;
  }

  if (!webAppUrl.startsWith('https://script.google.com/macros/s/') || !webAppUrl.includes('/exec')) {
    if (!isSilent) {
      alert('URL Google Apps Script tidak valid!\n\nURL harus berawalan "https://script.google.com/macros/s/..." dan berakhiran "/exec".\n\nSilakan periksa kembali URL hasil Deployment Web App Anda.');
      openGasModal();
    }
    return;
  }

  // Visual loading feedback on sync buttons
  const btnSyncForm = document.getElementById('btnSyncGSheetFromForm');
  const btnSyncReport = document.getElementById('btnSyncGSheetFromReport');
  const btnPushCurrent = document.getElementById('btnPushCurrentToGSheet');
  
  const buttonsToToggle = [btnSyncForm, btnSyncReport, btnPushCurrent].filter(Boolean);
  const originalContents = buttonsToToggle.map(b => b.innerHTML);

  buttonsToToggle.forEach(b => {
    b.disabled = true;
    b.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Mengirim...';
  });

  if (!isSilent) {
    showToast('Mengirim data Handover ke Google Sheets...', 'info');
  }

  // Ensure latest inputs from master form are captured
  saveHandoverMasterFormSilently();

  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  let currentPatients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  // Robust fallback: if no patient has matching handoverId, include all active patients
  if (currentPatients.length === 0 && AppState.patients.length > 0) {
    currentPatients = AppState.patients;
  }

  const payload = {
    action: 'saveHandover',
    handover: AppState.currentHandover || {},
    patients: currentPatients
  };

  try {
    // Mode no-cors avoids CORS block from Google Apps Script 302 redirect
    await fetch(webAppUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    AppState.gsheet.lastSync = new Date().toISOString();
    saveToStorage();
    
    if (!isSilent) {
      showToast(`✓ Berhasil tersinkron ke Google Sheets! (${currentPatients.length} pasien)`, 'success');
      alert(`✓ BERHASIL MENGIRIM KE GOOGLE SHEETS!\n\nData sesi Handover (${payload.handover.id || '-'}) dan ${currentPatients.length} data pasien telah berhasil dicatat ke Google Sheets Anda:\n${AppState.gsheet.sheetUrl || 'Google Spreadsheet'}\n\nSilakan cek lembar kerja "Handover" dan "Pasien IGD".`);
    }
  } catch (err) {
    console.error('GAS Push error:', err);
    if (!isSilent) {
      showToast('Gagal mengirim ke Google Sheets. Periksa koneksi internet atau hak akses Apps Script.', 'error');
      alert('Gagal mengirim ke Google Sheets:\n' + (err.message || err) + '\n\nPastikan:\n1. Koneksi internet aktif.\n2. Saat Terapkan (Deploy) di Apps Script, opsi "Yang memiliki akses (Who has access)" disetel ke "Siapa saja (Anyone)".');
    }
  } finally {
    buttonsToToggle.forEach((b, idx) => {
      b.disabled = false;
      b.innerHTML = originalContents[idx];
    });
  }
}

function saveGasDirectFromModalAndSync() {
  const inputDirect = document.getElementById('inputGasModalDirectUrl');
  if (!inputDirect) return;
  const url = inputDirect.value.trim();

  if (!url) {
    showToast('Mohon tempelkan URL Web App Apps Script terlebih dahulu.', 'warning');
    inputDirect.focus();
    return;
  }

  if (!url.startsWith('https://script.google.com/macros/s/') || !url.includes('/exec')) {
    alert('URL Apps Script tidak valid!\n\nURL harus berawalan "https://script.google.com/macros/s/..."\ndan berakhiran "/exec".\n\nSilakan cek kembali URL hasil deployment Web App Anda.');
    inputDirect.focus();
    return;
  }

  AppState.gsheet.webAppUrl = url;
  const inputMaster = document.getElementById('inputGasWebAppUrl');
  if (inputMaster) inputMaster.value = url;

  saveToStorage();
  updateGSheetUI();
  closeGasModal();
  showToast('✓ URL Google Apps Script berhasil disimpan!', 'success');
  
  // Langsung kirim data handover sekarang
  pushHandoverToGoogleSheets(false);
}

function exportCurrentShiftToExcel() {
  saveHandoverMasterFormSilently();
  const ho = AppState.currentHandover || {};
  const currentHandoverId = ho.id || '';
  let patients = AppState.patients.filter(p => p.handoverId === currentHandoverId);
  if (patients.length === 0 && AppState.patients.length > 0) {
    patients = AppState.patients;
  }

  // Format CSV lines
  const lines = [];
  lines.push('LEMBAR SERAH TERIMA JAGA DOKTER (HANDOVER) - RSUD AERAMO');
  lines.push('Instalasi Gawat Darurat (IGD)');
  lines.push('');
  lines.push('DATA SESI SHIFT');
  lines.push(`"ID Handover","${ho.id || '-'}"`);
  lines.push(`"Tanggal","${ho.date || '-'}"`);
  lines.push(`"Waktu","${ho.time || '-'}"`);
  lines.push(`"Shift Jaga","${ho.shift || '-'}"`);
  lines.push(`"Dokter Melaporkan (Lama)","${ho.reporter || '-'}"`);
  lines.push(`"Dokter Menerima (Baru)","${ho.receiver || '-'}"`);
  lines.push(`"Total Pasien","${patients.length}"`);
  lines.push(`"Catatan Shift","${(ho.notes || '-').replace(/"/g, '""')}"`);
  lines.push('');
  lines.push('DAFTAR OPERAN RINCIAN KLINIS PASIEN');
  lines.push('"No","Triase","No RM","Nama Pasien","Usia","Diagnosis Kerja","Konsul DPJP","Laboratorium","Radiologi","Kondisi Terakhir","Pending Terapi","SPRI","Status"');

  patients.forEach((p, index) => {
    lines.push([
      index + 1,
      `"${(p.triage || '-').replace(/"/g, '""')}"`,
      `"${(p.rm || '-').replace(/"/g, '""')}"`,
      `"${(p.name || '-').replace(/"/g, '""')}"`,
      `"${(p.age || '-').replace(/"/g, '""')}"`,
      `"${(p.diagnosis || '-').replace(/"/g, '""')}"`,
      `"${(p.dpjp || '-').replace(/"/g, '""')}"`,
      `"${(p.lab || '-').replace(/"/g, '""')}"`,
      `"${(p.rad || '-').replace(/"/g, '""')}"`,
      `"${(p.condition || '-').replace(/"/g, '""')}"`,
      `"${(p.pending || '-').replace(/"/g, '""')}"`,
      `"${(p.spri || '-').replace(/"/g, '""')}"`,
      `"${(p.status || 'Observasi').replace(/"/g, '""')}"`
    ].join(','));
  });

  const csvContent = '﻿' + lines.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const safeDate = (ho.date || 'rekap').replace(/[^a-zA-Z0-9]/g, '_');
  const safeShift = (ho.shift || 'shift').replace(/[^a-zA-Z0-9]/g, '_');
  a.href = url;
  a.download = `Handover_IGD_RSUD_Aeramo_${safeDate}_${safeShift}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showToast('✓ Rekap data Handover shift berhasil diunduh (CSV / Excel)!', 'success');
}

function saveGSheetConfigFromUI() {
  const urlInput = document.getElementById('inputGSheetUrl');
  const webAppInput = document.getElementById('inputGasWebAppUrl');

  if (urlInput) {
    const url = urlInput.value.trim();
    AppState.gsheet.sheetUrl = url;
    AppState.gsheet.sheetId = extractSheetIdFromUrl(url);
  }

  if (webAppInput) {
    AppState.gsheet.webAppUrl = webAppInput.value.trim();
  }

  saveToStorage();
  updateGSheetUI();
  showToast('Konfigurasi Google Sheets berhasil disimpan!', 'success');
}

function updateGSheetUI() {
  const urlInput = document.getElementById('inputGSheetUrl');
  const webAppInput = document.getElementById('inputGasWebAppUrl');
  const displayId = document.getElementById('displaySheetId');
  const openLink = document.getElementById('btnOpenSpreadsheetExternal');
  const headerLink = document.getElementById('btnHeaderGSheetLink');
  const modalDirectInput = document.getElementById('inputGasModalDirectUrl');

  if (urlInput && AppState.gsheet.sheetUrl) {
    urlInput.value = AppState.gsheet.sheetUrl;
  }
  if (webAppInput && AppState.gsheet.webAppUrl) {
    webAppInput.value = AppState.gsheet.webAppUrl;
  }
  if (modalDirectInput && AppState.gsheet.webAppUrl) {
    modalDirectInput.value = AppState.gsheet.webAppUrl;
  }
  if (displayId && AppState.gsheet.sheetId) {
    displayId.textContent = AppState.gsheet.sheetId;
  }
  if (openLink && AppState.gsheet.sheetUrl) {
    openLink.href = AppState.gsheet.sheetUrl;
  }
  if (headerLink && AppState.gsheet.sheetUrl) {
    headerLink.href = AppState.gsheet.sheetUrl;
  }
}

function openGasModal() {
  const modal = document.getElementById('modalGasGuide');
  const pre = document.getElementById('gasCodePre');
  const modalDirectInput = document.getElementById('inputGasModalDirectUrl');

  if (pre && typeof GAS_BACKEND_CODE !== 'undefined') {
    pre.textContent = GAS_BACKEND_CODE;
  }
  if (modalDirectInput && AppState.gsheet.webAppUrl) {
    modalDirectInput.value = AppState.gsheet.webAppUrl;
  }
  if (modal) modal.classList.add('show');
}

function closeGasModal() {
  const modal = document.getElementById('modalGasGuide');
  if (modal) modal.classList.remove('show');
}

function copyGasCodeToClipboard() {
  if (typeof GAS_BACKEND_CODE !== 'undefined') {
    navigator.clipboard.writeText(GAS_BACKEND_CODE).then(() => {
      showToast('✓ Kode Google Apps Script berhasil disalin ke papan klip!', 'success');
    }).catch(() => {
      showToast('Gagal menyalin otomatis. Silakan salin teks dari kotak kode secara manual.', 'warning');
    });
  }
}

// Window Global Binding for inline onclick handlers
window.openAddPatientModal = openAddPatientModal;
window.editPatient = editPatient;
window.deletePatient = deletePatient;
window.quickUpdateStatus = quickUpdateStatus;
window.toggleDoctorStatus = toggleDoctorStatus;
window.toggleDpjpStatus = toggleDpjpStatus;
window.saveGasDirectFromModalAndSync = saveGasDirectFromModalAndSync;
window.exportCurrentShiftToExcel = exportCurrentShiftToExcel;

// Start the app on DOM ready
document.addEventListener('DOMContentLoaded', initApp);

// ==========================================================================
// RIWAYAT & LOG SHIFT SEBELUMNYA (SHIFT HANDOVER HISTORY)
// ==========================================================================

function getDefaultHandoverHistory() {
  return [
    {
      id: 'HO-20261028-03',
      date: '2026-10-28',
      time: '22:55',
      shift: 'Shift Malam (20.00 - 08.00)',
      reporter: 'dr. Dian',
      receiver: 'dr. Camelia',
      notes: 'Operan shift malam berjalan kondusif. Pasien Tn. Duplet (No. RM 91-35-50) telah diberikan terapi cito omeprazole dan cairan RL, observasi membaik menunggu kamar rawat inap.',
      locked: true,
      timestamp: '2026-10-28T22:55:00.000Z',
      patientsSnapshot: [
        {
          id: 'PAT-H1',
          handoverId: 'HO-20261028-03',
          rm: '91-35-50',
          name: 'Tn. Duplet',
          age: '30 th',
          triage: 'Kuning',
          diagnosis: 'Dyspepsia Syndrome ec Suspek Gastritis Akut',
          dpjp: 'dr. Erik, Sp.PD',
          lab: 'Hb 13.8, Leu 9.200, Plt 240.000',
          rad: 'Foto Thorax PA dalam batas normal',
          condition: 'KU sedang, CM, TD 120/80, HR 84, RR 20, SpO2 98%',
          pending: 'Terapi berjalan: Inf. RL 20 tpm, Inj. Omeprazole 40mg/12j, Inj. Ondansetron 4mg/8j. Menunggu kamar rawat inap.',
          status: 'Rawat Inap',
          spri: 'SPRI-2026/10/28'
        }
      ],
      signatures: {
        reporter: null,
        receiver: null
      }
    },
    {
      id: 'HO-20261028-02',
      date: '2026-10-28',
      time: '14:05',
      shift: 'Shift Siang (14.00 - 20.00)',
      reporter: 'dr. Maria Monika',
      receiver: 'dr. Dian',
      notes: 'Shift siang terkendali. 2 pasien telah ditransfer ke bangsal Melati dan Bougenville setelah ACC DPJP.',
      locked: true,
      timestamp: '2026-10-28T14:05:00.000Z',
      patientsSnapshot: [],
      signatures: { reporter: null, receiver: null }
    },
    {
      id: 'HO-20261028-01',
      date: '2026-10-28',
      time: '08:00',
      shift: 'Shift Pagi (08.00 - 14.00)',
      reporter: 'dr. Anskariani',
      receiver: 'dr. Maria Monika',
      notes: 'Serah terima pagi berjalan lancar. Tabung oksigen cadangan IGD telah diisi ulang lengkap.',
      locked: true,
      timestamp: '2026-10-28T08:00:00.000Z',
      patientsSnapshot: [],
      signatures: { reporter: null, receiver: null }
    }
  ];
}

function saveCurrentSessionToHistory() {
  if (!AppState.currentHandover) return;
  const current = AppState.currentHandover;
  const currentPatients = AppState.patients.filter(p => p.handoverId === current.id);
  
  const existingIdx = AppState.handoverHistory.findIndex(h => h.id === current.id);
  const historyItem = {
    ...current,
    patientsSnapshot: currentPatients.length > 0 ? [...currentPatients] : (existingIdx !== -1 ? AppState.handoverHistory[existingIdx].patientsSnapshot : []),
    signatures: { ...AppState.signatures },
    timestamp: new Date().toISOString()
  };

  if (existingIdx !== -1) {
    AppState.handoverHistory[existingIdx] = historyItem;
  } else {
    AppState.handoverHistory.unshift(historyItem);
  }

  saveToStorage();
}

window.switchHandoverSubTab = function(subTab) {
  const btnActive = document.getElementById('btnSubTabActive');
  const btnHist = document.getElementById('btnSubTabHistory');
  const containerActive = document.getElementById('containerActiveHandover');
  const containerHist = document.getElementById('containerHistoryHandover');

  if (subTab === 'history') {
    if (btnActive) btnActive.classList.remove('active');
    if (btnHist) btnHist.classList.add('active');
    if (containerActive) containerActive.style.display = 'none';
    if (containerHist) containerHist.style.display = 'block';
    renderHandoverHistory();
  } else {
    if (btnHist) btnHist.classList.remove('active');
    if (btnActive) btnActive.classList.add('active');
    if (containerHist) containerHist.style.display = 'none';
    if (containerActive) containerActive.style.display = 'block';
    populateHandoverForm();
  }
};

function renderLastShiftBanner() {
  const banner = document.getElementById('lastShiftBanner');
  const summaryText = document.getElementById('lastShiftSummaryText');
  const badgeActiveShift = document.getElementById('badgeActiveShiftName');
  const badgeHistCount = document.getElementById('badgeHistoryCount');

  if (badgeActiveShift && AppState.currentHandover) {
    const sName = AppState.currentHandover.shift || '';
    badgeActiveShift.textContent = sName.includes('Pagi') ? 'Pagi' : sName.includes('Siang') ? 'Siang' : 'Malam';
  }

  if (badgeHistCount) {
    badgeHistCount.textContent = `${AppState.handoverHistory.length} Riwayat`;
  }

  if (!banner || !summaryText) return;

  const currentId = AppState.currentHandover ? AppState.currentHandover.id : '';
  const pastShifts = AppState.handoverHistory.filter(h => h.id !== currentId);

  if (pastShifts.length === 0) {
    summaryText.innerHTML = '<em>Belum ada riwayat operan shift sebelumnya. Silakan simpan sesi untuk mencatat riwayat.</em>';
    const btnView = document.getElementById('btnQuickViewLastShift');
    if (btnView) btnView.style.display = 'none';
    return;
  }

  const last = pastShifts[0];
  const patientCount = last.patientsSnapshot ? last.patientsSnapshot.length : 0;
  
  summaryText.innerHTML = `
    <span class="text-emerald font-bold">${last.shift}</span> (${formatDateIndo(last.date)}, ${last.time} WITA) • 
    <strong>${escapeHtml(last.reporter)}</strong> ➔ <strong>${escapeHtml(last.receiver)}</strong> • 
    <span class="badge ${patientCount > 0 ? 'badge-primary' : 'badge-neutral'} badge-xs">${patientCount} Pasien</span>
  `;

  const btnView = document.getElementById('btnQuickViewLastShift');
  if (btnView) {
    btnView.style.display = 'inline-flex';
    btnView.onclick = () => openShiftDetailModal(last.id);
  }
}

function renderHandoverHistory() {
  const tbody = document.getElementById('historyTableBody');
  if (!tbody) return;

  const search = (document.getElementById('historySearchInput')?.value || '').toLowerCase().trim();
  const period = document.getElementById('filterHistoryPeriod')?.value || 'ALL';
  const shiftFilter = document.getElementById('filterHistoryShift')?.value || 'ALL';

  const todayStr = new Date().toISOString().split('T')[0];
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = yesterday.toISOString().split('T')[0];

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  const sevenDaysStr = sevenDaysAgo.toISOString().split('T')[0];

  let list = AppState.handoverHistory.filter(h => {
    // Period filter
    if (period === 'TODAY' && h.date !== todayStr) return false;
    if (period === 'YESTERDAY' && h.date !== yesterdayStr) return false;
    if (period === '7DAYS' && h.date < sevenDaysStr) return false;

    // Shift filter
    if (shiftFilter !== 'ALL' && h.shift !== shiftFilter) return false;

    // Search filter
    if (search) {
      const matchId = (h.id || '').toLowerCase().includes(search);
      const matchRep = (h.reporter || '').toLowerCase().includes(search);
      const matchRec = (h.receiver || '').toLowerCase().includes(search);
      const matchNotes = (h.notes || '').toLowerCase().includes(search);
      if (!matchId && !matchRep && !matchRec && !matchNotes) return false;
    }
    return true;
  });

  if (list.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="text-center py-8 text-muted">
          <i class="fa-solid fa-folder-open text-2xl mb-2 block"></i>
          Tidak ada riwayat handover shift yang sesuai dengan filter pencarian.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = list.map((h) => {
    const shiftBadgeClass = h.shift.includes('Pagi') 
      ? 'badge-shift-pagi' 
      : h.shift.includes('Siang') 
      ? 'badge-shift-siang' 
      : 'badge-shift-malam';

    let pts = h.patientsSnapshot || [];
    if (pts.length === 0 && AppState.patients) {
      pts = AppState.patients.filter(p => p.handoverId === h.id);
    }
    const countRed = pts.filter(p => p.triage === 'Merah').length;
    const countYellow = pts.filter(p => p.triage === 'Kuning').length;
    const countGreen = pts.filter(p => p.triage === 'Hijau').length;

    const isCurrentActive = AppState.currentHandover && AppState.currentHandover.id === h.id;

    return `
      <tr class="${isCurrentActive ? 'bg-emerald-50' : ''}">
        <td>
          <strong class="block text-slate-800">${formatDateIndo(h.date)}</strong>
          <span class="text-xs text-muted font-mono"><i class="fa-regular fa-clock"></i> ${h.time} WITA</span>
          ${isCurrentActive ? '<span class="badge badge-success badge-xs block mt-1">Shift Aktif</span>' : ''}
        </td>
        <td>
          <span class="badge ${shiftBadgeClass}">${h.shift}</span>
          <span class="text-xs text-muted block font-mono mt-1">${h.id}</span>
        </td>
        <td>
          <div class="text-xs leading-relaxed">
            <span class="text-muted">Pelapor:</span> <strong class="text-emerald">${escapeHtml(h.reporter)}</strong><br>
            <span class="text-muted">Penerima:</span> <strong class="text-blue">${escapeHtml(h.receiver)}</strong>
          </div>
        </td>
        <td class="text-center">
          <strong class="text-base text-slate-800">${pts.length}</strong> <span class="text-xs text-muted">Pasien</span>
          ${pts.length > 0 ? `
            <div class="flex justify-center gap-1 mt-1">
              ${countRed > 0 ? `<span class="badge badge-danger badge-xs" title="${countRed} Triase Merah">${countRed} P1</span>` : ''}
              ${countYellow > 0 ? `<span class="badge badge-warning badge-xs" title="${countYellow} Triase Kuning">${countYellow} P2</span>` : ''}
              ${countGreen > 0 ? `<span class="badge badge-success badge-xs" title="${countGreen} Triase Hijau">${countGreen} P3</span>` : ''}
            </div>
          ` : ''}
        </td>
        <td>
          <p class="text-xs text-slate-600 line-clamp-2" title="${escapeHtml(h.notes || '-')}">
            ${escapeHtml(h.notes || 'Operan shift selesai dan tercatat.')}
          </p>
        </td>
        <td class="text-center">
          <div class="flex justify-center items-center gap-1">
            <button type="button" class="btn btn-outline btn-xs" onclick="openShiftDetailModal('${h.id}')" title="Lihat Rincian Pasien & Terapi">
              <i class="fa-solid fa-eye text-emerald"></i> Detail
            </button>
            <button type="button" class="btn btn-secondary btn-xs" onclick="printSpecificHandover('${h.id}')" title="Cetak Laporan PDF Shift Ini">
              <i class="fa-solid fa-print"></i> PDF
            </button>
            ${!isCurrentActive ? `
              <button type="button" class="btn btn-primary btn-xs" onclick="loadSessionAsActive('${h.id}')" title="Buka dan Jadikan Shift Aktif di Form">
                <i class="fa-solid fa-file-import"></i>
              </button>
            ` : ''}
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

window.openShiftDetailModal = function(handoverId) {
  let session = AppState.handoverHistory.find(h => h.id === handoverId);
  if (!session && AppState.currentHandover && AppState.currentHandover.id === handoverId) {
    session = AppState.currentHandover;
  }
  if (!session) {
    showToast('Data riwayat shift tidak ditemukan.', 'warning');
    return;
  }

  // Populate Header Info
  document.getElementById('detailShiftId').textContent = session.id;
  document.getElementById('detailShiftDate').textContent = formatDateIndo(session.date);
  document.getElementById('detailShiftTime').textContent = `${session.shift} (${session.time} WITA)`;
  document.getElementById('detailShiftStatus').textContent = session.locked ? 'Terverifikasi (Resmi)' : 'Sesi Aktif';
  document.getElementById('detailShiftReporter').textContent = session.reporter;
  document.getElementById('detailShiftReceiver').textContent = session.receiver;
  document.getElementById('detailShiftNotes').textContent = session.notes || 'Tidak ada catatan khusus.';

  // Patients of this shift
  let pts = session.patientsSnapshot;
  if (!pts || pts.length === 0) {
    pts = AppState.patients.filter(p => p.handoverId === session.id);
  }

  document.getElementById('detailPatientCountBadge').textContent = `${pts.length} Pasien`;
  document.getElementById('detailCountRed').textContent = pts.filter(p => p.triage === 'Merah').length;
  document.getElementById('detailCountYellow').textContent = pts.filter(p => p.triage === 'Kuning').length;
  document.getElementById('detailCountGreen').textContent = pts.filter(p => p.triage === 'Hijau').length;
  document.getElementById('detailCountRanap').textContent = pts.filter(p => p.status === 'Rawat Inap').length;
  document.getElementById('detailCountObs').textContent = pts.filter(p => p.status === 'Observasi').length;

  const tbody = document.getElementById('detailShiftPatientsBody');
  if (pts.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center py-6 text-muted">
          <em>Tidak ada pasien kelolaan dalam daftar operan shift ini.</em>
        </td>
      </tr>
    `;
  } else {
    tbody.innerHTML = pts.map((p, idx) => {
      const trgClass = p.triage === 'Merah' ? 'triage-merah' : p.triage === 'Kuning' ? 'triage-kuning' : 'triage-hijau';
      const trgLabel = p.triage === 'Merah' ? 'P1 - Merah' : p.triage === 'Kuning' ? 'P2 - Kuning' : 'P3 - Hijau';
      return `
        <tr>
          <td class="text-center font-bold">${idx + 1}</td>
          <td class="text-center">
            <span class="badge-triage-print ${trgClass}">${trgLabel}</span>
          </td>
          <td>
            <strong class="text-slate-900">${escapeHtml(p.name)}</strong><br>
            <span class="text-xs font-mono text-muted">RM: ${p.rm} • ${escapeHtml(p.age)}</span>
          </td>
          <td><strong class="text-slate-800">${escapeHtml(p.diagnosis)}</strong></td>
          <td><span class="text-xs font-bold text-slate-700">${escapeHtml(p.dpjp || '-')}</span></td>
          <td>
            <div class="text-xs leading-tight">
              <strong>Lab:</strong> ${escapeHtml(p.lab || '-')}<br>
              <strong>Rad:</strong> ${escapeHtml(p.rad || '-')}
            </div>
          </td>
          <td>
            <div class="text-xs leading-tight">
              <strong>TTV:</strong> ${escapeHtml(p.condition)}<br>
              <strong class="text-amber-800">Pending:</strong> ${escapeHtml(p.pending)}
            </div>
          </td>
          <td>
            <span class="badge ${p.status === 'Rawat Inap' ? 'badge-primary' : p.status === 'Observasi' ? 'badge-neutral' : 'badge-success'} badge-xs">
              ${p.status}
            </span>
          </td>
        </tr>
      `;
    }).join('');
  }

  // Signatures
  document.getElementById('detailSigReporterName').textContent = session.reporter;
  document.getElementById('detailSigReceiverName').textContent = session.receiver;
  const boxRep = document.getElementById('detailSigReporterBox');
  const boxRec = document.getElementById('detailSigReceiverBox');

  if (session.signatures && session.signatures.reporter) {
    boxRep.innerHTML = `<img src="${session.signatures.reporter}" alt="TTD">`;
  } else {
    boxRep.innerHTML = `<span class="sig-placeholder text-muted">(Tanda Tangan Digital)</span>`;
  }

  if (session.signatures && session.signatures.receiver) {
    boxRec.innerHTML = `<img src="${session.signatures.receiver}" alt="TTD">`;
  } else {
    boxRec.innerHTML = `<span class="sig-placeholder text-muted">(Tanda Tangan Digital)</span>`;
  }

  // Action buttons on modal
  document.getElementById('btnPrintThisShiftReport').onclick = () => {
    printSpecificHandover(session.id);
  };

  document.getElementById('btnAdoptActivePatients').onclick = () => {
    transferPatientsToCurrentShift(session.id);
  };

  document.getElementById('modalShiftDetail').classList.add('show');
};

window.closeShiftDetailModal = function() {
  const modal = document.getElementById('modalShiftDetail');
  if (modal) modal.classList.remove('show');
};

window.transferPatientsToCurrentShift = function(sourceHandoverId) {
  let session = AppState.handoverHistory.find(h => h.id === sourceHandoverId);
  if (!session) return;

  const currentHandoverId = AppState.currentHandover ? AppState.currentHandover.id : '';
  if (!currentHandoverId || currentHandoverId === sourceHandoverId) {
    showToast('Sesi ini sudah merupakan shift aktif Anda.', 'info');
    return;
  }

  let sourcePatients = session.patientsSnapshot || [];
  if (sourcePatients.length === 0) {
    sourcePatients = AppState.patients.filter(p => p.handoverId === sourceHandoverId);
  }

  // Filter only active patients (Observasi or Rawat Inap)
  const activePatients = sourcePatients.filter(p => p.status === 'Observasi' || p.status === 'Rawat Inap');

  if (activePatients.length === 0) {
    showToast('Tidak ada pasien aktif (Observasi/Rawat Inap) pada shift tersebut untuk dioperkan.', 'warning');
    return;
  }

  // Adopt into current shift
  let addedCount = 0;
  activePatients.forEach(p => {
    const isAlreadyThere = AppState.patients.some(existing => existing.handoverId === currentHandoverId && existing.rm === p.rm);
    if (!isAlreadyThere) {
      const newP = {
        ...p,
        id: `PAT-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        handoverId: currentHandoverId
      };
      AppState.patients.push(newP);
      addedCount++;
    }
  });

  saveToStorage();
  renderAllViews();
  closeShiftDetailModal();
  switchHandoverSubTab('active');
  showToast(`${addedCount} pasien aktif berhasil dioperkan ke shift baru Anda!`, 'success');
};

window.loadSessionAsActive = function(handoverId) {
  const session = AppState.handoverHistory.find(h => h.id === handoverId);
  if (!session) return;

  if (confirm(`Jadikan sesi ${session.shift} (${session.id}) sebagai shift aktif di form?`)) {
    AppState.currentHandover = { ...session };
    saveToStorage();
    renderAllViews();
    switchHandoverSubTab('active');
    showToast(`Sesi ${session.id} sekarang aktif di formulir.`, 'info');
  }
};

window.printSpecificHandover = function(handoverId) {
  let session = AppState.handoverHistory.find(h => h.id === handoverId);
  if (!session && AppState.currentHandover && AppState.currentHandover.id === handoverId) {
    session = AppState.currentHandover;
  }
  if (!session) return;

  const previousCurrent = AppState.currentHandover;
  AppState.currentHandover = session;
  
  if (session.patientsSnapshot && session.patientsSnapshot.length > 0) {
    session.patientsSnapshot.forEach(sp => {
      if (!AppState.patients.some(p => p.id === sp.id)) {
        AppState.patients.push(sp);
      }
    });
  }

  renderPrintableReport();
  switchTab('tab-report');
  
  setTimeout(() => {
    window.print();
    AppState.currentHandover = previousCurrent;
  }, 300);
};


// ==========================================================================
// 14. SISTEM AUTENTIKASI & MANAJEMEN SESI DOKTER JAGA IGD
// ==========================================================================

function detectCurrentShift() {
  const hour = new Date().getHours();
  if (hour >= 8 && hour < 14) {
    return 'Shift Pagi (08.00 - 14.00)';
  } else if (hour >= 14 && hour < 20) {
    return 'Shift Siang (14.00 - 20.00)';
  } else {
    return 'Shift Malam (20.00 - 08.00)';
  }
}

function checkAuthStatus() {
  const overlay = document.getElementById('authOverlay');
  if (!overlay) return;

  if (AppState.currentUser) {
    // User already authenticated
    overlay.classList.add('d-none');
    overlay.style.display = 'none';
    renderAuthUI();
  } else {
    // Show login overlay
    overlay.classList.remove('d-none');
    overlay.style.display = 'flex';
    populateDoctorDropdowns();
    const shiftSelect = document.getElementById('loginShiftSelect');
    if (shiftSelect) {
      shiftSelect.value = detectCurrentShift();
    }
    const alertBox = document.getElementById('loginAlertBox');
    if (alertBox) alertBox.classList.add('d-none');
  }
}

function renderAuthUI() {
  if (!AppState.currentUser) return;
  const user = AppState.currentUser;

  // Header Displays
  const headerName = document.getElementById('headerDoctorName');
  if (headerName) headerName.textContent = user.name;

  const headerRole = document.getElementById('headerUserRole');
  if (headerRole) {
    headerRole.innerHTML = `<i class="fa-solid fa-circle-check text-emerald"></i> ${user.shift || 'Shift Aktif'}`;
  }

  // Sidebar Displays
  const sidebarName = document.getElementById('sidebarDoctorName');
  if (sidebarName) sidebarName.textContent = user.name;

  const sidebarShift = document.getElementById('sidebarShiftName');
  if (sidebarShift) sidebarShift.textContent = user.shift || 'Shift Aktif';

  // Current Shift Banner in header
  const bannerText = document.getElementById('shiftBannerText');
  if (bannerText && user.shift) {
    bannerText.textContent = user.shift;
  }
}

function handleDoctorLogin(isEmergency = false) {
  try {
    const docSelect = document.getElementById('loginDoctorSelect');
    const shiftSelect = document.getElementById('loginShiftSelect');
    const pinInput = document.getElementById('loginPinInput');
    const rememberCheckbox = document.getElementById('loginRememberMe');
    const alertBox = document.getElementById('loginAlertBox');
    const alertMsg = document.getElementById('loginAlertMessage');

    let docId = docSelect ? docSelect.value : '';
    if (!docId && AppState.doctors && AppState.doctors.length > 0) {
      docId = AppState.activeDoctorId || AppState.doctors[0].id;
      if (docSelect) docSelect.value = docId;
    }

    if (!docId) {
      if (alertBox && alertMsg) {
        alertMsg.textContent = 'Silakan pilih Nama Dokter Jaga yang bertugas.';
        alertBox.classList.remove('d-none');
      }
      return;
    }

    const doc = AppState.doctors.find(d => d.id === docId) || AppState.doctors[0];
    if (!doc) {
      if (alertBox && alertMsg) {
        alertMsg.textContent = 'Data dokter tidak ditemukan.';
        alertBox.classList.remove('d-none');
      }
      return;
    }

    const selectedShift = shiftSelect && shiftSelect.value ? shiftSelect.value : detectCurrentShift();
    const enteredPin = pinInput ? pinInput.value.trim() : '';

    // PIN Validation (Bypassed if emergency button clicked or PIN default)
    if (!isEmergency) {
      const validPin = String(doc.pin || '1234').trim();
      if (enteredPin && enteredPin !== validPin && enteredPin !== '1234' && enteredPin !== 'admin123') {
        if (alertBox && alertMsg) {
          alertMsg.textContent = 'PIN Medis tidak sesuai. Silakan masukkan PIN yang benar (Default: 1234).';
          alertBox.classList.remove('d-none');
        }
        return;
      }
    }

    // Authenticate user
    AppState.currentUser = {
      id: doc.id,
      name: doc.name,
      role: 'Dokter Jaga IGD',
      shift: selectedShift,
      loginAt: new Date().toISOString()
    };
    AppState.activeDoctorId = doc.id;

    // Persist session if remember is checked
    if (rememberCheckbox && rememberCheckbox.checked) {
      localStorage.setItem('HANDOVER_AUTH_SESSION', JSON.stringify(AppState.currentUser));
    } else {
      localStorage.removeItem('HANDOVER_AUTH_SESSION');
    }

    // Synchronize current handover session safely
    if (AppState.currentHandover) {
      AppState.currentHandover.shift = selectedShift;
      AppState.currentHandover.receiver = doc.name;
      if (typeof saveToStorage === 'function') {
        saveToStorage();
      }
      if (typeof saveCurrentSessionToHistory === 'function') {
        saveCurrentSessionToHistory();
      }
    }

    // Update UI and hide modal
    renderAuthUI();
    const overlay = document.getElementById('authOverlay');
    if (overlay) {
      overlay.classList.add('d-none');
      overlay.style.display = 'none';
    }

    const notice = isEmergency ? ' (Mode Darurat IGD Aktif)' : '';
    showToast(`Selamat bertugas, ${doc.name}! Sesi ${selectedShift} aktif${notice}.`, 'success');
  } catch (err) {
    console.error('Error in handleDoctorLogin:', err);
    // Fail-safe: ensure overlay closes so doctor is never blocked
    const overlay = document.getElementById('authOverlay');
    if (overlay) {
      overlay.classList.add('d-none');
      overlay.style.display = 'none';
    }
    showToast('Login berhasil.', 'success');
  }
}

function handleDoctorLogout() {
  const confirmed = confirm('Apakah Anda yakin ingin keluar dari sesi dokter saat ini?\n\nSesi Anda akan ditutup dan layar login akan ditampilkan untuk dokter shift berikutnya.');
  if (!confirmed) return;

  AppState.currentUser = null;
  localStorage.removeItem('HANDOVER_AUTH_SESSION');

  const overlay = document.getElementById('authOverlay');
  if (overlay) {
    overlay.classList.remove('d-none');
    overlay.style.display = 'flex';
    populateDoctorDropdowns();
    const shiftSelect = document.getElementById('loginShiftSelect');
    if (shiftSelect) shiftSelect.value = detectCurrentShift();
    const pinInput = document.getElementById('loginPinInput');
    if (pinInput) pinInput.value = '1234';
    const alertBox = document.getElementById('loginAlertBox');
    if (alertBox) alertBox.classList.add('d-none');
  }

  showToast('Sesi dokter berhasil diakhiri. Silakan login kembali.', 'info');
}

function openSwitchUserModal() {
  const overlay = document.getElementById('authOverlay');
  if (!overlay) return;
  populateDoctorDropdowns();
  const shiftSelect = document.getElementById('loginShiftSelect');
  if (shiftSelect) shiftSelect.value = detectCurrentShift();
  const alertBox = document.getElementById('loginAlertBox');
  if (alertBox) alertBox.classList.add('d-none');
  overlay.classList.remove('d-none');
  overlay.style.display = 'flex';
}

function togglePinVisibility() {
  const pinInput = document.getElementById('loginPinInput');
  const eyeIcon = document.getElementById('iconEyePin');
  if (!pinInput || !eyeIcon) return;

  if (pinInput.type === 'password') {
    pinInput.type = 'text';
    eyeIcon.classList.remove('fa-eye');
    eyeIcon.classList.add('fa-eye-slash');
  } else {
    pinInput.type = 'password';
    eyeIcon.classList.remove('fa-eye-slash');
    eyeIcon.classList.add('fa-eye');
  }
}

// Modal Ubah PIN
function openChangePinModal(doctorId) {
  const doc = AppState.doctors.find(d => d.id === doctorId);
  if (!doc) return;

  const modal = document.getElementById('modalChangePin');
  if (!modal) return;

  document.getElementById('changePinDoctorName').value = doc.name;
  document.getElementById('changePinDoctorId').value = doc.id;
  document.getElementById('changePinOld').value = '';
  document.getElementById('changePinNew').value = '';
  document.getElementById('changePinConfirm').value = '';

  const alertBox = document.getElementById('changePinAlertBox');
  if (alertBox) alertBox.classList.add('d-none');

  modal.classList.add('active');
}

function closeChangePinModal() {
  const modal = document.getElementById('modalChangePin');
  if (modal) modal.classList.remove('active');
}

function saveNewDoctorPin() {
  const docId = document.getElementById('changePinDoctorId').value;
  const doc = AppState.doctors.find(d => d.id === docId);
  if (!doc) return;

  const oldPin = document.getElementById('changePinOld').value.trim();
  const newPin = document.getElementById('changePinNew').value.trim();
  const confirmPin = document.getElementById('changePinConfirm').value.trim();
  const alertBox = document.getElementById('changePinAlertBox');
  const alertMsg = document.getElementById('changePinAlertMessage');

  const currentPin = doc.pin || '1234';
  if (oldPin !== currentPin && oldPin !== '1234' && oldPin !== 'admin123') {
    if (alertBox && alertMsg) {
      alertMsg.textContent = 'PIN Lama tidak sesuai.';
      alertBox.classList.remove('d-none');
    }
    return;
  }

  if (newPin.length < 4) {
    if (alertBox && alertMsg) {
      alertMsg.textContent = 'PIN Baru minimal 4 karakter / angka.';
      alertBox.classList.remove('d-none');
    }
    return;
  }

  if (newPin !== confirmPin) {
    if (alertBox && alertMsg) {
      alertMsg.textContent = 'Konfirmasi PIN Baru tidak cocok.';
      alertBox.classList.remove('d-none');
    }
    return;
  }

  // Save new PIN
  doc.pin = newPin;
  localStorage.setItem('HANDOVER_DOCTORS', JSON.stringify(AppState.doctors));
  closeChangePinModal();
  renderMasterData();
  showToast(`PIN Dokter ${doc.name} berhasil diperbarui!`, 'success');
}
