# E-Hanter Gaid — Elektronik Handover Dokter Jaga IGD • RSUD Aeramo

Sistem Informasi dan Aplikasi Terintegrasi Handover Dokter IGD untuk digitalisasi serah terima shift jaga medis, monitoring triase secara real-time, penerbitan Surat Permintaan Rawat Inap (SPRI), tanda tangan digital dokter, serta integrasi langsung dengan **Google Sheets** dan **Cloudflare Pages**.

## 🌟 Fitur Utama
- **Dashboard & Triase Medis Real-Time:** Monitoring live triase Merah (Resusitasi P1), Kuning (Urgensi P2), Hijau (P3), dan status pasien observasi / SPRI rawat inap / pulang.
- **Formulir Handover Shift & Pasien:** Input data shift, dokter pelapor, dokter penerima, dan daftar pasien terkait (Child list).
- **Tanda Tangan Digital Dokter (Canvas Signature):** Verifikasi dua arah dokter pelapor dan dokter penerima langsung di layar (mouse & touchscreen tablet/HP).
- **Live Patient Tracker:** Filter pencarian instan No. RM, nama pasien, diagnosis kerja, dan DPJP spesialis konsultan.
- **Lembar Cetak Rekam Medis Resmi (`FORM-RM-IGD-07`):** Standar lampiran rekam medis IGD lengkap dengan kop rumah sakit, stempel verifikasi, dan matriks triase.
- **Integrasi Google Sheets (Two-Way Sync):** Tersambung ke Google Spreadsheet sebagai cloud database untuk penyimpanan riwayat shift dan daftar pasien secara otomatis.
- **Backend Google Apps Script (`google-apps-script.js`):** Endpoint API ringan untuk membaca dan menulis data ke Google Sheets tanpa hambatan login berulang.

## 🚀 Struktur Berkas
- `index.html` : Antarmuka pengguna (UI) aplikasi handover
- `app.js` : Logika state, sinkronisasi Google Sheets, tracker, dan canvas tanda tangan
- `style.css` : Sistem desain antarmuka medis (Hospital Clean White & Emerald Green accents)
- `google-apps-script.js` : Kode backend Google Apps Script untuk Google Sheets
- `Handover Dokter.xlsx` : Berkas master tabel referensi awal
- `Rancangan Aplikasi Handover Dokter IGD.pdf` : Dokumen spesifikasi fungsional sistem

## ☁️ Deploy di Cloudflare Pages
Aplikasi ini berbasis static web (HTML, CSS, JS) tanpa ketergantungan build kompleks:
1. Hubungkan repository ini ke **Cloudflare Pages** (Workers & Pages).
2. Framework Preset: `None` (Static HTML).
3. Build output directory: `/` (Root directory).
4. Klik **Save and Deploy**.

