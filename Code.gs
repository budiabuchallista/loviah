/**
 * KAS KELUARGA — Backend Google Apps Script
 * Sheet ini menyimpan seluruh data aplikasi (transaksi, kegiatan, hutang piutang)
 * sebagai satu blok JSON di sel A1, dan waktu update di sel B1.
 *
 * CARA PASANG:
 * 1. Buka https://sheet.new untuk membuat Google Sheet baru, beri nama "Kas Keluarga".
 * 2. Di sheet pertama, ganti namanya jadi "Data" (klik kanan tab di bawah > Rename).
 * 3. Menu Extensions > Apps Script.
 * 4. Hapus semua kode default, tempel seluruh isi file ini.
 * 5. Klik ikon Save (disket), lalu klik Deploy > New deployment.
 * 6. Pilih tipe "Web app". Isi:
 *      Execute as: Me
 *      Who has access: Anyone
 * 7. Klik Deploy, izinkan akses (Authorize access) dengan akun Google Anda.
 * 8. Salin URL Web App yang muncul (diakhiri /exec), tempel ke aplikasi Kas Keluarga
 *    di menu "Sinkronisasi Google Sheets" — gunakan URL yang sama di HP istri Anda.
 */

function doGet(e) {
  const sheet = getSheet_();
  const json = sheet.getRange('A1').getValue() ||
    '{"transactions":[],"kegiatan":[],"hutangPiutang":[]}';
  const updatedAt = sheet.getRange('B1').getValue() || '';
  return ContentService
    .createTextOutput(JSON.stringify({ data: json, updatedAt: String(updatedAt) }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = getSheet_();
  const body = JSON.parse(e.postData.contents);
  sheet.getRange('A1').setValue(JSON.stringify(body.data));
  const now = new Date().toISOString();
  sheet.getRange('B1').setValue(now);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, updatedAt: now }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data');
  if (!sheet) sheet = ss.insertSheet('Data');
  return sheet;
}
