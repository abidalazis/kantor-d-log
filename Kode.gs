// Kode.gs

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Log Pekerjaan Harian & Portofolio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Menyimpan data pekerjaan ke Google Sheets.
 * @param {Object} data - Objek yang berisi data dari form.
 * @returns {Object} - Status sukses atau gagal.
 */
function simpanData(data) {
  try {
    const sheet = getSheet();
    
    // Pastikan header sudah ada (Optional, untuk pertama kali)
    setupHeader(sheet);
    
    // Menyiapkan data array sesuai urutan kolom:
    // [Timestamp, Tanggal Pekerjaan, Kategori, Nama Pekerjaan, Link Hasil/Canva, Catatan]
    const timestamp = new Date(); // Waktu saat disubmit
    const rowData = [
      timestamp,
      data.tanggal,
      data.kategori,
      data.namaPekerjaan,
      data.linkCanva,
      data.catatan
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: 'Data berhasil disimpan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan data: ' + error.message };
  }
}

/**
 * Mengambil data riwayat pekerjaan dari Google Sheets.
 * @returns {Array} - Array of objects berisi histori pekerjaan.
 */
function getRiwayat() {
  try {
    const sheet = getSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) return []; // Hanya header atau kosong
    
    // Hapus header
    const data = values.slice(1);
    
    // Format data ke bentuk array of objects
    // Urutan: timestamp [0], tanggal [1], kategori [2], nama [3], link [4], catatan [5]
    const result = data.map(function(row) {
      // Memformat tanggal ke string bentuk YYYY-MM-DD agar mudah di-filter di UI
      let tanggalStr = '';
      if (row[1] instanceof Date) {
        tanggalStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        tanggalStr = row[1]; // fallback jika bentuk plain text
      }

      // Pastikan objek Date diubah ke primitive (string/number) karena google.script.run tidak mendukung Date object native
      let timestampVal = row[0];
      if (timestampVal instanceof Date) {
        timestampVal = timestampVal.toISOString(); 
      }

      return {
        timestamp: timestampVal,
        tanggal: String(tanggalStr),
        kategori: String(row[2] || ''),
        namaPekerjaan: String(row[3] || ''),
        linkCanva: String(row[4] || ''),
        catatan: String(row[5] || '')
      };
    });
    
    // Membalik urutan agar yang terbaru (baris bawah) tampil lebih dulu
    return result.reverse();
    
  } catch (error) {
    Logger.log(error);
    return []; // Return array kosong jika terjadi error
  }
}

/**
 * Fungsi helper untuk mendapatkan sheet aktif.
 */
function getSheet() {
  // Script ini diasumsikan berjalan sebagai web app yang menempel pada spreadsheet (container-bound)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('Log Pekerjaan');
  if (!sheet) {
    sheet = ss.insertSheet('Log Pekerjaan'); // Buat otomatis tab sheet jika belum ada
  }
  return sheet;
}

/**
 * Fungsi helper untuk membuat baris header jika sheet masih kosong.
 */
function setupHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp', 'Tanggal Pekerjaan', 'Kategori', 'Nama Pekerjaan', 'Link Hasil/Canva', 'Catatan'];
    sheet.appendRow(headers);
    // Menebalkan huruf dan memberi warna latar untuk baris header
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
  }
}
