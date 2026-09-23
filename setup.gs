function doGet(e) {
  // Ganti 'Sheet1' dengan nama sheet yang Anda gunakan jika berbeda
  var sheetName = 'Sheet1'; 
  
  // Menggunakan CacheService untuk mempercepat loading data
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get("ppdb_data");
  
  if (cachedData != null) {
    // Jika data ada di cache, langsung kembalikan (ini jauh lebih cepat)
    return ContentService.createTextOutput(cachedData).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Sheet tidak ditemukan. Pastikan nama sheet adalah '" + sheetName + "'."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  
  // Mengambil baris pertama sebagai header
  var headers = data[0];
  var rows = data.slice(1);
  
  var result = rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      var value = row[index];
      
      // Format tanggal jika nilai berupa objek Date
      if (value instanceof Date) {
        var day = ("0" + value.getDate()).slice(-2);
        var month = ("0" + (value.getMonth() + 1)).slice(-2);
        var year = value.getFullYear();
        var hours = ("0" + value.getHours()).slice(-2);
        var minutes = ("0" + value.getMinutes()).slice(-2);
        
        // Membedakan format timestamp dan tanggal lahir berdasarkan nama header
        if (header.toLowerCase().includes("timestamp")) {
           // Sesuai permintaan: DD-MM-YYYY hh:mm
           value = day + "-" + month + "-" + year + " " + hours + ":" + minutes;
        } else {
           // Sesuai permintaan: DD-MM-YYYY
           value = day + "-" + month + "-" + year;
        }
      }
      obj[header] = value;
    });
    return obj;
  });

  var jsonOutput = JSON.stringify({
    status: "success",
    data: result
  });
  
  // Simpan hasil ke cache selama 30 detik. 
  // Jika ada pendaftar baru, maksimal delay 30 detik sebelum muncul di web.
  cache.put("ppdb_data", jsonOutput, 30);

  return ContentService.createTextOutput(jsonOutput).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi dummy untuk memberikan izin saat pertama kali dijalankan
function setup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Setup berhasil. ID Spreadsheet: " + sheet.getId());
}
