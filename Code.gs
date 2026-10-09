function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Ayo Booking - Reservation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// 1. Ambil Daftar Ruangan dari Tab 'SPACE'
function getRuangan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPACE");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push(data[i][0].toString().trim());
    }
  }
  return list;
}

// Helper untuk format Jam/Tanggal ke Teks Rapi
function formatTeks(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  return val.toString().trim();
}

function formatTanggal(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return val.toString().trim();
}

// 2. Ambil Jadwal Booking Berdasarkan Tanggal
function getJadwalBooking(tanggalTarget) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BOOK");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    var tglRow = formatTanggal(data[i][2]);
    var status = data[i][5] ? data[i][5].toString().trim() : "";

    if (tglRow === tanggalTarget && status !== "Canceled") {
      list.push({
        pemesan: data[i][0] ? data[i][0].toString() : "",
        ruang: data[i][1] ? data[i][1].toString() : "",
        jamMulai: formatTeks(data[i][3]),
        jamSelesai: formatTeks(data[i][4]),
        status: status
      });
    }
  }
  return list;
}

// 3. Simpan Pesanan Baru & Peringatan Bentrok
function simpanBooking(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BOOK");
  var existingJadwal = getJadwalBooking(form.tanggal);
  
  // Validasi Bentrok
  var bentrok = existingJadwal.some(function(item) {
    if (item.ruang === form.ruang) {
      var startNew = form.jamMulai;
      var endNew = form.jamSelesai;
      var startExist = item.jamMulai;
      var endExist = item.jamSelesai;
      
      // Rumus overlapping jam
      return (startNew < endExist && endNew > startExist);
    }
    return false;
  });

  if (bentrok) {
    return { 
      success: false, 
      message: "⚠️ JADWAL PENUH!\n\nSpace/Ruangan '" + form.ruang + "' sudah terisi di jam tersebut pada tanggal " + form.tanggal + ". Silakan pilih jam atau ruangan lain." 
    };
  }

  // Simpan jika tidak bentrok
  sheet.appendRow([
    form.nama,
    form.ruang,
    form.tanggal,
    form.jamMulai,
    form.jamSelesai,
    "Confirmed"
  ]);

  return { success: true, message: "✅ BERHASIL!\n\nJadwal booking berhasil disimpan." };
}
