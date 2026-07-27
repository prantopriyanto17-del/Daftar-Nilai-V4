/**
 * Aplikasi Daftar Nilai Antigravity
 * Backend Google Apps Script (GAS)
 * Penulis: Senior Full-stack Developer (Antigravity)
 */

// Schema database untuk mendefinisikan lembar kerja dan kolom-kolomnya
const DB_SCHEMA = {
  "Daftar Siswa": ["No", "NISN", "Nama", "Kelas", "NoOrangTua", "Keterangan"],
  "Daftar Nilai Matematika": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Daftar Nilai B.Indonesia": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "IPAS": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Pendidikan Pancasila": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Seni": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "B Sunda": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Kokurikuler": ["No", "NISN", "Nama", "Nilai Harian/Tugas", "Nilai TP1", "TP2", "TP3", "TP4", "TP5", "Nilai LM1", "LM2", "LM3", "LM4", "LM5", "Nilai Formatif", "Nilai Sumatif", "Nilai Sikap", "Nilai Kehadiran", "Keterangan"],
  "Catatan Wali Kelas": ["No", "NISN", "Nama", "Kelas", "NoOrangTua", "Catatan"]
};

/**
 * Endpoint GET (doGet)
 * Melayani UI HTML atau request GET API
 */
function doGet(e) {
  e = e || { parameter: {} };
  var action = e.parameter.action;
  
  // Jika tidak ada parameter action, layani UI Index.html
  if (!action) {
    try {
      var html = HtmlService.createTemplateFromFile('Index').evaluate();
      html.setTitle('Daftar Nilai Antigravity');
      html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      return html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput("Error memuat UI: " + err.toString());
    }
  }
  
  // Tangani request API GET
  var request = {
    action: action,
    sheetName: e.parameter.sheetName,
    nisn: e.parameter.nisn,
    no: e.parameter.no
  };
  
  var result = executeAction(request);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Endpoint POST (doPost)
 * Menangani penulisan data (Create, Update, Delete)
 */
function doPost(e) {
  var request = {};
  try {
    if (e && e.postData && e.postData.contents) {
      request = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Gagal memparsing JSON: " + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Gabungkan parameter URL jika ada
  if (e && e.parameter) {
    for (var key in e.parameter) {
      request[key] = e.parameter[key];
    }
  }
  
  var result = executeAction(request);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fungsi Router Utama
 * Dipanggil oleh doGet, doPost, dan google.script.run
 */
function executeAction(request) {
  try {
    var action = request.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!ss) {
      // Coba buka spreadsheet default jika ini dijalankan di luar container
      try {
        ss = SpreadsheetApp.openByUrl(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL"));
      } catch (err) {
        throw new Error("Spreadsheet tidak terdeteksi. Hubungkan skrip ini ke Google Sheet atau setel properti SPREADSHEET_URL.");
      }
    }
    
    // Inisialisasi Database
    if (action === 'initDb') {
      return initializeDatabase(ss);
    }
    
    switch (action) {
      case 'read':
        return readData(ss, request.sheetName);
      case 'create':
        return createData(ss, request.sheetName, request.data);
      case 'update':
        return updateData(ss, request.sheetName, request.data);
      case 'delete':
        return deleteData(ss, request.sheetName, request.nisn || request.no || (request.data && (request.data.NISN || request.data.No)));
      case 'getSiswaList':
        return getSiswaList(ss);
      case 'getStudentSummary':
        return getStudentSummary(ss, request.nisn);
      case 'getDashboardStats':
        return getDashboardStats(ss);
      default:
        return { success: false, error: 'Aksi tidak dikenal: ' + action };
    }
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Membuat dan menyiapkan sheet sesuai DB_SCHEMA
 */
function initializeDatabase(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  
  for (var sheetName in DB_SCHEMA) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = DB_SCHEMA[sheetName];
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      // Format Header
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#4f46e5")
        .setFontColor("#ffffff")
        .setHorizontalAlignment("center");
      
      sheet.setFrozenRows(1);
      for (var col = 1; col <= headers.length; col++) {
        sheet.autoResizeColumn(col);
      }
      log.push("Sheet '" + sheetName + "' berhasil dibuat.");
    } else {
      // Cek apakah kolom sudah sesuai
      var lastCol = sheet.getLastColumn();
      var existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      var isHeaderMatch = true;
      
      if (existingHeaders.length < headers.length) {
        isHeaderMatch = false;
      } else {
        for (var i = 0; i < headers.length; i++) {
          if (existingHeaders[i] !== headers[i]) {
            isHeaderMatch = false;
            break;
          }
        }
      }
      
      if (!isHeaderMatch) {
        // Tulis ulang header yang benar
        sheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground("#4f46e5")
          .setFontColor("#ffffff")
          .setHorizontalAlignment("center");
        sheet.setFrozenRows(1);
        log.push("Header di sheet '" + sheetName + "' diperbarui.");
      } else {
        log.push("Sheet '" + sheetName + "' sudah ada dan sesuai.");
      }
    }
  }
  
  return { success: true, message: "Inisialisasi database selesai.", log: log };
}

/**
 * Membaca seluruh data dari Sheet tertentu
 */
function readData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: "Sheet '" + sheetName + "' tidak ditemukan." };
  }
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1) {
    return { success: true, data: [] }; // Hanya ada header atau kosong
  }
  
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var headers = values[0];
  var data = [];
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    obj._rowIndex = i + 1; // Menyimpan baris asli untuk update/delete
    data.push(obj);
  }
  
  return { success: true, data: data };
}

/**
 * Menambahkan baris baru ke Sheet
 */
function createData(ss, sheetName, rowData) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: "Sheet '" + sheetName + "' tidak ditemukan." };
  }
  
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getDataRange().getValues();
  
  // Validasi NISN Unik di Daftar Siswa atau sheet bersangkutan
  if (rowData.NISN) {
    var nisnIndex = headers.indexOf("NISN");
    if (nisnIndex !== -1) {
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][nisnIndex]).trim() === String(rowData.NISN).trim()) {
          return { success: false, error: "NISN " + rowData.NISN + " sudah terdaftar di '" + sheetName + "'." };
        }
      }
    }
  }
  
  // Auto-increment untuk No
  var nextNo = 1;
  if (values.length > 1) {
    var noIndex = headers.indexOf("No");
    if (noIndex !== -1) {
      var maxNo = 0;
      for (var i = 1; i < values.length; i++) {
        var val = parseInt(values[i][noIndex]);
        if (!isNaN(val) && val > maxNo) {
          maxNo = val;
        }
      }
      nextNo = maxNo + 1;
    }
  }
  rowData["No"] = nextNo;
  
  // Buat array data baris sesuai susunan header
  var rowToAppend = [];
  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var val = rowData[header];
    rowToAppend.push(val !== undefined && val !== null ? val : "");
  }
  
  sheet.appendRow(rowToAppend);
  return { success: true, message: "Data berhasil ditambahkan ke " + sheetName, data: rowData };
}

/**
 * Memperbarui baris data di Sheet berdasarkan NISN atau No
 */
function updateData(ss, sheetName, rowData) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: "Sheet '" + sheetName + "' tidak ditemukan." };
  }
  
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getDataRange().getValues();
  
  var nisnIndex = headers.indexOf("NISN");
  var noIndex = headers.indexOf("No");
  var rowIndex = -1;
  
  // Cari berdasarkan NISN
  if (rowData.NISN && nisnIndex !== -1) {
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][nisnIndex]).trim() === String(rowData.NISN).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
  }
  
  // Jika tidak ketemu, cari berdasarkan No
  if (rowIndex === -1 && rowData.No && noIndex !== -1) {
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][noIndex]).trim() === String(rowData.No).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, error: "Data tidak ditemukan di '" + sheetName + "' untuk di-update." };
  }
  
  // Update cell per cell sesuai headers yang dikirim
  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    if (header === "No" || (header === "NISN" && rowData.NISN === undefined)) {
      continue; // No dan NISN (jika kosong) tidak di-overwrite
    }
    if (rowData[header] !== undefined) {
      sheet.getRange(rowIndex, j + 1).setValue(rowData[header]);
    }
  }
  
  return { success: true, message: "Data berhasil diperbarui di " + sheetName };
}

/**
 * Menghapus baris data dari Sheet berdasarkan NISN atau No
 */
function deleteData(ss, sheetName, key) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, error: "Sheet '" + sheetName + "' tidak ditemukan." };
  }
  
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getDataRange().getValues();
  
  var nisnIndex = headers.indexOf("NISN");
  var noIndex = headers.indexOf("No");
  var rowIndex = -1;
  
  for (var i = 1; i < values.length; i++) {
    if ((nisnIndex !== -1 && String(values[i][nisnIndex]).trim() === String(key).trim()) ||
        (noIndex !== -1 && String(values[i][noIndex]).trim() === String(key).trim())) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, error: "Data tidak ditemukan untuk dihapus." };
  }
  
  sheet.deleteRow(rowIndex);
  return { success: true, message: "Data berhasil dihapus dari " + sheetName };
}

/**
 * Mengambil ringkasan siswa terdaftar
 */
function getSiswaList(ss) {
  var res = readData(ss, "Daftar Siswa");
  if (!res.success) return res;
  
  var list = res.data.map(function(s) {
    return {
      NISN: s.NISN,
      Nama: s.Nama,
      Kelas: s.Kelas,
      NoOrangTua: s.NoOrangTua
    };
  });
  return { success: true, data: list };
}

/**
 * Mengambil ringkasan lengkap nilai siswa untuk dilaporkan ke orang tua (WhatsApp)
 */
function getStudentSummary(ss, nisn) {
  // 1. Ambil data Siswa
  var siswaRes = readData(ss, "Daftar Siswa");
  if (!siswaRes.success) return siswaRes;
  
  var student = null;
  for (var i = 0; i < siswaRes.data.length; i++) {
    if (String(siswaRes.data[i].NISN).trim() === String(nisn).trim()) {
      student = siswaRes.data[i];
      break;
    }
  }
  
  if (!student) {
    return { success: false, error: "Siswa dengan NISN " + nisn + " tidak ditemukan." };
  }
  
  // 2. Ambil nilai-nilai dari 7 mata pelajaran
  var subjects = [
    "Daftar Nilai Matematika",
    "Daftar Nilai B.Indonesia",
    "IPAS",
    "Pendidikan Pancasila",
    "Seni",
    "B Sunda",
    "Kokurikuler"
  ];
  
  var grades = {};
  var totalSikaps = 0;
  var totalKehadirans = 0;
  var totalCognitives = 0;
  var subjectScoreCount = 0;
  var sikapCount = 0;
  var kehadiranCount = 0;
  
  subjects.forEach(function(subject) {
    var res = readData(ss, subject);
    var subName = subject.replace("Daftar Nilai ", ""); // Matematika, B.Indonesia, IPAS, dll.
    
    grades[subName] = null;
    if (res.success && res.data) {
      for (var i = 0; i < res.data.length; i++) {
        var row = res.data[i];
        if (String(row.NISN).trim() === String(nisn).trim()) {
          // Parsing nilai kognitif ke numerik
          var harian = parseFloat(row["Nilai Harian/Tugas"]) || 0;
          var tp1 = parseFloat(row["Nilai TP1"]) || 0;
          var tp2 = parseFloat(row["TP2"]) || 0;
          var tp3 = parseFloat(row["TP3"]) || 0;
          var tp4 = parseFloat(row["TP4"]) || 0;
          var tp5 = parseFloat(row["TP5"]) || 0;
          var lm1 = parseFloat(row["Nilai LM1"]) || 0;
          var lm2 = parseFloat(row["LM2"]) || 0;
          var lm3 = parseFloat(row["LM3"]) || 0;
          var lm4 = parseFloat(row["LM4"]) || 0;
          var lm5 = parseFloat(row["LM5"]) || 0;
          var formatif = parseFloat(row["Nilai Formatif"]) || 0;
          var sumatif = parseFloat(row["Nilai Sumatif"]) || 0;
          var sikap = parseFloat(row["Nilai Sikap"]) || 0;
          var kehadiran = parseFloat(row["Nilai Kehadiran"]) || 0;
          
          // Hitung rata-rata TP
          var tpCount = 0, tpSum = 0;
          [tp1, tp2, tp3, tp4, tp5].forEach(function(v) { if (v > 0) { tpSum += v; tpCount++; } });
          var tpAvg = tpCount > 0 ? (tpSum / tpCount) : 0;
          
          // Hitung rata-rata LM (Lingkup Materi)
          var lmCount = 0, lmSum = 0;
          [lm1, lm2, lm3, lm4, lm5].forEach(function(v) { if (v > 0) { lmSum += v; lmCount++; } });
          var lmAvg = lmCount > 0 ? (lmSum / lmCount) : 0;
          
          // Akumulasi nilai kognitif rata-rata mata pelajaran
          var cognitValues = [];
          if (harian > 0) cognitValues.push(harian);
          if (tpAvg > 0) cognitValues.push(tpAvg);
          if (lmAvg > 0) cognitValues.push(lmAvg);
          if (formatif > 0) cognitValues.push(formatif);
          if (sumatif > 0) cognitValues.push(sumatif);
          
          var avgCognitive = 0;
          if (cognitValues.length > 0) {
            var sum = 0;
            cognitValues.forEach(function(v) { sum += v; });
            avgCognitive = sum / cognitValues.length;
          }
          
          grades[subName] = {
            harian: harian,
            tpAvg: tpAvg.toFixed(1),
            lmAvg: lmAvg.toFixed(1),
            formatif: formatif,
            sumatif: sumatif,
            sikap: sikap,
            kehadiran: kehadiran,
            rataRata: avgCognitive.toFixed(1),
            keterangan: row["Keterangan"] || "-"
          };
          
          if (avgCognitive > 0) {
            totalCognitives += avgCognitive;
            subjectScoreCount++;
          }
          if (sikap > 0) {
            totalSikaps += sikap;
            sikapCount++;
          }
          if (kehadiran > 0) {
            totalKehadirans += kehadiran;
            kehadiranCount++;
          }
          break;
        }
      }
    }
  });
  
  // 3. Ambil Catatan Wali Kelas
  var catatanRes = readData(ss, "Catatan Wali Kelas");
  var catatan = "-";
  if (catatanRes.success && catatanRes.data) {
    for (var i = 0; i < catatanRes.data.length; i++) {
      if (String(catatanRes.data[i].NISN).trim() === String(nisn).trim()) {
        catatan = catatanRes.data[i].Catatan || "-";
        break;
      }
    }
  }
  
  var classAverage = subjectScoreCount > 0 ? (totalCognitives / subjectScoreCount).toFixed(1) : "0";
  var sikapAverage = sikapCount > 0 ? (totalSikaps / sikapCount).toFixed(1) : "0";
  var kehadiranAverage = kehadiranCount > 0 ? (totalKehadirans / kehadiranCount).toFixed(1) : "0";
  
  return {
    success: true,
    data: {
      siswa: student,
      nilai: grades,
      catatanWaliKelas: catatan,
      ringkasan: {
        rataRataKognitif: classAverage,
        rataRataSikap: sikapAverage,
        rataRataKehadiran: kehadiranAverage
      }
    }
  };
}

/**
 * Menghitung statistik global untuk Dashboard
 */
function getDashboardStats(ss) {
  var siswaRes = readData(ss, "Daftar Siswa");
  if (!siswaRes.success) return siswaRes;
  
  var totalSiswa = siswaRes.data.length;
  
  var subjects = [
    "Daftar Nilai Matematika",
    "Daftar Nilai B.Indonesia",
    "IPAS",
    "Pendidikan Pancasila",
    "Seni",
    "B Sunda",
    "Kokurikuler"
  ];
  
  var totalCognitivesSum = 0;
  var totalScoreCount = 0;
  var subjectAverages = {};
  
  subjects.forEach(function(subject) {
    var res = readData(ss, subject);
    var subName = subject.replace("Daftar Nilai ", "");
    
    var subSum = 0;
    var subCount = 0;
    
    if (res.success && res.data) {
      res.data.forEach(function(row) {
        var harian = parseFloat(row["Nilai Harian/Tugas"]) || 0;
        var formatif = parseFloat(row["Nilai Formatif"]) || 0;
        var sumatif = parseFloat(row["Nilai Sumatif"]) || 0;
        
        // Hitung rata-rata TP (1 s.d 5)
        var tpCount = 0, tpSum = 0;
        ["Nilai TP1", "TP2", "TP3", "TP4", "TP5"].forEach(function(col) {
          var v = parseFloat(row[col]) || 0;
          if (v > 0) { tpSum += v; tpCount++; }
        });
        var tpAvg = tpCount > 0 ? (tpSum / tpCount) : 0;
        
        // Hitung rata-rata LM (1 s.d 5)
        var lmCount = 0, lmSum = 0;
        ["Nilai LM1", "LM2", "LM3", "LM4", "LM5"].forEach(function(col) {
          var v = parseFloat(row[col]) || 0;
          if (v > 0) { lmSum += v; lmCount++; }
        });
        var lmAvg = lmCount > 0 ? (lmSum / lmCount) : 0;
        
        var cognitiveTerms = [];
        if (harian > 0) cognitiveTerms.push(harian);
        if (tpAvg > 0) cognitiveTerms.push(tpAvg);
        if (lmAvg > 0) cognitiveTerms.push(lmAvg);
        if (formatif > 0) cognitiveTerms.push(formatif);
        if (sumatif > 0) cognitiveTerms.push(sumatif);
        
        if (cognitiveTerms.length > 0) {
          var avgCognitive = 0;
          cognitiveTerms.forEach(function(v) { avgCognitive += v; });
          avgCognitive = avgCognitive / cognitiveTerms.length;
          
          subSum += avgCognitive;
          subCount++;
        }
      });
    }
    
    var avgScore = subCount > 0 ? (subSum / subCount) : 0;
    subjectAverages[subName] = avgScore.toFixed(1);
    
    if (avgScore > 0) {
      totalCognitivesSum += avgScore;
      totalScoreCount++;
    }
  });
  
  var classGeneralAverage = totalScoreCount > 0 ? (totalCognitivesSum / totalScoreCount).toFixed(1) : "0";
  
  // Hitung jumlah bimbingan (Catatan wali kelas)
  var catatanRes = readData(ss, "Catatan Wali Kelas");
  var totalCatatan = 0;
  if (catatanRes.success && catatanRes.data) {
    catatanRes.data.forEach(function(row) {
      if (row.Catatan && row.Catatan.trim() !== "" && row.Catatan.trim() !== "-") {
        totalCatatan++;
      }
    });
  }
  
  return {
    success: true,
    data: {
      totalSiswa: totalSiswa,
      rataRataKelas: classGeneralAverage,
      totalBimbingan: totalCatatan,
      grafikRataRata: subjectAverages
    }
  };
}
