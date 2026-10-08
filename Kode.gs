/**
 * ==============================================================================
 * APLIKASI WEB IURAN ANGGOTA & KAS SETORAN
 * File: Code.gs (Backend Engine) - Fixed Version
 * ==============================================================================
 */

// ID Spreadsheet Utama
const SPREADSHEET_ID = "1-K4mHbMoLgxkFtSZKZk9hiapmyqTKlaI0_7w4s2AWHg";

// Global Sheet Name Constants
const SHEET_ANGGOTA = "Master_Anggota";
const SHEET_PAGU = "Master_Pagu";
const SHEET_MASUK = "Transaksi_Masuk";
const SHEET_SETOR = "Transaksi_Setor";
const SHEET_PENGATURAN = "Pengaturan";

/**
 * FUNGSI OTORISASI PERIZINAN GOOGLE DRIVE & SPREADSHEET
 * Jalankan fungsi ini sekali di Editor Apps Script jika memerlukan dialog izin.
 */
function dummyAuthorizeDriveScope() {
  var folder = DriveApp.getFolderById("1fHDGNAMQFtmcOCl3oId-R7rF2seo25_g");
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var testBlob = Utilities.newBlob("test", "text/plain", "auth_test.txt");
  var tempFile = folder.createFile(testBlob);
  tempFile.setTrashed(true);
  Logger.log("Drive Access Authorized Successfully: " + folder.getName() + " - " + ss.getName());
}

/**
 * FUNGSI PEMBANTU: Membuka Spreadsheet Secara Pasti & Aman
 */
function getSS() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getId()) return ss;
  } catch (e) {
    // Apabila berjalan di konteks Web App
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * 1. FUNGSI UTAMA WEB APP (ENTRY POINT)
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle("Dana Sosial Setia Kawan Kecamatan Secang")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper untuk menyertakan file HTML lain jika diperlukan
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 2. AUTENTIKASI ADMIN & PENGATURAN
 */
function verifyAdminPin(pin) {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName(SHEET_PENGATURAN);
    
    if (!sheet) {
      return { success: false, message: "Tab 'Pengaturan' tidak ditemukan di Spreadsheet!" };
    }
    
    var data = sheet.getDataRange().getValues();
    var pinSaved = "123456"; // Fallback default
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "PIN_ADMIN") {
        pinSaved = String(data[i][1]).trim();
        break;
      }
    }
    
    if (String(pin).trim() === pinSaved) {
      return { success: true, message: "Login Berhasil" };
    } else {
      return { success: false, message: "PIN Admin Salah!" };
    }
  } catch (err) {
    return { success: false, message: "Terjadi kesalahan: " + err.message };
  }
}

/**
 * 3. PEMBACAAN DATA MATRIKS PAGU
 */
function getMatrixPaguMap() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEET_PAGU);
  if (!sheet) return {};
  
  var data = sheet.getDataRange().getValues();
  var paguMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var tahun = data[i][0];
    if (!tahun) continue;
    for (var bln = 1; bln <= 12; bln++) {
      var key = tahun + "_" + bln;
      paguMap[key] = Number(data[i][bln]) || 0;
    }
  }
  return paguMap;
}

/**
 * HELPER: Menghitung Detail Status & Tunggakan Anggota
 */
function buildDetailStatusAnggota(foundRow, paguMap) {
  var ss = getSS();
  var idAnggota = foundRow[0];
  var namaAnggota = foundRow[1];
  var nip = foundRow[2];
  var cutoffBayarBln = Number(foundRow[8]) || 0;
  var cutoffBayarThn = Number(foundRow[9]) || 0;

  paguMap = paguMap || getMatrixPaguMap();
  var now = new Date();
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth() + 1; // 1 - 12

  var startThn = cutoffBayarThn;
  var startBln = cutoffBayarBln + 1;

  if (cutoffBayarThn === 0 || cutoffBayarBln === 0) {
    var paguYears = Object.keys(paguMap).map(function(k) { return parseInt(k.split("_")[0]); }).filter(Boolean);
    startThn = paguYears.length > 0 ? Math.min.apply(null, paguYears) : currentYear;
    startBln = 1;
  } else if (startBln > 12) {
    startBln = 1;
    startThn = cutoffBayarThn + 1;
  }

  var totalTunggakan = 0;
  var rincianPerTahun = {};

  var currY = startThn;
  var currM = startBln;

  while (currY < currentYear || (currY === currentYear && currM <= currentMonth)) {
    var key = currY + "_" + currM;
    var paguNominal = Number(paguMap[key]);
    if (isNaN(paguNominal) || paguNominal <= 0) {
      paguNominal = 30000;
    }

    totalTunggakan += paguNominal;

    if (!rincianPerTahun[currY]) {
      rincianPerTahun[currY] = {
        tahun: currY,
        total: 0,
        bulanDari: currM,
        bulanSampai: currM,
        detailBulan: []
      };
    }
    rincianPerTahun[currY].total += paguNominal;
    rincianPerTahun[currY].bulanSampai = currM;
    rincianPerTahun[currY].detailBulan.push({
      bulan: currM,
      namaBulan: getNamaBulan(currM),
      nominal: paguNominal
    });

    currM++;
    if (currM > 12) {
      currM = 1;
      currY++;
    }
  }

  var rincianTunggakanList = Object.keys(rincianPerTahun).map(function(y) {
    var r = rincianPerTahun[y];
    var rentangBulan = (r.bulanDari === r.bulanSampai) 
      ? getNamaBulan(r.bulanDari) 
      : (getNamaBulan(r.bulanDari) + " - " + getNamaBulan(r.bulanSampai));
    return {
      tahun: r.tahun,
      total: r.total,
      rentangBulan: rentangBulan,
      detailBulan: r.detailBulan
    };
  });

  var tglLahirVal = "-";
  if (foundRow[3]) {
    try {
      var dTgl = (foundRow[3] instanceof Date) ? foundRow[3] : new Date(foundRow[3]);
      if (!isNaN(dTgl.getTime())) {
        tglLahirVal = dTgl.getDate() + " " + getNamaBulan(dTgl.getMonth() + 1) + " " + dTgl.getFullYear();
      } else {
        tglLahirVal = String(foundRow[3]);
      }
    } catch(eTgl) {
      tglLahirVal = String(foundRow[3]);
    }
  }

  var tmtVal = "-";
  if (foundRow[5]) {
    try {
      var dTmt = (foundRow[5] instanceof Date) ? foundRow[5] : new Date(foundRow[5]);
      if (!isNaN(dTmt.getTime())) {
        tmtVal = dTmt.getDate() + " " + getNamaBulan(dTmt.getMonth() + 1) + " " + dTmt.getFullYear();
      } else {
        tmtVal = String(foundRow[5]);
      }
    } catch(eTmt) {
      tmtVal = String(foundRow[5]);
    }
  }

  // Calculate status string
  var statusOverride = foundRow[12] ? String(foundRow[12]).trim() : "";
  var statusCalculated = "";
  if (statusOverride === "Wafat" || statusOverride === "Mengundurkan Diri" || statusOverride === "Keluar") {
    statusCalculated = statusOverride;
  } else if (cutoffBayarThn >= currentYear) {
    statusCalculated = "Aktif";
  } else if (cutoffBayarThn === (currentYear - 1)) {
    statusCalculated = "Aktif " + cutoffBayarThn;
  } else if (cutoffBayarThn > 0) {
    statusCalculated = "Aktif " + cutoffBayarThn;
  } else {
    statusCalculated = "Belum Bayar";
  }

  return {
    id: idAnggota,
    nama: namaAnggota,
    nip: nip ? String(nip) : "-",
    tglLahir: tglLahirVal,
    alamat: foundRow[4] ? String(foundRow[4]) : "-",
    tmt: tmtVal,
    noHp: foundRow[6] ? String(foundRow[6]) : "-",
    noKode: foundRow[13] ? String(foundRow[13]) : "-",
    statusStr: statusCalculated,
    lunasSampai: (cutoffBayarBln > 0 && cutoffBayarThn > 0) ? (getNamaBulan(cutoffBayarBln) + " " + cutoffBayarThn) : "Belum Ada Catatan",
    totalTunggakan: totalTunggakan,
    rincianTunggakan: rincianTunggakanList,
    bulanIniStr: getNamaBulan(currentMonth) + " " + currentYear
  };
}

/**
 * 4. PENCARIAN & PEMERIKSAAN STATUS ANGGOTA (PUBLIK / USER)
 */
function checkStatusAnggota(keyword) {
  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    
    if (!sheetAnggota) {
      return { success: false, message: "Tab Master_Anggota tidak ditemukan." };
    }
    
    var dataAnggota = sheetAnggota.getDataRange().getValues();
    if (dataAnggota.length <= 1) {
      return { success: false, message: "Data anggota masih kosong." };
    }
    
    var matchedRows = [];
    var searchStr = String(keyword).trim().toLowerCase();
    
    for (var i = 1; i < dataAnggota.length; i++) {
      var id = String(dataAnggota[i][0]).toLowerCase();
      var nama = String(dataAnggota[i][1]).toLowerCase();
      var nip = String(dataAnggota[i][2]).toLowerCase();
      
      if (id === searchStr || nip === searchStr || nama.indexOf(searchStr) !== -1) {
        matchedRows.push(dataAnggota[i]);
      }
    }
    
    if (matchedRows.length === 0) {
      return { success: false, message: "Anggota tidak ditemukan. Periksa kembali nama/NIP yang diketik." };
    }
    
    var paguMap = getMatrixPaguMap();

    // Jika hanya ditemukan 1 orang
    if (matchedRows.length === 1) {
      return {
        success: true,
        type: "single",
        data: buildDetailStatusAnggota(matchedRows[0], paguMap)
      };
    }

    // Jika ditemukan beberapa orang yang cocok (multiple matches)
    var listHasil = matchedRows.map(function(row) {
      return {
        id: row[0],
        nama: row[1],
        nip: row[2] ? String(row[2]) : "Tanpa NIP",
        lunasSampai: (Number(row[8]) > 0 && Number(row[9]) > 0) ? (getNamaBulan(Number(row[8])) + " " + Number(row[9])) : "Belum Ada Catatan"
      };
    });

    return {
      success: true,
      type: "multiple",
      count: listHasil.length,
      keyword: keyword,
      list: listHasil
    };

  } catch (err) {
    return { success: false, message: "Gagal memuat status: " + err.message };
  }
}

/**
 * 4B. AMBIL DETAIL ANGGOTA BY ID (DARI DAFTAR PILIHAN NAMA)
 */
function getDetailAnggotaById(idAnggota) {
  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheetAnggota) return { success: false, message: "Tab Master_Anggota tidak ditemukan." };
    
    var dataAnggota = sheetAnggota.getDataRange().getValues();
    var foundRow = null;
    
    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        foundRow = dataAnggota[i];
        break;
      }
    }
    
    if (!foundRow) return { success: false, message: "Data anggota tidak ditemukan." };
    
    return {
      success: true,
      data: buildDetailStatusAnggota(foundRow, getMatrixPaguMap())
    };
  } catch (err) {
    return { success: false, message: "Gagal memuat detail anggota: " + err.message };
  }
}

/**
 * 5. MENGAMBIL DAFTAR SELURUH ANGGOTA (UNTUK DROPDOWN/ADMIN)
 */
function getAllAnggota() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEET_ANGGOTA);
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var tglLahirVal = "";
      if (data[i][3]) {
        if (data[i][3] instanceof Date) {
          tglLahirVal = Utilities.formatDate(data[i][3], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } else {
          tglLahirVal = String(data[i][3]);
        }
      }

      var tmtVal = "";
      if (data[i][5]) {
        if (data[i][5] instanceof Date) {
          tmtVal = Utilities.formatDate(data[i][5], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        } else {
          tmtVal = String(data[i][5]);
        }
      }

      result.push({
        id: data[i][0],
        nama: data[i][1],
        nip: data[i][2] ? String(data[i][2]) : "",
        tglLahir: tglLahirVal,
        alamat: data[i][4] ? String(data[i][4]) : "",
        tmt: tmtVal,
        noHp: data[i][6] ? String(data[i][6]) : "",
        saldoDeposit: Number(data[i][7]) || 0,
        bayarBln: Number(data[i][8]) || 0,
        bayarThn: Number(data[i][9]) || 0,
        setorBln: Number(data[i][10]) || 0,
        setorThn: Number(data[i][11]) || 0,
        statusOverride: data[i][12] ? String(data[i][12]).trim() : "",
        noKode: data[i][13] ? String(data[i][13]) : ""
      });
    }
  }
  return result;
}

/**
  * FUNGSI SIMPAN ANGGOTA BARU (CREATE)
  */
function simpanAnggotaBaru(form) {
  try {
    var authCheck = verifyAdminPin(form ? form.adminPin : "");
    if (!authCheck.success) {
      return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };
    }

    var ss = getSS();
    var sheet = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheet) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var data = sheet.getDataRange().getValues();
    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      var valId = parseInt(data[i][0], 10);
      if (!isNaN(valId) && valId > maxId) maxId = valId;
    }
    var newId = maxId + 1;

    // Col A (0): ID, B (1): Nama, C (2): NIP, D (3): Tgl Lahir, E (4): Alamat, F (5): TMT, G (6): No HP,
    // H (7): Saldo (0), I (8): BayarBln (0), J (9): BayarThn (0), K (10): SetorBln (0), L (11): SetorThn (0),
    // M (12): Status Override, N (13): No Kode
    sheet.appendRow([
      newId,
      form.nama || "",
      form.nip || "",
      form.tglLahir || "",
      form.alamat || "",
      form.tmt || "",
      form.noHp || "",
      0, 0, 0, 0, 0,
      form.statusOverride || "",
      form.noKode || ""
    ]);

    return { success: true, message: "Anggota baru berhasil ditambahkan! (ID: " + newId + ")" };
  } catch (err) {
    return { success: false, message: "Gagal menyimpan anggota: " + err.message };
  }
}

/**
  * FUNGSI UPDATE ANGGOTA (UPDATE)
  */
function updateAnggota(form) {
  try {
    var authCheck = verifyAdminPin(form ? form.adminPin : "");
    if (!authCheck.success) {
      return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };
    }

    var ss = getSS();
    var sheet = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheet) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(form.id)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }

    // Update Kolom B..G (Index 2..7 di spreadsheet)
    sheet.getRange(targetRow, 2).setValue(form.nama || "");
    sheet.getRange(targetRow, 3).setValue(form.nip || "");
    sheet.getRange(targetRow, 4).setValue(form.tglLahir || "");
    sheet.getRange(targetRow, 5).setValue(form.alamat || "");
    sheet.getRange(targetRow, 6).setValue(form.tmt || "");
    sheet.getRange(targetRow, 7).setValue(form.noHp || "");

    // Update Kolom M (13): Status Override, Kolom N (14): No Kode
    sheet.getRange(targetRow, 13).setValue(form.statusOverride || "");
    sheet.getRange(targetRow, 14).setValue(form.noKode || "");

    return { success: true, message: "Data Anggota berhasil diperbarui!" };
  } catch (err) {
    return { success: false, message: "Gagal memperbarui data: " + err.message };
  }
}

/**
  * FUNGSI HAPUS ANGGOTA (DELETE)
  */
function hapusAnggota(pin, idAnggota) {
  try {
    var authCheck = verifyAdminPin(pin);
    if (!authCheck.success) {
      return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };
    }

    var ss = getSS();
    var sheet = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheet) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idAnggota)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }

    sheet.deleteRow(targetRow);
    return { success: true, message: "Anggota ID " + idAnggota + " berhasil dihapus!" };
  } catch (err) {
    return { success: false, message: "Gagal menghapus anggota: " + err.message };
  }
}

/**
 * 6. PROSES SIMPAN TRANSAKSI MASUK (BAYAR IURAN)
 */
function simpanTransaksiMasuk(form) {
  try {
    // Verifikasi Otorisasi PIN Admin
    var authCheck = verifyAdminPin(form ? form.adminPin : "");
    if (!authCheck.success) {
      return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };
    }

    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    var sheetMasuk = ss.getSheetByName(SHEET_MASUK);
    
    if (!sheetAnggota || !sheetMasuk) {
      return { success: false, message: "Lembar data Master_Anggota atau Transaksi_Masuk tidak ditemukan!" };
    }
    
    var idAnggota = form.idAnggota;
    var nominalDiterima = Number(form.nominalDiterima) || 0;
    var tahunDari = Number(form.tahunDari || form.tahunPeriode);
    var bulanDari = Number(form.bulanDari);
    var tahunSampai = Number(form.tahunSampai || form.tahunPeriode);
    var bulanSampai = Number(form.bulanSampai);
    var iuranTerpakai = Number(form.iuranTerpakai) || 0;
    var kembalianCash = Number(form.kembalianCash) || 0;
    var masukDeposit = Number(form.masukDeposit) || 0;
    var catatan = form.catatan || "-";
    
    var tgl = new Date();
    var tglStr = Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    var idTrx = "TRX-" + Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyyMMdd") + "-" + Math.floor(100 + Math.random() * 900);
    
    var dataAnggota = sheetAnggota.getDataRange().getValues();
    var namaAnggota = "";
    var nipAnggota = "-";
    var targetRowIdx = -1;
    var currentDeposit = 0;
    
    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        namaAnggota = dataAnggota[i][1];
        nipAnggota = dataAnggota[i][2] ? String(dataAnggota[i][2]) : "-";
        currentDeposit = Number(dataAnggota[i][7]) || 0;
        targetRowIdx = i + 1;
        break;
      }
    }
    
    if (targetRowIdx === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }
    
    var jmlBulan = (tahunSampai - tahunDari) * 12 + (bulanSampai - bulanDari) + 1;
    var rincianBulanStr = "Iuran Dana Sosial Setia Kawan Bulan " + getNamaBulan(bulanDari) + " " + tahunDari + 
      ((bulanDari !== bulanSampai || tahunDari !== tahunSampai) ? (" s.d. " + getNamaBulan(bulanSampai) + " " + tahunSampai) : "") + 
      " (" + jmlBulan + " bulan)";

    var periodeStr = (tahunDari === tahunSampai) 
      ? (tahunDari + " (" + getNamaBulan(bulanDari) + " - " + getNamaBulan(bulanSampai) + ")")
      : (getNamaBulan(bulanDari) + " " + tahunDari + " - " + getNamaBulan(bulanSampai) + " " + tahunSampai);

    sheetMasuk.appendRow([
      idTrx, tglStr, idAnggota, namaAnggota, nominalDiterima,
      tahunSampai, bulanDari, bulanSampai, iuranTerpakai,
      kembalianCash, masukDeposit, catatan + " [" + periodeStr + "]"
    ]);
    
    var newDeposit = currentDeposit + masukDeposit;
    // Batch update: Saldo Deposit (kolom 8), Cutoff Bayar Bulan (kolom 9), Cutoff Bayar Tahun (kolom 10)
    sheetAnggota.getRange(targetRowIdx, 8, 1, 3).setValues([[newDeposit, bulanSampai, tahunSampai]]);
    
    return { 
      success: true, 
      message: "Transaksi Pembayaran Berhasil Disimpan (" + periodeStr + ")! ID: " + idTrx,
      kuitansiData: {
        idTrx: idTrx,
        tgl: tglStr,
        nama: namaAnggota,
        nip: nipAnggota,
        nominal: nominalDiterima,
        rincianBulan: rincianBulanStr
      }
    };
  } catch (err) {
    return { success: false, message: "Gagal menyimpan transaksi: " + err.message };
  }
}

/**
 * 7. PROSES SIMPAN TRANSAKSI SETORAN KAS (KE BENDAHARA)
 */
function simpanTransaksiSetor(form) {
  try {
    // Verifikasi Otorisasi PIN Admin
    var authCheck = verifyAdminPin(form ? form.adminPin : "");
    if (!authCheck.success) {
      return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };
    }

    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    var sheetSetor = ss.getSheetByName(SHEET_SETOR);
    
    if (!sheetAnggota || !sheetSetor) {
      return { success: false, message: "Lembar data Master_Anggota atau Transaksi_Setor tidak ditemukan!" };
    }
    
    var idAnggota = form.idAnggota;
    var nominalDisetor = Number(form.nominalDisetor) || 0;
    var tahunDari = Number(form.tahunDari || form.tahunPeriode);
    var bulanDari = Number(form.bulanDari);
    var tahunSampai = Number(form.tahunSampai || form.tahunPeriode);
    var bulanSampai = Number(form.bulanSampai);
    var penerimaSetor = form.penerimaSetor || "Bendahara";
    var catatan = form.catatan || "-";
    
    var tgl = new Date();
    var tglStr = Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    var idSetor = "STR-" + Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyyMMdd") + "-" + Math.floor(100 + Math.random() * 900);
    
    var dataAnggota = sheetAnggota.getDataRange().getValues();
    var namaAnggota = "";
    var targetRowIdx = -1;
    
    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        namaAnggota = dataAnggota[i][1];
        targetRowIdx = i + 1;
        break;
      }
    }
    
    if (targetRowIdx === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }
    
    var periodeStr = (tahunDari === tahunSampai) 
      ? (tahunDari + " (" + getNamaBulan(bulanDari) + " - " + getNamaBulan(bulanSampai) + ")")
      : (getNamaBulan(bulanDari) + " " + tahunDari + " - " + getNamaBulan(bulanSampai) + " " + tahunSampai);

    sheetSetor.appendRow([
      idSetor, tglStr, idAnggota, namaAnggota, nominalDisetor,
      tahunSampai, bulanDari, bulanSampai, penerimaSetor, catatan + " [" + periodeStr + "]"
    ]);
    
    // Batch update: Cutoff Setor Bulan (kolom 11), Cutoff Setor Tahun (kolom 12)
    sheetAnggota.getRange(targetRowIdx, 11, 1, 2).setValues([[bulanSampai, tahunSampai]]);
    
    return { success: true, message: "Setoran Kas Berhasil Dicatat (" + periodeStr + ")! ID: " + idSetor };
  } catch (err) {
    return { success: false, message: "Gagal mencatat setoran: " + err.message };
  }
}

/**
 * 8. MENAMPILKAN DATA LENGKAP ANGGOTA UNTUK TAB DATA ANGGOTA ADMIN
 */
function getAdminDataAnggota(pin) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var list = getAllAnggota();
    return { success: true, data: list };
  } catch (err) {
    return { success: false, message: "Gagal mengambil data anggota: " + err.message };
  }
}

/**
 * 9. REKAP LAPORAN LENGKAP KAS & IURAN UNTUK CETAK LAPORAN
 */
function getAdminRekapLaporan(pin, filterTahun) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetMasuk = ss.getSheetByName(SHEET_MASUK);
    var sheetSetor = ss.getSheetByName(SHEET_SETOR);
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);

    var nipMap = {};
    if (sheetAnggota) {
      var dataA = sheetAnggota.getDataRange().getValues();
      for (var k = 1; k < dataA.length; k++) {
        if (dataA[k][0]) {
          nipMap[String(dataA[k][0])] = dataA[k][2] ? String(dataA[k][2]) : "-";
        }
      }
    }
    
    var listMasuk = [];
    var totalIuranMasuk = 0;
    if (sheetMasuk) {
      var dataM = sheetMasuk.getDataRange().getValues();
      for (var i = 1; i < dataM.length; i++) {
        var r = dataM[i];
        if (!r[0]) continue;
        var thn = Number(r[5]);
        if (!filterTahun || filterTahun === "semua" || thn === Number(filterTahun)) {
          var nom = Number(r[4]) || 0;
          totalIuranMasuk += nom;
          var bDari = Number(r[6]);
          var bSampai = Number(r[7]);
          var jmlBulan = (bSampai - bDari) + 1;
          var rincianBlnStr = "Iuran Dana Sosial Setia Kawan Bulan " + getNamaBulan(bDari) + " " + thn + (bDari !== bSampai ? (" s.d. " + getNamaBulan(bSampai) + " " + thn) : "") + " (" + jmlBulan + " bulan)";

          listMasuk.push({
            idTrx: r[0],
            tanggal: r[1] ? Utilities.formatDate(new Date(r[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "-",
            idAnggota: r[2],
            nama: r[3],
            nip: nipMap[String(r[2])] || "-",
            nominal: nom,
            tahun: thn,
            periode: getNamaBulan(bDari) + " - " + getNamaBulan(bSampai),
            rincianBulan: rincianBlnStr,
            catatan: r[11] || "-"
          });
        }
      }
    }

    var listSetor = [];
    var totalSetorKas = 0;
    if (sheetSetor) {
      var dataS = sheetSetor.getDataRange().getValues();
      for (var j = 1; j < dataS.length; j++) {
        var s = dataS[j];
        if (!s[0]) continue;
        var thnS = Number(s[5]);
        if (!filterTahun || filterTahun === "semua" || thnS === Number(filterTahun)) {
          var nomS = Number(s[4]) || 0;
          totalSetorKas += nomS;
          listSetor.push({
            idSetor: s[0],
            tanggal: s[1] ? Utilities.formatDate(new Date(s[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "-",
            idAnggota: s[2],
            nama: s[3],
            nominal: nomS,
            tahun: thnS,
            periode: getNamaBulan(s[6]) + " - " + getNamaBulan(s[7]),
            penerima: s[8] || "-",
            catatan: s[9] || "-"
          });
        }
      }
    }

    return {
      success: true,
      summary: {
        totalIuranMasuk: totalIuranMasuk,
        totalSetorKas: totalSetorKas,
        saldoKasDisimpan: totalIuranMasuk - totalSetorKas,
        jumlahTransaksiMasuk: listMasuk.length,
        jumlahSetoranKas: listSetor.length
      },
      transaksiMasuk: listMasuk,
      transaksiSetor: listSetor
    };
  } catch (err) {
    return { success: false, message: "Gagal membuat rekap laporan: " + err.message };
  }
}

/**
 * 9B. AMBIL RIWAYAT PEMBAYARAN IURAN LENGKAP UNTUK MENU RIWAYAT
 */
function getRiwayatPembayaranAdmin(pin, filterTahun) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetMasuk = ss.getSheetByName(SHEET_MASUK);
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);

    var nipMap = {};
    if (sheetAnggota) {
      var dataA = sheetAnggota.getDataRange().getValues();
      for (var k = 1; k < dataA.length; k++) {
        if (dataA[k][0]) {
          nipMap[String(dataA[k][0])] = dataA[k][2] ? String(dataA[k][2]) : "-";
        }
      }
    }

    // Map file PDF tersimpan di Drive folder berdasarkan ID Trx
    var pdfMap = {};
    try {
      var folder = DriveApp.getFolderById(KUITANSI_FOLDER_ID);
      if (folder) {
        var files = folder.getFiles();
        while (files.hasNext()) {
          var f = files.next();
          var fn = f.getName(); // Format: Kuitansi_TRX-20261007-245_COBA_DANSOS.pdf
          var parts = fn.split("_");
          if (parts.length >= 2) {
            // parts[1] adalah ID TRX (misal TRX-20261007-245)
            var cleanTrxId = parts[1].trim().toUpperCase();
            pdfMap[cleanTrxId] = f.getUrl();
          }
        }
      }
    } catch(eDrive) {}

    var listRiwayat = [];
    var totalNominal = 0;

    if (sheetMasuk) {
      var dataM = sheetMasuk.getDataRange().getValues();
      for (var i = dataM.length - 1; i >= 1; i--) { // Urutan terbaru di atas
        var r = dataM[i];
        if (!r[0]) continue;
        var thn = Number(r[5]);
        if (!filterTahun || filterTahun === "semua" || thn === Number(filterTahun)) {
          var nom = Number(r[4]) || 0;
          totalNominal += nom;
          var bDari = Number(r[6]);
          var bSampai = Number(r[7]);
          var jmlBulan = (bSampai - bDari) + 1;
          var rincianBlnStr = "Iuran Dana Sosial Setia Kawan Bulan " + getNamaBulan(bDari) + " " + thn + (bDari !== bSampai ? (" s.d. " + getNamaBulan(bSampai) + " " + thn) : "") + " (" + jmlBulan + " bulan)";

          var idTrxStr = String(r[0]);
          listRiwayat.push({
            idTrx: idTrxStr,
            tanggal: r[1] ? Utilities.formatDate(new Date(r[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "-",
            idAnggota: r[2],
            nama: r[3],
            nip: nipMap[String(r[2])] || "-",
            nominal: nom,
            tahun: thn,
            periode: getNamaBulan(bDari) + " - " + getNamaBulan(bSampai),
            rincianBulan: rincianBlnStr,
            catatan: r[11] || "-",
            pdfUrl: pdfMap[idTrxStr.toUpperCase()] || null
          });
        }
      }
    }

    return {
      success: true,
      totalNominal: totalNominal,
      count: listRiwayat.length,
      data: listRiwayat
    };
  } catch (err) {
    return { success: false, message: "Gagal mengambil riwayat pembayaran: " + err.message };
  }
}

/**
 * 9C. AMBIL MATRIX DATA PEMBAYARAN IURAN PER ANGGOTA PER BULAN
 */
function getDataPembayaranMatrix(pin, targetTahun) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    var sheetMasuk = ss.getSheetByName(SHEET_MASUK);

    if (!sheetAnggota) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var now = new Date();
    var currentYear = now.getFullYear();
    var selTahun = Number(targetTahun) || currentYear;

    var paguMap = getMatrixPaguMap();

    // 1. Ambil data seluruh anggota
    var dataA = sheetAnggota.getDataRange().getValues();
    var memberMap = {};
    var memberOrder = [];

    for (var i = 1; i < dataA.length; i++) {
      var row = dataA[i];
      var id = String(row[0]);
      if (!id) continue;

      var cutoffBln = Number(row[8]) || 0;
      var cutoffThn = Number(row[9]) || 0;

      memberMap[id] = {
        id: id,
        nama: row[1] || "-",
        nip: row[2] ? String(row[2]) : "-",
        noKode: row[13] ? String(row[13]) : "-",
        tmt: row[5] ? (row[5] instanceof Date ? Utilities.formatDate(row[5], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : String(row[5])) : "-",
        cutoffBln: cutoffBln,
        cutoffThn: cutoffThn,
        bulanMap: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 }
      };
      memberOrder.push(id);
    }

    // 2. Hitung nominal terbayar per bulan per anggota
    // A. Terlebih dahulu isi dari Master_Anggota (cutoffBln & cutoffThn) berdasarkan Master_Pagu
    for (var mId in memberMap) {
      var mem = memberMap[mId];
      if (mem.cutoffThn > 0 && mem.cutoffBln > 0) {
        // Jika cutoffThn > selTahun: Berarti anggota SUDAH LUNAS PENUH untuk seluruh 12 bulan di selTahun
        if (mem.cutoffThn > selTahun) {
          for (var b = 1; b <= 12; b++) {
            var pKey = selTahun + "_" + b;
            mem.bulanMap[b] = Number(paguMap[pKey]) || 30000;
          }
        } 
        // Jika cutoffThn === selTahun: Berarti anggota sudah lunas dari bulan 1 s/d cutoffBln di selTahun
        else if (mem.cutoffThn === selTahun) {
          for (var b = 1; b <= mem.cutoffBln; b++) {
            var pKey = selTahun + "_" + b;
            mem.bulanMap[b] = Number(paguMap[pKey]) || 30000;
          }
        }
      }
    }

    // B. Tambahkan/timpa dengan rincian transaksi riil dari Transaksi_Masuk jika ada
    if (sheetMasuk) {
      var dataM = sheetMasuk.getDataRange().getValues();
      for (var j = 1; j < dataM.length; j++) {
        var m = dataM[j];
        if (!m[0]) continue;
        var idAng = String(m[2]);
        var nominalTrx = Number(m[4]) || 0;
        var thnSampai = Number(m[5]);
        var bDari = Number(m[6]);
        var bSampai = Number(m[7]);

        if (!memberMap[idAng] || nominalTrx <= 0) continue;

        var startYear = thnSampai;
        var startMonth = bDari;
        var totalMonths = 0;

        if (bDari <= bSampai) {
          totalMonths = (bSampai - bDari) + 1;
          startYear = thnSampai;
        } else {
          startYear = thnSampai - 1;
          totalMonths = (12 - bDari + 1) + bSampai;
        }

        var curY = startYear;
        var curM = startMonth;
        for (var step = 0; step < totalMonths; step++) {
          if (curY === selTahun && curM >= 1 && curM <= 12) {
            // Nilai isian bulan disesuaikan dengan pagu tarif di Master_Pagu untuk bulan/tahun tersebut
            var pKey = selTahun + "_" + curM;
            var monthPagu = Number(paguMap[pKey]) || 30000;
            memberMap[idAng].bulanMap[curM] = monthPagu;
          }
          curM++;
          if (curM > 12) {
            curM = 1;
            curY++;
          }
        }
      }
    }

    // 3. Susun data baris dan hitung total per bulan
    var rows = [];
    var totalPerBulan = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    var grandTotal = 0;

    for (var k = 0; k < memberOrder.length; k++) {
      var id = memberOrder[k];
      var m = memberMap[id];
      var lastPaidStr = (m.cutoffBln > 0 && m.cutoffThn > 0) ? (getNamaBulan(m.cutoffBln) + " " + m.cutoffThn) : "Belum Ada";

      var rowTotal = 0;
      var bulanVal = {};
      for (var b = 1; b <= 12; b++) {
        var val = m.bulanMap[b] || 0;
        bulanVal[b] = val;
        rowTotal += val;
        totalPerBulan[b] += val;
      }
      grandTotal += rowTotal;

      rows.push({
        no: k + 1,
        id: m.id,
        nama: m.nama,
        nip: m.nip,
        noKode: m.noKode,
        tmt: m.tmt,
        pembayaranTerakhir: lastPaidStr,
        jan: bulanVal[1],
        feb: bulanVal[2],
        mar: bulanVal[3],
        apr: bulanVal[4],
        mei: bulanVal[5],
        jun: bulanVal[6],
        jul: bulanVal[7],
        agt: bulanVal[8],
        sep: bulanVal[9],
        okt: bulanVal[10],
        nov: bulanVal[11],
        des: bulanVal[12],
        jumlah: rowTotal
      });
    }

    return {
      success: true,
      tahun: selTahun,
      currentYear: currentYear,
      data: rows,
      totalPerBulan: totalPerBulan,
      grandTotal: grandTotal
    };
  } catch (err) {
    return { success: false, message: "Gagal mengambil data pembayaran: " + err.message };
  }
}

/**
 * 10. TAMBAH ANGGOTA BARU (KHUSUS ADMIN)
 */
function simpanAnggotaBaru(form) {
  var authCheck = verifyAdminPin(form ? form.adminPin : "");
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheetAnggota) return { success: false, message: "Sheet Master_Anggota tidak ada!" };

    var data = sheetAnggota.getDataRange().getValues();
    var nextId = "ANG-" + String(data.length).padStart(3, '0');
    
    sheetAnggota.appendRow([
      nextId, form.nama, form.nip || "-", form.gender || "-", form.tglLahir || "-", form.alamat || "-", form.noHp || "-", 0, 0, 0, 0, 0
    ]);

    return { success: true, message: "Anggota baru berhasil ditambahkan! ID: " + nextId };
  } catch (err) {
    return { success: false, message: "Gagal menambah anggota: " + err.message };
  }
}

/**
 * HELPER: Nama Bulan Bahasa Indonesia
 */
function getNamaBulan(index) {
  var bulanMap = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return bulanMap[index] || "";
}

/**
 * 11. GENERATE PDF KUITANSI (19cm x 10cm) & SIMPAN KE GOOGLE DRIVE FOLDER
 */
const KUITANSI_FOLDER_ID = "1fHDGNAMQFtmcOCl3oId-R7rF2seo25_g";

function generateAndSaveKuitansiPDF(data) {
  try {
    data = data || {};
    var idTrx = data.idTrx || "TRX-TEST";
    var nama = data.nama || "Tes Anggota";
    var nip = data.nip || "123456";
    var nominal = data.nominal || 90000;
    var terbilangTeks = data.terbilangTeks || "Sembilan puluh ribu rupiah";
    var rincianTeks = data.rincianTeks || "iuran dansos pensiun bulan Januari s/d Maret 2026.";
    var tglFormatted = data.tglFormatted || "1 April 2026";
    var penandatangan = data.penandatangan || "Puji Purnomo";
    var isKwitansiMode = (data.printMode === 'kwitansi');

    var folder = DriveApp.getFolderById(KUITANSI_FOLDER_ID);
    if (!folder) {
      return { success: false, message: "Folder Google Drive penyimpan PDF tidak ditemukan!" };
    }

    var filename = "Kuitansi_" + idTrx + "_" + nama.replace(/[^a-zA-Z0-9]/g, "_") + ".pdf";

    var htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          @page {
            size: 190mm 100mm;
            margin: 0;
          }
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            width: 190mm;
            height: 100mm;
            position: relative;
            box-sizing: border-box;
            color: #0f172a;
            ${isKwitansiMode ? 'border: none;' : 'border: 2px solid #1F4E79;'}
          }
          /* Label HVS Static */
          .kw-label {
            display: ${isKwitansiMode ? 'none' : 'block'};
            position: absolute;
            color: #1F4E79;
            font-weight: bold;
            font-size: 11pt;
          }
          .kw-lbl-no       { top: 11mm; left: 45mm; }
          .kw-lbl-diterima { top: 20mm; left: 45mm; }
          .kw-lbl-uang     { top: 28mm; left: 45mm; }
          .kw-lbl-guna     { top: 35mm; left: 45mm; }
          .kw-lbl-terbilang { top: 78mm; left: 45mm; }

          /* Field Values (Koordinat Presisi Absolut X, Y dalam mm) */
          .kw-pos-no {
            position: absolute;
            top: 11mm;
            left: 55mm;
            font-size: 10pt;
            font-weight: normal;
          }
          .kw-pos-diterima {
            position: absolute;
            top: 20mm;
            left: 92mm;
            width: 90mm;
            font-size: 11pt;
            font-weight: bold;
          }
          .kw-pos-uang {
            position: absolute;
            top: 28mm;
            left: 90mm;
            width: 92mm;
            font-size: 10.5pt;
            font-style: italic;
          }
          .kw-pos-guna {
            position: absolute;
            top: 35mm;
            left: 45mm;
            width: 137mm;
            font-size: 10.5pt;
            line-height: 7.7mm;
            text-indent: 40mm;
            word-wrap: break-word;
            word-break: normal;
          }
          .kw-pos-tanggal {
            position: absolute;
            top: 60mm;
            left: 132mm;
            width: 55mm;
            font-size: 11pt;
            text-align: left;
          }
          .kw-pos-nominal {
            position: absolute;
            top: 78mm;
            left: 83mm;
            font-size: 13pt;
            font-weight: bold;
            ${isKwitansiMode ? '' : 'border-top: 2px solid #000; border-bottom: 2px solid #000; padding: 2px 10px;'}
          }
          .kw-pos-ttd {
            position: absolute;
            top: 85mm;
            left: 140mm;
            width: 45mm;
            font-size: 11pt;
            font-weight: bold;
            text-align: left;
          }
        </style>
      </head>
      <body>
        <!-- Static Labels for HVS Mode -->
        <span class="kw-label kw-lbl-no">No.</span>
        <span class="kw-label kw-lbl-diterima">Telah diterima dari :</span>
        <span class="kw-label kw-lbl-uang">Uang sebanyak :</span>
        <span class="kw-label kw-lbl-guna">Guna membayar :</span>
        ${isKwitansiMode ? '' : '<span class="kw-label kw-lbl-terbilang">Terbilang</span>'}

        <!-- Values positioned absolutely -->
        <div class="kw-pos-no">${idTrx}</div>
        <div class="kw-pos-diterima">${nama} ${nip && nip !== '-' ? (' - NIP: ' + nip) : ''}</div>
        <div class="kw-pos-uang">${terbilangTeks}</div>
        <div class="kw-pos-guna">${rincianTeks}</div>
        <div class="kw-pos-tanggal">Secang, ${tglFormatted}</div>
        <div class="kw-pos-nominal">Rp. ${Number(nominal).toLocaleString('id-ID')},-</div>
        <div class="kw-pos-ttd">( ${penandatangan} )</div>
      </body>
      </html>
    `;

    var htmlOutput = HtmlService.createHtmlOutput(htmlContent);
    var pdfBlob = htmlOutput.getAs('application/pdf').setName(filename);
    
    // Cek apakah file sudah ada sebelumnya di folder
    var existingFiles = folder.getFilesByName(filename);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    var file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var downloadUrl = file.getDownloadUrl();
    var viewUrl = file.getUrl();

    return {
      success: true,
      message: "File PDF Kuitansi berhasil dibuat & disimpan!",
      fileId: file.getId(),
      fileName: filename,
      downloadUrl: downloadUrl,
      viewUrl: viewUrl
    };
  } catch (err) {
    return { success: false, message: "Gagal membuat PDF: " + err.message };
  }
}

/**
 * REKAP DATA ANGGOTA BELUM SETOR (IURAN ANGGOTA TERBAYAR TAPI BELUM DISETORKAN)
 * Dibaca dari Master_Anggota:
 * - Column I (index 8): bayar_terakhir_bulan
 * - Column J (index 9): bayar_terakhir_tahun
 * - Column K (index 10): setor_terakhir_bulan
 * - Column L (index 11): setor_terakhir_tahun
 */
function getDataBelumSetor(pin) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheetAnggota) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var paguMap = getMatrixPaguMap();
    var dataA = sheetAnggota.getDataRange().getValues();
    var resultList = [];
    var totalSeluruhTunggakan = 0;

    for (var i = 1; i < dataA.length; i++) {
      var row = dataA[i];
      var id = String(row[0]);
      if (!id) continue;

      var nama = row[1] || "-";
      var nip = row[2] ? String(row[2]) : "-";

      var bayarBln = Number(row[8]) || 0;
      var bayarThn = Number(row[9]) || 0;
      var setorBln = Number(row[10]) || 0;
      var setorThn = Number(row[11]) || 0;

      // Jika anggota belum pernah membayar iuran sama sekali, tidak ada dana terbayar yang belum disetor
      if (bayarThn === 0 || bayarBln === 0) continue;

      // Hitung indeks bulan absolut
      var bayarVal = bayarThn * 12 + bayarBln;
      var setorVal = (setorThn > 0 && setorBln > 0) ? (setorThn * 12 + setorBln) : 0;

      // Jika setor_terakhir >= bayar_terakhir, berarti seluruh iuran terbayar SUDAH disetorkan (0 belum disetor)
      if (setorVal >= bayarVal) continue;

      // Tentukan bulan awal belum disetor (setor_terakhir + 1 bulan)
      var startBln = 1;
      var startThn = bayarThn;

      if (setorVal > 0) {
        startBln = setorBln + 1;
        startThn = setorThn;
        if (startBln > 12) {
          startBln = 1;
          startThn = setorThn + 1;
        }
      } else {
        startThn = bayarThn;
        startBln = 1;
      }

      // Tentukan bulan akhir belum disetor (bayar_terakhir)
      var endThn = bayarThn;
      var endBln = bayarBln;

      var totalIuran = 0;
      var jumlahBulan = 0;
      var bulanAwalStr = "";
      var bulanAkhirStr = "";

      var currY = startThn;
      var currM = startBln;

      while (currY < endThn || (currY === endThn && currM <= endBln)) {
        var key = currY + "_" + currM;
        var paguNominal = Number(paguMap[key]) || 30000;

        totalIuran += paguNominal;
        jumlahBulan++;

        if (jumlahBulan === 1) {
          bulanAwalStr = getNamaBulan(currM) + " " + currY;
        }
        bulanAkhirStr = getNamaBulan(currM) + " " + currY;

        currM++;
        if (currM > 12) {
          currM = 1;
          currY++;
        }
      }

      // HANYA tampilkan anggota yang memiliki dana terbayar tetapi belum disetor (jumlahBulan > 0)
      if (jumlahBulan > 0 && totalIuran > 0) {
        totalSeluruhTunggakan += totalIuran;
        resultList.push({
          id: id,
          nama: nama,
          nip: nip,
          bulanAwal: bulanAwalStr,
          bulanAkhir: bulanAkhirStr,
          jumlahBulan: jumlahBulan,
          jumlahIuran: totalIuran
        });
      }
    }

    return {
      success: true,
      data: resultList,
      totalAnggota: resultList.length,
      totalSeluruhTunggakan: totalSeluruhTunggakan
    };
  } catch (err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

/**
 * REKAP DATA ANGGOTA UNTUK INPUT SETORAN KAS
 * Menampilkan seluruh anggota beserta detail iuran terbayar yang belum disetor dan status setor terakhir.
 */
function getDataInputSetor(pin) {
  var authCheck = verifyAdminPin(pin);
  if (!authCheck.success) return { success: false, message: "Akses Ditolak!" };

  try {
    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    if (!sheetAnggota) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var paguMap = getMatrixPaguMap();
    var dataA = sheetAnggota.getDataRange().getValues();
    var resultList = [];
    var totalSeluruhBelumSetor = 0;
    var totalAnggotaBelumSetor = 0;

    for (var i = 1; i < dataA.length; i++) {
      var row = dataA[i];
      var id = String(row[0]);
      if (!id) continue;

      var nama = row[1] || "-";
      var nip = row[2] ? String(row[2]) : "-";
      var noKode = row[13] ? String(row[13]) : "-";

      var bayarBln = Number(row[8]) || 0;
      var bayarThn = Number(row[9]) || 0;
      var setorBln = Number(row[10]) || 0;
      var setorThn = Number(row[11]) || 0;

      var bayarVal = bayarThn * 12 + bayarBln;
      var setorVal = (setorThn > 0 && setorBln > 0) ? (setorThn * 12 + setorBln) : 0;

      var lastPaidStr = (bayarThn > 0 && bayarBln > 0) ? (getNamaBulan(bayarBln) + " " + bayarThn) : "Belum Ada";
      var lastSetorStr = (setorThn > 0 && setorBln > 0) ? (getNamaBulan(setorBln) + " " + setorThn) : "Belum Ada";

      var totalIuran = 0;
      var jumlahBulan = 0;
      var bulanAwalStr = "-";
      var bulanAkhirStr = "-";
      var detailBulanBelumSetor = [];

      if (bayarVal > 0 && setorVal < bayarVal) {
        // Tentukan bulan awal (setor_terakhir + 1 bulan)
        var startBln = 1;
        var startThn = bayarThn;

        if (setorVal > 0) {
          startBln = setorBln + 1;
          startThn = setorThn;
          if (startBln > 12) {
            startBln = 1;
            startThn = setorThn + 1;
          }
        } else {
          startThn = bayarThn;
          startBln = 1;
        }

        var endThn = bayarThn;
        var endBln = bayarBln;

        var currY = startThn;
        var currM = startBln;

        while (currY < endThn || (currY === endThn && currM <= endBln)) {
          var key = currY + "_" + currM;
          var paguNominal = Number(paguMap[key]) || 30000;

          totalIuran += paguNominal;
          jumlahBulan++;

          if (jumlahBulan === 1) {
            bulanAwalStr = getNamaBulan(currM) + " " + currY;
          }
          bulanAkhirStr = getNamaBulan(currM) + " " + currY;

          detailBulanBelumSetor.push({
            bulan: currM,
            tahun: currY,
            namaBulan: getNamaBulan(currM),
            label: getNamaBulan(currM) + " " + currY,
            jumlahBulanKe: jumlahBulan,
            subtotal: totalIuran,
            nominal: paguNominal
          });

          currM++;
          if (currM > 12) {
            currM = 1;
            currY++;
          }
        }
      }

      if (jumlahBulan > 0) {
        totalSeluruhBelumSetor += totalIuran;
        totalAnggotaBelumSetor++;
      }

      resultList.push({
        id: id,
        nama: nama,
        nip: nip,
        noKode: noKode,
        bayarBln: bayarBln,
        bayarThn: bayarThn,
        setorBln: setorBln,
        setorThn: setorThn,
        lastPaidStr: lastPaidStr,
        lastSetorStr: lastSetorStr,
        bulanAwal: bulanAwalStr,
        bulanAkhir: bulanAkhirStr,
        jumlahBulan: jumlahBulan,
        jumlahIuran: totalIuran,
        isLunasSetor: (setorVal >= bayarVal && bayarVal > 0),
        detailBulanBelumSetor: detailBulanBelumSetor
      });
    }

    return {
      success: true,
      data: resultList,
      totalAnggota: resultList.length,
      totalAnggotaBelumSetor: totalAnggotaBelumSetor,
      totalSeluruhBelumSetor: totalSeluruhBelumSetor
    };
  } catch (err) {
    return { success: false, message: "Error: " + err.toString() };
  }
}

/**
 * PROSES SETOR KAS ANGGOTA (Setor Seluruhnya / Setor Bulanan)
 */
function processSetorKasAnggota(pin, payload) {
  try {
    var authCheck = verifyAdminPin(pin);
    if (!authCheck.success) return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };

    payload = payload || {};
    var idAnggota = payload.idAnggota;
    if (!idAnggota) return { success: false, message: "ID Anggota tidak valid!" };

    var targetBln = Number(payload.targetBln) || 0;
    var targetThn = Number(payload.targetThn) || 0;
    var nominalDisetor = Number(payload.nominal) || 0;
    var bDari = Number(payload.bDari) || targetBln;
    var bSampai = Number(payload.bSampai) || targetBln;
    var thnDari = Number(payload.thnDari) || targetThn;
    var thnSampai = Number(payload.thnSampai) || targetThn;
    var penerima = payload.penerima || "Bendahara";
    var catatan = payload.catatan || "Setoran Kas Anggota";

    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    var sheetSetor = ss.getSheetByName(SHEET_SETOR);
    if (!sheetAnggota || !sheetSetor) {
      return { success: false, message: "Sheet Master_Anggota atau Transaksi_Setor tidak ditemukan!" };
    }

    var dataAnggota = sheetAnggota.getDataRange().getValues();
    var namaAnggota = "";
    var targetRowIdx = -1;

    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        namaAnggota = dataAnggota[i][1];
        targetRowIdx = i + 1;
        break;
      }
    }

    if (targetRowIdx === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }

    // Catat ke Transaksi_Setor
    var tgl = new Date();
    var tglStr = Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    var idSetor = "STR-" + Utilities.formatDate(tgl, ss.getSpreadsheetTimeZone(), "yyyyMMdd") + "-" + Math.floor(100 + Math.random() * 900);

    var periodeStr = (thnDari === thnSampai && bDari === bSampai)
      ? (getNamaBulan(bDari) + " " + thnDari)
      : (getNamaBulan(bDari) + " " + thnDari + " s.d. " + getNamaBulan(bSampai) + " " + thnSampai);

    sheetSetor.appendRow([
      idSetor, tglStr, idAnggota, namaAnggota, nominalDisetor,
      thnSampai, bDari, bSampai, penerima, catatan + " [" + periodeStr + "]"
    ]);

    // Update Master_Anggota: Setor Terakhir Bulan (kolom 11), Setor Terakhir Tahun (kolom 12)
    sheetAnggota.getRange(targetRowIdx, 11, 1, 2).setValues([[targetBln, targetThn]]);

    return {
      success: true,
      message: "Setoran kas berhasil disimpan untuk " + namaAnggota + " (" + periodeStr + ")!"
    };
  } catch (err) {
    return { success: false, message: "Gagal memproses setoran: " + err.toString() };
  }
}

/**
 * PROSES BATAL SETOR KAS ANGGOTA (MENGEMBALIKAN POSISI SETOR TERAKHIR)
 */
function batalSetorKasAnggota(pin, payload) {
  try {
    var authCheck = verifyAdminPin(pin);
    if (!authCheck.success) return { success: false, message: "Akses Ditolak: PIN Admin tidak valid!" };

    payload = payload || {};
    var idAnggota = payload.idAnggota;
    if (!idAnggota) return { success: false, message: "ID Anggota tidak valid!" };

    var ss = getSS();
    var sheetAnggota = ss.getSheetByName(SHEET_ANGGOTA);
    var sheetSetor = ss.getSheetByName(SHEET_SETOR);
    if (!sheetAnggota) return { success: false, message: "Sheet Master_Anggota tidak ditemukan!" };

    var dataAnggota = sheetAnggota.getDataRange().getValues();
    var namaAnggota = "";
    var targetRowIdx = -1;

    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        namaAnggota = dataAnggota[i][1];
        targetRowIdx = i + 1;
        break;
      }
    }

    if (targetRowIdx === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }

    var prevBln = 0;
    var prevThn = 0;

    // Cari riwayat penyetoran anggota di Transaksi_Setor
    if (sheetSetor) {
      var dataSetor = sheetSetor.getDataRange().getValues();
      var lastTrxRowIdx = -1;

      // Cari baris transaksi setoran paling akhir untuk anggota ini
      for (var r = dataSetor.length - 1; r >= 1; r--) {
        if (String(dataSetor[r][2]) === String(idAnggota)) {
          lastTrxRowIdx = r + 1;
          break;
        }
      }

      // Hapus transaksi penyetoran paling akhir tersebut (rollback transaksi)
      if (lastTrxRowIdx > -1) {
        sheetSetor.deleteRow(lastTrxRowIdx);
      }

      // Cari transaksi penyetoran SEBELUMNYA jika ada
      dataSetor = sheetSetor.getDataRange().getValues(); // Refresh data setelah deleteRow
      for (var r2 = dataSetor.length - 1; r2 >= 1; r2--) {
        if (String(dataSetor[r2][2]) === String(idAnggota)) {
          // Kolom 5 = Tahun Sampai (index 5), Kolom 7 = Bulan Sampai (index 7)
          prevThn = Number(dataSetor[r2][5]) || 0;
          prevBln = Number(dataSetor[r2][7]) || 0;
          break;
        }
      }
    }

    // Jika payload menyediakan targetBln/targetThn yang spesifik > 0, gunakan nilai tersebut
    var finalBln = (Number(payload.targetBln) > 0) ? Number(payload.targetBln) : prevBln;
    var finalThn = (Number(payload.targetThn) > 0) ? Number(payload.targetThn) : prevThn;

    // Update Master_Anggota: Setor Terakhir Bulan (kolom 11), Setor Terakhir Tahun (kolom 12)
    sheetAnggota.getRange(targetRowIdx, 11, 1, 2).setValues([[finalBln, finalThn]]);

    var statusStr = (finalThn > 0 && finalBln > 0) ? (getNamaBulan(finalBln) + " " + finalThn) : "Belum Ada Setoran";

    return {
      success: true,
      message: "Setoran kas anggota " + namaAnggota + " berhasil dibatalkan dan dikembalikan ke posisi: " + statusStr
    };
  } catch (err) {
    return { success: false, message: "Gagal membatalkan setoran: " + err.toString() };
  }
}