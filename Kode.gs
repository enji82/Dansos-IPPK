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
      .setTitle("Sistem Iuran Anggota & Setoran Kas")
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

  return {
    id: idAnggota,
    nama: namaAnggota,
    nip: nip,
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
      result.push({
        id: data[i][0],
        nama: data[i][1],
        nip: data[i][2],
        saldoDeposit: Number(data[i][7]) || 0,
        bayarBln: Number(data[i][8]) || 0,
        bayarThn: Number(data[i][9]) || 0,
        setorBln: Number(data[i][10]) || 0,
        setorThn: Number(data[i][11]) || 0
      });
    }
  }
  return result;
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
    var tahunPeriode = Number(form.tahunPeriode);
    var bulanDari = Number(form.bulanDari);
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
    var targetRowIdx = -1;
    var currentDeposit = 0;
    
    for (var i = 1; i < dataAnggota.length; i++) {
      if (String(dataAnggota[i][0]) === String(idAnggota)) {
        namaAnggota = dataAnggota[i][1];
        currentDeposit = Number(dataAnggota[i][7]) || 0;
        targetRowIdx = i + 1;
        break;
      }
    }
    
    if (targetRowIdx === -1) {
      return { success: false, message: "Data Anggota tidak ditemukan!" };
    }
    
    sheetMasuk.appendRow([
      idTrx, tglStr, idAnggota, namaAnggota, nominalDiterima,
      tahunPeriode, bulanDari, bulanSampai, iuranTerpakai,
      kembalianCash, masukDeposit, catatan
    ]);
    
    var newDeposit = currentDeposit + masukDeposit;
    // Batch update: Saldo Deposit (kolom 8), Cutoff Bayar Bulan (kolom 9), Cutoff Bayar Tahun (kolom 10)
    sheetAnggota.getRange(targetRowIdx, 8, 1, 3).setValues([[newDeposit, bulanSampai, tahunPeriode]]);
    
    return { success: true, message: "Transaksi Pembayaran Berhasil Disimpan! ID: " + idTrx };
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
    var tahunPeriode = Number(form.tahunPeriode);
    var bulanDari = Number(form.bulanDari);
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
    
    sheetSetor.appendRow([
      idSetor, tglStr, idAnggota, namaAnggota, nominalDisetor,
      tahunPeriode, bulanDari, bulanSampai, penerimaSetor, catatan
    ]);
    
    // Batch update: Cutoff Setor Bulan (kolom 11), Cutoff Setor Tahun (kolom 12)
    sheetAnggota.getRange(targetRowIdx, 11, 1, 2).setValues([[bulanSampai, tahunPeriode]]);
    
    return { success: true, message: "Setoran Kas Berhasil Dicatat! ID: " + idSetor };
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
          listMasuk.push({
            idTrx: r[0],
            tanggal: r[1] ? Utilities.formatDate(new Date(r[1]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "-",
            idAnggota: r[2],
            nama: r[3],
            nominal: nom,
            tahun: thn,
            periode: getNamaBulan(r[6]) + " - " + getNamaBulan(r[7]),
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