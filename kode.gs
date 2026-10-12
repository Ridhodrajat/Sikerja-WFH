const SPREADSHEET_ID = ""; // Isi jika script standalone. Kosongkan jika bound ke Spreadsheet.
const SHEET_PEGAWAI = "MasterPegawai";
const SHEET_MAGANG = "MasterMagang";
const SHEET_PRESENSI = "PresensiWFH";
const APP_TIMEZONE = "Asia/Makassar";

// ======================= ADMIN LOGIN =======================
// PIN admin disimpan di Script Properties, bukan di file HTML.
// Cara set:
// Apps Script > Project Settings > Script Properties
// Tambahkan property: ADMIN_PIN  dengan value PIN pilihan Anda.
function verifikasiAdmin(pin) {
  try {
    const savedPin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");

    if (!savedPin) {
      return {
        success: false,
        setupRequired: true,
        error: "ADMIN_PIN belum diatur di Script Properties.",
      };
    }

    const input = String(pin || "").trim();
    if (!input) return { success: false };

    return { success: input === String(savedPin).trim() };
  } catch (error) {
    console.error(error);
    return {
      success: false,
      error: error.message || String(error),
    };
  }
}

function getSpreadsheet() {
  if (SPREADSHEET_ID && String(SPREADSHEET_ID).trim() !== "") {
    return SpreadsheetApp.openById(String(SPREADSHEET_ID).trim());
  }

  // Coba ambil dari Script Properties jika SPREADSHEET_ID di file kosong
  const savedId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (savedId && String(savedId).trim() !== "") {
    return SpreadsheetApp.openById(String(savedId).trim());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // Jika standalone dan SPREADSHEET_ID belum ada, buat Otomatis Google Sheets baru
  const newSs = SpreadsheetApp.create("SIKERJA WFH - Database Kehadiran");
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", newSs.getId());
  return newSs;
}

function doGet() {
  // Index.html tidak memakai scriptlet <?= ?> apa pun, jadi tidak perlu lewat
  // createTemplateFromFile().evaluate(). Untuk HTML berukuran besar, jalur evaluate()
  // pernah memicu bug Apps Script sendiri saat mem-bundle halaman ("Failed to execute
  // 'write' on 'Document': Unexpected identifier ..."). createHtmlOutputFromFile()
  // mengirim HTML apa adanya tanpa melalui mesin templating itu, jadi lebih aman.
  return HtmlService.createHtmlOutputFromFile("Index").setTitle("SIKERJA WFH - Kanwil Ditjenpas Kalimantan Selatan").addMetaTag("viewport", "width=device-width, initial-scale=1.0").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupDatabase() {
  const ss = getSpreadsheet();

  let shPegawai = ss.getSheetByName(SHEET_PEGAWAI);
  if (!shPegawai) {
    shPegawai = ss.insertSheet(SHEET_PEGAWAI);
    shPegawai
      .getRange(1, 1, 1, 6)
      .setValues([["ID", "Nama Lengkap", "NIP", "Golongan", "Jabatan", "Status"]])
      .setFontWeight("bold");
  }

  let shMagang = ss.getSheetByName(SHEET_MAGANG);
  if (!shMagang) {
    shMagang = ss.insertSheet(SHEET_MAGANG);
    shMagang
      .getRange(1, 1, 1, 6)
      .setValues([["ID", "Nama Lengkap", "NIM", "Instansi", "Posisi", "Status"]])
      .setFontWeight("bold");
  } else {
    // Menyamakan header versi lama tanpa menghapus data.
    const headers = shMagang.getRange(1, 1, 1, Math.max(6, shMagang.getLastColumn())).getDisplayValues()[0];
    if (headers[3] === "Golongan") shMagang.getRange(1, 4).setValue("Instansi");
    if (headers[4] === "Jabatan") shMagang.getRange(1, 5).setValue("Posisi");
  }

  let shPresensi = ss.getSheetByName(SHEET_PRESENSI);
  if (!shPresensi) {
    shPresensi = ss.insertSheet(SHEET_PRESENSI);
    shPresensi
      .getRange(1, 1, 1, 11)
      .setValues([["Timestamp", "Tanggal", "Kategori", "ID_User", "Nama Lengkap", "NIP_NIM", "Golongan", "Jabatan", "Status_Kehadiran", "Uraian_Kegiatan", "Foto_URL"]])
      .setFontWeight("bold");
  }

  // Format hanya area yang dipakai agar proses tidak lambat.
  const pegRows = Math.max(shPegawai.getLastRow(), 2);
  const magRows = Math.max(shMagang.getLastRow(), 2);
  const preRows = Math.max(shPresensi.getLastRow(), 2);
  shPegawai.getRange(1, 3, pegRows, 1).setNumberFormat("@");
  shMagang.getRange(1, 3, magRows, 1).setNumberFormat("@");
  shPresensi.getRange(1, 6, preRows, 1).setNumberFormat("@");
}

function getInitialData() {
  try {
    setupDatabase();
    const ss = getSpreadsheet();
    const todayStr = formatDate_(new Date(), "yyyy-MM-dd");

    const pegawai = readPegawai_(ss.getSheetByName(SHEET_PEGAWAI));
    const magang = readMagang_(ss.getSheetByName(SHEET_MAGANG));

    // PENTING: JANGAN kirim seluruh riwayat PresensiWFH di sini. Sheet ini terus bertambah
    // setiap hari (bisa ratusan/ribuan baris seiring waktu), dan mengirim semuanya lewat
    // google.script.run pernah memicu bug internal Apps Script saat mem-bundle respons ke
    // browser ("Failed to execute 'write' on 'Document': Unexpected identifier ...") begitu
    // ukurannya melewati ambang batas tertentu. Halaman utama hanya butuh data HARI INI
    // (statistik, dropdown presensi, tabel admin); riwayat/rekap tanggal lain diambil
    // on-demand lewat getRekapData() saat tab Rekap benar-benar dipakai.
    const presensiHariIni = readPresensi_(ss.getSheetByName(SHEET_PRESENSI)).filter((p) => p.tanggal === todayStr);

    return {
      success: true,
      todayStr: todayStr,
      pegawai: pegawai,
      magang: magang,
      presensi: presensiHariIni,
      presensiHariIni: presensiHariIni,
    };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  }
}

// Alias untuk kompatibilitas dengan HTML versi lama.
function getAppData() {
  return getInitialData();
}

function simpanPresensi(formData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupDatabase();

    if (!formData || !formData.idUser || !formData.namaLengkap) {
      throw new Error("Data presensi tidak lengkap. Silakan pilih nama terlebih dahulu.");
    }

    // Cuti dan Dinas Luar hanya boleh dicatat melalui menu Admin.
    const requestedStatus = String(formData.statusKehadiran || "").trim();
    if (requestedStatus === "Cuti" || requestedStatus === "Dinas Luar") {
      throw new Error("Status Cuti dan Dinas Luar hanya dapat ditandai oleh Admin.");
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRESENSI);
    const now = new Date();
    const timestampStr = formatDate_(now, "yyyy-MM-dd HH:mm:ss");
    const tanggalStr = formatDate_(now, "yyyy-MM-dd");

    let fotoUrl = "";
    if (formData.fotoBase64) {
      fotoUrl = saveImageToDrive(formData.fotoBase64, formData.namaLengkap, tanggalStr);
    }

    // Satu user hanya satu presensi per hari (tidak boleh mengisi ulang).
    // Dicocokkan lewat Tanggal + Kategori + ID_User -- Kategori wajib ikut dibandingkan
    // karena ID Pegawai dan ID Magang penomorannya terpisah (masing-masing mulai dari 1),
    // jadi ID yang sama persis bisa dipakai orang berbeda di kategori yang berbeda.
    const values = sheet.getDataRange().getDisplayValues();
    const idUserNorm = normalizeId_(formData.idUser);
    const kategoriNorm = String(formData.kategori || "").trim();
    let rowToUpdate = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]).trim() === tanggalStr && String(values[i][2]).trim() === kategoriNorm && normalizeId_(values[i][3]) === idUserNorm) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate > 0) {
      const statusSebelumnya = String(values[rowToUpdate - 1][8] || "").trim();
      throw new Error(statusSebelumnya ? "Anda sudah tercatat presensi hari ini (" + statusSebelumnya + "). Presensi hanya dapat dilakukan satu kali per hari." : "Anda sudah tercatat presensi hari ini. Presensi hanya dapat dilakukan satu kali per hari.");
    }

    const rowData = [timestampStr, tanggalStr, String(formData.kategori || "").trim(), idUserNorm, String(formData.namaLengkap || "").trim(), String(formData.nipNim || "").trim(), String(formData.golongan || "").trim(), String(formData.jabatan || "").trim(), String(formData.statusKehadiran || "").trim(), String(formData.uraianKegiatan || "").trim(), fotoUrl];

    sheet.appendRow(rowData);

    SpreadsheetApp.flush();
    return { success: true, data: getInitialData() };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function saveImageToDrive(base64Data, namaUser, tanggal) {
  try {
    const matches = String(base64Data).match(/^data:(.+?);base64,(.+)$/);
    if (!matches) throw new Error("Format foto tidak valid.");

    const contentType = matches[1];
    const bytes = Utilities.base64Decode(matches[2]);
    const cleanNama = String(namaUser || "user").replace(/[^a-zA-Z0-9]/g, "_");
    const ext = contentType.includes("png") ? "png" : contentType.includes("webp") ? "webp" : "jpg";
    const blob = Utilities.newBlob(bytes, contentType, `WFH_${cleanNama}_${tanggal}.${ext}`);

    const folder = getOrCreateFolder_("SIKERJA_WFH_FOTO");
    const file = folder.createFile(blob);

    // Jika kebijakan Workspace melarang ANYONE_WITH_LINK, file tetap tersimpan dan URL tetap dikembalikan.
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {}

    // file.getUrl() ("/file/d/ID/view") hanya berupa halaman viewer Drive, BUKAN URL gambar
    // langsung, sehingga gagal dipakai sebagai <img src> di halaman web (tampak putus/alt text
    // saja). Format "uc?export=view&id=" di bawah ini yang bisa dirender langsung sebagai gambar.
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (error) {
    throw new Error("Gagal menyimpan foto: " + (error.message || String(error)));
  }
}

// Folder khusus agar file tidak tercecer di root Drive. Dipakai bersama untuk foto
// dokumentasi maupun berkas PDF hasil rekap.
function getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function tambahMasterPegawai(pegawaiData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupDatabase();

    if (!pegawaiData || !pegawaiData.nama || !pegawaiData.nip) {
      throw new Error("Nama dan NIP wajib diisi.");
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PEGAWAI);
    const existing = readPegawai_(sheet);
    const nip = String(pegawaiData.nip).trim();

    if (existing.some((p) => p.nip === nip)) {
      throw new Error("NIP tersebut sudah terdaftar.");
    }

    const newId = makeId_("PEG");
    sheet.appendRow([newId, String(pegawaiData.nama).trim(), nip, String(pegawaiData.golongan || "").trim(), String(pegawaiData.jabatan || "").trim(), "Aktif"]);

    return { success: true, data: getInitialData() };
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function tambahMasterMagang(magangData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupDatabase();

    if (!magangData || !magangData.nama || !magangData.nim) {
      throw new Error("Nama dan NIM/ID Magang wajib diisi.");
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_MAGANG);
    const existing = readMagang_(sheet);
    const nim = String(magangData.nim).trim();

    if (existing.some((m) => m.nim === nim)) {
      throw new Error("NIM/ID Magang tersebut sudah terdaftar.");
    }

    const newId = makeId_("MGG");
    sheet.appendRow([newId, String(magangData.nama).trim(), nim, String(magangData.instansi || magangData.golongan || "").trim(), String(magangData.posisi || magangData.jabatan || "").trim(), "Aktif"]);

    return { success: true, data: getInitialData() };
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function hapusMasterData(type, id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupDatabase();

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(type === "pegawai" ? SHEET_PEGAWAI : SHEET_MAGANG);
    if (!sheet) throw new Error("Sheet master tidak ditemukan.");

    const values = sheet.getDataRange().getDisplayValues();
    const idNorm = normalizeId_(id);
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (normalizeId_(values[i][0]) === idNorm) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow < 0) throw new Error("Data yang akan dihapus tidak ditemukan.");
    sheet.deleteRow(targetRow);

    return { success: true, data: getInitialData() };
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function setStatusAdminCepat(userId, status) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    userId = normalizeId_(userId);
    status = String(status || "").trim();

    if (!userId) throw new Error("ID pegawai tidak valid.");
    if (status && status !== "Cuti" && status !== "Dinas Luar") {
      throw new Error("Status tidak valid.");
    }

    const ss = getSpreadsheet();
    const shPegawai = ss.getSheetByName(SHEET_PEGAWAI);
    const shPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!shPegawai || !shPresensi) throw new Error("Sheet database belum lengkap.");

    // Cari 1 pegawai langsung dari kolom ID. Dibandingkan lewat normalizeId_ (bukan
    // TextFinder literal) supaya ID lama yang terformat angka (mis. "14" vs "14.0") tetap cocok.
    const pegLastRow = shPegawai.getLastRow();
    if (pegLastRow < 2) throw new Error("Master pegawai kosong.");

    const pegValues = shPegawai.getRange(2, 1, pegLastRow - 1, 6).getDisplayValues();
    const pegRowIdx = pegValues.findIndex((r) => normalizeId_(r[0]) === userId);
    if (pegRowIdx < 0) throw new Error("Pegawai tidak ditemukan.");

    const p = pegValues[pegRowIdx];
    const peg = {
      id: normalizeId_(p[0]),
      nama: String(p[1]).trim(),
      nip: String(p[2]).trim(),
      golongan: String(p[3]).trim(),
      jabatan: String(p[4]).trim(),
    };

    const now = new Date();
    const tanggalStr = formatDate_(now, "yyyy-MM-dd");
    const timestampStr = formatDate_(now, "yyyy-MM-dd HH:mm:ss");
    const waktuStr = formatDate_(now, "HH:mm:ss");

    // Cari hanya baris milik user ini di kolom ID_User, lalu cocokkan tanggal.
    let existingRow = -1;
    let existingStatus = "";
    const lastRow = shPresensi.getLastRow();

    if (lastRow >= 2) {
      const presValues = shPresensi.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
      for (let i = 0; i < presValues.length; i++) {
        const row = presValues[i];
        const tanggal = String(row[1]).trim(); // col B
        const kategori = String(row[2]).trim(); // col C
        const uid = normalizeId_(row[3]); // col D
        const st = String(row[8]).trim(); // col I
        if (tanggal === tanggalStr && kategori === "Pegawai" && uid === userId) {
          existingRow = i + 2;
          existingStatus = st;
          break;
        }
      }
    }

    // Jangan timpa orang yang sudah mengisi Hadir WFH sendiri.
    if (existingRow > 0 && existingStatus === "Hadir WFH") {
      throw new Error("Pegawai ini sudah mengisi presensi WFH dan tidak dapat ditandai Cuti/Dinas Luar.");
    }

    if (!status) {
      if (existingRow > 0 && (existingStatus === "Cuti" || existingStatus === "Dinas Luar")) {
        shPresensi.getRange(existingRow, 1, 1, 11).clearContent();
      }
      return {
        success: true,
        status: "",
        timestamp: timestampStr,
        waktu: waktuStr,
      };
    }

    const rowData = [[timestampStr, tanggalStr, "Pegawai", peg.id, peg.nama, peg.nip, peg.golongan, peg.jabatan, status, status === "Cuti" ? "Cuti - ditandai oleh Admin" : "Dinas Luar - ditandai oleh Admin", ""]];

    if (existingRow > 0) {
      shPresensi.getRange(existingRow, 1, 1, 11).setValues(rowData);
    } else {
      const targetRow = Math.max(shPresensi.getLastRow() + 1, 2);
      shPresensi.getRange(targetRow, 1, 1, 11).setValues(rowData);
      existingRow = targetRow;
    }

    return {
      success: true,
      id: "PRES_ROW_" + existingRow,
      status: status,
      timestamp: timestampStr,
      waktu: waktuStr,
    };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function simpanStatusAdmin(statusRows) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!Array.isArray(statusRows)) {
      throw new Error("Data status admin tidak valid.");
    }

    const ss = getSpreadsheet();
    const shPegawai = ss.getSheetByName(SHEET_PEGAWAI);
    const shPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!shPegawai || !shPresensi) throw new Error("Sheet database belum lengkap.");

    const pegawai = readPegawai_(shPegawai);
    const pegawaiById = {};
    pegawai.forEach((p) => (pegawaiById[String(p.id)] = p));

    const now = new Date();
    const tanggalStr = formatDate_(now, "yyyy-MM-dd");
    const timestampStr = formatDate_(now, "yyyy-MM-dd HH:mm:ss");

    const lastRow = Math.max(shPresensi.getLastRow(), 1);
    const existing = lastRow >= 2 ? shPresensi.getRange(2, 1, lastRow - 1, 11).getDisplayValues() : [];

    // Map pilihan admin dari browser.
    const selectedById = {};
    statusRows.forEach((item) => {
      const id = normalizeId_(item && item.idUser);
      const status = String((item && item.status) || "").trim();
      if (!id) return;
      selectedById[id] = status === "Cuti" || status === "Dinas Luar" ? status : "";
    });

    const rowsToWrite = [];
    const rowsToClear = [];
    const existingTodayIds = new Set();

    // Proses baris yang sudah ada tanpa getRange berulang-ulang.
    existing.forEach((row, idx) => {
      const tanggal = String(row[1] || "").trim();
      const kategori = String(row[2] || "").trim();
      const userId = normalizeId_(row[3]);
      const existingStatus = String(row[8] || "").trim();

      if (tanggal !== tanggalStr || kategori !== "Pegawai" || !userId) return;
      existingTodayIds.add(userId);

      // Jangan pernah menimpa Hadir WFH yang diisi pegawai sendiri.
      if (existingStatus === "Hadir WFH") return;

      if (!Object.prototype.hasOwnProperty.call(selectedById, userId)) return;

      const newStatus = selectedById[userId];
      const sheetRow = idx + 2;

      if (!newStatus) {
        // Hanya hapus baris status yang dibuat admin.
        if (existingStatus === "Cuti" || existingStatus === "Dinas Luar") {
          rowsToClear.push(sheetRow);
        }
        return;
      }

      const p = pegawaiById[userId];
      if (!p) return;

      rowsToWrite.push({
        row: sheetRow,
        values: [[timestampStr, tanggalStr, "Pegawai", p.id, p.nama, p.nip, p.golongan, p.jabatan, newStatus, newStatus === "Cuti" ? "Cuti - ditandai oleh Admin" : "Dinas Luar - ditandai oleh Admin", ""]],
      });
    });

    // Tambahkan pilihan baru yang belum punya baris hari ini.
    const rowsToAppend = [];
    Object.keys(selectedById).forEach((userId) => {
      const status = selectedById[userId];
      if (!status || existingTodayIds.has(userId)) return;
      const p = pegawaiById[userId];
      if (!p) return;

      rowsToAppend.push([timestampStr, tanggalStr, "Pegawai", p.id, p.nama, p.nip, p.golongan, p.jabatan, status, status === "Cuti" ? "Cuti - ditandai oleh Admin" : "Dinas Luar - ditandai oleh Admin", ""]);
    });

    // Tulis perubahan. Biasanya hanya 1-5 pegawai sehingga sangat cepat.
    rowsToWrite.forEach((item) => {
      shPresensi.getRange(item.row, 1, 1, 11).setValues(item.values);
    });

    // Kosongkan baris yang status adminnya dibatalkan.
    rowsToClear.forEach((rowNum) => {
      shPresensi.getRange(rowNum, 1, 1, 11).clearContent();
    });

    if (rowsToAppend.length) {
      const startRow = shPresensi.getLastRow() + 1;
      shPresensi.getRange(startRow, 1, rowsToAppend.length, 11).setValues(rowsToAppend);
    }

    SpreadsheetApp.flush();

    // Kembalikan data tanpa menjalankan setupDatabase lagi.
    // Sama seperti getInitialData(): jangan kirim seluruh riwayat presensi, cukup hari ini.
    const presensiHariIni = readPresensi_(shPresensi).filter((p) => p.tanggal === tanggalStr);
    const magang = readMagang_(ss.getSheetByName(SHEET_MAGANG));

    return {
      success: true,
      data: {
        success: true,
        todayStr: tanggalStr,
        pegawai: pegawai,
        magang: magang,
        presensi: presensiHariIni,
        presensiHariIni: presensiHariIni,
      },
    };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function getRekapData(filter) {
  try {
    setupDatabase();
    const ss = getSpreadsheet();
    let data = readPresensi_(ss.getSheetByName(SHEET_PRESENSI));
    filter = filter || {};

    if (filter.tipe === "harian" && filter.tanggal) {
      data = data.filter((p) => p.tanggal === String(filter.tanggal));
    } else if (filter.tipe === "bulanan" && filter.bulan) {
      data = data.filter((p) => p.tanggal && p.tanggal.startsWith(String(filter.bulan)));
    }

    if (filter.kategori && filter.kategori !== "Semua") {
      data = data.filter((p) => p.kategori === filter.kategori);
    }

    return { success: true, data: data };
  } catch (error) {
    return { success: false, error: error.message || String(error) };
  }
}

// ================= REKAP LAPORAN RESMI (GOOGLE DOCS & PDF) =================
// payload dikirim dari client (openRekapGoogleDocs()/openRekapPDF() di index.html), sudah
// berisi data harian yang difilter & diurutkan (urutPegawaiRekap: Kepala Kantor Wilayah lalu
// pangkat/golongan tertinggi ke terendah). Struktur dokumen mengikuti format contoh resmi
// "LAPORAN PELAKSANAAN WORK FROM HOME (WFH) DI LINGKUNGAN KANTOR WILAYAH DITJENPAS KALSEL".

function buatRekapGoogleDocs(payload) {
  let hasil;
  try {
    hasil = buatDokumenRekap_(payload);
    const file = DriveApp.getFileById(hasil.id);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (_) {}
    return { success: true, url: hasil.url };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  }
}

function buatRekapPDF(payload) {
  let hasil;
  try {
    hasil = buatDokumenRekap_(payload);
    const docFile = DriveApp.getFileById(hasil.id);
    const pdfBlob = docFile.getAs("application/pdf").setName(namaFileRekap_(payload));

    const folder = getOrCreateFolder_("SIKERJA_WFH_LAPORAN_PDF");
    const pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {}

    // Dokumen Google Docs sementara hanya alat bantu ekspor, tidak perlu disimpan.
    try {
      docFile.setTrashed(true);
    } catch (_) {}

    return { success: true, url: pdfFile.getUrl() };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  }
}

function namaFileRekap_(payload) {
  const tgl = String((payload && payload.tanggal) || formatDate_(new Date(), "yyyy-MM-dd"));
  return "Laporan WFH Kanwil Ditjenpas Kalsel " + tgl;
}

// Menyusun isi dokumen (dipakai bersama oleh jalur Google Docs maupun PDF) lalu menutupnya
// (saveAndClose) agar aman diekspor/dibagikan. Mengembalikan id & url yang sudah di-cache
// karena beberapa method Document tidak bisa dipanggil lagi setelah ditutup.
function buatDokumenRekap_(payload) {
  if (!payload || !payload.tanggal) throw new Error("Data rekap tidak valid. Pilih tanggal/bulan lalu klik Tampilkan Rekap terlebih dahulu.");

  const doc = DocumentApp.create(namaFileRekap_(payload));
  const body = doc.getBody();
  body.clear();

  // Ukuran halaman A4 (595.3 x 841.9 pt) PORTRAIT (Potret) sesuai instruksi prompt
  body.setPageWidth(595.3).setPageHeight(841.9);
  body.setMarginTop(42.5).setMarginBottom(70.85).setMarginLeft(56.65).setMarginRight(56.65);

  const tanggalLengkap = String(payload.tanggalLengkap || payload.tanggal);
  const tanggalBiasa = String(payload.tanggalBiasa || payload.tanggal);
  // Dikirim client sudah dalam format "JUM'AT, 4 September 2026" (hanya nama hari huruf
  // besar, sisanya huruf biasa) -- fallback ke uppercase penuh kalau payload lama/tidak
  // mengirim field ini (mis. dari klien yang belum diperbarui).
  const tanggalJudul = String(payload.tanggalJudul || tanggalLengkap.toUpperCase());
  const jumlahPegawai = Number(payload.jumlahPegawai || 0);
  const jumlahWFH = Number(payload.jumlahWFH || 0);
  const jumlahCuti = Number(payload.jumlahCuti || 0);
  const jumlahDinasLuar = Number(payload.jumlahDinasLuar || 0);
  const jumlahMagang = Number(payload.jumlahMagang || 0);
  const pegawaiCuti = Array.isArray(payload.pegawaiCuti) ? payload.pegawaiCuti : [];
  const pegawaiDinasLuar = Array.isArray(payload.pegawaiDinasLuar) ? payload.pegawaiDinasLuar : [];
  // Tabel Dokumentasi Kegiatan HANYA berisi pegawai Hadir WFH -- Cuti dan Dinas Luar tidak
  // melaksanakan kegiatan WFH untuk didokumentasikan (tidak ada foto/uraian kegiatan), jadi
  // mereka cukup disebutkan di ringkasan "B. PELAKSANAAN KEGIATAN" (baris "Pegawai Cuti"/
  // "Pegawai Dinas Luar" di tambahTabelKeterangan_ di atas), TIDAK digabung ke tabel ini.
  const pegawaiWFH = Array.isArray(payload.pegawaiWFH) ? payload.pegawaiWFH : [];
  const magang = Array.isArray(payload.magang) ? payload.magang : [];
  const kakanwil = payload.kakanwil || {};

  // ---------- KOP SURAT (samakan dengan letterhead di #printableArea / contoh laporan) ----------
  tambahKopSurat_(body, payload.logoDataUri);
  // body.clear() di atas menyisakan 1 paragraf kosong bawaan (lihat catatan di
  // hapusParagrafKosongAwal_) yang sekarang jadi child pertama SEBELUM tabel kop surat --
  // dihapus di sini setelah kop surat (child kedua) ada, supaya dokumen tidak dimulai
  // dengan baris kosong.
  hapusParagrafKosongAwal_(body);

  // ---------- JUDUL ----------
  // Diperkecil dari 12pt ke 11pt (masih sama dengan ukuran isi surat) supaya judul 4 baris
  // ini muat rapi di kop tanpa terasa kebesaran, sesuai laporan balik user.
  tambahParagrafTengah_(body, "LAPORAN PELAKSANAAN WORK FROM HOME (WFH)", true, 11);
  tambahParagrafTengah_(body, "DI LINGKUNGAN KANTOR WILAYAH DIREKTORAT JENDERAL", true, 11);
  tambahParagrafTengah_(body, "PEMASYARAKATAN KALIMANTAN SELATAN", true, 11);
  tambahParagrafTengah_(body, tanggalJudul, true, 11);
  body.appendParagraph("");

  // ---------- A. PENDAHULUAN ----------
  tambahJudulBagianHuruf_(body, "A", "PENDAHULUAN");
  tambahSubJudulNomor_(body, "1", "Umum");
  tambahParagrafIsi_(body, "Menindaklanjuti Surat Edaran Menteri Imigrasi dan Pemasyarakatan Nomor 2 Tahun 2026 tentang Pelaksanaan Tugas Kedinasan bagi Pegawai Aparatur Sipil Negara melalui Penerapan Work From Home dalam rangka mendukung percepatan transformasi tata kelola penyelenggaraan pemerintahan di lingkungan Kementerian Imigrasi dan Pemasyarakatan, serta Surat Sekretaris Direktorat Jenderal Pemasyarakatan Nomor PAS.1-SA.08.05-1410 tanggal 08 April 2026 tentang penyampaian pelaksanaan Work From Home (WFH), perlu dilakukan tindak lanjut pelaksanaan di lingkungan Kantor Wilayah dan Unit Pelaksana Teknis Pemasyarakatan.");
  tambahSubJudulNomor_(body, "2", "Dasar Pelaksanaan");
  tambahListNomor_(body, ["Surat Edaran Menteri Imigrasi dan Pemasyarakatan Nomor 2 Tahun 2026 tentang Pelaksanaan Tugas Kedinasan bagi Pegawai Aparatur Sipil Negara melalui Penerapan Work From Home dalam rangka mendukung percepatan transformasi tata kelola penyelenggaraan pemerintahan.", "Surat Sekretaris Direktorat Jenderal Pemasyarakatan Nomor PAS.1-SA.08.05-1410 tanggal 08 April 2026 tentang Penyampaian Pelaksanaan Work From Home (WFH) di lingkungan Direktorat Jenderal Pemasyarakatan."]);

  // ---------- B. PELAKSANAAN KEGIATAN ----------
  tambahJudulBagianHuruf_(body, "B", "PELAKSANAAN KEGIATAN");
  tambahParagrafIsi_(body, "Pelaksanaan Work From Home (WFH) di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan dilaksanakan pada:");

  const jumlahPegawaiKet = jumlahPegawai + " orang (" + jumlahWFH + " WFH" + (jumlahCuti ? ", " + jumlahCuti + " Cuti" : "") + (jumlahDinasLuar ? ", " + jumlahDinasLuar + " Dinas Luar" : "") + ")";
  const barisKeterangan = [
    ["Hari/Tanggal", tanggalLengkap],
    ["Jumlah Pegawai", jumlahPegawaiKet],
  ];
  if (jumlahMagang > 0) barisKeterangan.push(["Jumlah peserta magang", jumlahMagang + " orang"]);
  // Nilai berupa array (bukan string) dirender sebagai daftar bernomor 1./2./dst oleh
  // tambahTabelKeterangan_ (satu nama per baris), bukan digabung koma dalam satu baris
  // panjang -- dan BUKAN dikosongkan lalu ditambah daftar terpisah setelah tabel seperti
  // sebelumnya (tampak seperti nama "nyasar" di luar tabel keterangan).
  barisKeterangan.push(["Pegawai Cuti", pegawaiCuti.length ? pegawaiCuti : "Tidak ada"]);
  barisKeterangan.push(["Pegawai Dinas Luar", pegawaiDinasLuar.length ? pegawaiDinasLuar : "Tidak ada"]);
  tambahTabelKeterangan_(body, barisKeterangan);

  body.appendParagraph("");
  tambahParagrafIsi_(body, "Adapun pelaksanaannya sebagai berikut:");
  tambahListNomor_(body, ["WFH dilaksanakan pada " + tanggalLengkap + ".", "Seluruh Pegawai ASN Kanwil yang melaksanakan fungsi dukungan manajemen bekerja secara fleksibel dari lokasi masing-masing.", "Pegawai melakukan presensi secara daring melalui aplikasi STAR-ASN.", "Pegawai menyampaikan laporan harian kegiatan kepada atasan langsung.", "Koordinasi dan komunikasi kedinasan dilakukan melalui media daring.", "Pengawasan dan pemantauan capaian kinerja dilakukan oleh atasan langsung."]);

  // ---------- C. HASIL YANG DI CAPAI ----------
  tambahJudulBagianHuruf_(body, "C", "HASIL YANG DI CAPAI");
  tambahListHuruf_(body, ["Pelaksanaan WFH berjalan tertib dan lancar.", "Presensi pegawai tercatat melalui sistem daring.", "Laporan harian kegiatan tersampaikan dengan baik.", "Tugas dan fungsi Kanwil tetap berjalan efektif tanpa kendala yang berarti."]);

  // ---------- D. SIMPULAN ----------
  tambahJudulBagianHuruf_(body, "D", "SIMPULAN");
  tambahParagrafIsi_(body, "Pelaksanaan penyesuaian tugas kedinasan melalui Work From Home (WFH) di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan pada " + tanggalBiasa + " telah dilaksanakan sesuai ketentuan dan berjalan dengan baik serta tetap menjamin ketercapaian kinerja organisasi.");

  // ---------- E. PENUTUP ----------
  tambahJudulBagianHuruf_(body, "E", "PENUTUP");
  tambahParagrafIsi_(body, "Demikian laporan pelaksanaan penyesuaian tugas kedinasan melalui Work From Home (WFH) ini disusun sebagai bentuk pertanggungjawaban dan dokumentasi pelaksanaan kebijakan pimpinan di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan.");

  // ---------- TANDA TANGAN PEJABAT ----------
  // Menggunakan tabel 2 kolom tanpa border agar letak "Banjarmasin, ...", nama pejabat, dan NIP
  // rata kiri terhadap kolomnya sendiri di sebelah kanan (sejajar vertikal), tidak tergeser
  // oleh indentasi manual atau alignment paragraph semata.
  body.appendParagraph("");
  const tableTtd = body.appendTable();
  tableTtd.setBorderWidth(0);
  const rowTtd = tableTtd.appendTableRow();
  
  // Kolom kiri kosong (pemisah / penyeimbang)
  const cellKiri = rowTtd.appendTableCell("");
  cellKiri.setWidth(260);
  
  // Kolom kanan berisi detail TTD Kakanwil
  const cellKanan = rowTtd.appendTableCell("");
  cellKanan.setWidth(220);
  cellKanan.clear();
  
  const rapikanTtd_ = (p, bold, underline) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    p.editAsText().setBold(bold).setUnderline(underline).setFontSize(11);
    return p;
  };

  rapikanTtd_(cellKanan.appendParagraph("Banjarmasin, " + tanggalBiasa), false, false);
  rapikanTtd_(cellKanan.appendParagraph("Kepala Kantor Wilayah,"), false, false);
  rapikanTtd_(cellKanan.appendParagraph(""), false, false);
  rapikanTtd_(cellKanan.appendParagraph(""), false, false);
  rapikanTtd_(cellKanan.appendParagraph(""), false, false);
  rapikanTtd_(cellKanan.appendParagraph(String(kakanwil.nama || "-")), true, true);
  rapikanTtd_(cellKanan.appendParagraph("NIP. " + String(kakanwil.nip || "-")), false, false);
  hapusParagrafKosongAwal_(cellKanan);

  // ---------- DOKUMENTASI KEGIATAN: PEGAWAI ----------
  // Heading dokumentasi di contoh laporan resmi TIDAK dibesarkan (tetap 11pt bold, sama
  // dengan judul bagian A/B/C/D/E) -- jangan dinaikkan lagi ke 12-13pt seperti sebelumnya.
  body.appendPageBreak();
  tambahParagrafTengah_(body, "DOKUMENTASI KEGIATAN", true, 11);
  tambahParagrafTengah_(body, "WFH PADA " + String(tanggalBiasa).toUpperCase(), true, 11);
  body.appendParagraph("");
  tambahTabelDokumentasi_(body, pegawaiWFH, "pegawai");

  // ---------- DOKUMENTASI KEGIATAN: MAGANG ----------
  if (magang.length) {
    body.appendPageBreak();
    tambahParagrafTengah_(body, "DOKUMENTASI KEGIATAN", true, 11);
    tambahParagrafTengah_(body, "WFH PESERTA MAGANG PADA " + String(tanggalBiasa).toUpperCase(), true, 11);
    body.appendParagraph("");
    tambahTabelDokumentasi_(body, magang, "magang");
  }

  doc.saveAndClose();
  // Kalau blok tanda tangan kepanjangan/tidak cukup di sisa halaman, dorong utuh ke halaman
  // berikutnya (bukan terpotong nama di satu halaman, NIP-nya di halaman lain) -- lihat
  // jagaTandaTanganTetapUtuh_() untuk kenapa ini butuh Google Docs API, bukan DocumentApp.
  jagaTandaTanganTetapUtuh_(doc.getId(), "Banjarmasin, " + tanggalBiasa);
  return { id: doc.getId(), url: doc.getUrl() };
}

// DocumentApp (dipakai di seluruh buatDokumenRekap_ di atas) TIDAK punya API "keep
// together"/"keep with next" -- sudah dicek ke referensi resmi Apps Script, Class Paragraph
// tidak mengekspos method seperti itu. Fitur ini HANYA tersedia lewat Google Docs API
// (advanced service terpisah, request updateParagraphStyle dengan field keepWithNext), jadi
// dipanggil di sini setelah dokumen dasarnya selesai dibangun DocumentApp & disimpan.
//
// PENTING -- WAJIB diaktifkan manual sekali di Apps Script sebelum ini bisa jalan:
// Extensions/Editor Apps Script > ikon "+" di samping "Services" > cari "Google Docs API" >
// Tambahkan. Kalau belum diaktifkan (atau gagal karena sebab lain), fungsi ini gagal SENYAP
// (try/catch) -- dokumen tetap jadi seperti biasa, cuma tanpa jaminan blok tanda tangan utuh
// (perilaku lama: bisa saja terpotong 2 halaman kalau kebetulan pas di batas halaman).
function jagaTandaTanganTetapUtuh_(docId, teksBanjarmasin) {
  try {
    const dokumen = Docs.Documents.get(docId);
    const content = (dokumen.body && dokumen.body.content) || [];

    // Kumpulkan SEMUA paragraf (top-level maupun di dalam sel tabel) karena blok tanda tangan
    // sekarang berada di dalam tabel 2 kolom. Setiap elemen menyimpan startIndex/endIndex dan
    // nomor urutnya agar blok tanda tangan bisa dijaga utuh.
    const paragraf = [];
    const kumpulkan_ = (elements) => {
      (elements || []).forEach((el) => {
        if (el.paragraph) {
          const teks = (el.paragraph.elements || [])
            .map((e) => (e.textRun && e.textRun.content) || "")
            .join("")
            .trim();
          paragraf.push({ teks: teks, startIndex: el.startIndex, endIndex: el.endIndex });
        } else if (el.table) {
          (el.table.tableRows || []).forEach((row) => {
            (row.tableCells || []).forEach((cell) => kumpulkan_(cell.content));
          });
        }
      });
    };
    kumpulkan_(content);

    // Cari paragraf "Banjarmasin, <tanggal laporan>" sebagai penanda -- isinya unik (memuat
    // tanggal laporan spesifik), tidak mungkin kebetulan sama dengan teks lain di dokumen.
    let idxAnchor = -1;
    for (let i = 0; i < paragraf.length; i++) {
      if (paragraf[i].teks === teksBanjarmasin) {
        idxAnchor = i;
        break;
      }
    }
    if (idxAnchor < 0) return; // Anchor tidak ketemu (struktur berubah?) -- jangan paksa apa-apa.

    // Blok yang dijaga tetap utuh: paragraf "Banjarmasin, ..." sampai paragraf NIP (anchor+6:
    // "Kepala Kantor Wilayah," + 3 baris kosong + nama + NIP).
    const idxMulai = idxAnchor;
    const idxAkhir = Math.min(paragraf.length - 1, idxAnchor + 6);
    const elMulai = paragraf[idxMulai];
    const elAkhir = paragraf[idxAkhir];

    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            updateParagraphStyle: {
              range: { startIndex: elMulai.startIndex, endIndex: elAkhir.endIndex },
              paragraphStyle: { keepWithNext: true },
              fields: "keepWithNext",
            },
          },
        ],
      },
      docId,
    );
  } catch (error) {
    console.error("Gagal menjaga blok tanda tangan tetap utuh (Docs API belum aktif?): " + (error.message || error));
  }
}

// PENYEBAB "jarak Enter"/baris kosong nyasar sebelum konten di sel tabel (kop surat, tabel
// keterangan, tabel dokumentasi): TableCell (juga Body) di DocumentApp SELALU punya minimal
// 1 paragraf; appendTableCell("") membuat cell dengan 1 paragraf KOSONG, dan .clear() TIDAK
// benar-benar menghilangkan paragraf itu -- ia cuma mengosongkan isinya, paragrafnya sendiri
// tetap ada. Jadi pola umum `cell.clear(); cell.appendParagraph(teks);` menghasilkan 2
// paragraf (satu kosong duluan, baru isinya), yang tampak sebagai baris kosong SEBELUM
// konten sebenarnya -- persis pola yang dikeluhkan user (gap selalu 1x di awal tiap sel,
// tidak pernah di antara baris-baris berikutnya). Helper ini menghapus paragraf kosong sisa
// itu SETELAH konten sungguhan ditambahkan (dipanggil di akhir tiap fungsi yang memakai pola
// clear()+append di atas) -- container tidak boleh dibiarkan 0 child sama sekali, makanya
// baru dihapus setelah ada child lain untuk menggantikannya.
function hapusParagrafKosongAwal_(cell) {
  try {
    if (cell.getNumChildren() > 1) {
      const anak0 = cell.getChild(0);
      if (anak0.getType() === DocumentApp.ElementType.PARAGRAPH && anak0.asParagraph().getText() === "") {
        cell.removeChild(anak0);
      }
    }
  } catch (_) {}
}

// Meniru letterhead di #printableArea (index.html): logo di kiri, nama instansi & alamat/
// kontak rata TENGAH di kolom teksnya sendiri (bukan rata kiri) -- kolom teks mengisi sisa
// lebar baris (2 kolom: logo, teks; tanpa kolom spacer kanan terpisah), lalu teks di dalamnya
// di-center relatif terhadap kolom itu.
function tambahKopSurat_(body, logoDataUri) {
  const table = body.appendTable();
  table.setBorderWidth(0);
  const tr = table.appendTableRow();

  const cellLogo = tr.appendTableCell("");
  cellLogo.setWidth(90);
  cellLogo.clear();
  if (logoDataUri) sisipkanLogoKeSel_(cellLogo, logoDataUri);
  hapusParagrafKosongAwal_(cellLogo);

  const cellTeks = tr.appendTableCell("");
  cellTeks.clear();
  // Hanya nama unit eselon I (baris terakhir) yang tebal & lebih besar -- dua baris nama
  // kementerian/direktorat di atasnya huruf biasa, sesuai kop surat resmi.
  [
    ["KEMENTERIAN IMIGRASI DAN PEMASYARAKATAN REPUBLIK INDONESIA", false, 10],
    ["DIREKTORAT JENDERAL PEMASYARAKATAN", false, 10],
    ["KANTOR WILAYAH KALIMANTAN SELATAN", true, 12],
  ].forEach(([teks, bold, size]) => {
    const p = cellTeks.appendParagraph(teks);
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.editAsText().setBold(bold).setUnderline(false).setFontSize(size);
    // Baris-baris kop surat harus rapat, bukan jarak default Google Docs (~10pt after).
    p.setSpacingAfter(0);
  });
  // Alamat & kontak digabung jadi 1 baris (sebelumnya 2 baris terpisah) supaya kop 1 baris
  // lebih ringkas sebelum garis horizontal.
  const pAlamat = cellTeks.appendParagraph("Jalan Jenderal A. Yani Km. 5,5 No. 24, Banjarmasin, Kalimantan Selatan   Telepon 085252502005   Pos-el: kanwilditjenpaskalsel@gmail.com");
  pAlamat.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pAlamat.editAsText().setBold(false).setUnderline(false).setFontSize(7.5);
  pAlamat.setSpacingAfter(0);
  hapusParagrafKosongAwal_(cellTeks);

  const hr = body.appendHorizontalRule();
  const pGaris = hr.getParent().asParagraph();
  pGaris.setSpacingBefore(0).setSpacingAfter(0);
  // Apps Script (bukan Google Docs pada umumnya) punya kuirk: appendTable() bisa menyisakan
  // paragraf kosong tambahan tepat setelah tabel, TERPISAH dari paragraf yang baru dibuat
  // appendHorizontalRule() di atas untuk menampung garisnya -- percobaan pertama (cek child
  // terakhir SEBELUM appendHorizontalRule()) ternyata salah target (masih kena tabelnya
  // sendiri), garis tetap kelihatan berjarak. Diperbaiki dengan menyapu MUNDUR dari paragraf
  // garis: hapus semua paragraf kosong yang nyangkut tepat sebelum paragraf garis (berhenti
  // otomatis begitu ketemu tabel kop, karena tabel bukan ElementType.PARAGRAPH).
  let idx = body.getChildIndex(pGaris) - 1;
  while (idx >= 0) {
    const anak = body.getChild(idx);
    if (anak.getType() === DocumentApp.ElementType.PARAGRAPH && anak.asParagraph().getText() === "") {
      body.removeChild(anak);
      idx = body.getChildIndex(pGaris) - 1;
    } else {
      break;
    }
  }
}

function tambahParagrafTengah_(body, text, bold, size) {
  const p = body.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.editAsText()
    .setBold(!!bold)
    .setUnderline(false)
    .setFontSize(size || 11);
  // Spasi tunggal, tanpa jarak ekstra setelah paragraf -- sesuai contoh (`w:line="240"`,
  // `w:after="0"`) untuk judul laporan & heading "DOKUMENTASI KEGIATAN"/"WFH PADA ...".
  p.setLineSpacing(1).setSpacingAfter(0);
  return p;
}

// Jarak baris (line spacing) 1,5 dipakai di seluruh isi surat (judul bagian, sub-judul,
// paragraf isi, list) -- diukur langsung dari `w:line="360"` (360/240 = 1,5) di XML contoh
// laporan resmi (datasets/*.docx), BUKAN 1 (default DocumentApp) atau 2 (spasi ganda).
// Bagian di luar isi surat (judul laporan, tanda tangan, heading "DOKUMENTASI KEGIATAN")
// sengaja TIDAK dipakaikan ini karena di contoh mereka memakai `w:line="240"` (spasi tunggal).
const SPASI_BARIS_ISI_SURAT = 1.5;

// Judul bagian utama bersurat: "A. PENDAHULUAN", "B. PELAKSANAAN KEGIATAN", dst,
// mengikuti penomoran huruf pada #printableArea & contoh laporan resmi.
function tambahJudulBagianHuruf_(body, huruf, text) {
  body.appendParagraph("");
  const p = body.appendParagraph(huruf + ". " + text);
  p.editAsText().setBold(true).setUnderline(false);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingAfter(0);
  return p;
}

// Sub-judul bernomor di dalam satu bagian, mis. "1. Umum" / "2. Dasar Pelaksanaan".
function tambahSubJudulNomor_(body, nomor, text) {
  const p = body.appendParagraph(nomor + ". " + text);
  p.editAsText().setBold(true).setUnderline(false);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingAfter(0);
  return p;
}

// Semua teks isi (paragraf biasa, item list, sel tabel) HARUS ditegaskan bold(false) dan
// underline(false) secara eksplisit di sini, jangan mengandalkan "default" DocumentApp --
// pernah kejadian gaya bold dari heading sebelumnya, dan gaya underline dari blok tanda
// tangan (nama Kakanwil pakai .setUnderline(true)), ikut terbawa ke teks isi berikutnya
// kalau tidak ditegaskan ulang.
function tambahParagrafIsi_(body, text) {
  const p = body.appendParagraph(text);
  p.editAsText().setBold(false).setUnderline(false);
  // Rata kanan-kiri (justify) supaya sisi kanan paragraf rapi sejajar seperti surat resmi --
  // sebelumnya default DocumentApp adalah rata kiri saja (LEFT), sisi kanan bergerigi.
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingAfter(0);
  return p;
}

// Tabel keterangan tanpa garis (Hari/Tanggal, Jumlah Pegawai, dst) supaya kolom nilainya
// rapi sejajar, menggantikan perataan pakai spasi manual yang gampang meleset.
// `value` boleh berupa string biasa (satu baris, mis. "Jum'at, 7 Agustus 2026"), ATAU array
// nama (dipakai untuk "Pegawai Cuti"/"Pegawai Dinas Luar") -- kalau array, dirender sebagai
// daftar bernomor 1./2./dst, satu nama per baris, bukan digabung koma dalam satu baris panjang.
function tambahTabelKeterangan_(body, rows) {
  const table = body.appendTable();
  table.setBorderWidth(0);

  // Paragraf baru di Google Docs default punya "space after" (~10pt, dari gaya Normal text
  // bawaan) yang KELIHATAN seperti baris kosong kalau tidak ditegaskan ke 0 -- ini yang
  // sebelumnya bikin tiap baris (label, dan tiap nama bernomor di sel nilai) terlihat renggang
  // dengan jarak seperti ada Enter kosong di antaranya, padahal cuma ganti paragraf biasa.
  const rapikanParagraf_ = (p) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    p.editAsText().setBold(false).setUnderline(false);
    return p;
  };

  rows.forEach(([label, value]) => {
    const tr = table.appendTableRow();

    const cLabel = tr.appendTableCell("");
    cLabel.setWidth(140);
    cLabel.clear();
    rapikanParagraf_(cLabel.appendParagraph(label));
    hapusParagrafKosongAwal_(cLabel);

    const cValue = tr.appendTableCell("");
    cValue.clear();
    if (Array.isArray(value) && value.length) {
      value.forEach((nama, i) => {
        // Baris pertama diberi ": " di depan nomornya, baris berikutnya nomor lanjut tanpa
        // tanda titik dua lagi (indentasi manual 2 spasi supaya kira-kira sejajar).
        const prefix = i === 0 ? ": " : "  ";
        rapikanParagraf_(cValue.appendParagraph(prefix + (i + 1) + ". " + nama));
      });
    } else {
      rapikanParagraf_(cValue.appendParagraph(": " + (Array.isArray(value) ? "Tidak ada" : value)));
    }
    hapusParagrafKosongAwal_(cValue);
  });
}

function tambahListNomor_(body, items) {
  (items || []).forEach((text) => {
    const item = body.appendListItem(String(text));
    item.setGlyphType(DocumentApp.GlyphType.NUMBER);
    item.editAsText().setBold(false).setUnderline(false);
    // Rata kanan-kiri, sama seperti tambahParagrafIsi_ -- item list juga bagian isi surat.
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    item.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingAfter(0);
  });
}

// List berhuruf kecil (a./b./c./d.) untuk bagian "HASIL YANG DI CAPAI", mengikuti gaya
// penomoran di contoh laporan resmi.
function tambahListHuruf_(body, items) {
  (items || []).forEach((text) => {
    const item = body.appendListItem(String(text));
    item.setGlyphType(DocumentApp.GlyphType.LATIN_LOWER);
    item.editAsText().setBold(false).setUnderline(false);
    // Rata kanan-kiri, sama seperti tambahParagrafIsi_ -- item list juga bagian isi surat.
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    item.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingAfter(0);
  });
}

// Tabel dokumentasi: 2 orang per baris (No.|Dokumentasi|Keterangan diulang 2x), meniru contoh
// laporan resmi "14 Agustus 2026" di datasets/ yang menaruh dokumentasi magang berdampingan
// 2 kolom per baris untuk hemat halaman (di sana dicapai lewat page-section landscape + kolom
// koran khusus Word; Google Docs/DocumentApp TIDAK punya fitur setara -- Class Body tidak
// mengekspos multi-kolom maupun ganti orientasi halaman per section -- jadi efek "2 kolom"
// di sini didekati dengan tabel 6-kolom per baris (2 set No/Dokumentasi/Keterangan) yang tetap
// di halaman A4 portrait, dipakai untuk tabel Pegawai maupun Magang atas permintaan user).
function tambahTabelDokumentasi_(body, rows, tipe) {
  if (!rows || !rows.length) {
    body.appendParagraph("Tidak ada data.").editAsText().setBold(false);
    return;
  }

  // Ukuran font sel tabel sengaja dibuat bertingkat kecil (header 10pt, uraian 9pt, nama 10pt,
  // NIP/NIM 9pt, jabatan 8pt) meniru contoh laporan resmi (datasets/*.docx) -- di sana nama
  // pegawai memang tidak dibesarkan seperti teks isi surat, dan NIP/jabatan malah lebih kecil
  // lagi, supaya tabel yang berulang untuk puluhan orang tidak terlihat kebesaran/berat.
  const FONT_HEADER = 10;
  const FONT_NO = 10;
  const FONT_URAIAN_LABEL = 9;
  const FONT_URAIAN = 9;
  const FONT_NAMA = 10;
  const FONT_NIP = 9;
  const FONT_JABATAN = 8;

  // Lebar kolom per set (No./Dokumentasi/Keterangan) dibagi 2 dari lebar konten body A4 portrait.
  // Body width = 595.3 pt, margin kiri 56.65, margin kanan 56.65 -> Lebar konten = ~482 pt.
  // Dibagi dua orang per baris: masing-masing ~241 pt (No: 20 pt, Foto: 80 pt, Keterangan: ~141 pt).
  const LEBAR_NO = 20;
  const LEBAR_FOTO = 80;
  const LEBAR_TARGET_GAMBAR = 75;

  const table = body.appendTable();
  const header = table.appendTableRow();
  for (let i = 0; i < 2; i++) {
    const cNo = header.appendTableCell("No");
    const cDok = header.appendTableCell("Dokumentasi");
    const cKet = header.appendTableCell("Keterangan");
    // .setUnderline(false) WAJIB ditegaskan di sini dan di semua teks tabel di bawah -- blok
    // tanda tangan sebelumnya (pNamaTtd) memakai .setUnderline(true) untuk nama Kakanwil, dan
    // tanpa ditegaskan ulang, gaya underline itu ikut terbawa ke teks tabel Dokumentasi
    // Kegiatan yang ditulis setelahnya (bug yang sama seperti catatan bold di atas, tapi
    // untuk underline).
    [cNo, cDok, cKet].forEach((c) => c.editAsText().setBold(true).setUnderline(false).setFontSize(FONT_HEADER));
  }

  const rapikan_ = (p, bold, size) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    p.editAsText().setBold(bold).setUnderline(false).setFontSize(size);
    return p;
  };

  // isiSatuOrang_ menulis satu set No./Dokumentasi/Keterangan ke 3 sel yang sudah ada di baris
  // `tr` -- dipanggil 2x per baris (nomor genap-ganjil), atau dibiarkan kosong kalau jumlah
  // orang ganjil (slot ke-2 di baris terakhir).
  const isiSatuOrang_ = (tr, nomor, row) => {
    const cellNo = tr.appendTableCell(nomor != null ? String(nomor) : "");
    cellNo.setWidth(LEBAR_NO);
    cellNo.editAsText().setBold(true).setUnderline(false).setFontSize(FONT_NO);

    const cellFoto = tr.appendTableCell("");
    cellFoto.setWidth(LEBAR_FOTO);
    if (row) {
      cellFoto.clear();
      sisipkanGambarKeSel_(cellFoto, row.foto, row.nama, LEBAR_TARGET_GAMBAR);
      hapusParagrafKosongAwal_(cellFoto);
    }

    // rapikan_() menegaskan setSpacingAfter(0) supaya baris "Uraian Kegiatan :", isi uraian,
    // nama, NIP, dan jabatan menempel rapat -- tanpa ini tiap baris punya jarak ekstra bawaan
    // Google Docs (~10pt) yang terlihat seperti ada Enter kosong di antaranya. Baris kosong di
    // tengah (setelah uraian) SENGAJA dipertahankan sebagai spacer sebelum nama -- yang
    // dihapus hapusParagrafKosongAwal_() hanya paragraf kosong SISA .clear() di awal cell.
    const cellKet = tr.appendTableCell("");
    if (row) {
      cellKet.clear();
      rapikan_(cellKet.appendParagraph("Uraian Kegiatan :"), true, FONT_URAIAN_LABEL);
      rapikan_(cellKet.appendParagraph(String(row.uraian || "-")), false, FONT_URAIAN);
      rapikan_(cellKet.appendParagraph(""), false, FONT_URAIAN);
      rapikan_(cellKet.appendParagraph(String(row.nama || "-")), true, FONT_NAMA);
      const labelId = tipe === "pegawai" ? "NIP. " : "NIP/NIM. ";
      rapikan_(cellKet.appendParagraph(labelId + String(row.nipNim || "-")), false, FONT_NIP);
      rapikan_(cellKet.appendParagraph(String(row.jabatan || "-")), false, FONT_JABATAN);
      hapusParagrafKosongAwal_(cellKet);
    }
  };

  for (let i = 0; i < rows.length; i += 2) {
    const tr = table.appendTableRow();
    isiSatuOrang_(tr, i + 1, rows[i]);
    isiSatuOrang_(tr, rows[i + 1] ? i + 2 : null, rows[i + 1] || null);
  }

  // Lebar kolom No./Dokumentasi diulang untuk kedua set (indeks 0-2 set pertama, 3-5 set
  // kedua); kolom Keterangan (indeks 2 & 5) sengaja dibiarkan otomatis mengisi sisa lebar
  // (DocumentApp meregangkan tabel penuh ke lebar konten body).
  table.setColumnWidth(0, LEBAR_NO);
  table.setColumnWidth(1, LEBAR_FOTO);
  table.setColumnWidth(3, LEBAR_NO);
  table.setColumnWidth(4, LEBAR_FOTO);
}

function sisipkanGambarKeSel_(cell, fotoUrl, namaUntukLog, lebarTarget) {
  try {
    const fileId = ekstrakDriveFileId_(fotoUrl);
    if (!fileId) throw new Error("URL foto tidak dikenali.");

    const blob = DriveApp.getFileById(fileId).getBlob();
    const image = cell.appendImage(blob);
    // Lebar-tetap (default 75pt, sesuai kolom Dokumentasi yang sudah dipersempit untuk layout
    // 2-orang-per-baris) dengan tinggi proporsional -- dibatasi tinggi maksimum di
    // resizeGambarProporsional_() supaya foto potret (portrait, rasio tinggi > lebar) tidak
    // membuat baris tabel melar berlebihan dibanding baris lain yang fotonya landscape/persegi.
    resizeGambarProporsional_(image, lebarTarget || 75, 130);
  } catch (error) {
    console.error("Gagal menyisipkan foto untuk " + namaUntukLog + ": " + (error.message || error));
    cell.appendParagraph("(Foto tidak tersedia)").editAsText().setBold(false).setFontSize(9);
  }
}

// Menyisipkan logo (dikirim client sebagai data URI base64) ke sel kop surat.
function sisipkanLogoKeSel_(cell, dataUri) {
  try {
    const matches = String(dataUri || "").match(/^data:(.+?);base64,(.+)$/);
    if (!matches) return;
    const bytes = Utilities.base64Decode(matches[2]);
    const blob = Utilities.newBlob(bytes, matches[1], "logo.png");
    const image = cell.appendImage(blob);
    resizeGambarProporsional_(image, 82);
  } catch (_) {}
}

// `tinggiMaks` opsional -- kalau diisi dan foto berorientasi potret (portrait, tinggi jauh
// lebih besar dari lebar setelah diskalakan ke `lebarTarget`), skala ulang berdasarkan tinggi
// supaya tidak melebihi `tinggiMaks`, baru lebar menyesuaikan (tetap proporsional, jadi lebih
// sempit dari `lebarTarget` -- bukan terpotong/gepeng). Dipakai untuk foto dokumentasi kegiatan
// yang banyak dikirim dalam bentuk potret (foto selfie WFH), supaya baris tabel tidak melar
// berlebihan dibanding baris lain yang fotonya landscape/persegi.
function resizeGambarProporsional_(image, lebarTarget, tinggiMaks) {
  const lebarAsli = image.getWidth();
  const tinggiAsli = image.getHeight();
  const rasio = lebarAsli > 0 ? tinggiAsli / lebarAsli : 1;
  let lebar = lebarTarget;
  let tinggi = Math.round(rasio * lebar);
  if (tinggiMaks && tinggi > tinggiMaks) {
    tinggi = tinggiMaks;
    lebar = rasio > 0 ? Math.round(tinggi / rasio) : lebarTarget;
  }
  image.setWidth(lebar);
  image.setHeight(tinggi);
}

function ekstrakDriveFileId_(url) {
  const s = String(url || "");
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  return "";
}

// Dipanggil dari client (renderDokumentasiTable di index.html) untuk menampilkan foto di
// preview web. Sengaja TIDAK mengandalkan <img src> langsung ke URL Drive, karena banyak
// kebijakan Google Workspace (terutama instansi pemerintah) memblokir/menonaktifkan sharing
// "Anyone with link" sehingga foto gagal tampil bagi siapa pun di luar akun pemilik file.
// DriveApp di sini berjalan sebagai pemilik skrip sehingga selalu punya akses, apa pun
// pengaturan sharing filenya, lalu dikirim ke client sebagai data URI base64.
function getFotoDataUri(fotoUrl) {
  try {
    const fileId = ekstrakDriveFileId_(fotoUrl);
    if (!fileId) return "";
    const blob = DriveApp.getFileById(fileId).getBlob();
    return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    console.error("Gagal memuat foto: " + (error.message || error));
    return "";
  }
}

// ------------------------- HELPERS -------------------------

function readPegawai_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  return rows
    .filter((r) => r.some((v) => String(v).trim() !== ""))
    .map((r) => ({
      id: normalizeId_(r[0]),
      nama: String(r[1]).trim(),
      nip: String(r[2]).trim(),
      golongan: String(r[3]).trim(),
      jabatan: String(r[4]).trim(),
      status: String(r[5] || "Aktif").trim(),
    }));
}

function readMagang_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  return rows
    .filter((r) => r.some((v) => String(v).trim() !== ""))
    .map((r) => ({
      id: normalizeId_(r[0]),
      nama: String(r[1]).trim(),
      nim: String(r[2]).trim(),
      instansi: String(r[3]).trim(),
      posisi: String(r[4]).trim(),
      status: String(r[5] || "Aktif").trim(),
    }));
}

function readPresensi_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getDisplayValues();

  return rows
    .filter((r) => r.some((v) => String(v).trim() !== ""))
    .map((r, idx) => {
      const timestamp = String(r[0]).trim();
      const waktu = timestamp.includes(" ") ? timestamp.split(" ")[1] : timestamp;
      return {
        id: "PRES_ROW_" + (idx + 2),
        timestamp: timestamp,
        tanggal: String(r[1]).trim(),
        waktu: waktu || "-",
        kategori: String(r[2]).trim(),
        userId: normalizeId_(r[3]),
        nama: String(r[4]).trim(),
        nipNim: String(r[5]).trim(),
        golongan: String(r[6]).trim(),
        jabatan: String(r[7]).trim(),
        status: String(r[8]).trim(),
        uraian: String(r[9]).trim(),
        foto: String(r[10]).trim(),
      };
    })
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, APP_TIMEZONE, pattern);
}

function makeId_(prefix) {
  return prefix + "-" + Utilities.getUuid().split("-")[0].toUpperCase();
}

// Sheet sering menampilkan ID lama yang diketik manual sebagai angka (mis. "14" tersimpan/
// terformat menjadi "14.0"), sementara ID baru dari makeId_() berupa teks ("PEG-XXXXXXXX").
// Tanpa normalisasi, perbandingan ID antar sheet bisa gagal dan menyebabkan data terlihat
// tidak sinkron (mis. statistik kehadiran magang meleset). Fungsi ini menyeragamkan bentuknya.
function normalizeId_(value) {
  const s = String(value == null ? "" : value).trim();
  return /^\d+\.0+$/.test(s) ? s.replace(/\.0+$/, "") : s;
}