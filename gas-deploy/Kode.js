const SPREADSHEET_ID = "1WvufYa4Yr8JB53cs6uMIP8fQH0ccDGhziYuQYGEYUa8"; // ID Spreadsheet database SIKERJA WFH.
const SHEET_PEGAWAI = "MasterPegawai";
const SHEET_MAGANG = "MasterMagang";
const SHEET_PRESENSI = "PresensiWFH";
const APP_TIMEZONE = "Asia/Makassar";
const DEFAULT_LOGO_URL = "https://kemenimipas.go.id/images/logo/Kementerian-Hukum-Dan-Ham-Kemenkumham-Logo-Vector.png";
const MAX_FOTO_BYTES = 5 * 1024 * 1024; // Batas ukuran foto dokumentasi: 5 MB
const BARIS_KOSONG_TTD = 5; // Jumlah baris kosong untuk ruang tanda tangan Kepala Kanwil (sebelumnya 3)
// Ukuran foto dalam POINT (pt). PENTING: setWidth()/setHeight() Apps Script memakai PIKSEL (1 pt = 4/3 px),
// karena itu nilai pt di bawah dikonversi dulu. Kolom Dokumentasi = 150 pt, tetapi garis tabel menyita sebagian
// ruang di sisi kiri-kanan, sehingga lebar foto dibuat 146.25 pt (= 195 px): hampir menyentuh garis
// namun tidak menutupi garis di sisi kanan. Jika ingin lebih rapat/renggang, ubah nilai ini (kelipatan 0.75 pt).
const FOTO_LEBAR_MAKS_PT = 146.25;
const FOTO_TINGGI_MAKS_PT = 240; // batas tinggi agar 2 baris foto tegak tetap muat dalam 1 halaman landscape
const PX_PER_PT = 4 / 3;
// Logo resmi diambil dari file di Google Drive (sumber utama untuk website dan header PDF/Docs).
// ID = bagian panjang pada link Drive: /file/d/<ID>/view
const LOGO_DRIVE_FILE_ID = "1od2aUJBtlSWVDT1x0QPQ_ZR-oDNnMmHy";

// ======================= ADMIN LOGIN =======================
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

// Logo permanen: dipakai web (cadangan bila browser gagal memuat URL) dan dokumen laporan.
function getLogoDataUri() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("LOGO_DATA_URI_V3");
    if (cached) return cached;

    const blob = ambilBlobLogoResmi_("");
    if (!blob) return "";
    const uri = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    if (uri.length < 90000) cache.put("LOGO_DATA_URI_V3", uri, 21600);
    return uri;
  } catch (error) {
    console.warn("Gagal memuat logo: " + (error.message || error));
    return "";
  }
}

function getSpreadsheet() {
  if (SPREADSHEET_ID && String(SPREADSHEET_ID).trim() !== "") {
    return SpreadsheetApp.openById(String(SPREADSHEET_ID).trim());
  }

  const savedId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (savedId && String(savedId).trim() !== "") {
    return SpreadsheetApp.openById(String(savedId).trim());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const newSs = SpreadsheetApp.create("SIKERJA WFH - Database Kehadiran");
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", newSs.getId());
  return newSs;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("SIKERJA WFH - Kanwil Ditjenpas Kalimantan Selatan")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

    const requestedStatus = String(formData.statusKehadiran || "").trim();
    if (requestedStatus === "Cuti" || requestedStatus === "Dinas Luar") {
      throw new Error("Status Cuti dan Dinas Luar hanya dapat ditandai oleh Admin.");
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRESENSI);
    const now = new Date();
    const timestampStr = formatDate_(now, "yyyy-MM-dd HH:mm:ss");
    const tanggalStr = formatDate_(now, "yyyy-MM-dd");

    const values = sheet.getDataRange().getDisplayValues();
    const idUserNorm = normalizeId_(formData.idUser);
    const kategoriNorm = String(formData.kategori || "").trim();
    const nipNimNorm = String(formData.nipNim || "").trim();
    let rowToUpdate = -1;
    for (let i = 1; i < values.length; i++) {
      const sameDate = String(values[i][1]).trim() === tanggalStr;
      const sameKat = String(values[i][2]).trim() === kategoriNorm;
      if (!sameDate || !sameKat) continue;

      // NIP/NIM adalah kunci utama. ID hanya dipakai bila NIP/NIM tidak tersedia,
      // karena ID di master bisa berubah/dipakai ulang untuk orang lain.
      const cocok = nipNimNorm !== ""
        ? String(values[i][5]).trim() === nipNimNorm
        : (idUserNorm !== "" && normalizeId_(values[i][3]) === idUserNorm);

      if (cocok) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate > 0) {
      const statusSebelumnya = String(values[rowToUpdate - 1][8] || "").trim();
      throw new Error(
        statusSebelumnya
          ? "Anda sudah tercatat presensi hari ini (" + statusSebelumnya + "). Presensi hanya dapat dilakukan satu kali per hari."
          : "Anda sudah tercatat presensi hari ini. Presensi hanya dapat dilakukan satu kali per hari."
      );
    }

    let fotoUrl = "";
    if (formData.fotoBase64) {
      fotoUrl = saveImageToDrive(formData.fotoBase64, formData.namaLengkap, tanggalStr);
    }

    const rowData = [
      timestampStr,
      tanggalStr,
      kategoriNorm,
      idUserNorm,
      String(formData.namaLengkap || "").trim(),
      nipNimNorm,
      String(formData.golongan || "").trim(),
      String(formData.jabatan || "").trim(),
      requestedStatus,
      String(formData.uraianKegiatan || "").trim(),
      fotoUrl,
    ];

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
    if (bytes.length > MAX_FOTO_BYTES) {
      throw new Error("Ukuran foto melebihi batas maksimal 5 MB.");
    }
    const cleanNama = String(namaUser || "user").replace(/[^a-zA-Z0-9]/g, "_");
    const ext = contentType.includes("png") ? "png" : contentType.includes("webp") ? "webp" : "jpg";
    const blob = Utilities.newBlob(bytes, contentType, `WFH_${cleanNama}_${tanggal}.${ext}`);

    const folder = getOrCreateFolder_("SIKERJA_WFH_FOTO");
    const file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {}

    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (error) {
    throw new Error("Gagal menyimpan foto: " + (error.message || String(error)));
  }
}

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

function setStatusAdminCepat(userId, status, nipNim) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    userId = normalizeId_(userId);
    nipNim = String(nipNim || "").trim();
    status = String(status || "").trim();

    if (!userId && !nipNim) throw new Error("Identitas pegawai tidak valid.");
    if (status && status !== "Cuti" && status !== "Dinas Luar") {
      throw new Error("Status tidak valid.");
    }

    const ss = getSpreadsheet();
    const shPegawai = ss.getSheetByName(SHEET_PEGAWAI);
    const shPresensi = ss.getSheetByName(SHEET_PRESENSI);
    if (!shPegawai || !shPresensi) throw new Error("Sheet database belum lengkap.");

    const pegLastRow = shPegawai.getLastRow();
    if (pegLastRow < 2) throw new Error("Master pegawai kosong.");

    const pegValues = shPegawai.getRange(2, 1, pegLastRow - 1, 6).getDisplayValues();
    // NIP/NIM adalah kunci utama; ID hanya cadangan bila NIP/NIM tidak tersedia.
    const pegRowIdx = nipNim
      ? pegValues.findIndex((r) => String(r[2]).trim() === nipNim)
      : pegValues.findIndex((r) => userId !== "" && normalizeId_(r[0]) === userId);
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

    let existingRow = -1;
    let existingStatus = "";
    const lastRow = shPresensi.getLastRow();

    if (lastRow >= 2) {
      const presValues = shPresensi.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
      for (let i = 0; i < presValues.length; i++) {
        const row = presValues[i];
        const tanggal = String(row[1]).trim();
        const kategori = String(row[2]).trim();
        const uid = normalizeId_(row[3]);
        const rowNip = String(row[5] || "").trim();
        const st = String(row[8]).trim();
        // NIP/NIM didahulukan; ID hanya dipakai bila NIP/NIM tidak tersedia.
        const cocok = peg.nip
          ? rowNip === peg.nip
          : (userId !== "" && uid === userId);
        if (tanggal === tanggalStr && kategori === "Pegawai" && cocok) {
          existingRow = i + 2;
          existingStatus = st;
          break;
        }
      }
    }

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

    const selectedByKey = {};
    statusRows.forEach((item) => {
      const id = normalizeId_(item && item.idUser);
      const p = id ? pegawaiById[id] : null;
      const nip = p ? String(p.nip || "").trim() : "";
      const key = nip ? "NIP:" + nip : (id ? "ID:" + id : "");
      const status = String((item && item.status) || "").trim();
      if (!key) return;
      selectedByKey[key] = {
        status: status === "Cuti" || status === "Dinas Luar" ? status : "",
        pegawai: p
      };
    });

    const rowsToWrite = [];
    const rowsToClear = [];
    const existingTodayKeys = new Set();

    existing.forEach((row, idx) => {
      const tanggal = String(row[1] || "").trim();
      const kategori = String(row[2] || "").trim();
      const userId = normalizeId_(row[3]);
      const rowNip = String(row[5] || "").trim();
      const existingStatus = String(row[8] || "").trim();

      if (tanggal !== tanggalStr || kategori !== "Pegawai") return;
      if (!rowNip && !userId) return;

      // NIP/NIM didahulukan; ID hanya dipakai untuk baris lama yang tidak menyimpan NIP/NIM.
      const key = rowNip ? "NIP:" + rowNip : "ID:" + userId;
      existingTodayKeys.add(key);

      if (existingStatus === "Hadir WFH") return;

      if (!Object.prototype.hasOwnProperty.call(selectedByKey, key)) return;

      const newStatus = selectedByKey[key].status;
      const sheetRow = idx + 2;

      if (!newStatus) {
        if (existingStatus === "Cuti" || existingStatus === "Dinas Luar") {
          rowsToClear.push(sheetRow);
        }
        return;
      }

      const p = selectedByKey[key].pegawai;
      if (!p) return;

      rowsToWrite.push({
        row: sheetRow,
        values: [[timestampStr, tanggalStr, "Pegawai", p.id, p.nama, p.nip, p.golongan, p.jabatan, newStatus, newStatus === "Cuti" ? "Cuti - ditandai oleh Admin" : "Dinas Luar - ditandai oleh Admin", ""]],
      });
    });

    const rowsToAppend = [];
    Object.keys(selectedByKey).forEach((key) => {
      const sel = selectedByKey[key];
      if (!sel.status || existingTodayKeys.has(key)) return;
      const p = sel.pegawai;
      if (!p) return;

      rowsToAppend.push([timestampStr, tanggalStr, "Pegawai", p.id, p.nama, p.nip, p.golongan, p.jabatan, sel.status, sel.status === "Cuti" ? "Cuti - ditandai oleh Admin" : "Dinas Luar - ditandai oleh Admin", ""]);
    });

    rowsToWrite.forEach((item) => {
      shPresensi.getRange(item.row, 1, 1, 11).setValues(item.values);
    });

    rowsToClear.forEach((rowNum) => {
      shPresensi.getRange(rowNum, 1, 1, 11).clearContent();
    });

    if (rowsToAppend.length) {
      const startRow = shPresensi.getLastRow() + 1;
      shPresensi.getRange(startRow, 1, rowsToAppend.length, 11).setValues(rowsToAppend);
    }

    SpreadsheetApp.flush();

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

// Membatalkan (menghapus) satu baris presensi HARI INI atas perintah Admin.
// Dipakai bila nama tidak muncul di form absen / rekap namun tercatat di monitoring,
// sehingga orang tersebut dapat mengisi presensi ulang dengan data yang benar.
function batalkanPresensi(presensiId, userId, nipNim, nama) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupDatabase();

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRESENSI);
    if (!sheet) throw new Error("Sheet presensi tidak ditemukan.");

    const tanggalStr = formatDate_(new Date(), "yyyy-MM-dd");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data presensi.");

    const values = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    const uid = normalizeId_(userId);
    const nipNorm = String(nipNim || "").trim();
    const namaNorm = String(nama || "").trim().toLowerCase();

    const rowUid = (i) => normalizeId_(values[i][3]);
    const rowNip = (i) => String(values[i][5] || "").trim();
    const rowNama = (i) => String(values[i][4] || "").trim().toLowerCase();
    const isToday = (i) => String(values[i][1] || "").trim() === tanggalStr;

    // NIP/NIM adalah kunci utama. ID hanya dipakai bila NIP/NIM tidak tersedia,
    // agar baris orang lain tidak ikut terhapus saat ID master berubah/dipakai ulang.
    const cocokIdentitas = (i) => {
      if (nipNorm) return rowNip(i) === nipNorm;
      if (uid) return rowUid(i) === uid;
      return !!namaNorm && rowNama(i) === namaNorm;
    };

    let targetRow = -1;

    // 1) Baris yang dipilih admin (PRES_ROW_n), diverifikasi dulu agar tidak salah hapus bila baris bergeser.
    const match = String(presensiId || "").match(/^PRES_ROW_(\d+)$/);
    if (match) {
      const idx = Number(match[1]) - 2;
      if (idx >= 0 && idx < values.length && isToday(idx) && cocokIdentitas(idx)) {
        targetRow = idx + 2;
      }
    }

    // 2) Bila baris tidak cocok lagi, cari baris presensi hari ini berdasarkan NIP/NIM (ID/nama hanya bila kosong).
    if (targetRow < 0) {
      for (let i = values.length - 1; i >= 0; i--) {
        if (isToday(i) && cocokIdentitas(i)) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow < 0) throw new Error("Data presensi hari ini untuk nama tersebut tidak ditemukan atau sudah dibatalkan.");

    sheet.deleteRow(targetRow);
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

const FONT_DOKUMEN = "Arial";
const SPASI_BARIS_ISI_SURAT = 1.5;

function buatRekapGoogleDocs(payload) {
  let hasil;
  try {
    hasil = buatDokumenRekap_(payload);
    const file = DriveApp.getFileById(hasil.id);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (_) {}
    return { success: true, id: hasil.id, url: hasil.url, warning: peringatanLogo_(hasil) };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  }
}

// Catatan langkah dokumen yang gagal (agar tidak gagal diam-diam). Ditampilkan sebagai peringatan di web.
const CATATAN_DOKUMEN = [];
function catatDokumen_(pesan) {
  console.error(pesan);
  if (CATATAN_DOKUMEN.indexOf(pesan) === -1 && CATATAN_DOKUMEN.length < 6) CATATAN_DOKUMEN.push(pesan);
}

function peringatanLogo_(hasil) {
  const daftar = [];
  if (hasil && hasil.logoOk === false) {
    daftar.push("Logo tidak berhasil dimuat ke dokumen (cek LOGO_DRIVE_FILE_ID di kode_1.gs).");
  }
  CATATAN_DOKUMEN.forEach((c) => daftar.push(c));
  return daftar.join("\n");
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

    try {
      docFile.setTrashed(true);
    } catch (_) {}

    return { success: true, id: pdfFile.getId(), url: pdfFile.getUrl(), warning: peringatanLogo_(hasil) };
  } catch (error) {
    console.error(error);
    return { success: false, error: error.message || String(error) };
  }
}

function namaFileRekap_(payload) {
  const tgl = String((payload && payload.tanggal) || formatDate_(new Date(), "yyyy-MM-dd"));
  return "Laporan WFH Kanwil Ditjenpas Kalsel " + tgl;
}

function buatDokumenRekap_(payload) {
  CATATAN_DOKUMEN.length = 0;
  if (!payload || !payload.tanggal) throw new Error("Data rekap tidak valid. Pilih tanggal terlebih dahulu.");

  const doc = DocumentApp.create(namaFileRekap_(payload));
  const body = doc.getBody();
  body.clear();

  // Ukuran A4 Portrait (595.3 x 841.9 pt) -> Lebar efektif konten = 595.3 - 56.65 - 56.65 = 482 pt
  body.setPageWidth(595.3).setPageHeight(841.9);
  body.setMarginTop(42.5).setMarginBottom(56.65).setMarginLeft(56.65).setMarginRight(56.65);

  const tanggalLengkap = String(payload.tanggalLengkap || payload.tanggal);
  const tanggalBiasa = String(payload.tanggalBiasa || payload.tanggal);
  const tanggalJudul = String(payload.tanggalJudul || tanggalLengkap).toUpperCase();
  const tanggalDokumentasi = String(payload.tanggalDokumentasi || tanggalBiasa.toUpperCase());
  const jumlahPegawai = Number(payload.jumlahPegawai || 0);
  const jumlahWFH = Number(payload.jumlahWFH || 0);
  const jumlahCuti = Number(payload.jumlahCuti || 0);
  const jumlahDinasLuar = Number(payload.jumlahDinasLuar || 0);
  const jumlahMagang = Number(payload.jumlahMagang || 0);
  const pegawaiCuti = Array.isArray(payload.pegawaiCuti) ? payload.pegawaiCuti : [];
  const pegawaiDinasLuar = Array.isArray(payload.pegawaiDinasLuar) ? payload.pegawaiDinasLuar : [];
  const pegawaiWFH = Array.isArray(payload.pegawaiWFH) ? payload.pegawaiWFH : [];
  const magang = Array.isArray(payload.magang) ? payload.magang : [];
  
  // PENYESUAIAN 1: Memaksa nama Kakanwil selalu baru di Google Docs & PDF (Jika frontend masih bawa data lama)
  const kakanwil = payload.kakanwil || {};
  if (!kakanwil.nama || kakanwil.nama.includes("Erwedi")) {
    kakanwil.nama = "Bagus Kurniawan, M.A., S.Sos.";
    kakanwil.nip = "197103291994031001";
  }

  // ---------- KOP SURAT ----------
  let logoOk = tambahKopSurat_(body, payload.logoDataUri);
  hapusParagrafKosongAwal_(body);

  // ---------- JUDUL LAPORAN ----------
  tambahParagrafTengah_(body, "LAPORAN PELAKSANAAN WORK FROM HOME (WFH)", true, 11);
  tambahParagrafTengah_(body, "DI LINGKUNGAN KANTOR WILAYAH DIREKTORAT JENDERAL", true, 11);
  tambahParagrafTengah_(body, "PEMASYARAKATAN KALIMANTAN SELATAN", true, 11);
  const pTanggalJudul = tambahParagrafTengah_(body, tanggalJudul, true, 11);
  pTanggalJudul.setSpacingAfter(10);

  // ---------- A. PENDAHULUAN ----------
  tambahJudulBagianHuruf_(body, "A", "PENDAHULUAN");
  tambahSubJudulNomor_(body, "1", "Umum", 18);
  tambahParagrafIsi_(
    body,
    "Menindaklanjuti Surat Edaran Menteri Imigrasi dan Pemasyarakatan Nomor 2 Tahun 2026 tentang Pelaksanaan Tugas Kedinasan bagi Pegawai Aparatur Sipil Negara melalui Penerapan Work From Home dalam rangka mendukung percepatan transformasi tata kelola penyelenggaraan pemerintahan di lingkungan Kementerian Imigrasi dan Pemasyarakatan, serta Surat Sekretaris Direktorat Jenderal Pemasyarakatan Nomor PAS.1-SA.08.05-1410 tanggal 08 April 2026 tentang penyampaian pelaksanaan Work From Home (WFH), perlu dilakukan tindak lanjut pelaksanaan di lingkungan Kantor Wilayah dan Unit Pelaksana Teknis Pemasyarakatan.",
    18,
    18
  );
  tambahSubJudulNomor_(body, "2", "Dasar Pelaksanaan", 18);
  tambahListNomor_(
    body,
    [
      "Surat Edaran Menteri Imigrasi dan Pemasyarakatan Nomor 2 Tahun 2026 tentang Pelaksanaan Tugas Kedinasan bagi Pegawai Aparatur Sipil Negara melalui Penerapan Work From Home dalam rangka mendukung percepatan transformasi tata kelola penyelenggaraan pemerintahan.",
      "Surat Sekretaris Direktorat Jenderal Pemasyarakatan Nomor PAS.1-SA.08.05-1410 tanggal 08 April 2026 tentang Penyampaian Pelaksanaan Work From Home (WFH) di lingkungan Direktorat Jenderal Pemasyarakatan.",
    ],
    18
  );

  // ---------- B. PELAKSANAAN KEGIATAN ----------
  tambahJudulBagianHuruf_(body, "B", "PELAKSANAAN KEGIATAN");
  const pPengantarB = tambahParagrafIsi_(
    body,
    "Pelaksanaan Work From Home (WFH) di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan dilaksanakan pada:",
    18,
    18
  );
  pPengantarB.setSpacingAfter(4);

  const jumlahPegawaiKet =
    jumlahPegawai +
    " orang (" +
    jumlahWFH +
    " WFH" +
    (jumlahCuti ? ", " + jumlahCuti + " Cuti" : "") +
    (jumlahDinasLuar ? ", " + jumlahDinasLuar + " Dinas Luar" : "") +
    ")";
  const barisKeterangan = [
    ["Hari/Tanggal", ": " + tanggalLengkap],
    ["Jumlah Pegawai", ": " + jumlahPegawaiKet],
  ];
  if (jumlahMagang > 0) barisKeterangan.push(["Jumlah peserta magang", ": " + jumlahMagang + " orang"]);
  barisKeterangan.push(["Pegawai Cuti", pegawaiCuti.length ? pegawaiCuti : ": Tidak ada"]);
  barisKeterangan.push(["Pegawai Dinas Luar", pegawaiDinasLuar.length ? pegawaiDinasLuar : ": Tidak ada"]);
  tambahTabelKeterangan_(body, barisKeterangan, 18);

  // "Adapun pelaksanaannya sebagai berikut:" dimulai di halaman 2. Dengan Docs API dipakai properti
  // "pageBreakBefore" pada paragraf itu sendiri (tanpa paragraf pemisah halaman, jadi tidak ada halaman kosong).
  const pakaiDocsApi = typeof Docs !== "undefined" && Docs.Documents;
  if (!pakaiDocsApi) body.appendPageBreak();
  const pAdapun = tambahParagrafIsi_(body, "Adapun pelaksanaannya sebagai berikut:", 0, 0);
  pAdapun.setSpacingBefore(0);
  tambahListNomor_(
    body,
    [
      "WFH dilaksanakan pada " + tanggalLengkap + ".",
      "Seluruh Pegawai ASN Kanwil yang melaksanakan fungsi dukungan manajemen bekerja secara fleksibel dari lokasi masing-masing.",
      "Pegawai melakukan presensi secara daring melalui aplikasi STAR-ASN.",
      "Pegawai menyampaikan laporan harian kegiatan kepada atasan langsung.",
      "Koordinasi dan komunikasi kedinasan dilakukan melalui media daring.",
      "Pengawasan dan pemantauan capaian kinerja dilakukan oleh atasan langsung.",
    ],
    0
  );

  // ---------- C. HASIL YANG DI CAPAI ----------
  tambahJudulBagianHuruf_(body, "C", "HASIL YANG DICAPAI");
  tambahListHuruf_(
    body,
    [
      "Pelaksanaan WFH berjalan tertib dan lancar.",
      "Presensi pegawai tercatat melalui sistem daring.",
      "Laporan harian kegiatan tersampaikan dengan baik.",
      "Tugas dan fungsi Kanwil tetap berjalan efektif tanpa kendala yang berarti.",
    ],
    0
  );

  // ---------- D. SIMPULAN ----------
  tambahJudulBagianHuruf_(body, "D", "SIMPULAN");
  tambahParagrafIsi_(
    body,
    "Pelaksanaan penyesuaian tugas kedinasan melalui Work From Home (WFH) di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan pada " +
      tanggalBiasa +
      " telah dilaksanakan sesuai ketentuan dan berjalan dengan baik serta tetap menjamin ketercapaian kinerja organisasi.",
    0,
    28
  );

  // ---------- E. PENUTUP ----------
  tambahJudulBagianHuruf_(body, "E", "PENUTUP");
  const pPenutup = tambahParagrafIsi_(
    body,
    "Demikian laporan pelaksanaan penyesuaian tugas kedinasan melalui Work From Home (WFH) ini disusun sebagai bentuk pertanggungjawaban dan dokumentasi pelaksanaan kebijakan pimpinan di lingkungan Kantor Wilayah Direktorat Jenderal Pemasyarakatan Kalimantan Selatan.",
    0,
    28
  );
  pPenutup.setSpacingAfter(14);

  // ---------- TANDA TANGAN PEJABAT ----------
  const tableTtd = body.appendTable();
  tableTtd.setBorderWidth(0);
  const rowTtd = tableTtd.appendTableRow();

  const cellKiri = rowTtd.appendTableCell("");
  cellKiri.setWidth(250);
  cellKiri.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);

  const cellKanan = rowTtd.appendTableCell("");
  cellKanan.setWidth(232);
  cellKanan.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cellKanan.clear();

  const rapikanTtd_ = (p, bold, underline) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
    p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(bold).setUnderline(underline).setFontSize(11);
    return p;
  };

  rapikanTtd_(cellKanan.appendParagraph("Banjarmasin, " + tanggalBiasa), false, false);
  rapikanTtd_(cellKanan.appendParagraph("Kepala Kantor Wilayah,"), false, false);
  for (let i = 0; i < BARIS_KOSONG_TTD; i++) {
    rapikanTtd_(cellKanan.appendParagraph(""), false, false);
  }
  rapikanTtd_(cellKanan.appendParagraph(String(kakanwil.nama || "-")), false, false);
  rapikanTtd_(cellKanan.appendParagraph("NIP. " + String(kakanwil.nip || "-")), false, false);
  hapusParagrafKosongAwal_(cellKanan);

  // ---------- DOKUMENTASI KEGIATAN: PEGAWAI ----------
  tambahParagrafTengah_(body, "DOKUMENTASI KEGIATAN", true, 11);
  const pSubDokPeg = tambahParagrafTengah_(body, "WFH PADA " + tanggalDokumentasi, true, 11);
  pSubDokPeg.setSpacingAfter(8);
  tambahTabelDokumentasi_(body, pegawaiWFH, "pegawai");

  // ---------- DOKUMENTASI KEGIATAN: MAGANG ----------
  if (magang.length) {
    tambahParagrafTengah_(body, "DOKUMENTASI KEGIATAN", true, 11);
    const pSubDokMag = tambahParagrafTengah_(body, "WFH PESERTA MAGANG PADA " + tanggalDokumentasi, true, 11);
    pSubDokMag.setSpacingAfter(8);
    tambahTabelDokumentasi_(body, magang, "magang");
  }

  // Paragraf kosong di ujung dokumen (wajib ada setelah tabel) dikecilkan agar tidak memunculkan halaman kosong
  try {
    const terakhir = body.getChild(body.getNumChildren() - 1);
    if (terakhir.getType() === DocumentApp.ElementType.PARAGRAPH && terakhir.asParagraph().getText() === "") {
      terakhir.asParagraph().setLineSpacing(1).setSpacingBefore(0).setSpacingAfter(0);
      terakhir.asParagraph().setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 });
    }
  } catch (_) {}

  doc.saveAndClose();

  if (!logoOk && pakaiDocsApi) logoOk = sisipkanLogoViaDocsApi_(doc.getId());

  const halamanBaruOk = pakaiDocsApi ? mulaiHalamanBaruAdapun_(doc.getId()) : true;
  if (!halamanBaruOk) sisipkanPemisahHalamanAdapun_(doc.getId());
  sesuaikanOrientasiDanFotoDokumentasi_(doc.getId());
  jagaTandaTanganTetapUtuh_(doc.getId(), "Banjarmasin, " + tanggalBiasa);
  return { id: doc.getId(), url: doc.getUrl(), logoOk: logoOk };
}

const TEKS_ADAPUN = "Adapun pelaksanaannya sebagai berikut:";

// Menandai paragraf "Adapun..." agar selalu mulai di awal halaman (halaman 2), tanpa paragraf kosong tambahan.
function mulaiHalamanBaruAdapun_(docId) {
  try {
    const dokumen = Docs.Documents.get(docId);
    const content = (dokumen.body && dokumen.body.content) || [];
    for (let i = 0; i < content.length; i++) {
      const el = content[i];
      if (!el.paragraph) continue;
      const teks = (el.paragraph.elements || []).map((e) => (e.textRun && e.textRun.content) || "").join("").trim();
      if (teks === TEKS_ADAPUN) {
        Docs.Documents.batchUpdate(
          {
            requests: [
              {
                updateParagraphStyle: {
                  range: { startIndex: el.startIndex, endIndex: el.endIndex },
                  paragraphStyle: { pageBreakBefore: true },
                  fields: "pageBreakBefore",
                },
              },
            ],
          },
          docId
        );
        return true;
      }
    }
  } catch (error) {
    catatDokumen_("Halaman 2 gagal diatur (aktifkan layanan Docs API): " + (error.message || error));
  }
  return false;
}

// Cadangan bila properti pageBreakBefore tidak dapat dipakai: sisipkan pemisah halaman tepat sebelum paragraf "Adapun...".
function sisipkanPemisahHalamanAdapun_(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    for (let i = 0; i < body.getNumChildren(); i++) {
      const anak = body.getChild(i);
      if (anak.getType() === DocumentApp.ElementType.PARAGRAPH && anak.asParagraph().getText().trim() === TEKS_ADAPUN) {
        body.insertPageBreak(i);
        break;
      }
    }
    doc.saveAndClose();
  } catch (error) {
    catatDokumen_("Pemisah halaman gagal disisipkan: " + (error.message || error));
  }
}

function jagaTandaTanganTetapUtuh_(docId, teksBanjarmasin) {
  try {
    if (typeof Docs === "undefined" || !Docs.Documents) return;
    const dokumen = Docs.Documents.get(docId);
    const content = (dokumen.body && dokumen.body.content) || [];

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

    let idxAnchor = -1;
    for (let i = 0; i < paragraf.length; i++) {
      if (paragraf[i].teks === teksBanjarmasin) {
        idxAnchor = i;
        break;
      }
    }
    if (idxAnchor < 0) return;

    const idxMulai = idxAnchor;
    // Banjarmasin + Kepala Kantor Wilayah + baris kosong + nama + NIP
    const idxAkhir = Math.min(paragraf.length - 1, idxAnchor + 2 + BARIS_KOSONG_TTD + 1);
    const elMulai = paragraf[idxMulai];
    // Paragraf terakhir (NIP) tidak diberi keepWithNext agar blok tidak "menempel" ke bagian setelahnya
    const elAkhir = paragraf[Math.max(idxMulai, idxAkhir - 1)];

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
      docId
    );
  } catch (error) {
    catatDokumen_("Blok tanda tangan gagal dijaga utuh: " + (error.message || error));
  }
}

function sesuaikanOrientasiDanFotoDokumentasi_(docId) {
  try {
    if (typeof Docs === "undefined" || !Docs.Documents) {
      return;
    }

    const dokumen = Docs.Documents.get(docId);
    const content = (dokumen.body && dokumen.body.content) || [];

    const posisiHeading = [];
    const kumpulkanParagraf_ = (elements) => {
      (elements || []).forEach((el) => {
        if (el.paragraph) {
          const teks = (el.paragraph.elements || [])
            .map((e) => (e.textRun && e.textRun.content) || "")
            .join("")
            .trim();
          if (teks === "DOKUMENTASI KEGIATAN") {
            posisiHeading.push({ startIndex: el.startIndex, endIndex: el.endIndex });
          }
        } else if (el.table) {
          (el.table.tableRows || []).forEach((row) => {
            (row.tableCells || []).forEach((cell) => kumpulkanParagraf_(cell.content));
          });
        }
      });
    };
    kumpulkanParagraf_(content);

    if (!posisiHeading.length) return;

    const requestsBreak = [];
    [...posisiHeading].reverse().forEach((h) => {
      requestsBreak.push({
        insertSectionBreak: {
          location: { index: h.startIndex },
          sectionType: "NEXT_PAGE",
        },
      });
    });

    Docs.Documents.batchUpdate({ requests: requestsBreak }, docId);

    const dokumen2 = Docs.Documents.get(docId);
    const content2 = (dokumen2.body && dokumen2.body.content) || [];

    const posisiHeading2 = [];
    const cariHeading_ = (elements) => {
      (elements || []).forEach((el) => {
        if (el.paragraph) {
          const teks = (el.paragraph.elements || [])
            .map((e) => (e.textRun && e.textRun.content) || "")
            .join("")
            .trim();
          if (teks === "DOKUMENTASI KEGIATAN") {
            posisiHeading2.push({ startIndex: el.startIndex, endIndex: el.endIndex });
          }
        } else if (el.table) {
          (el.table.tableRows || []).forEach((row) => {
            (row.tableCells || []).forEach((cell) => cariHeading_(cell.content));
          });
        }
      });
    };
    cariHeading_(content2);

    const sectionStyleLandscape = {
      flipPageOrientation: true,
      marginTop: { magnitude: 42.5, unit: "PT" },
      marginBottom: { magnitude: 42.5, unit: "PT" },
      marginLeft: { magnitude: 56.65, unit: "PT" },
      marginRight: { magnitude: 56.65, unit: "PT" },
    };

    const requestsStyle = [];

    posisiHeading2.forEach((h) => {
      requestsStyle.push({
        updateSectionStyle: {
          range: { startIndex: h.startIndex, endIndex: h.endIndex },
          sectionStyle: sectionStyleLandscape,
          fields: "flipPageOrientation,marginTop,marginBottom,marginLeft,marginRight",
        },
      });
    });

    if (requestsStyle.length > 0) {
      Docs.Documents.batchUpdate({ requests: requestsStyle }, docId);
    }

    // KUNCI LEBAR KOLOM: No 24 pt, Dokumentasi 150 pt, Keterangan 190 pt (2 set = 728 pt, pas lebar halaman landscape)
    const requestsLock = [];
    const kunciKolom_ = (startIndex, kolom, lebarPt) => {
      requestsLock.push({
        updateTableColumnProperties: {
          tableStartLocation: { index: startIndex },
          columnIndices: kolom,
          tableColumnProperties: { widthType: "FIXED_WIDTH", width: { magnitude: lebarPt, unit: "PT" } },
          fields: "widthType,width",
        },
      });
    };
    const cariTabel_ = (elements) => {
      elements.forEach((el) => {
        if (el.table && el.table.columns === 6) {
          kunciKolom_(el.startIndex, [0, 3], 24);
          kunciKolom_(el.startIndex, [1, 4], 150);
          kunciKolom_(el.startIndex, [2, 5], 190);
        }
      });
    };
    cariTabel_(content2);

    if (requestsLock.length > 0) {
      Docs.Documents.batchUpdate({ requests: requestsLock }, docId);
    }
  } catch (error) {
    catatDokumen_("Halaman landscape / lebar kolom dokumentasi gagal diatur (aktifkan layanan Docs API): " + (error.message || error));
  }
}

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

function tambahKopSurat_(body, logoDataUri) {
  // REVISI HEADER RATA TENGAH PRESISI:
  // Lebar konten A4 Portrait = 482 pt.
  // Kolom Logo (78 pt) + Kolom Teks Tengah (354 pt) + Kolom Penyeimbang Kanan (50 pt) = 482 pt,
  // dengan indentasi paragraf dikunci 0 sehingga teks alamat/jalan berada persis di tengah.
  const table = body.appendTable();
  table.setBorderWidth(0);
  const tr = table.appendTableRow();

  const cellLogo = tr.appendTableCell("");
  cellLogo.setWidth(78);
  cellLogo.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cellLogo.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cellLogo.clear();
  const logoOk = sisipkanLogoKeSel_(cellLogo, logoDataUri);
  hapusParagrafKosongAwal_(cellLogo);
  if (cellLogo.getNumChildren() > 0 && cellLogo.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    cellLogo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER).setIndentStart(0).setIndentFirstLine(0);
  }

  const cellTeks = tr.appendTableCell("");
  cellTeks.setWidth(374);
  cellTeks.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cellTeks.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cellTeks.clear();

  const tambahBarisKop_ = (teks, bold, size, spaceBefore = 0) => {
    const p = cellTeks.appendParagraph(teks);
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.setIndentStart(0).setIndentFirstLine(0).setIndentEnd(0);
    p.setLineSpacing(1.05).setSpacingBefore(spaceBefore).setSpacingAfter(0);
    p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(bold).setUnderline(false).setFontSize(size);
    return p;
  };

  tambahBarisKop_("KEMENTERIAN IMIGRASI DAN PEMASYARAKATAN REPUBLIK INDONESIA", false, 9.5, 0);
  tambahBarisKop_("DIREKTORAT JENDERAL PEMASYARAKATAN", false, 9.5, 0);
  tambahBarisKop_("KANTOR WILAYAH KALIMANTAN SELATAN", true, 11.5, 0);

  // Baris alamat jalan & baris telepon/email dipisah menjadi 2 paragraf CENTER mandiri agar benar-benar simetris
  tambahBarisKop_("Jalan Jenderal A. Yani Km. 5,5 No. 24, Banjarmasin, Kalimantan Selatan", false, 8, 2);

  const pKontak = tambahBarisKop_("Telepon 085252502005   Pos-el: kanwilditjenpaskalsel@gmail.com", false, 8, 0);
  const textKontak = pKontak.editAsText();
  const emailStr = "kanwilditjenpaskalsel@gmail.com";
  const pText = pKontak.getText();
  const emailIdx = pText.indexOf(emailStr);
  if (emailIdx !== -1) {
    textKontak.setForegroundColor(emailIdx, emailIdx + emailStr.length - 1, "#0000FF");
    textKontak.setItalic(emailIdx, emailIdx + emailStr.length - 1, true);
  }

  hapusParagrafKosongAwal_(cellTeks);

  // Kolom ke-3 sebagai penyeimbang kanan agar titik tengah teks kop sejajar dengan judul di bawahnya
  const cellKanan = tr.appendTableCell("");
  cellKanan.setWidth(30);
  cellKanan.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cellKanan.clear();
  hapusParagrafKosongAwal_(cellKanan);

  const hr = body.appendHorizontalRule();
  const pGaris = hr.getParent().asParagraph();
  pGaris.setSpacingBefore(2).setSpacingAfter(8);
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
  return logoOk;
}

function tambahParagrafTengah_(body, text, bold, size) {
  const p = body.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setIndentStart(0).setIndentFirstLine(0).setIndentEnd(0);
  p.editAsText()
    .setFontFamily(FONT_DOKUMEN)
    .setBold(!!bold)
    .setUnderline(false)
    .setFontSize(size || 11);
  p.setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
  return p;
}

function tambahJudulBagianHuruf_(body, huruf, text) {
  const p = body.appendParagraph(huruf + ". " + text);
  p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(true).setUnderline(false).setFontSize(11);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingBefore(10).setSpacingAfter(2);
  return p;
}

function tambahSubJudulNomor_(body, nomor, text, leftIndent = 0) {
  const p = body.appendParagraph(nomor + ". " + text);
  if (leftIndent > 0) {
    p.setIndentStart(leftIndent);
    p.setIndentFirstLine(leftIndent);
  }

  p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(true).setUnderline(false).setFontSize(11);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingBefore(2).setSpacingAfter(0);
  return p;
}

function tambahParagrafIsi_(body, text, leftIndent = 0, firstLineIndent = 0) {
  const p = body.appendParagraph(text);

  if (leftIndent > 0) {
    p.setIndentStart(leftIndent);
  }
  if (firstLineIndent > 0) {
    p.setIndentFirstLine(leftIndent + firstLineIndent);
  } else if (firstLineIndent === 0 && leftIndent > 0) {
    p.setIndentFirstLine(leftIndent);
  }

  p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setUnderline(false).setFontSize(11);
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  p.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingBefore(0).setSpacingAfter(2);
  return p;
}

function tambahTabelKeterangan_(body, rows, indent = 18) {
  const table = body.appendTable();
  table.setBorderWidth(0);

  const rapikanParagraf_ = (p) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.15);
    p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setUnderline(false).setFontSize(11);
    return p;
  };

  rows.forEach(([label, value]) => {
    const tr = table.appendTableRow();

    const cLabel = tr.appendTableCell("");
    cLabel.setWidth(165);
    cLabel.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(indent).setPaddingRight(0);
    cLabel.clear();
    rapikanParagraf_(cLabel.appendParagraph(label));
    hapusParagrafKosongAwal_(cLabel);

    const cValue = tr.appendTableCell("");
    cValue.setWidth(317);
    cValue.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(0);
    cValue.clear();
    if (Array.isArray(value) && value.length === 1) {
      // Hanya 1 orang: langsung nama, tanpa nomor
      rapikanParagraf_(cValue.appendParagraph(": " + value[0]));
    } else if (Array.isArray(value) && value.length > 1) {
      // Lebih dari 1 orang: diurutkan per nomor
      value.forEach((nama, i) => {
        const prefix = i === 0 ? ": " : "  ";
        rapikanParagraf_(cValue.appendParagraph(prefix + (i + 1) + ". " + nama));
      });
    } else {
      rapikanParagraf_(cValue.appendParagraph(Array.isArray(value) ? ": Tidak ada" : value));
    }
    hapusParagrafKosongAwal_(cValue);
  });
}

function tambahListNomor_(body, items, indent = 0) {
  (items || []).forEach((text) => {
    const item = body.appendListItem(String(text));
    item.setGlyphType(DocumentApp.GlyphType.NUMBER);

    if (indent > 0) {
      item.setIndentStart(indent + 18);
      item.setIndentFirstLine(indent);
    }

    item.editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setUnderline(false).setFontSize(11);
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    item.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingBefore(0).setSpacingAfter(0);
  });
}

function tambahListHuruf_(body, items, indent = 0) {
  (items || []).forEach((text) => {
    const item = body.appendListItem(String(text));
    item.setGlyphType(DocumentApp.GlyphType.LATIN_LOWER);

    if (indent > 0) {
      item.setIndentStart(indent + 18);
      item.setIndentFirstLine(indent);
    }

    item.editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setUnderline(false).setFontSize(11);
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    item.setLineSpacing(SPASI_BARIS_ISI_SURAT).setSpacingBefore(0).setSpacingAfter(0);
  });
}

function tambahTabelDokumentasi_(body, rows, tipe) {
  if (!rows || !rows.length) {
    body.appendParagraph("Tidak ada data.").editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setFontSize(10);
    return;
  }

  const FONT_HEADER = 10;
  const FONT_NO = 10;
  const FONT_URAIAN_LABEL = 9;
  const FONT_URAIAN = 9;
  const FONT_NAMA = 10;
  const FONT_NIP = 9;
  const FONT_JABATAN = 8;

  const table = body.appendTable();
  const header = table.appendTableRow();
  for (let i = 0; i < 2; i++) {
    const cNo = header.appendTableCell("No");
    const cDok = header.appendTableCell("Dokumentasi");
    const cKet = header.appendTableCell("Keterangan");

    [cNo, cDok, cKet].forEach((c) => {
      c.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4);
      c.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      c.editAsText().setFontFamily(FONT_DOKUMEN).setBold(true).setUnderline(false).setFontSize(FONT_HEADER);
      c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
    });
  }

  const rapikan_ = (p, bold, size) => {
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.1);
    p.editAsText().setFontFamily(FONT_DOKUMEN).setBold(bold).setUnderline(false).setFontSize(size);
    return p;
  };

  const isiSatuOrang_ = (tr, nomor, row) => {
    const cellNo = tr.appendTableCell(nomor != null ? String(nomor) : "");
    cellNo.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(2).setPaddingRight(2);
    cellNo.editAsText().setFontFamily(FONT_DOKUMEN).setBold(true).setUnderline(false).setFontSize(FONT_NO);
    cellNo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cellNo.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

    const cellFoto = tr.appendTableCell("");
    if (row) {
      cellFoto.clear();
      cellFoto.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
      cellFoto.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

      // Foto memenuhi lebar kolom dokumentasi (lihat sisipkanGambarKeSel_)
      sisipkanGambarKeSel_(cellFoto, row.foto, row.nama);
      
      // Hapus jarak bawaan paragraf Google Docs agar tidak ada sela putih
      if (cellFoto.getNumChildren() > 0) {
        const p = cellFoto.getChild(0).asParagraph();
        p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1)
         .setIndentStart(0).setIndentFirstLine(0).setIndentEnd(0)
         .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
      hapusParagrafKosongAwal_(cellFoto);
    }

    const cellKet = tr.appendTableCell("");
    cellKet.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(5).setPaddingRight(5);
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

  // Lebar total Landscape A4 = 728 pt -> 2 set x (24 + 150 + 190 = 364 pt) = 728 pt
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() === 6) {
      row.getCell(0).setWidth(24);
      row.getCell(1).setWidth(150);
      row.getCell(2).setWidth(190);
      row.getCell(3).setWidth(24);
      row.getCell(4).setWidth(150);
      row.getCell(5).setWidth(190);
    }
  }
}

// Membaca ukuran asli gambar (piksel) langsung dari data JPEG/PNG/GIF, agar rasio selalu benar (tidak gepeng).
function ukuranGambarDariBlob_(blob) {
  try {
    const b = blob.getBytes();
    const u = (i) => b[i] & 0xff;
    if (u(0) === 0x89 && u(1) === 0x50 && u(2) === 0x4e && u(3) === 0x47) {
      return { w: (u(16) << 24) | (u(17) << 16) | (u(18) << 8) | u(19), h: (u(20) << 24) | (u(21) << 16) | (u(22) << 8) | u(23) };
    }
    if (u(0) === 0x47 && u(1) === 0x49 && u(2) === 0x46) {
      return { w: u(6) | (u(7) << 8), h: u(8) | (u(9) << 8) };
    }
    if (u(0) === 0xff && u(1) === 0xd8) {
      let i = 2;
      while (i + 9 < b.length) {
        if (u(i) !== 0xff) { i++; continue; }
        const m = u(i + 1);
        if (m === 0xff) { i++; continue; }
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
          return { w: (u(i + 7) << 8) | u(i + 8), h: (u(i + 5) << 8) | u(i + 6) };
        }
        if (m === 0x01 || (m >= 0xd0 && m <= 0xd9)) { i += 2; continue; }
        i += 2 + ((u(i + 2) << 8) | u(i + 3));
      }
    }
  } catch (_) {}
  return null;
}

function sisipkanGambarKeSel_(cell, fotoUrl, namaUntukLog) {
  try {
    const fileId = ekstrakDriveFileId_(fotoUrl);
    if (!fileId) throw new Error("URL foto tidak dikenali.");

    const blob = DriveApp.getFileById(fileId).getBlob();
    const image = cell.appendImage(blob);

    // Rasio diambil dari data gambar itu sendiri; cadangan: ukuran yang dilaporkan Docs
    const dim = ukuranGambarDariBlob_(blob);
    const lebarAsli = (dim && dim.w) || image.getWidth() || 1;
    const tinggiAsli = (dim && dim.h) || image.getHeight() || 1;

    // Foto diperbesar sampai selebar kolom (mentok garis kiri-kanan). Jika foto sangat tinggi (portrait),
    // yang dibatasi adalah tingginya. Lebar dan tinggi SELALU diskalakan bersama, sehingga tidak gepeng.
    const kotakLebarPx = FOTO_LEBAR_MAKS_PT * PX_PER_PT;
    const kotakTinggiPx = FOTO_TINGGI_MAKS_PT * PX_PER_PT;
    const skala = Math.min(kotakLebarPx / lebarAsli, kotakTinggiPx / tinggiAsli);
    const lebar = Math.max(1, Math.round(lebarAsli * skala));
    const tinggi = Math.max(1, Math.round(tinggiAsli * skala));

    cell.setPaddingLeft(0).setPaddingRight(0).setPaddingTop(0).setPaddingBottom(0);

    image.setWidth(lebar);
    image.setHeight(tinggi);

    // Bersihkan indentasi & jarak paragraf tempat gambar berada (tanpa sela putih di atas/bawah foto)
    const p = image.getParent().asParagraph();
    p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    p.setIndentFirstLine(0).setIndentStart(0).setIndentEnd(0);
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    try {
      // Ukuran huruf paragraf dikecilkan agar tidak ada ruang kosong di bawah gambar
      p.setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 });
    } catch (_) {}
  } catch (error) {
    catatDokumen_("Foto gagal dimasukkan (" + (namaUntukLog || "-") + "): " + (error.message || error));
    cell.appendParagraph("(Foto tidak tersedia)").editAsText().setFontFamily(FONT_DOKUMEN).setBold(false).setFontSize(9);
  }
}

// Memastikan blob benar-benar gambar (PNG/JPEG/GIF) berdasarkan isi file, bukan sekadar status HTTP 200.
function blobGambarValid_(blob) {
  try {
    if (!blob) return null;
    const b = blob.getBytes();
    if (!b || b.length < 100) return null;
    const u = (i) => b[i] & 0xff;
    let tipe = "";
    if (u(0) === 0x89 && u(1) === 0x50 && u(2) === 0x4e && u(3) === 0x47) tipe = "image/png";
    else if (u(0) === 0xff && u(1) === 0xd8) tipe = "image/jpeg";
    else if (u(0) === 0x47 && u(1) === 0x49 && u(2) === 0x46) tipe = "image/gif";
    if (!tipe) {
      try {
        const png = blob.getAs("image/png");
        const bp = png && png.getBytes();
        if (bp && bp.length > 100 && (bp[0] & 0xff) === 0x89) {
          return png.setContentType("image/png").setName("logo.png");
        }
      } catch (_) {}
      return null;
    }
    return blob.setContentType(tipe).setName("logo." + (tipe === "image/png" ? "png" : tipe === "image/gif" ? "gif" : "jpg"));
  } catch (_) {
    return null;
  }
}

function blobDariDataUri_(dataUri) {
  try {
    const m = String(dataUri || "").match(/^data:(.+?);base64,([\s\S]+)$/);
    if (!m) return null;
    return blobGambarValid_(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], "logo.png"));
  } catch (_) {
    return null;
  }
}

function ambilLogoDariDrive_() {
  try {
    if (LOGO_DRIVE_FILE_ID && String(LOGO_DRIVE_FILE_ID).trim()) {
      const b = blobGambarValid_(DriveApp.getFileById(String(LOGO_DRIVE_FILE_ID).trim()).getBlob());
      if (b) return b;
    }
    const files = getOrCreateFolder_("SIKERJA_WFH_ASET").getFilesByName("logo_resmi");
    if (files.hasNext()) return blobGambarValid_(files.next().getBlob());
  } catch (e) {
    console.warn("Logo cadangan di Drive tidak dapat dibaca: " + (e.message || e));
  }
  return null;
}

function unduhLogoDariUrl_() {
  const enc = encodeURIComponent(DEFAULT_LOGO_URL);
  // Situs resmi membatasi akses otomatis, sehingga disediakan jalur cadangan lewat proxy gambar publik.
  const daftarUrl = [
    DEFAULT_LOGO_URL,
    "https://wsrv.nl/?output=png&url=" + enc,
    "https://images.weserv.nl/?output=png&url=" + enc,
    "https://images1-focus-opensocial.googleusercontent.com/gadgets/proxy?container=focus&gadget=a&no_expand=1&rewriteMime=image/*&url=" + enc,
  ];
  const opsi = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
      Accept: "image/png,image/jpeg,image/*;q=0.8,*/*;q=0.5",
      Referer: "https://kemenimipas.go.id/",
    },
  };
  for (let i = 0; i < daftarUrl.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(daftarUrl[i], opsi);
      if (resp.getResponseCode() === 200) {
        const b = blobGambarValid_(resp.getBlob());
        if (b) return b;
        console.warn("Jalur logo " + (i + 1) + ": respons bukan gambar valid.");
      } else {
        console.warn("Jalur logo " + (i + 1) + ": HTTP " + resp.getResponseCode());
      }
    } catch (e) {
      console.warn("Jalur logo " + (i + 1) + " gagal: " + (e.message || e));
    }
  }
  return null;
}

// Urutan: (1) file logo di Google Drive, (2) gambar kiriman browser, (3) unduh dari URL resmi/proxy lalu simpan ke Drive.
function ambilBlobLogoResmi_(dataUri) {
  const dariDrive = ambilLogoDariDrive_();
  if (dariDrive) return dariDrive;

  const dariBrowser = blobDariDataUri_(dataUri);
  if (dariBrowser) return dariBrowser;

  const dariUrl = unduhLogoDariUrl_();
  if (dariUrl) {
    try {
      const folder = getOrCreateFolder_("SIKERJA_WFH_ASET");
      if (!folder.getFilesByName("logo_resmi").hasNext()) {
        folder.createFile(dariUrl.copyBlob().setName("logo_resmi"));
      }
    } catch (e) {
      console.warn("Logo tidak dapat disimpan ke Drive: " + (e.message || e));
    }
    return dariUrl;
  }
  return null;
}

// Mengembalikan true bila logo berhasil dimasukkan.
function sisipkanLogoKeSel_(cell, dataUri) {
  try {
    const blob = ambilBlobLogoResmi_(dataUri);
    if (!blob) {
      console.error("Logo tidak tersedia dari browser, Drive, maupun URL resmi.");
      return false;
    }
    const image = cell.appendImage(blob);

    const dim = ukuranGambarDariBlob_(blob);
    const lebarAsli = (dim && dim.w) || image.getWidth() || 1;
    const tinggiAsli = (dim && dim.h) || image.getHeight() || 1;
    const rasio = tinggiAsli / lebarAsli;
    const lebarLogoPx = Math.round(72 * PX_PER_PT); // lebar logo 72 pt
    image.setWidth(lebarLogoPx);
    image.setHeight(Math.round(rasio * lebarLogoPx));
    return true;
  } catch (error) {
    console.error("Gagal menyisipkan logo: " + (error.message || error));
    return false;
  }
}

function ekstrakDriveFileId_(url) {
  const s = String(url || "");
  let m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  return "";
}

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

  // Nomor baris dihitung SEBELUM baris kosong dibuang, agar PRES_ROW_n selalu menunjuk baris sheet yang benar.
  return rows
    .map((r, idx) => {
      if (!r.some((v) => String(v).trim() !== "")) return null;
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
    .filter(Boolean)
    .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, APP_TIMEZONE, pattern);
}

function makeId_(prefix) {
  return prefix + "-" + Utilities.getUuid().split("-")[0].toUpperCase();
}

function normalizeId_(value) {
  const s = String(value == null ? "" : value).trim();
  return /^\d+\.0+$/.test(s) ? s.replace(/\.0+$/, "") : s;
}

// Cadangan terakhir: Google Docs sendiri yang mengambil gambar dari URL logo resmi (tanpa melewati UrlFetchApp).
function sisipkanLogoViaDocsApi_(docId) {
  try {
    const dokumen = Docs.Documents.get(docId);
    const content = (dokumen.body && dokumen.body.content) || [];
    for (let i = 0; i < content.length; i++) {
      const tabel = content[i].table;
      if (!tabel) continue;
      const sel = tabel.tableRows[0].tableCells[0];
      const para = sel.content[0];
      Docs.Documents.batchUpdate(
        {
          requests: [
            {
              updateParagraphStyle: {
                range: { startIndex: para.startIndex, endIndex: para.endIndex },
                paragraphStyle: { alignment: "CENTER" },
                fields: "alignment",
              },
            },
            {
              insertInlineImage: {
                location: { index: para.startIndex },
                uri: DEFAULT_LOGO_URL,
                objectSize: { width: { magnitude: 72, unit: "PT" } },
              },
            },
          ],
        },
        docId
      );
      return true;
    }
  } catch (error) {
    console.error("Logo via Docs API gagal: " + (error.message || error));
  }
  return false;
}

// Jalankan manual dari editor Apps Script (pilih fungsi ini lalu Run, buka Execution log)
// untuk melihat jalur logo mana yang berhasil.
function diagnosaLogo() {
  const dariDrive = ambilLogoDariDrive_();
  Logger.log("Logo dari Drive (" + LOGO_DRIVE_FILE_ID + "): " + (dariDrive ? "BERHASIL (" + dariDrive.getBytes().length + " byte, " + dariDrive.getContentType() + ")" : "GAGAL - cek ID file / format gambar"));
  Logger.log("Docs API aktif: " + (typeof Docs !== "undefined" && !!Docs.Documents));
  Logger.log("Hasil akhir: " + (ambilBlobLogoResmi_("") ? "logo tersedia" : "logo TIDAK tersedia -> isi LOGO_DRIVE_FILE_ID"));
}