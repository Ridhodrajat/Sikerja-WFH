# SIKERJA WFH

Sistem Kehadiran & Monitoring Kinerja Pegawai — Kanwil Ditjenpas Kalimantan Selatan.

Aplikasi web presensi WFH berbasis Google Apps Script: satu halaman HTML (`Index.html`) sebagai antarmuka dan `kode.gs` sebagai backend. Google Sheets dipakai sebagai database, Google Drive untuk menyimpan foto dokumentasi dan hasil laporan PDF.

## Fitur

- Presensi harian Pegawai dan Magang dengan status **Hadir WFH**, **Cuti**, atau **Dinas Luar**.
- Foto dokumentasi wajib diunggah (maksimal 5 MB), otomatis disimpan ke folder Drive `SIKERJA_WFH_FOTO`.
- Pencegahan presensi ganda: satu orang hanya bisa tercatat sekali per hari per kategori.
- Mode Admin dengan PIN.
- Monitoring status hari ini, ubah status cepat, dan pembatalan presensi yang salah.
- Master data pegawai dan magang (tambah/hapus).
- Rekapitulasi harian/bulanan dengan filter kategori, plus cetak PDF resmi (kop surat, logo, tabel dokumentasi foto, ruang tanda tangan Kepala Kanwil).
- Pencocokan identitas berbasis **NIP/NIM** sebagai kunci utama.

## Struktur Repository

```
Index.html              Antarmuka web (nama file di Apps Script harus "Index")
kode.gs                 Backend Apps Script
gas-deploy/
  Index.html            Salinan siap deploy via clasp
  Kode.js               Isi sama dengan kode.gs (ekstensi .js untuk clasp)
  .clasp.json
  appsscript.json
Sikerja Work Version/   Arsip versi lama, tidak dipakai
```

Root `Index.html` + `kode.gs` dan `gas-deploy/` harus selalu berisi kode yang sama.

## Struktur Database (Google Sheets)

**MasterPegawai** — `ID | Nama Lengkap | NIP | Golongan | Jabatan | Status`

**MasterMagang** — `ID | Nama Lengkap | NIM | Instansi | Posisi | Status`

**PresensiWFH** — `Timestamp | Tanggal | Kategori | ID_User | Nama Lengkap | NIP_NIM | Golongan | Jabatan | Status_Kehadiran | Uraian_Kegiatan | Foto_URL`

Kolom `Tanggal` disimpan sebagai teks `yyyy-mm-dd`, dan kolom `NIP` / `NIM` / `NIP_NIM` diformat sebagai teks agar angka nol di depan tidak hilang.

## Cara Pemasangan (manual)

1. Buka <https://script.google.com> lalu buat project baru.
2. Salin isi `kode.gs` ke file `Code.gs`.
3. Buat file HTML baru dengan nama `Index`, lalu salin isi `Index.html`.
4. Buka **Project Settings → Script Properties**, tambahkan properti `ADMIN_PIN` dengan PIN admin pilihan Anda.
5. Jalankan fungsi `setupDatabase()` satu kali dan izinkan akses. Sheet `MasterPegawai`, `MasterMagang`, dan `PresensiWFH` akan dibuat otomatis bila belum ada.
6. **Deploy → New deployment → Web app**, dengan *Execute as*: **Me** dan *Who has access*: **Anyone**.
7. Buka URL web app, masuk **Mode Admin**, lalu isi data pegawai dan magang.

## Deploy via clasp (opsional)

```bash
cd gas-deploy
clasp login
clasp push
clasp deploy
```

## Konfigurasi

| Nama | Lokasi | Keterangan |
|---|---|---|
| `SPREADSHEET_ID` | `kode.gs` (baris awal) | ID spreadsheet database. Jika `SpreadsheetApp.getActiveSpreadsheet()` tersedia, nilai kosong akan memakai spreadsheet aktif |
| `ADMIN_PIN` | Script Properties | PIN login admin. **Wajib diisi**, jika kosong login admin ditolak |
| `APP_TIMEZONE` | `kode.gs` | `Asia/Makassar` (WITA) |
| `LOGO_DRIVE_FILE_ID` | `kode.gs` | ID file logo di Drive untuk header web dan kop laporan |
| `MAX_FOTO_BYTES` | `kode.gs` | Batas ukuran foto dokumentasi (5 MB) |

## Aturan Pencocokan Identitas

Presensi dan master data dicocokkan dengan urutan berikut:

1. **NIP/NIM** — kunci utama karena nilainya unik dan tidak berubah.
2. **ID_User** — hanya dipakai bila NIP/NIM tidak tersedia.
3. **Nama lengkap** — hanya untuk data lama yang tidak punya NIP/NIM maupun ID.

`ID_User` tidak dipakai sebagai kunci utama karena nomor ID pada master data dapat berubah atau dipakai ulang setelah penomoran ulang, sehingga bisa membuat pegawai lain terblokir "sudah presensi" atau muncul sebagai orang yang salah di rekap.

## Alur Admin

1. Klik **Mode Admin** dan masukkan PIN.
2. Tab **Status** menampilkan pegawai/magang yang belum dan sudah presensi hari ini.
3. Gunakan tombol status cepat untuk mencatat Cuti/Dinas Luar tanpa menunggu pegawai mengisi form.
4. Gunakan tombol **Batalkan** pada baris presensi yang salah; data hari itu untuk orang tersebut dihapus sehingga ia bisa presensi ulang.
5. Tab **Rekap** untuk mencetak laporan PDF harian atau bulanan.

## Catatan

- Zona waktu aplikasi: `Asia/Makassar` (WITA).
- PIN admin hanya disimpan di Script Properties, tidak ditulis di dalam kode.
- Web app disetel dapat diakses siapa saja yang memiliki link (lihat `gas-deploy/appsscript.json`). PIN admin membatasi menu admin, bukan akses ke fungsi backend, jadi jangan menyebarkan URL aplikasi ke publik luas.
- Jika memakai spreadsheet sendiri, ganti nilai `SPREADSHEET_ID` di `kode.gs` sebelum deploy.
- Setiap perubahan kode harus di-deploy ulang lewat **Deploy → Manage deployments → Edit → New version** agar URL yang dipakai pegawai ikut terbarui.
