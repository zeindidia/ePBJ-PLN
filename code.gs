// ID Spreadsheet Mas Zendid
const SHEET_ID = '1aZhbOu9pVJHpJ5d9ElVrIZBUy5jTX0bduoLZXQILpiQ';

// 1. FUNGSI RENDER HTML
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('e-PBJ UPT Malang V1.2')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Tambahan krusial
}

// 2. FUNGSI UNTUK MENGAMBIL DATA DARI SHEET (DIPANGGIL OLEH FRONTEND)
function getDataFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Data_PBJ');
    
    // Jika sheet belum ada, kembalikan array kosong
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Mapping data Sheet ke dalam bentuk Array of Objects (seperti mockDataPBJ)
    const result = rows.map(row => {
      return {
        id: row[0] || '',
        judul: row[1] || '',
        jenis: row[2] || '',
        pos: row[3] || '',
        status: row[4] || '',
        vendor: row[5] || '',
        no_spk: row[6] || ''
      };
    });
    
    return result;
  } catch (error) {
    Logger.log(error);
    return [];
  }
}

// 3. FUNGSI SETUP DATABASE (JALANKAN SEKALI SAJA DI EDITOR GAS)
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const schema = {
    'Data_PBJ': ['ID_Rencana', 'Judul_Pekerjaan', 'Jenis_Proses', 'Pos_Anggaran', 'Status_Tahapan', 'Vendor_Terpilih', 'No_SPK'],
    'Master_SKK': ['Jenis_SKK', 'No_Surat_SKK', 'Tanggal_SKK', 'Link_Bukti'],
    'Data_PRK': ['No_SKK_Asal', 'No_PRK', 'Nama_Pekerjaan', 'Pagu_Anggaran'],
    'Master_Vendor': ['ID_Vendor', 'Nama_Vendor', 'Kualifikasi', 'Rating'],
    'Bank_Data_RAB': ['Uraian', 'Satuan', 'Harga_Satuan'],
    'Setting_Master': ['Kategori', 'Opsi_Value'],
    'Setting_Lokasi_Pejabat': ['Kategori', 'Nama_Lokasi_Jabatan', 'Detail']
  };

  for (const sheetName in schema) {
    let sheet = ss.getSheetByName(sheetName);
    
    // Jika sheet belum ada, buat baru
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Set Header
    const headers = schema[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Styling Header agar tebal dan ber-background
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#00a2b9')
         .setFontColor('white');
  }
  
  // Hapus "Sheet1" bawaan jika ada dan kosong
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
  
  Logger.log("Setup Database Selesai! Semua sheet dan header berhasil dibuat.");
}

// 4. FUNGSI MENYIMPAN DATA INISIASI BARU KE SHEET 'Data_PBJ'
function simpanInisiasiPBJ(dataForm) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Data_PBJ');
    
    if (!sheet) throw new Error("Sheet 'Data_PBJ' tidak ditemukan. Jalankan setupDatabase terlebih dahulu.");

    const rowBaru = [
      dataForm.idRencana,
      dataForm.judul,
      dataForm.jenisProses,
      dataForm.posAnggaran,
      'Inisiasi', // Status default awal
      '',         // Vendor kosong di awal
      ''          // SPK kosong di awal
    ];
    
    sheet.appendRow(rowBaru);
    
    return { success: true, message: 'Data Inisiasi berhasil disimpan ke Spreadsheet!' };
  } catch (error) {
    Logger.log(error);
    return { success: false, message: error.toString() };
  }
}

// =========================================================================
// --- FUNGSI CRUD DATA SETTING (LOKASI UPT & GI) ---
// =========================================================================

// 5. Simpan atau Update Data Setting Lokasi
function simpanSettingLokasi(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Lokasi_Pejabat');
    if (!sheet) throw new Error("Sheet 'Setting_Lokasi_Pejabat' tidak ditemukan.");
    
    const data = sheet.getDataRange().getValues();
    
    // Mode EDIT: Jika payload membawa oldNama
    if (payload.oldNama) {
      for (let i = 1; i < data.length; i++) {
        // Cocokkan Kategori (Col A) dan Nama Lama (Col B)
        if (data[i][0] === payload.kategori && data[i][1] === payload.oldNama) {
          sheet.getRange(i + 1, 2).setValue(payload.namaBaru);
          if (payload.alamatBaru !== undefined) {
             sheet.getRange(i + 1, 3).setValue(payload.alamatBaru);
          }
          return { success: true, message: 'Data ' + payload.kategori + ' berhasil diperbarui!' };
        }
      }
    }
    
    // Mode TAMBAH BARU
    // Skema: ['Kategori', 'Nama_Lokasi_Jabatan', 'Detail']
    sheet.appendRow([payload.kategori, payload.namaBaru, payload.alamatBaru || '']);
    
    return { success: true, message: 'Data ' + payload.kategori + ' berhasil ditambahkan!' };
    
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan: ' + error.toString() };
  }
}

// 6. Baca Semua Data Setting Lokasi (Untuk render ulang UI Frontend)
function getSemuaSettingLokasi() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Lokasi_Pejabat');
    if (!sheet) return { success: true, data: { upt: [], gi: [], jabatan: [] } };

    const data = sheet.getDataRange().getValues();
    let result = { upt: [], gi: [], jabatan: [] };
    
    // Mulai dari baris index 1 (mengabaikan header di index 0)
    for (let i = 1; i < data.length; i++) {
      let kategori = data[i][0];
      let item = { nama: data[i][1], alamat: data[i][2] };
      
      if (kategori === 'Lokasi UPT') result.upt.push(item);
      else if (kategori === 'Lokasi Gardu Induk') result.gi.push(item);
      else if (kategori === 'Jabatan & Email Direksi') result.jabatan.push(item); // Ini tambahannya
    }
    
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 7. Hapus Data Setting Lokasi
function hapusSettingLokasi(kategori, nama) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Lokasi_Pejabat');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // Cocokkan Kategori (Col A) dan Nama (Col B)
      if (data[i][0] === kategori && data[i][1] === nama) {
        sheet.deleteRow(i + 1); // +1 karena index array mulai dari 0, row sheet mulai dari 1
        return { success: true, message: 'Data berhasil dihapus!' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// =========================================================================
// --- FUNGSI CRUD SUMBER ANGGARAN (SKK & PRK) ---
// =========================================================================

// 1a. Simpan SKK Baru (Perbaikan Logika Base64 Upload Drive)
function simpanSKK(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Master_SKK');
    const folderId = '1MAq85OEpSCwDa8gVaQOU02JUSz37AlhX'; 
    let fileUrl = '';

    if (payload.fileData && payload.fileName) {
      try {
        let base64Data = payload.fileData.split(',')[1]; // Ekstrak murni data Base64
        let blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', payload.fileName);
        
        let folder = DriveApp.getFolderById(folderId);
        let file = folder.createFile(blob);
        fileUrl = file.getUrl(); // Dapatkan link Google Drive-nya
      } catch (driveError) {
        // JIKA GAGAL, TAMPILKAN ERROR ASLINYA KE SPREADSHEET AGAR KITA TAHU PENYEBABNYA
        fileUrl = "Error Upload: " + driveError.message;
      }
    }

    sheet.appendRow([payload.jenis, payload.no, payload.tanggal, fileUrl]);
    return { success: true, message: payload.jenis + ' berhasil disimpan!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 1b. Edit SKK
function editSKK(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Master_SKK');
    const data = sheet.getDataRange().getValues();
    const folderId = '1MAq85OEpSCwDa8gVaQOU02JUSz37AlhX'; 
    let newFileUrl = null;

    if (payload.fileData && payload.fileName) {
      let base64Data = payload.fileData.split(',')[1];
      let blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', payload.fileName);
      let folder = DriveApp.getFolderById(folderId);
      let file = folder.createFile(blob);
      newFileUrl = file.getUrl();
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.jenis && data[i][1] === payload.oldNo) {
        sheet.getRange(i + 1, 2).setValue(payload.newNo);
        sheet.getRange(i + 1, 3).setValue(payload.newTanggal);
        if (newFileUrl) {
            sheet.getRange(i + 1, 4).setValue(newFileUrl); // Timpa URL lama jika ada upload baru
        }
        return { success: true, message: 'Data SKK berhasil diupdate!' };
      }
    }
    return { success: false, message: 'Data lama tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 1c. Hapus SKK
function hapusSKK(noSKK) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Master_SKK');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === noSKK) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'SKK ' + noSKK + ' berhasil dihapus!' };
      }
    }
    return { success: false, message: 'Data SKK tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 2. Ambil Semua Data SKK (Perbaikan Format Tanggal)
function getSemuaSKK() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Master_SKK');
    if (!sheet) return { success: true, data: { skki: [], skko: [] } };
    
    const data = sheet.getDataRange().getValues();
    let result = { skki: [], skko: [] };
    
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (row[0] && row[1]) { 
        // Konversi format tanggal ke text agar tidak error saat dikirim ke HTML
        let tgl = row[2];
        if (tgl instanceof Date) {
            tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd MMM yyyy');
        }
        
        let item = { 
            jenis: row[0].toString(), 
            no: row[1].toString(), 
            tanggal: tgl.toString(), 
            file: (row[3] || '').toString() 
        };
        
        if (item.jenis === 'SKKI') result.skki.push(item);
        else if (item.jenis === 'SKKO') result.skko.push(item);
      }
    }
    
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 3. Simpan PRK Baru
function simpanPRK(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PRK');
    sheet.appendRow([payload.skkAsal, payload.noPRK, payload.nama, payload.pagu]);
    return { success: true, message: 'PRK ' + payload.noPRK + ' berhasil disimpan!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 4. Ambil Semua Data PRK
function getSemuaPRK() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PRK');
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    let result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (row[0] && row[1]) { 
         result.push({ 
             skkAsal: row[0].toString(), 
             no: row[1].toString(), 
             nama: row[2].toString(), 
             pagu: row[3].toString() 
         });
      }
    }
    
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Fungsi Pancingan untuk Memaksa Izin Upload File
function pancingIzinPenuh() {
   const folderId = '1MAq85OEpSCwDa8gVaQOU02JUSz37AlhX';
   var folder = DriveApp.getFolderById(folderId);
   var file = folder.createFile('tes_izin.txt', 'Berhasil dapat izin', MimeType.PLAIN_TEXT);
   file.setTrashed(true); // Langsung dihapus agar tidak nyampah
}

function simpanAtauEditPRK(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PRK');
    const data = sheet.getDataRange().getValues();

    if (payload.oldNoPRK) {
      // Mode Edit
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === payload.oldNoPRK) {
          sheet.getRange(i + 1, 1).setValue(payload.skkAsal);
          sheet.getRange(i + 1, 2).setValue(payload.noPRK);
          sheet.getRange(i + 1, 3).setValue(payload.nama);
          sheet.getRange(i + 1, 4).setValue(payload.pagu);
          return { success: true, message: 'PRK berhasil diupdate!' };
        }
      }
    } else {
      // Mode Tambah Baru
      sheet.appendRow([payload.skkAsal, payload.noPRK, payload.nama, payload.pagu]);
      return { success: true, message: 'PRK ' + payload.noPRK + ' berhasil disimpan!' };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function hapusPRK(noPRK) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PRK');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === noPRK) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'PRK berhasil dihapus!' };
      }
    }
    return { success: false, message: 'Data PRK tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function copyDataPRK(sumberSKK, tujuanSKK) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PRK');
    const data = sheet.getDataRange().getValues();
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sumberSKK) {
        sheet.appendRow([tujuanSKK, data[i][1] + "-COPY", data[i][2], data[i][3]]);
        count++;
      }
    }
    return { success: true, message: `Berhasil mencopy ${count} PRK!` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// REPLACE getSettingMaster di Code.gs agar support keterangan (kolom ke-3)
function getSettingMaster() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Master');
    if (!sheet) return { success: true, data: {} };
    
    const data = sheet.getDataRange().getValues();
    
    let result = { 
      syaratBayar: [], alatTeknik: [], alatK3: [], jenisKontrak: [], 
      klasifPenyedia: [], klasifBU: [], bidang: [], pendidikan: [], 
      jabatan: [], serkom: [], hariLibur: [], syaratUmum: [] // Tambahan syaratUmum
    };
    
    for (let i = 1; i < data.length; i++) {
      let kategori = data[i][0] ? data[i][0].toString().trim() : '';
      let value = data[i][1];
      let ket = data[i][2] ? data[i][2].toString().trim() : '';
      
      // Jika kosong, lewati
      if (!kategori || value === "" || value === null) continue;

      // PENGAMANAN TANGGAL 100% AMAN (Konversi Manual)
      if (value instanceof Date) {
          let yyyy = value.getFullYear();
          let mm = String(value.getMonth() + 1).padStart(2, '0');
          let dd = String(value.getDate()).padStart(2, '0');
          value = `${yyyy}-${mm}-${dd}`;
      } else {
          value = value.toString().trim();
      }
      
      if (kategori === 'Syarat Bayar') result.syaratBayar.push({val: value, ket: ket});
      else if (kategori === 'Alat Teknik') result.alatTeknik.push({val: value, ket: ket});
      else if (kategori === 'Alat K3') result.alatK3.push({val: value, ket: ket});
      else if (kategori === 'Jenis Kontrak') result.jenisKontrak.push({val: value, ket: ket});
      else if (kategori === 'Klasifikasi Penyedia') result.klasifPenyedia.push({val: value, ket: ket});
      else if (kategori === 'Klasifikasi BU') result.klasifBU.push({val: value, ket: ket});
      else if (kategori === 'Bidang Pekerjaan') result.bidang.push({val: value, ket: ket});
      else if (kategori === 'Pendidikan') result.pendidikan.push({val: value, ket: ket});
      else if (kategori === 'Jabatan CV') result.jabatan.push({val: value, ket: ket});
      else if (kategori === 'Serkom') result.serkom.push({val: value, ket: ket});
      else if (kategori === 'Hari Libur') result.hariLibur.push({val: value, ket: ket});
      else if (kategori === 'Syarat Umum') result.syaratUmum.push({val: value, ket: ket}); // Tambahan ini
    }
    
    return { success: true, data: result };
  } catch (error) { 
    return { success: false, message: error.toString() }; 
  }
}

function simpanAtauEditSettingMaster(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Master');
    const data = sheet.getDataRange().getValues();

    if (payload.oldVal) {
      for (let i = 1; i < data.length; i++) {
        let sheetVal = data[i][1];
        if (payload.kategori === 'Hari Libur' && sheetVal instanceof Date) {
            sheetVal = Utilities.formatDate(sheetVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }

        if (data[i][0] === payload.kategori && sheetVal === payload.oldVal) {
          sheet.getRange(i + 1, 2).setValue(payload.newVal);
          if (payload.newKet !== undefined) sheet.getRange(i + 1, 3).setValue(payload.newKet);
          return { success: true, message: 'Data diperbarui' };
        }
      }
    } else {
      sheet.appendRow([payload.kategori, payload.newVal, payload.newKet || '']);
      return { success: true };
    }
  } catch(e) { return { success: false, message: e.toString() }; }
}

function hapusDataSettingMaster(kategori, val) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Setting_Master');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let sheetVal = data[i][1];
      if (kategori === 'Hari Libur' && sheetVal instanceof Date) {
          sheetVal = Utilities.formatDate(sheetVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      if (data[i][0] === kategori && sheetVal === val) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Data tidak ditemukan" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// --- FUNGSI HAPUS DATA INISIASI PEKERJAAN ---
function deleteDataPekerjaan(idRencana) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data_PBJ');
    if (!sheet) throw new Error("Sheet 'Data_PBJ' tidak ditemukan.");
    
    const data = sheet.getDataRange().getValues();
    
    // Looping mencari baris dengan ID_Rencana yang cocok (kolom index 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idRencana) {
        sheet.deleteRow(i + 1); // +1 karena index sheet mulai dari 1, sedangkan array mulai 0
        return { status: 'success', message: 'Data inisiasi berhasil dihapus dari Spreadsheet!' };
      }
    }
    
    return { status: 'error', message: 'Data inisiasi dengan ID tersebut tidak ditemukan.' };
  } catch (error) {
    Logger.log(error);
    return { status: 'error', message: error.toString() };
  }
}
