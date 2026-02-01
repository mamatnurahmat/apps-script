// Global variables
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TIME_ZONE = 'GMT';
const SHEET_NAMES = {
  INVENTORY: 'Inventory',
  ALAT_MASUK: 'Alat Masuk',
  ALAT_KELUAR: 'Alat Keluar',
  SUPPLIER: 'Supplier',
  KATEGORI: 'Kategori Alat',
  USER: 'User',
  LAPORAN: 'Laporan'
};

function doGet(e) {
  if (e.parameter.page === 'admin') {
      return HtmlService.createHtmlOutputFromFile('LaporanMasuk')
        .setTitle('Laporan Masuk - Simantools')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  if (e.parameter.page === 'atem') {
      return HtmlService.createHtmlOutputFromFile('Atem')
        .setTitle('ATEM - Laporan Masuk')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  if (e.parameter.page === 'index') {
       return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistem Inventory')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default routing ke halaman Lapor (publik)
  return HtmlService.createHtmlOutputFromFile('Lapor')
      .setTitle('Lapor Kerusakan - Simantools')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Index') // Pastikan ini 'Index'
    .setTitle('Sistem Inventory')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Initialize the application
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sistem Inventory')
    .addItem('Buka Aplikasi', 'showSidebar')
    .addToUi();
}

// Initialize sheets with sample data
function initializeSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheets if they don't exist
  for (const sheetName of Object.values(SHEET_NAMES)) {
    if (!spreadsheet.getSheetByName(sheetName)) {
      spreadsheet.insertSheet(sheetName);
    }
  }
  
  // Setup headers
  setupHeaders();
  
  // Generate sample data (DINONAKTIFKAN)
  // generateSampleData(); 
  
  // Setup first admin if not exists
  setupFirstAdmin();
  
  return { success: true, message: "Sheets initialized successfully" };
}

// Setup headers for all sheets
function setupHeaders() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Inventory sheet headers
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  inventorySheet.getRange("A1:G1").setValues([["Kode Alat", "Nama Alat", "Kategori", "Stok", "Satuan", "Harga Beli", "Harga Jual"]]);
  inventorySheet.getRange("A1:G1").setFontWeight("bold");
  
  // Alat Masuk sheet headers
  const alatMasukSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
  alatMasukSheet.getRange("A1:F1").setValues([["ID Transaksi", "Tanggal", "Kode Alat", "Nama Alat", "Jumlah", "Supplier"]]);
  alatMasukSheet.getRange("A1:F1").setFontWeight("bold");
  
  // Alat Keluar sheet headers
  const alatKeluarSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  alatKeluarSheet.getRange("A1:E1").setValues([["ID Transaksi", "Tanggal", "Kode Alat", "Nama Alat", "Jumlah"]]);
  alatKeluarSheet.getRange("A1:E1").setFontWeight("bold");
  
  // Supplier sheet headers
  const supplierSheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  supplierSheet.getRange("A1:D1").setValues([["ID Supplier", "Nama Supplier", "Alamat", "Telepon"]]);
  supplierSheet.getRange("A1:D1").setFontWeight("bold");
  
  // Kategori sheet headers
  const kategoriSheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
  kategoriSheet.getRange("A1:B1").setValues([["ID Kategori", "Nama Kategori"]]);
  kategoriSheet.getRange("A1:B1").setFontWeight("bold");

  // Laporan sheet headers (Updated for new column)
  const laporanSheet = spreadsheet.getSheetByName(SHEET_NAMES.LAPORAN);
  // Pastikan header lengkap: A-M
  // A:ID, B:Alat, C:Kode, D:Keluhan, E:Pelapor, F:Ruangan, G:Tanggal, H:Status, I:Waktu Pengerjaan, J:Identifikasi, K:Tindakan, L:Rekomendasi, M:Catatan
  if (laporanSheet) {
      if (laporanSheet.getLastColumn() < 13) {
          laporanSheet.getRange("A1:M1").setValues([["ID Laporan", "Alat", "Kode", "Keluhan", "Pelapor", "Ruangan", "Tanggal Laporan", "Status", "Waktu Pengerjaan", "Identifikasi", "Tindakan", "Rekomendasi", "Catatan"]]);
          laporanSheet.getRange("A1:M1").setFontWeight("bold");
      }
  }

  
  // User sheet headers
  const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  userSheet.getRange("A1:D1").setValues([["Username", "Password", "Nama Lengkap", "Role"]]);
  userSheet.getRange("A1:D1").setFontWeight("bold");
}

// Setup first admin
function setupFirstAdmin() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  const userData = userSheet.getDataRange().getValues();
  
  // Check if admin already exists
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][3] === "Admin") {
      return { success: false, message: "Admin already exists" };
    }
  }
  
  // Create admin user
  const adminData = ["admin", "admin123", "Administrator", "Admin"];
  userSheet.appendRow(adminData);
  
  // Create manager user
  const managerData = ["manager", "manager123", "Manager", "Manajemen"];
  userSheet.appendRow(managerData);
  
  return { success: true, message: "Admin and Manager users created" };
}

function authenticateUser(username, password) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  const userData = userSheet.getDataRange().getValues();
  
  // Trim input dari user untuk menghilangkan spasi di awal/akhir
  const cleanUsername = username.trim();
  const cleanPassword = password.trim();
  
  for (let i = 1; i < userData.length; i++) {
    // Trim data dari sheet juga untuk memastikan tidak ada spasi tersembunyi
    const sheetUsername = String(userData[i][0]).trim();
    const sheetPassword = String(userData[i][1]).trim();
    
    if (sheetUsername === cleanUsername && sheetPassword === cleanPassword) {
      // PERBAIKAN: Ambil indeks yang benar untuk Nama Lengkap dan Role
      return {
        success: true,
        username: userData[i][0],
        fullName: userData[i][2], // BENAR: Indeks 2 untuk Nama Lengkap
        role: userData[i][3]      // BENAR: Indeks 3 untuk Role
      };
    }
  }
  
  return { success: false, message: "Invalid username or password" };
}

// Get all data for dropdowns
function getDropdownData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get suppliers
  const supplierSheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  const supplierData = supplierSheet.getDataRange().getValues();
  const suppliers = [];
  for (let i = 1; i < supplierData.length; i++) {
    suppliers.push(supplierData[i][1]); // Nama Supplier
  }
  
  // Get categories
  const kategoriSheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
  const kategoriData = kategoriSheet.getDataRange().getValues();
  const categories = [];
  for (let i = 1; i < kategoriData.length; i++) {
    categories.push(kategoriData[i][1]); // Nama Kategori
  }
  
  // Get inventory items
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  const inventoryData = inventorySheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < inventoryData.length; i++) {
    items.push({
      code: inventoryData[i][0], // Kode Alat
      name: inventoryData[i][1], // Nama Alat
      category: inventoryData[i][2], // Kategori
      stock: inventoryData[i][3], // Stok
      unit: inventoryData[i][4], // Satuan
      buyPrice: inventoryData[i][5], // Harga Beli
      sellPrice: inventoryData[i][6]  // Harga Jual
    });
  }
  
  return {
    suppliers: suppliers,
    categories: categories,
    items: items
  };
}

// CRUD functions for Inventory
function getInventoryData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  const data = sheet.getDataRange().getValues();
  
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i,
      code: data[i][0],
      name: data[i][1],
      category: data[i][2],
      stock: data[i][3],
      unit: data[i][4],
      buyPrice: data[i][5],
      sellPrice: data[i][6]
    });
  }
  
  return result;
}

function addInventoryItem(item) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
  // Generate new code if not provided
  if (!item.code) {
    const data = sheet.getDataRange().getValues();
    const lastCode = data.length > 1 ? data[data.length - 1][0] : "BRG000";
    const codeNumber = parseInt(lastCode.substring(3)) + 1;
    item.code = "BRG" + codeNumber.toString().padStart(3, '0');
  }
  
  // Add new row
  sheet.appendRow([
    item.code,
    item.name,
    item.category,
    item.stock || 0,
    item.unit,
    item.buyPrice || 0,
    item.sellPrice || 0
  ]);
  
  return { success: true, message: "Item added successfully", code: item.code };
}

function updateInventoryItem(id, item) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    const rowToUpdate = parseInt(id) + 1;
    
    Logger.log(`Memperbarui Inventory di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      // Gunakan setValues untuk MEMPERBARUI, bukan appendRow
      sheet.getRange(rowToUpdate, 1, 1, 7).setValues([[
        item.code,
        item.name,
        item.category,
        item.stock,
        item.unit,
        item.buyPrice,
        item.sellPrice
      ]]);
      return { success: true, message: "Item berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui inventory: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui item: " + e.message };
  }
}

function deleteInventoryItem(id) {
  Logger.log(`--- DEBUG: deleteInventoryItem DIMULAI dengan id=${id} ---`);
  Logger.log(`PERINGATAN: Fungsi deleteInventoryItem dipanggil. Ini menghapus dari MASTER ALAT.`);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
    
    const rowToDelete = parseInt(id) + 1;
    
    Logger.log(`DEBUG: Baris yang akan dihapus di sheet INVENTORY adalah: ${rowToDelete}`);
    
    if (rowToDelete > 1) {
      Logger.log(`DEBUG: Akan memanggil sheet.deleteRow(${rowToDelete}) pada sheet INVENTORY`);
      sheet.deleteRow(rowToDelete);
      Logger.log(`DEBUG: Berhasil menghapus baris ${rowToDelete} dari sheet INVENTORY`);
      return { success: true, message: "Item berhasil dihapus" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
  } catch (e) {
    Logger.log(`ERROR: Exception di deleteInventoryItem: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus item: " + e.message };
  }
}

function getAlatMasukData(startDate, endDate) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
  
  if (!sheet) {
    Logger.log('Sheet "Alat Masuk" tidak ditemukan!');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (!data || data.length <= 1) {
    Logger.log('Sheet "Alat Masuk" kosong atau hanya memiliki header.');
    return [];
  }
  
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { 
      // --- PERBAIKAN: Format tanggal yang dibaca dari sheet ---
      const itemDate = data[i][1]; // Ini adalah objek Date dari sheet
      const formattedDate = Utilities.formatDate(itemDate, timeZone, 'yyyy-MM-dd');
      
      if (startDate && endDate) {
        if (formattedDate >= startDate && formattedDate <= endDate) {
          result.push({
            id: i,
            transactionId: data[i][0] || '',
            date: formattedDate,
            itemCode: data[i][2] || '',
            itemName: data[i][3] || '',
            quantity: data[i][4] || 0,
            supplier: data[i][5] || ''
          });
        }
      } else {
        result.push({
          id: i,
          transactionId: data[i][0] || '',
          date: formattedDate,
          itemCode: data[i][2] || '',
          itemName: data[i][3] || '',
          quantity: data[i][4] || 0,
          supplier: data[i][5] || ''
        });
      }
    }
  }
  
  return result;
}

function addAlatMasuk(transaction) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
  
  // Generate new transaction ID if not provided
  if (!transaction.transactionId) {
    const data = sheet.getDataRange().getValues();
    const lastId = data.length > 1 ? data[data.length - 1][0] : "TRX000";
    const idNumber = parseInt(lastId.substring(3)) + 1;
    transaction.transactionId = "TRX" + idNumber.toString().padStart(3, '0');
  }
  
  // --- PERBAIKAN UTAMA: Format tanggal dengan benar ---
  const dateParts = transaction.date.split('-');
  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10) - 1; // Bulan JavaScript dimulai dari 0 (Januari=0)
  const day = parseInt(dateParts[2], 10);
  const dateObject = new Date(year, month, day);
  
  // Format tanggal agar sesuai dengan timezone spreadsheet
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const formattedDate = Utilities.formatDate(dateObject, timeZone, 'yyyy-MM-dd');

  // Add new row
  sheet.appendRow([
    transaction.transactionId,
    formattedDate, // Gunakan tanggal yang sudah benar
    transaction.itemCode,
    transaction.itemName,
    transaction.quantity,
    transaction.supplier
  ]);
  
  // Update inventory stock
  updateInventoryStock(transaction.itemCode, transaction.quantity, "in");
  
  return { success: true, message: "Transaction added successfully", transactionId: transaction.transactionId };
}

function updateAlatMasuk(id, transaction) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
    
    const rowToUpdate = parseInt(id) + 1;
    
    // Get old quantity to adjust stock
    const oldData = sheet.getRange(rowToUpdate, 1, 1, 6).getValues()[0];
    const oldQuantity = oldData[4];
    const itemCode = oldData[2];
    
    Logger.log(`Memperbarui Alat Masuk di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      // --- PERBAIKAN UTAMA: Format tanggal dengan benar ---
      const dateParts = transaction.date.split('-');
      const year = parseInt(dateParts[0], 10);
      const month = parseInt(dateParts[1], 10) - 1;
      const day = parseInt(dateParts[2], 10);
      const dateObject = new Date(year, month, day);
      
      const timeZone = spreadsheet.getSpreadsheetTimeZone();
      const formattedDate = Utilities.formatDate(dateObject, timeZone, 'yyyy-MM-dd');

      // Gunakan setValues untuk MEMPERBARUI
      sheet.getRange(rowToUpdate, 1, 1, 6).setValues([[
        transaction.transactionId,
        formattedDate, // Gunakan tanggal yang sudah benar
        transaction.itemCode,
        transaction.itemName,
        transaction.quantity,
        transaction.supplier
      ]]);
      
      // Adjust inventory stock
      const quantityDiff = transaction.quantity - oldQuantity;
      updateInventoryStock(itemCode, quantityDiff, "in");
      
      return { success: true, message: "Transaksi berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui alat masuk: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui transaksi: " + e.message };
  }
}

function deleteAlatMasuk(id) {
  Logger.log(`--- DEBUG: deleteAlatMasuk DIMULAI dengan id=${id} ---`);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
    
    const rowToDelete = parseInt(id) + 1;
    
    Logger.log(`DEBUG: Baris yang akan dihapus di sheet ALAT MASUK adalah: ${rowToDelete}`);
    
    if (rowToDelete <= 1) {
      Logger.log(`ERROR: ID baris tidak valid (${rowToDelete})`);
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
    
    // Ambil data dari baris yang akan dihapus SEBELUM dihapus
    const data = sheet.getRange(rowToDelete, 1, 1, 6).getValues()[0];
    const quantity = parseInt(data[4]) || 0;
    const itemCode = data[2];
    
    Logger.log(`DEBUG: Data dari baris yang dihapus: ${JSON.stringify(data)}`);
    Logger.log(`DEBUG: ItemCode=${itemCode}, Quantity=${quantity}`);
    
    // HAPUS baris transaksi dari sheet ALAT MASUK
    Logger.log(`DEBUG: Akan memanggil sheet.deleteRow(${rowToDelete}) pada sheet ALAT MASUK`);
    sheet.deleteRow(rowToDelete);
    Logger.log(`DEBUG: Berhasil menghapus baris ${rowToDelete} dari sheet ALAT MASUK`);

    // Perbarui stok di sheet INVENTORY dengan mengurangi jumlah alat yang masuk
    Logger.log(`DEBUG: Akan memanggil updateInventoryStock('${itemCode}', ${-quantity}, 'in')`);
    const stockUpdateResult = updateInventoryStock(itemCode, -quantity, "in");
    Logger.log(`DEBUG: Hasil dari updateInventoryStock: ${JSON.stringify(stockUpdateResult)}`);

    if (!stockUpdateResult.success) {
      Logger.log(`ERROR: Gagal memperbarui stok saat menghapus transaksi masuk: ${stockUpdateResult.message}`);
      return { success: true, message: "Transaksi dihapus, tetapi terjadi kesalahan saat memperbarui stok: " + stockUpdateResult.message };
    }
    
    Logger.log(`--- DEBUG: deleteAlatMasuk SELESAI ---`);
    return { success: true, message: "Transaksi berhasil dihapus dan stok dikembalikan." };

  } catch (e) {
    Logger.log(`ERROR: Exception di deleteAlatMasuk: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus transaksi: " + e.message };
  }
}

function getAlatKeluarData(startDate, endDate) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  
  if (!sheet) {
    Logger.log('Sheet "Alat Keluar" tidak ditemukan!');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (!data || data.length <= 1) {
    Logger.log('Sheet "Alat Keluar" kosong atau hanya memiliki header.');
    return [];
  }
  
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { 
      // --- PERBAIKAN: Format tanggal yang dibaca dari sheet ---
      const itemDate = data[i][1]; // Ini adalah objek Date dari sheet
      const formattedDate = Utilities.formatDate(itemDate, timeZone, 'yyyy-MM-dd');

      let shouldInclude = true;

      if (startDate && endDate) {
        if (formattedDate < startDate || formattedDate > endDate) {
          shouldInclude = false;
        }
      }
      
      if (shouldInclude) {
        result.push({
          id: i,
          transactionId: data[i][0] || '',
          date: formattedDate,
          itemCode: data[i][2] || '',
          itemName: data[i][3] || '',
          quantity: data[i][4] || 0
        });
      }
    }
  }
  
  return result;
}

function addAlatKeluar(transaction) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  
  // Generate new transaction ID if not provided
  if (!transaction.transactionId) {
    const data = sheet.getDataRange().getValues();
    const lastId = data.length > 1 ? data[data.length - 1][0] : "TRX000";
    const idNumber = parseInt(lastId.substring(3)) + 1;
    transaction.transactionId = "TRX" + idNumber.toString().padStart(3, '0');
  }
  
  // --- PERBAIKAN UTAMA: Format tanggal dengan benar ---
  const dateParts = transaction.date.split('-');
  const year = parseInt(dateParts[0], 10);
  const month = parseInt(dateParts[1], 10) - 1;
  const day = parseInt(dateParts[2], 10);
  const dateObject = new Date(year, month, day);
  
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const formattedDate = Utilities.formatDate(dateObject, timeZone, 'yyyy-MM-dd');

  // Check if stock is sufficient
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  const inventoryData = inventorySheet.getDataRange().getValues();
  
  let currentStock = 0;
  for (let i = 1; i < inventoryData.length; i++) {
    if (inventoryData[i][0] === transaction.itemCode) {
      currentStock = inventoryData[i][3];
      break;
    }
  }
  
  if (currentStock < transaction.quantity) {
    return { success: false, message: "Insufficient stock. Current stock: " + currentStock };
  }
  
  // Add new row
  sheet.appendRow([
    transaction.transactionId,
    formattedDate, // Gunakan tanggal yang sudah benar
    transaction.itemCode,
    transaction.itemName,
    transaction.quantity
  ]);
  
  // Update inventory stock
  updateInventoryStock(transaction.itemCode, transaction.quantity, "out");

  return { success: true, message: "Transaction added successfully", transactionId: transaction.transactionId };
}

function updateAlatKeluar(id, transaction) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
    
    const rowToUpdate = parseInt(id) + 1;
    
    // Get old quantity to adjust stock
    const oldData = sheet.getRange(rowToUpdate, 1, 1, 5).getValues()[0];
    const oldQuantity = oldData[4];
    const itemCode = oldData[2];
    
    Logger.log(`Memperbarui Alat Keluar di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      // --- PERBAIKAN UTAMA: Format tanggal dengan benar ---
      const dateParts = transaction.date.split('-');
      const year = parseInt(dateParts[0], 10);
      const month = parseInt(dateParts[1], 10) - 1;
      const day = parseInt(dateParts[2], 10);
      const dateObject = new Date(year, month, day);
      
      const timeZone = spreadsheet.getSpreadsheetTimeZone();
      const formattedDate = Utilities.formatDate(dateObject, timeZone, 'yyyy-MM-dd');

      // Gunakan setValues untuk MEMPERBARUI
      sheet.getRange(rowToUpdate, 1, 1, 5).setValues([[
        transaction.transactionId,
        formattedDate, // Gunakan tanggal yang sudah benar
        transaction.itemCode,
        transaction.itemName,
        transaction.quantity
      ]]);
      
      // Adjust inventory stock
      const quantityDiff = transaction.quantity - oldQuantity;
      updateInventoryStock(itemCode, quantityDiff, "out");

      return { success: true, message: "Transaksi berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui alat keluar: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui transaksi: " + e.message };
  }
}

function deleteAlatKeluar(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
    
    const rowToDelete = parseInt(id) + 1;

    if (rowToDelete <= 1) {
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
    
    // Kembalikan stok sebelum menghapus
    const data = sheet.getRange(rowToDelete, 1, 1, 5).getValues()[0];
    const quantity = parseInt(data[4]) || 0;
    const itemCode = data[2];
    
    Logger.log(`Menghapus Alat Keluar di baris: ${rowToDelete}, Item: ${itemCode}, Qty: ${quantity}`);
    
    // Hapus baris transaksi
    sheet.deleteRow(rowToDelete);

    // Perbarui stok (tambah kembali stok yang keluar)
    const stockUpdateResult = updateInventoryStock(itemCode, quantity, "out");

    if (!stockUpdateResult.success) {
      Logger.log(`Gagal memperbarui stok saat menghapus transaksi keluar: ${stockUpdateResult.message}`);
      return { success: true, message: "Transaksi dihapus, tetapi terjadi kesalahan saat memperbarui stok: " + stockUpdateResult.message };
    }
    
    return { success: true, message: "Transaksi berhasil dihapus dan stok dikembalikan." };

  } catch (e) {
    Logger.log(`Error menghapus alat keluar: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus transaksi: " + e.message };
  }
}

function updateInventoryStock(itemCode, quantity, type) {
  Logger.log(`--- DEBUG: updateInventoryStock DIMULAI ---`);
  Logger.log(`DEBUG: itemCode=${itemCode}, quantity=${quantity}, type=${type}`);
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);

  if (!sheet) {
    Logger.log('ERROR: Sheet "Inventory" tidak ditemukan saat mencoba memperbarui stok untuk ' + itemCode);
    return { success: false, message: "Sheet Inventory tidak ditemukan" };
  }
  
  const data = sheet.getDataRange().getValues();
  let itemFound = false;
  let finalStock = 0;
  let rowToUpdate = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === itemCode) {
      itemFound = true;
      rowToUpdate = i + 1; // Simpan nomor baris aktual
      const currentStock = parseInt(data[i][3]) || 0;
      let newStock;

      if (type === "in") {
        newStock = currentStock + quantity;
      } else {
        newStock = currentStock - quantity;
      }
      
      if (newStock < 0) {
        Logger.log(`WARNING: Stok untuk ${itemCode} akan menjadi negatif (${newStock}). Stok diset ke 0.`);
        newStock = 0;
      }

      finalStock = newStock;
      
      Logger.log(`DEBUG: Ditemukan item ${itemCode} di baris ${rowToUpdate}. Stok saat ini: ${currentStock}, stok baru: ${newStock}.`);
      
      // PERBAIKAN: HANYA MEMPERBARUI SATU SEL (KOLOM STOK), BUKAN MENGHAPUS BARIS
      // Ini adalah operasi yang aman dan tidak akan menghapus data alat.
      sheet.getRange(rowToUpdate, 4).setValue(newStock);
      Logger.log(`DEBUG: Berhasil memperbarui stok di baris ${rowToUpdate}, kolom 4.`);
      break;
    }
  }

  if (!itemFound) {
    Logger.log(`ERROR: Alat dengan kode '${itemCode}' tidak ditemukan di sheet Inventory. Stok tidak diperbarui.`);
    return { success: false, message: `Alat dengan kode ${itemCode} tidak ditemukan di master data.` };
  }

  Logger.log(`--- DEBUG: updateInventoryStock SELESAI ---`);
  return { success: true, message: `Stok untuk ${itemCode} berhasil diperbarui menjadi ${finalStock}.` };
}

// CRUD functions for Supplier
function getSupplierData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  const data = sheet.getDataRange().getValues();
  
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i,
      supplierId: data[i][0],
      name: data[i][1],
      address: data[i][2],
      phone: data[i][3]
    });
  }
  
  return result;
}

function addSupplier(supplier) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  
  // Generate new supplier ID if not provided
  if (!supplier.supplierId) {
    const data = sheet.getDataRange().getValues();
    const lastId = data.length > 1 ? data[data.length - 1][0] : "SUP000";
    const idNumber = parseInt(lastId.substring(3)) + 1;
    supplier.supplierId = "SUP" + idNumber.toString().padStart(3, '0');
  }
  
  // Add new row
  sheet.appendRow([
    supplier.supplierId,
    supplier.name,
    supplier.address,
    supplier.phone
  ]);
  
  return { success: true, message: "Supplier added successfully", supplierId: supplier.supplierId };
}

function updateSupplier(id, supplier) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
    
    const rowToUpdate = parseInt(id) + 1;
    
    Logger.log(`Memperbarui Supplier di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      sheet.getRange(rowToUpdate, 1, 1, 4).setValues([[
        supplier.supplierId,
        supplier.name,
        supplier.address,
        supplier.phone
      ]]);
      return { success: true, message: "Supplier berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui supplier: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui supplier: " + e.message };
  }
}

function deleteSupplier(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
    
    const rowToDelete = parseInt(id) + 1;
    
    Logger.log(`Menghapus Supplier di baris: ${rowToDelete}`);
    
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "Supplier berhasil dihapus" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
  } catch (e) {
    Logger.log(`Error menghapus supplier: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus supplier: " + e.message };
  }
}

// CRUD functions for Laporan
function getLaporanData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LAPORAN);
  
  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      // Perbaikan: Format tanggal dengan benar
      let formattedDate = '';
      if (data[i][6]) {
         formattedDate = Utilities.formatDate(new Date(data[i][6]), timeZone, 'yyyy-MM-dd HH:mm');
      }

      let formattedWaktuPengerjaan = '-';
      // Cek kolom "Waktu Pengerjaan" di indeks 8 (kolom I)
      if (data[i][8]) {
          formattedWaktuPengerjaan = Utilities.formatDate(new Date(data[i][8]), timeZone, 'yyyy-MM-dd HH:mm');
      }

      result.push({
        id: i,
        laporanId: data[i][0],
        alat: data[i][1],
        kode: data[i][2],
        keluhan: data[i][3],
        pelapor: data[i][4],
        ruangan: data[i][5],
        tanggal: formattedDate,
        status: data[i][7],
        waktuPengerjaan: formattedWaktuPengerjaan,
        identifikasi: data[i][9] || '',
        tindakan: data[i][10] || '',
        rekomendasi: data[i][11] || '',
        catatan: data[i][12] || ''
      });
    }
  }
  
  return result;
}

function addLaporan(laporan) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LAPORAN);
  
  // Generate new Laporan ID
  if (!laporan.laporanId) {
    const data = sheet.getDataRange().getValues();
    const lastId = data.length > 1 ? data[data.length - 1][0] : "RPT000";
    // Check if lastId is valid format
    let idNumber = 1;
    if (lastId && lastId.length > 3 && lastId.startsWith("RPT")) {
         idNumber = parseInt(lastId.substring(3)) + 1;
    }
    laporan.laporanId = "RPT" + idNumber.toString().padStart(3, '0');
  }
  
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const timestamp = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
  
  // Default status Waiting
  laporan.status = 'Waiting';

  // Add new row
  sheet.appendRow([
    laporan.laporanId,
    laporan.alat,
    laporan.kode,
    laporan.keluhan,
    laporan.pelapor,
    laporan.ruangan,
    timestamp,
    laporan.status
  ]);
  
  return { success: true, message: "Laporan berhasil dikirim", laporanId: laporan.laporanId };
}

function updateLaporanStatus(id, newStatus) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LAPORAN);
        const rowToUpdate = parseInt(id) + 1;

        if (rowToUpdate > 1) {
            sheet.getRange(rowToUpdate, 8).setValue(newStatus); // Column H is Status (8)
            
            // Jika status approval (On Process), catat waktu
            if (newStatus === 'On Process') {
                 const timeZone = spreadsheet.getSpreadsheetTimeZone();
                 const timestamp = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
                 sheet.getRange(rowToUpdate, 9).setValue(timestamp); // Column I is Waktu Pengerjaan (9)
                 return { success: true, message: "Laporan di-approve dan status menjadi On Process" };
            }
            
             return { success: true, message: "Status laporan berhasil diperbarui menjadi " + newStatus };
        } else {
             return { success: false, message: "ID laporan tidak valid" };
        }
    } catch (e) {
        return { success: false, message: "Gagal update status: " + e.message };
    }
}

function updateLaporanDetails(id, details) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.LAPORAN);
    const rowToUpdate = parseInt(id) + 1;

    if (rowToUpdate > 1) {
      // Update columns J, K, L, M (9, 10, 11, 12 in 0-indexed, but 10,11,12,13 in 1-indexed)
      // J=10, K=11, L=12, M=13
      sheet.getRange(rowToUpdate, 10).setValue(details.identifikasi);
      sheet.getRange(rowToUpdate, 11).setValue(details.tindakan);
      sheet.getRange(rowToUpdate, 12).setValue(details.rekomendasi);
      sheet.getRange(rowToUpdate, 13).setValue(details.catatan);

      // Otomatis ubah status jadi Done jika belum
      if(details.markAsDone) {
          sheet.getRange(rowToUpdate, 8).setValue('Done');
      }

      return { success: true, message: "Detail laporan berhasil disimpan" };
    } else {
      return { success: false, message: "ID laporan tidak valid" };
    }
  } catch(e) {
      return { success: false, message: "Gagal menyimpan detail: " + e.message };
  }
}


// CRUD functions for Kategori
function getKategoriData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
  const data = sheet.getDataRange().getValues();
  
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i,
      kategoriId: data[i][0],
      name: data[i][1]
    });
  }
  
  return result;
}

function addKategori(kategori) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
  
  // Generate new kategori ID if not provided
  if (!kategori.kategoriId) {
    const data = sheet.getDataRange().getValues();
    const lastId = data.length > 1 ? data[data.length - 1][0] : "KTG000";
    const idNumber = parseInt(lastId.substring(3)) + 1;
    kategori.kategoriId = "KTG" + idNumber.toString().padStart(3, '0');
  }
  
  // Add new row
  sheet.appendRow([
    kategori.kategoriId,
    kategori.name
  ]);
  
  return { success: true, message: "Kategori added successfully", kategoriId: kategori.kategoriId };
}

function updateKategori(id, kategori) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
    
    const rowToUpdate = parseInt(id) + 1;
    
    Logger.log(`Memperbarui Kategori di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      sheet.getRange(rowToUpdate, 1, 1, 2).setValues([[
        kategori.kategoriId,
        kategori.name
      ]]);
      return { success: true, message: "Kategori berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui kategori: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui kategori: " + e.message };
  }
}

function deleteKategori(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
    
    const rowToDelete = parseInt(id) + 1;
    
    Logger.log(`Menghapus Kategori di baris: ${rowToDelete}`);
    
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "Kategori berhasil dihapus" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
  } catch (e) {
    Logger.log(`Error menghapus kategori: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus kategori: " + e.message };
  }
}

function getUserData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  const data = sheet.getDataRange().getValues();
  
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i,
      username: data[i][0],
      password: data[i][1],
      fullName: data[i][2],
      role: data[i][3]
    });
  }
  
  return result;
}

function addUser(user) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  
  // Add new row
  sheet.appendRow([
    user.username,
    user.password,
    user.fullName,
    user.role
  ]);
  
  return { success: true, message: "User added successfully" };
}

function updateUser(id, user) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
    
    const rowToUpdate = parseInt(id) + 1;
    
    Logger.log(`Memperbarui User di baris: ${rowToUpdate}`);
    
    if (rowToUpdate > 1) {
      sheet.getRange(rowToUpdate, 1, 1, 4).setValues([[
        user.username,
        user.password,
        user.fullName,
        user.role
      ]]);
      return { success: true, message: "User berhasil diperbarui" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk diperbarui" };
    }
  } catch (e) {
    Logger.log(`Error memperbarui user: ${e.toString()}`);
    return { success: false, message: "Gagal memperbarui user: " + e.message };
  }
}

function deleteUser(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
    
    const rowToDelete = parseInt(id) + 1;
    
    Logger.log(`Menghapus User di baris: ${rowToDelete}`);
    
    if (rowToDelete > 1) {
      sheet.deleteRow(rowToDelete);
      return { success: true, message: "User berhasil dihapus" };
    } else {
      return { success: false, message: "ID baris tidak valid untuk dihapus" };
    }
  } catch (e) {
    Logger.log(`Error menghapus user: ${e.toString()}`);
    return { success: false, message: "Gagal menghapus user: " + e.message };
  }
}


function getSummaryData(filterType, startDate, endDate) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all data
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  const inventoryData = inventorySheet.getDataRange().getValues();
  
  const alatMasukSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
  const alatMasukData = alatMasukSheet.getDataRange().getValues();
  
  const alatKeluarSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  const alatKeluarData = alatKeluarSheet.getDataRange().getValues();
  
  const supplierSheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  const supplierData = supplierSheet.getDataRange().getValues();
  
  // Filter data based on date range
  let filteredAlatMasuk = [];
  let filteredAlatKeluar = [];
  
  if (startDate && endDate) {
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    for (let i = 1; i < alatMasukData.length; i++) {
      const date = new Date(alatMasukData[i][1]);
      if (date >= start && date <= end) {
        filteredAlatMasuk.push(alatMasukData[i]);
      }
    }
    
    for (let i = 1; i < alatKeluarData.length; i++) {
      const date = new Date(alatKeluarData[i][1]);
      if (date >= start && date <= end) {
        filteredAlatKeluar.push(alatKeluarData[i]);
      }
    }
  } else {
    filteredAlatMasuk = alatMasukData.slice(1);
    filteredAlatKeluar = alatKeluarData.slice(1);
  }
  
  // PERBAIKAN: Hitung total item dengan benar
  // Filter hanya item yang memiliki kode alat (tidak kosong)
  const validItems = inventoryData.filter((row, Index) => {
    // Lewati baris header
    if (Index === 0) return false;
    
    // Hanya hitung baris yang memiliki kode alat
    return row[0] && String(row[0]).trim() !== '';
  });
  
  // PERBAIKAN: Gunakan Set untuk memastikan tidak ada duplikasi
  const uniqueItems = new Set();
  validItems.forEach(item => {
    uniqueItems.add(item[0]); // Gunakan kode alat sebagai identifier unik
  });
  
  // PERBAIKAN: Hitung total jumlah alat masuk dan keluar
  let totalQuantityIn = 0;
  let totalQuantityOut = 0;
  
  // Hitung total jumlah alat masuk
  for (const item of filteredAlatMasuk) {
    totalQuantityIn += item[4]; // Kolom 5 adalah jumlah alat
  }
  
  // Hitung total jumlah alat keluar
  for (const item of filteredAlatKeluar) {
    totalQuantityOut += item[4]; // Kolom 5 adalah jumlah alat
  }
  
  // Calculate summaries
  const summary = {
    totalItems: uniqueItems.size, // PERBAIKAN: Gunakan ukuran Set untuk jumlah item unik
    totalStock: 0,
    totalValue: 0,
    totalIn: totalQuantityIn, // PERBAIKAN: Total jumlah alat masuk
    totalOut: totalQuantityOut, // PERBAIKAN: Total jumlah alat keluar
    totalSuppliers: supplierData.length - 1,
    topItems: [],
    topSuppliers: []
  };
  
  // Calculate total stock and value
  for (let i = 1; i < inventoryData.length; i++) {
    // PERBAIKAN: Pastikan item memiliki kode alat sebelum dihitung
    if (inventoryData[i][0] && String(inventoryData[i][0]).trim() !== '') {
      summary.totalStock += inventoryData[i][3];
      summary.totalValue += inventoryData[i][3] * inventoryData[i][6]; // Stock * Harga Jual
    }
  }
  
  // Get top items (most transactions)
  const itemTransactions = {};
  
  for (const item of filteredAlatMasuk) {
    const itemCode = item[2];
    if (!itemTransactions[itemCode]) {
      itemTransactions[itemCode] = {
        name: item[3],
        in: 0,
        out: 0,
        total: 0
      };
    }
    itemTransactions[itemCode].in += item[4];
    itemTransactions[itemCode].total += item[4];
  }
  
  for (const item of filteredAlatKeluar) {
    const itemCode = item[2];
    if (!itemTransactions[itemCode]) {
      itemTransactions[itemCode] = {
        name: item[3],
        in: 0,
        out: 0,
        total: 0
      };
    }
    itemTransactions[itemCode].out += item[4];
    itemTransactions[itemCode].total += item[4];
  }
  
  // Sort and get top 5 items
  const sortedItems = Object.entries(itemTransactions)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 5);
  
  for (const [code, data] of sortedItems) {
    summary.topItems.push({
      code: code,
      name: data.name,
      in: data.in,
      out: data.out,
      total: data.total
    });
  }
  
  // Get top suppliers (most transactions)
  const supplierTransactions = {};
  
  for (const item of filteredAlatMasuk) {
    const supplier = item[5];
    if (!supplierTransactions[supplier]) {
      supplierTransactions[supplier] = {
        transactions: 0,
        items: []
      };
    }
    supplierTransactions[supplier].transactions += 1;
    supplierTransactions[supplier].items.push({
      code: item[2],
      name: item[3],
      quantity: item[4]
    });
  }
  
  // Sort and get top 5 suppliers
  const sortedSuppliers = Object.entries(supplierTransactions)
    .sort((a, b) => b[1].transactions - a[1].transactions)
    .slice(0, 5);
  
  for (const [name, data] of sortedSuppliers) {
    summary.topSuppliers.push({
      name: name,
      transactions: data.transactions,
      items: data.items
    });
  }
  
  return summary;
}

function getChartData(chartType, filterType, startDate, endDate) {
  try {
    Logger.log(`getChartData called with chartType=${chartType}, filterType=${filterType}, startDate=${startDate}, endDate=${endDate}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get all data
    const alatMasukSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
    const alatMasukData = alatMasukSheet.getDataRange().getValues();
    
    const alatKeluarSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
    const alatKeluarData = alatKeluarSheet.getDataRange().getValues();
    
    // Filter data based on date range
    let filteredAlatMasuk = [];
    let filteredAlatKeluar = [];
    
    if (startDate && endDate) {
      const start = new Date(startDate);
      const end = new Date(endDate);
      
      for (let i = 1; i < alatMasukData.length; i++) {
        const date = new Date(alatMasukData[i][1]);
        if (date >= start && date <= end) {
          filteredAlatMasuk.push(alatMasukData[i]);
        }
      }
      
      for (let i = 1; i < alatKeluarData.length; i++) {
        const date = new Date(alatKeluarData[i][1]);
        if (date >= start && date <= end) {
          filteredAlatKeluar.push(alatKeluarData[i]);
        }
      }
    } else {
      filteredAlatMasuk = alatMasukData.slice(1);
      filteredAlatKeluar = alatKeluarData.slice(1);
    }
    
    // Process data based on chart type
    let chartData = {
      labels: [],
      datasets: []
    };
    
    if (chartType === 'transaction') {
      // Group by date, week, month, or year
      const groupedData = {};
      
      for (const item of filteredAlatMasuk) {
        const date = new Date(item[1]);
        let key;
        
        if (filterType === 'daily') {
          key = date.toISOString().split('T')[0];
        } else if (filterType === 'weekly') {
          const weekStart = new Date(date);
          weekStart.setDate(date.getDate() - date.getDay());
          key = weekStart.toISOString().split('T')[0];
        } else if (filterType === 'monthly') {
          key = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
        } else if (filterType === 'yearly') {
          key = `${date.getFullYear()}`;
        }
        
        if (!groupedData[key]) {
          groupedData[key] = {
            in: 0,
            out: 0
          };
        }
        
        groupedData[key].in += item[4];
      }
      
      for (const item of filteredAlatKeluar) {
        const date = new Date(item[1]);
        let key;
        
        if (filterType === 'daily') {
          key = date.toISOString().split('T')[0];
        } else if (filterType === 'weekly') {
          const weekStart = new Date(date);
          weekStart.setDate(date.getDate() - date.getDay());
          key = weekStart.toISOString().split('T')[0];
        } else if (filterType === 'monthly') {
          key = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
        } else if (filterType === 'yearly') {
          key = `${date.getFullYear()}`;
        }
        
        if (!groupedData[key]) {
          groupedData[key] = {
            in: 0,
            out: 0
          };
        }
        
        groupedData[key].out += item[4];
      }
      
      // Sort keys
      const sortedKeys = Object.keys(groupedData).sort();
      
      // PERBAIKAN: Format label dengan benar
      chartData.labels = sortedKeys;
      chartData.datasets = [
        {
          label: 'Alat Masuk',
          data: sortedKeys.map(key => groupedData[key].in),
          backgroundColor: 'rgba(75, 192, 192, 0.2)',
          borderColor: 'rgba(75, 192, 192, 1)',
          borderWidth: 1
        },
        {
          label: 'Alat Keluar',
          data: sortedKeys.map(key => groupedData[key].out),
          backgroundColor: 'rgba(255, 99, 132, 0.2)',
          borderColor: 'rgba(255, 99, 132, 1)',
          borderWidth: 1
        }
      ];
    } else if (chartType === 'category') {
      // Ambil semua kategori dari sheet kategori
      const kategoriSheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
      const kategoriData = kategoriSheet.getDataRange().getValues();
      
      // Buat objek untuk menyimpan data per kategori
      const groupedData = {};
      
      // Inisialisasi semua kategori dengan nilai 0
      for (let i = 1; i < kategoriData.length; i++) {
        const kategori = kategoriData[i][1]; // Nama kategori
        groupedData[kategori] = {
          in: 0,
          out: 0
        };
      }
      
      // Hitung alat masuk per kategori
      for (const item of filteredAlatMasuk) {
        const itemCode = item[2];
        
        // Get category from inventory
        const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
        const inventoryData = inventorySheet.getDataRange().getValues();
        
        for (let i = 1; i < inventoryData.length; i++) {
          if (inventoryData[i][0] === itemCode) {
            const category = inventoryData[i][2];
            
            if (groupedData[category]) {
              groupedData[category].in += item[4];
            }
            break;
          }
        }
      }
      
      // Hitung alat keluar per kategori
      for (const item of filteredAlatKeluar) {
        const itemCode = item[2];
        
        // Get category from inventory
        const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
        const inventoryData = inventorySheet.getDataRange().getValues();
        
        for (let i = 1; i < inventoryData.length; i++) {
          if (inventoryData[i][0] === itemCode) {
            const category = inventoryData[i][2];
            
            if (groupedData[category]) {
              groupedData[category].out += item[4];
            }
            break;
          }
        }
      }
      
      // Sort keys based on filter type
      let sortedKeys = Object.keys(groupedData);
      
      if (filterType === 'top') {
        // Sort by total transaction (descending)
        sortedKeys = sortedKeys.sort((a, b) => {
          const totalA = groupedData[a].in + groupedData[a].out;
          const totalB = groupedData[b].in + groupedData[b].out;
          return totalB - totalA;
        }).slice(0, 10); // Top 10
      } else if (filterType === 'low') {
        // Sort by total transaction (ascending)
        sortedKeys = sortedKeys.sort((a, b) => {
          const totalA = groupedData[a].in + groupedData[a].out;
          const totalB = groupedData[b].in + groupedData[b].out;
          return totalA - totalB;
        }).slice(0, 10); // Bottom 10
      } else {
        // Default: sort alphabetically
        sortedKeys = sortedKeys.sort();
      }
      
      // Prepare chart data
      chartData.labels = sortedKeys;
      chartData.datasets = [
        {
          label: 'Alat Masuk',
          data: sortedKeys.map(key => groupedData[key].in),
          backgroundColor: 'rgba(75, 192, 192, 0.2)',
          borderColor: 'rgba(75, 192, 192, 1)',
          borderWidth: 1
        },
        {
          label: 'Alat Keluar',
          data: sortedKeys.map(key => groupedData[key].out),
          backgroundColor: 'rgba(255, 99, 132, 0.2)',
          borderColor: 'rgba(255, 99, 132, 1)',
          borderWidth: 1
        }
      ];
    } else if (chartType === 'supplier') {
      // Group by supplier
      const groupedData = {};
      
      for (const item of filteredAlatMasuk) {
        const supplier = item[5];
        
        if (!groupedData[supplier]) {
          groupedData[supplier] = 0;
        }
        
        groupedData[supplier] += item[4];
      }
      
      // Sort keys
      const sortedKeys = Object.keys(groupedData)
        .sort((a, b) => groupedData[b] - groupedData[a])
        .slice(0, 10); // Top 10 suppliers
      
      // Prepare chart data
      chartData.labels = sortedKeys;
      chartData.datasets = [
        {
          label: 'Jumlah Alat',
          data: sortedKeys.map(key => groupedData[key]),
          backgroundColor: [
            'rgba(255, 99, 132, 0.2)',
            'rgba(54, 162, 235, 0.2)',
            'rgba(255, 206, 86, 0.2)',
            'rgba(75, 192, 192, 0.2)',
            'rgba(153, 102, 255, 0.2)',
            'rgba(255, 159, 64, 0.2)',
            'rgba(199, 199, 199, 0.2)',
            'rgba(83, 102, 255, 0.2)',
            'rgba(255, 99, 255, 0.2)',
            'rgba(99, 255, 132, 0.2)'
          ],
          borderColor: [
            'rgba(255, 99, 132, 1)',
            'rgba(54, 162, 235, 1)',
            'rgba(255, 206, 86, 1)',
            'rgba(75, 192, 192, 1)',
            'rgba(153, 102, 255, 1)',
            'rgba(255, 159, 64, 1)',
            'rgba(199, 199, 199, 1)',
            'rgba(83, 102, 255, 1)',
            'rgba(255, 99, 255, 1)',
            'rgba(99, 255, 132, 1)'
          ],
          borderWidth: 1
        }
      ];
    }
    
    Logger.log(`Returning chart data with ${chartData.labels.length} labels`);
    return chartData;
  } catch (error) {
    Logger.log(`Error in getChartData: ${error.toString()}`);
    throw error;
  }
}

// Search functions
function searchInventory(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  // PERBAIKAN 2.0: Pastikan sheet ada sebelum melanjutkan
  if (!sheet) {
    Logger.log('Sheet "Inventory" tidak ditemukan dalam fungsi searchInventory!');
    return []; // Kembalikan array kosong
  }
  
  const data = sheet.getDataRange().getValues();
  
  const results = [];
  const lowerQuery = query.toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]).toLowerCase().includes(lowerQuery) || // Kode Alat
      String(data[i][1]).toLowerCase().includes(lowerQuery) || // Nama Alat
      String(data[i][2]).toLowerCase().includes(lowerQuery)    // Kategori
    ) {
      results.push({
        id: i,
        code: data[i][0],
        name: data[i][1],
        category: data[i][2],
        stock: data[i][3],
        unit: data[i][4],
        buyPrice: data[i][5],
        sellPrice: data[i][6]
      });
    }
  }
  
  return results;
}

function searchAlatMasuk(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_MASUK);
  // PERBAIKAN 2.0: Pastikan sheet ada sebelum melanjutkan
  if (!sheet) {
    Logger.log('Sheet "Alat Masuk" tidak ditemukan dalam fungsi searchAlatMasuk!');
    return []; // Kembalikan array kosong
  }
  
  const data = sheet.getDataRange().getValues();
  
  const results = [];
  const lowerQuery = query.toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]).toLowerCase().includes(lowerQuery) || // ID Transaksi
      String(data[i][2]).toLowerCase().includes(lowerQuery) || // Kode Alat
      String(data[i][3]).toLowerCase().includes(lowerQuery) || // Nama Alat
      String(data[i][5]).toLowerCase().includes(lowerQuery)    // Supplier
    ) {
      const itemDate = new Date(data[i][1]);
      const formattedDate = itemDate.getFullYear() + '-' + 
                            ('0' + (itemDate.getMonth() + 1)).slice(-2) + '-' + 
                            ('0' + itemDate.getDate()).slice(-2);

      results.push({
        id: i,
        transactionId: data[i][0],
        date: formattedDate,
        itemCode: data[i][2],
        itemName: data[i][3],
        quantity: data[i][4],
        supplier: data[i][5]
      });
    }
  }
  
  return results;
}

function searchAlatKeluar(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  // PERBAIKAN 2.0: Pastikan sheet ada sebelum melanjutkan
  if (!sheet) {
    Logger.log('Sheet "Alat Keluar" tidak ditemukan dalam fungsi searchAlatKeluar!');
    return []; // Kembalikan array kosong
  }
  
  const data = sheet.getDataRange().getValues();
  
  const results = [];
  const lowerQuery = query.toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]).toLowerCase().includes(lowerQuery) || // ID Transaksi
      String(data[i][2]).toLowerCase().includes(lowerQuery) || // Kode Alat
      String(data[i][3]).toLowerCase().includes(lowerQuery)    // Nama Alat
    ) {
      const itemDate = new Date(data[i][1]);
      const formattedDate = itemDate.getFullYear() + '-' + 
                            ('0' + (itemDate.getMonth() + 1)).slice(-2) + '-' + 
                            ('0' + itemDate.getDate()).slice(-2);

      results.push({
        id: i,
        transactionId: data[i][0],
        date: formattedDate,
        itemCode: data[i][2],
        itemName: data[i][3],
        quantity: data[i][4]
      });
    }
  }
  
  return results;
}

function searchSupplier(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.SUPPLIER);
  if (!sheet) { return []; }
  const data = sheet.getDataRange().getValues();
  const results = [];
  const lowerQuery = query.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().includes(lowerQuery) || String(data[i][1]).toLowerCase().includes(lowerQuery) || String(data[i][2]).toLowerCase().includes(lowerQuery) || String(data[i][3]).toLowerCase().includes(lowerQuery)) {
      results.push({ id: i, supplierId: data[i][0], name: data[i][1], address: data[i][2], phone: data[i][3] });
    }
  }
  return results;
}

function searchKategori(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.KATEGORI);
  if (!sheet) { return []; }
  const data = sheet.getDataRange().getValues();
  const results = [];
  const lowerQuery = query.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().includes(lowerQuery) || String(data[i][1]).toLowerCase().includes(lowerQuery)) {
      results.push({ id: i, kategoriId: data[i][0], name: data[i][1] });
    }
  }
  return results;
}

function searchUser(query) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USER);
  if (!sheet) { return []; }
  const data = sheet.getDataRange().getValues();
  const results = [];
  const lowerQuery = query.toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().includes(lowerQuery) || String(data[i][2]).toLowerCase().includes(lowerQuery) || String(data[i][3]).toLowerCase().includes(lowerQuery)) {
      results.push({ id: i, username: data[i][0], password: data[i][1], fullName: data[i][2], role: data[i][3] });
    }
  }
  return results;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInventorySummary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  
  if (!sheet) {
    Logger.log('Sheet "Inventory" tidak ditemukan!');
    return { 
      totalNilaiBeli: 0, 
      totalNilaiJual: 0,
      totalItems: 0  // PERBAIKAN: Tambahkan total item
    };
  }
  
  const data = sheet.getDataRange().getValues();
  
  // PERBAIKAN: Filter item yang valid
  const validItems = data.filter((row, Index) => {
    if (Index === 0) return false; // Lewati header
    return row[0] && String(row[0]).trim() !== ''; // Hanya item dengan kode alat
  });
  
  // PERBAIKAN: Gunakan Set untuk memastikan tidak ada duplikasi
  const uniqueItems = new Set();
  validItems.forEach(item => {
    uniqueItems.add(item[0]); // Gunakan kode alat sebagai identifier unik
  });
  
  // Indeks kolom (A=0, B=1, C=2, ...)
  const stockIndex = 3;    // Kolom Stok
  const buyPriceIndex = 5; // Kolom Harga Beli
  const sellPriceIndex = 6; // Kolom Harga Jual
  
  let totalNilaiBeli = 0;
  let totalNilaiJual = 0;
  
  // Mulai dari i = 1 untuk melewati baris header
  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Pastikan item memiliki kode alat sebelum dihitung
    if (data[i][0] && String(data[i][0]).trim() !== '') {
      const stock = data[i][stockIndex];
      const buyPrice = data[i][buyPriceIndex];
      const sellPrice = data[i][sellPriceIndex];
      
      // Pastikan nilainya adalah angka
      if (typeof stock === 'number' && typeof buyPrice === 'number') {
        totalNilaiBeli += stock * buyPrice;
      }
      if (typeof stock === 'number' && typeof sellPrice === 'number') {
        totalNilaiJual += stock * sellPrice;
      }
    }
  }
  
  return {
    totalNilaiBeli: totalNilaiBeli,
    totalNilaiJual: totalNilaiJual,
    totalItems: uniqueItems.size  // PERBAIKAN: Kembalikan jumlah item unik
  };
}
function exportToExcel(sheetName, data, pageTitle) {
  try {
    // Logger untuk debugging
    Logger.log(`Memulai export ke Excel. Sheet: ${sheetName}, Judul: ${pageTitle}, Jumlah Data: ${data.length}`);
    
    // Buat spreadsheet baru dengan nama unik
    const fileName = `Export ${pageTitle} - ${new Date().toLocaleDateString('id-ID')}`;
    const newSpreadsheet = SpreadsheetApp.create(fileName);
    const newSheet = newSpreadsheet.getActiveSheet();
    newSheet.setName(sheetName);
    
    // Tentukan header berdasarkan sheetName
    let headers = [];
    switch(sheetName) {
      case 'Inventory':
        headers = ["Kode Alat", "Nama Alat", "Kategori", "Stok", "Satuan", "Harga Beli", "Harga Jual"];
        break;
      case 'Alat Masuk':
        headers = ["ID Transaksi", "Tanggal", "Kode Alat", "Nama Alat", "Jumlah", "Supplier"];
        break;
      case 'Alat Keluar':
        headers = ["ID Transaksi", "Tanggal", "Kode Alat", "Nama Alat", "Jumlah"];
        break;
      case 'Supplier':
        headers = ["ID Supplier", "Nama Supplier", "Alamat", "Telepon"];
        break;
      case 'Kategori Alat':
        headers = ["ID Kategori", "Nama Kategori"];
        break;
      case 'User':
        headers = ["Username", "Password", "Nama Lengkap", "Role"];
        break;
      default:
        // Jika sheetName tidak dikenali, gunakan kunci dari objek data pertama
        if (data.length > 0) {
          headers = Object.keys(data[0]);
        } else {
          headers = ["Data"];
        }
    }
    
    // Tambahkan baris header
    newSheet.appendRow(headers);
    
    // Format header (bold dan latar belakang abu-abu)
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold").setBackground("#f0f0f0");
    
    // Tambahkan baris data
    if (data && data.length > 0) {
      data.forEach(item => {
        let row = [];
        switch(sheetName) {
          case 'Inventory':
            row = [item.code, item.name, item.category, item.stock, item.unit, item.buyPrice, item.sellPrice];
            break;
          case 'Alat Masuk':
            row = [item.transactionId, item.date, item.itemCode, item.itemName, item.quantity, item.supplier];
            break;
          case 'Alat Keluar':
            row = [item.transactionId, item.date, item.itemCode, item.itemName, item.quantity];
            break;
          case 'Supplier':
            row = [item.supplierId, item.name, item.address, item.phone];
            break;
          case 'Kategori Alat':
            row = [item.kategoriId, item.name];
            break;
          case 'User':
            row = [item.username, item.password, item.fullName, item.role];
            break;
          default:
            // Default: ambil semua nilai dari objek sesuai urutan header
            headers.forEach(header => {
              row.push(item[header] || '');
            });
        }
        newSheet.appendRow(row);
      });
      
      // Auto-resize lebar kolom agar sesuai dengan isinya
      for (let i = 1; i <= headers.length; i++) {
        newSheet.autoResizeColumn(i);
      }
      
      // PERBAIKAN KRUSIAL: Format tanggal untuk kolom tanggal
      if (sheetName === 'Alat Masuk' || sheetName === 'Alat Keluar') {
        const dateColumnIndex = headers.IndexOf("Tanggal") + 1; // Dapatkan indeks kolom Tanggal (dimulai dari 1)
        if (dateColumnIndex > 0) {
          const lastRow = newSheet.getLastRow();
          if (lastRow > 1) {
            // PERBAIKAN: Range harus dimulai dari baris ke-2 (data pertama), bukan baris terakhir
            const dateRange = newSheet.getRange(2, dateColumnIndex, lastRow - 1, 1);
            dateRange.setNumberFormat("dd/mm/yyyy");
          }
        }
      }
    }
    
    // Dapatkan URL file yang baru dibuat
    const url = newSpreadsheet.getUrl();
    
    Logger.log(`Berhasil membuat spreadsheet. URL: ${url}`);
    
    return {
      success: true,
      message: "File Excel berhasil dibuat",
      url: url
    };
  } catch (e) {
    // Logger yang lebih detail untuk debugging
    Logger.log(`Error saat mengekspor ke Excel: ${e.toString()}`);
    Logger.log(`Stack Trace: ${e.stack}`);
    
    return {
      success: false,
      message: "Gagal membuat file Excel: " + e.message
    };
  }
}

function getRingkasanKeuntungan() {
  return {
    totalKeuntungan: calculateTotalKeuntunganFromSheet()
  };
}

function calculateTotalKeuntunganFromSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  const alatKeluarSheet = spreadsheet.getSheetByName(SHEET_NAMES.ALAT_KELUAR);
  if (!alatKeluarSheet) return 0;
  const alatKeluarData = alatKeluarSheet.getDataRange().getValues();
  
  const inventorySheet = spreadsheet.getSheetByName(SHEET_NAMES.INVENTORY);
  if (!inventorySheet) return 0;
  const inventoryData = inventorySheet.getDataRange().getValues();
  
  // Buat peta (map) untuk pencarian harga beli dan jual dengan cepat
  const inventoryMap = new Map();
  for (let i = 1; i < inventoryData.length; i++) {
    // Map Kode Alat -> { buyPrice: ..., sellPrice: ... }
    inventoryMap.set(inventoryData[i][0], {
      buyPrice: inventoryData[i][5], // Indeks 5 untuk Harga Beli
      sellPrice: inventoryData[i][6]  // Indeks 6 untuk Harga Jual
    }); 
  }
  
  let totalKeuntungan = 0;
  
  // Lewati baris header (i dimulai dari 1)
  for (let i = 1; i < alatKeluarData.length; i++) {
    const itemCode = alatKeluarData[i][2]; // Kode Alat
    const quantity = alatKeluarData[i][4]; // Jumlah
    
    const prices = inventoryMap.get(itemCode);
    
    if (prices) {
      // Hitung keuntungan per item dan kalikan dengan jumlah
      const profitPerItem = prices.sellPrice - prices.buyPrice;
      totalKeuntungan += quantity * profitPerItem;
    }
  }
  
  return totalKeuntungan;
}

// Fungsi pembantu untuk format tanggal untuk spreadsheet
function formatDateForSpreadsheet(date) {
  if (!date) return '';
  
  // Jika date adalah string, konversi ke objek Date
  if (typeof date === 'string') {
    date = new Date(date);
  }
  
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
