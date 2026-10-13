 * =========================================================================
 * RELOVE MARKETPLACE - GOOGLE APPS SCRIPT (Code.gs)
 * Platform Jual Beli Preloved: Praktis, Terpercaya, dan Berkelanjutan
 * =========================================================================
 */

// 1. ENTRY POINT WEB APP
function doGet(e) {
  // Pastikan sheet database sudah siap
  initDatabaseIfNotExists();

  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('RELOVE - Preloved Marketplace & Escrow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. MENU KHUSUS PADA SPREADSHEET
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🌸 RELOVE Admin')
    .addItem('🚀 Setup Ulang Database Sheets', 'setupDatabaseSheets')
    .addSeparator()
    .addItem('💰 Hitung Rekap Escrow & GMV', 'hitungRekapKeuangan')
    .addItem('🌿 Rekap Dampak Lingkungan (SDG 12)', 'hitungDampakLingkungan')
    .addToUi();
}

// 3. INISIALISASI SHEET DATABASE
function initDatabaseIfNotExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Products')) {
    setupDatabaseSheets();
  }
}

function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. Sheet Products
  let sheetProducts = ss.getSheetByName('Products');
  if (!sheetProducts) sheetProducts = ss.insertSheet('Products');
  sheetProducts.clear();
  const headersProducts = [
    'ID Produk', 'Judul Barang', 'Kategori', 'Harga Jual (Rp)', 'Harga Baru (Rp)',
    'Kondisi Fisik', 'Skor Kondisi (%)', 'Kota Lokasi', 'Penjual (Seller)',
    'Bahan / Material', 'Warna', 'Brand / Merek', 'Catatan Minus / Flaws',
    'Status Verifikasi', 'Promoted', 'Hemat CO2 (kg)', 'Hemat Air (Liter)',
    'URL Foto', 'Tanggal Upload'
  ];
  sheetProducts.appendRow(headersProducts);
  formatHeaderRow(sheetProducts, '#1E3F20', '#FFFFFF');

  sheetProducts.appendRow([
    'prod-1', 'Oversized Hoodie Washed Grey', 'Pakaian', 75000, 289000,
    '90% Like New', 90, 'Malang', 'sarahh_preloved',
    '100% Cotton Fleece', 'Washed Grey', 'Cottonology', 'Warna pekat, tidak ada sobek, rib kencang',
    'VERIFIED', true, 4.8, 2600,
    'https://images.unsplash.com/photo-1556905055-8f358a7a47b2?w=800', new Date()
  ]);
  sheetProducts.appendRow([
    'prod-2', 'Sneakers Adidas Spezial Cream Blue', 'Alas Kaki', 120000, 1199000,
    '85% Good Condition', 85, 'Surabaya', 'dimas_sneakers',
    'Suede & Rubber', 'Cream Vintage', 'Adidas Originals', 'Suede ada bekas pemakaian wajar, sol tebal',
    'VERIFIED', true, 9.5, 4100,
    'https://images.unsplash.com/photo-1595950653106-6c9ebd614d3a?w=800', new Date()
  ]);
  sheetProducts.appendRow([
    'prod-3', 'Tas Shoulder Bag Vintage Leather Tan', 'Tas & Ransel', 85000, 320000,
    '90% Like New', 90, 'Bandung', 'rin_archive',
    'Vegan PU Leather', 'Caramel Tan', 'Vintage Collection', 'Resleting lancar mulus, furing bersih',
    'VERIFIED', false, 3.2, 1800,
    'https://images.unsplash.com/photo-1590874103328-eac38a683ce7?w=800', new Date()
  ]);
  sheetProducts.appendRow([
    'prod-4', 'Vintage Corduroy Jacket Forest Green', 'Pakaian', 95000, 450000,
    '90% Like New', 90, 'Malang', 'sarahh_preloved',
    'Wide Wale Corduroy', 'Forest Green', 'Uniqlo Vintage', 'Kancing lengkap, warna pekat tidak berbulu',
    'VERIFIED', true, 6.8, 3400,
    'https://images.unsplash.com/photo-1544441893-675973e31985?w=800', new Date()
  ]);

  // B. Sheet Orders
  let sheetOrders = ss.getSheetByName('Orders');
  if (!sheetOrders) sheetOrders = ss.insertSheet('Orders');
  sheetOrders.clear();
  const headersOrders = [
    'ID Order', 'Nomor Pesanan', 'ID Produk', 'Judul Produk', 'Nama Pembeli',
    'Penjual', 'Alamat Pengiriman', 'Ekspedisi Kurir', 'Harga Produk (Rp)',
    'Ongkos Kirim (Rp)', 'Biaya Escrow RELOVE (Rp)', 'Total Bayar (Rp)',
    'Metode Pembayaran', 'Status Rekening Bersama', 'Status Pengiriman',
    'Nomor Resi', 'Waktu Transaksi'
  ];
  sheetOrders.appendRow(headersOrders);
  formatHeaderRow(sheetOrders, '#1E3F20', '#FFFFFF');

  sheetOrders.appendRow([
    'ord-101', 'RLV-20261012-9842', 'prod-1', 'Oversized Hoodie Washed Grey',
    'Fahmi Abdillah', 'sarahh_preloved', 'Jl. Cempaka No. 12, Malang 65145',
    'J&T Express', 75000, 12000, 2500, 89500,
    'DANA', 'IN_ESCROW', 'SHIPPED', 'JT90281140029', new Date()
  ]);

  // C. Sheet Verifikasi_Kurasi
  let sheetVerif = ss.getSheetByName('Verifikasi_Kurasi');
  if (!sheetVerif) sheetVerif = ss.insertSheet('Verifikasi_Kurasi');
  sheetVerif.clear();
  const headersVerif = [
    'ID Log', 'ID Produk', 'Judul Produk', 'Status Kurasi',
    'Kurator Admin', 'Catatan Evaluasi / Minus', 'Akurasi AI (%)', 'Waktu Keputusan'
  ];
  sheetVerif.appendRow(headersVerif);
  formatHeaderRow(sheetVerif, '#1E3F20', '#FFFFFF');

  sheetVerif.appendRow([
    'vrf-1', 'prod-1', 'Oversized Hoodie Washed Grey', 'APPROVED',
    'Admin Fahmi', 'Foto kerah dan rib terbukti mulus tanpa noda', 94.8, new Date()
  ]);

  // D. Sheet Dashboard_Ringkasan
  let sheetDash = ss.getSheetByName('Dashboard_Ringkasan');
  if (!sheetDash) sheetDash = ss.insertSheet('Dashboard_Ringkasan');
  sheetDash.clear();
  sheetDash.appendRow(['METRIK UTAMA RELOVE 2026', 'NILAI / JUMLAH', 'SATUAN / KETERANGAN']);
  formatHeaderRow(sheetDash, '#1E3F20', '#FFFFFF');

  sheetDash.appendRow(['Total Barang Aktif di Katalog', '=COUNTA(Products!A2:A)', 'Barang']);
  sheetDash.appendRow(['Total Transaksi Selesai & Berjalan', '=COUNTA(Orders!A2:A)', 'Pesanan']);
  sheetDash.appendRow(['Gross Merchandise Value (GMV)', '=SUM(Orders!L2:L)', 'Rupiah']);
  sheetDash.appendRow(['Total Dana di Rekening Bersama (Escrow)', '=SUMIFS(Orders!L2:L, Orders!N2:N, "IN_ESCROW")', 'Rupiah']);
  sheetDash.appendRow(['Pendapatan Biaya Perlindungan Escrow', '=SUM(Orders!K2:K)', 'Rupiah']);
  sheetDash.appendRow(['Total Penghematan Karbon (CO2e)', '=SUM(Products!P2:P)', 'Kilogram CO2e']);
  sheetDash.appendRow(['Total Air Bersih Dihemat', '=SUM(Products!Q2:Q)', 'Liter Air']);

  sheetDash.autoResizeColumns(1, 3);
  sheetProducts.autoResizeColumns(1, 19);
  sheetOrders.autoResizeColumns(1, 17);
  sheetVerif.autoResizeColumns(1, 8);
}

function formatHeaderRow(sheet, bgColor, textColor) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground(bgColor)
       .setFontColor(textColor)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
}

// 4. FUNGSI SERVER YANG DIPANGGIL CLIENT (google.script.run)
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initDatabaseIfNotExists();

  // Ambil Products
  const sheetProd = ss.getSheetByName('Products');
  const prodRows = sheetProd.getDataRange().getValues();
  const prodHeaders = prodRows[0];
  const products = prodRows.slice(1).map(r => ({
    id: r[0],
    title: r[1],
    category: r[2],
    price: Number(r[3]),
    originalPrice: Number(r[4]),
    condition: r[5],
    conditionPercent: Number(r[6]),
    location: r[7],
    sellerName: r[8],
    material: r[9],
    color: r[10],
    brand: r[11],
    flaws: r[12] ? r[12].toString().split(',').map(s => s.trim()) : [],
    verificationStatus: (r[13] || '').toString().toLowerCase(),
    isPromoted: r[14] === true || r[14] === 'TRUE',
    sustainability: {
      co2SavedKg: Number(r[15]) || 4.5,
      waterSavedLiter: Number(r[16]) || 2500
    },
    images: [r[17] || 'https://images.unsplash.com/photo-1556905055-8f358a7a47b2?w=800']
  }));

  // Ambil Orders
  const sheetOrd = ss.getSheetByName('Orders');
  const ordRows = sheetOrd.getDataRange().getValues();
  const orders = ordRows.slice(1).map(r => ({
    id: r[0],
    orderNumber: r[1],
    productId: r[2],
    productTitle: r[3],
    buyerName: r[4],
    sellerName: r[5],
    address: r[6],
    courier: r[7],
    productPrice: Number(r[8]),
    shippingCost: Number(r[9]),
    transactionFee: Number(r[10]),
    totalPayment: Number(r[11]),
    paymentMethod: r[12],
    escrowStatus: r[13],
    orderStatus: r[14],
    trackingNumber: r[15],
    createdAt: r[16] ? r[16].toString() : ''
  }));

  return {
    products: products,
    orders: orders
  };
}

// Simpan Checkout Pesanan ke Spreadsheet
function processCheckoutOrder(order) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrd = ss.getSheetByName('Orders');

  const orderId = order.id || ('ord-' + new Date().getTime());
  const orderNum = order.orderNumber || ('RLV-' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd-HHmm'));
  const tracking = order.trackingNumber || ('RESI-' + Math.floor(Math.random() * 1000000000));

  sheetOrd.appendRow([
    orderId,
    orderNum,
    order.productId,
    order.productTitle,
    order.buyerName,
    order.sellerName,
    order.address,
    order.courier,
    order.productPrice,
    order.shippingCost,
    order.transactionFee,
    order.totalPayment,
    order.paymentMethod,
    'IN_ESCROW',
    'PAID',
    tracking,
    new Date()
  ]);

  return {
    success: true,
    orderNumber: orderNum,
    trackingNumber: tracking
  };
}

// Tambah Produk Hasil AI Scanner atau Form Seller
function saveNewProduct(product) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Products');

  const newId = product.id || ('prod-' + new Date().getTime());
  sheet.appendRow([
    newId,
    product.title,
    product.category,
    product.price,
    product.originalPrice || 250000,
    product.condition,
    product.conditionPercent || 90,
    product.location || 'Malang',
    product.sellerName || 'sarahh_preloved',
    product.material || 'Katun',
    product.color || 'Netral',
    product.brand || 'Thrift Curated',
    Array.isArray(product.flaws) ? product.flaws.join(', ') : (product.flaws || ''),
    'VERIFIED', // Set verified for instant prototype gratification
    product.isPromoted ? true : false,
    product.co2Saved || 4.5,
    product.waterSaved || 2500,
    product.imageUrl || 'https://images.unsplash.com/photo-1556905055-8f358a7a47b2?w=800',
    new Date()
  ]);

  return { success: true, id: newId };
}

// Keputusan Verifikasi Admin
function updateProductVerification(productId, status, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Products');
  const sheetVerif = ss.getSheetByName('Verifikasi_Kurasi');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == productId) {
      sheet.getRange(i + 1, 14).setValue(status.toUpperCase());

      sheetVerif.appendRow([
        'vrf-' + new Date().getTime(),
        productId,
        data[i][1],
        status.toUpperCase(),
        'Super Admin',
        note || 'Inspeksi fisik terbukti akurat',
        95.0,
        new Date()
      ]);
      return { success: true };
    }
  }
  return { success: false, message: 'Produk tidak ditemukan' };
}

// Rekap Keuangan Manual
function hitungRekapKeuangan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrders = ss.getSheetByName('Orders');
  const data = sheetOrders.getDataRange().getValues();
  
  let totalGMV = 0, totalEscrow = 0, totalFee = 0;
  for (let i = 1; i < data.length; i++) {
    const totalBayar = Number(data[i][11]) || 0;
    const fee = Number(data[i][10]) || 0;
    totalGMV += totalBayar;
    totalFee += fee;
    if (data[i][13] === 'IN_ESCROW') totalEscrow += totalBayar;
  }

  const formatIDR = num => 'Rp ' + num.toLocaleString('id-ID');
  SpreadsheetApp.getUi().alert(
    '📊 Rekap Keuangan & Escrow RELOVE:\n\n' +
    '• Total GMV: ' + formatIDR(totalGMV) + '\n' +
    '• Dana Tertahan di Escrow: ' + formatIDR(totalEscrow) + '\n' +
    '• Pendapatan Biaya Layanan: ' + formatIDR(totalFee)
  );
}

function hitungDampakLingkungan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProd = ss.getSheetByName('Products');
  const data = sheetProd.getDataRange().getValues();

  let totalCO2 = 0, totalAir = 0;
  for (let i = 1; i < data.length; i++) {
    totalCO2 += Number(data[i][15]) || 0;
    totalAir += Number(data[i][16]) || 0;
  }

  SpreadsheetApp.getUi().alert(
    '🌿 Laporan Dampak Keberlanjutan RELOVE (SDG 12):\n\n' +
    '• Emisi Karbon Dicegah: ' + totalCO2.toFixed(1) + ' kg CO2e\n' +
    '• Air Bersih Dihemat: ' + totalAir.toLocaleString('id-ID') + ' Liter Air'
  );
}
 path from prompt to production with Gemini.</p>

  <a href="https://aistudio.google.com/apps">Start building</a>

</div>
