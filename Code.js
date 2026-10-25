/**
 * SELLA'S FINANCIAL TRACKING - Backend (Google Apps Script)
 * Versi Lengkap:
 * - General Ledger (Buku Besar) Engine untuk Ekspor PDF
 * - Modul Manajemen Tabungan & Portofolio Imbal Hasil (Savings & Investment)
 * - Estimasi Ketahanan Kas (Cash Runway)
 * - Proteksi Concurrency dengan LockService & Proteksi Data Sheet
 */

const SH = {
  W: 'Wallets',
  T: 'Transactions',
  C: 'Categories',
  S: 'Settings',
  SI: 'SavingsInstruments' // Sheet untuk portofolio tabungan & imbal hasil
};

const ACCOUNT_TYPES = ['Harian', 'Tabungan'];

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.seed === '1') {
      const res = seedDummyData();
      return ContentService.createTextOutput(JSON.stringify(res, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.test === '1') {
      const data = getDashboardData('all');
      return ContentService.createTextOutput(JSON.stringify(data, null, 2)).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', error: String(err), stack: err.stack }, null, 2)).setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Sella's Financial Tracking")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function seedDummyData() {
  const ss = getSpreadsheet_();
  
  // 1. Wallets
  let wSheet = ss.getSheetByName(SH.W);
  if (!wSheet) wSheet = ss.insertSheet(SH.W);
  wSheet.clear();
  wSheet.appendRow(['ID', 'Nama', 'JenisAkun']);
  wSheet.appendRow([1, 'BCA Harian', 'Harian']);
  wSheet.appendRow([2, 'Cash Tunai', 'Harian']);
  wSheet.appendRow([3, 'Seabank', 'Harian']);
  wSheet.appendRow([4, 'Bank Jago', 'Tabungan']);
  try { wSheet.setFrozenRows(1); } catch (e) {}

  // 2. Categories
  let cSheet = ss.getSheetByName(SH.C);
  if (!cSheet) cSheet = ss.insertSheet(SH.C);
  cSheet.clear();
  cSheet.appendRow(['ID', 'Nama']);
  const cats = ['Gaji', 'Makan', 'Transport', 'Belanja', 'Hiburan', 'Kesehatan', 'Tabungan', 'Lainnya', 'Transfer Antar Bank'];
  cats.forEach((c, idx) => cSheet.appendRow([idx + 1, c]));
  try { cSheet.setFrozenRows(1); } catch (e) {}

  // 3. Transactions
  let tSheet = ss.getSheetByName(SH.T);
  if (!tSheet) tSheet = ss.insertSheet(SH.T);
  tSheet.clear();
  tSheet.appendRow(['ID', 'Tanggal', 'Jenis', 'DompetAsal', 'DompetTujuan', 'Kategori', 'Keterangan', 'Jumlah']);
  tSheet.appendRow([1, '2026-10-01', 'masuk', 'BCA Harian', '', 'Gaji', 'Gaji Bulanan Oktober', 9500000]);
  tSheet.appendRow([2, '2026-10-02', 'transfer', 'BCA Harian', 'Bank Jago', 'Tabungan', 'Nabung Awal Bulan', 3000000]);
  tSheet.appendRow([3, '2026-10-03', 'keluar', 'BCA Harian', '', 'Makan', 'Belanja Mingguan', 450000]);
  tSheet.appendRow([4, '2026-10-04', 'keluar', 'Cash Tunai', '', 'Transport', 'Bensin & Tol', 120000]);
  tSheet.appendRow([5, '2026-10-05', 'keluar', 'BCA Harian', '', 'Belanja', 'Keperluan Rumah', 350000]);
  tSheet.appendRow([6, '2026-09-25', 'masuk', 'BCA Harian', '', 'Gaji', 'Gaji September', 9000000]);
  tSheet.appendRow([7, '2026-09-28', 'keluar', 'BCA Harian', '', 'Makan', 'Makan & Cafe', 800000]);
  try { tSheet.setFrozenRows(1); } catch (e) {}

  // 4. Settings
  let sSheet = ss.getSheetByName(SH.S);
  if (!sSheet) sSheet = ss.insertSheet(SH.S);
  sSheet.clear();
  sSheet.appendRow(['Key', 'Value']);
  sSheet.appendRow(['BudgetPercent', 15]);
  try { sSheet.setFrozenRows(1); } catch (e) {}

  // 5. Savings
  let siSheet = ss.getSheetByName(SH.SI);
  if (!siSheet) siSheet = ss.insertSheet(SH.SI);
  siSheet.clear();
  siSheet.appendRow(['ID', 'Nama', 'Kategori', 'NominalPokok', 'BungaPertahun', 'Catatan']);
  siSheet.appendRow([1, 'Deposito Seabank', 'Deposito Bank', 10000000, 6.0, 'Bunga cair bulanan']);
  siSheet.appendRow([2, 'Kantong Terkunci Jago', 'Tabungan Digital', 5000000, 5.0, 'Dana darurat']);
  siSheet.appendRow([3, 'SBN ORI026', 'Surat Berharga Negara (SBN)', 10000000, 6.4, 'Kupon tetap per bulan']);
  siSheet.appendRow([4, 'RDPU Sucorinvest', 'Reksadana Pasar Uang', 3500000, 4.8, 'Likuid sewaktu-waktu']);
  try { siSheet.setFrozenRows(1); } catch (e) {}

  return { status: 'success', message: 'Data dummy berhasil dimasukkan ke Spreadsheet!' };
}


const SPREADSHEET_ID = '1qnvZ8icqcCm9OfNfrMySF-Ta58tHOcG_dHTJMkv5HeI';

function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function autoInitSheet_(ss, name) {
  let s = ss.getSheetByName(name);
  if (s) return s;
  s = ss.insertSheet(name);
  if (name === SH.W) {
    s.appendRow(['ID', 'Nama', 'JenisAkun']);
    s.appendRow([1, 'BCA Harian', 'Harian']);
    s.appendRow([2, 'Cash Tunai', 'Harian']);
    s.appendRow([3, 'Seabank', 'Harian']);
    s.appendRow([4, 'Bank Jago', 'Tabungan']);
  } else if (name === SH.C) {
    s.appendRow(['ID', 'Nama']);
    const cats = ['Gaji', 'Makan', 'Transport', 'Belanja', 'Hiburan', 'Kesehatan', 'Tabungan', 'Lainnya', 'Transfer Antar Bank'];
    cats.forEach((c, idx) => s.appendRow([idx + 1, c]));
  } else if (name === SH.T) {
    s.appendRow(['ID', 'Tanggal', 'Jenis', 'DompetAsal', 'DompetTujuan', 'Kategori', 'Keterangan', 'Jumlah']);
  } else if (name === SH.S) {
    s.appendRow(['Key', 'Value']);
    s.appendRow(['BudgetPercent', 15]);
  } else if (name === SH.SI) {
    s.appendRow(['ID', 'Nama', 'Kategori', 'NominalPokok', 'BungaPertahun', 'Catatan']);
    s.appendRow([1, 'Deposito Seabank', 'Deposito Bank', 10000000, 6.0, 'Bunga cair bulanan']);
    s.appendRow([2, 'Kantong Terkunci Jago', 'Tabungan Digital', 5000000, 5.0, 'Dana darurat']);
    s.appendRow([3, 'SBN ORI026', 'Surat Berharga Negara (SBN)', 10000000, 6.4, 'Kupon tetap per bulan']);
    s.appendRow([4, 'RDPU Sucorinvest', 'Reksadana Pasar Uang', 3500000, 4.8, 'Likuid sewaktu-waktu']);
  }
  try { s.setFrozenRows(1); } catch (e) {}
  return s;
}

function sheet_(name) {
  const ss = getSpreadsheet_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = autoInitSheet_(ss, name);
  }
  return s;
}

/**
 * Jalankan fungsi ini SEKALI saat pertama kali memasang aplikasi.
 * Proteksi Keamanan: Tidak akan menghapus jika sheet Transactions sudah berisi mutasi riil.
 */
function setupSheets() {
  const ss = getSpreadsheet_();
  const txSheet = ss.getSheetByName(SH.T);
  
  if (txSheet && txSheet.getLastRow() > 1) {
    throw new Error('Dibatalkan demi keamanan: Sheet transaksi sudah memiliki data! Untuk reset ulang, hapus isi sheet secara manual.');
  }

  const init = (name, header, rows) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    s.clear();
    s.appendRow(header);
    if (rows && rows.length) {
      rows.forEach(r => s.appendRow(r));
    }
    s.setFrozenRows(1);
  };

  // 1. Master Dompet / Rekening
  init(SH.W, ['ID', 'Nama', 'JenisAkun'], [
    [1, 'BCA Harian', 'Harian'],
    [2, 'Cash Tunai', 'Harian'],
    [3, 'Seabank', 'Harian'],
    [4, 'Bank Jago', 'Tabungan']
  ]);

  // 2. Master Kategori
  init(SH.C, ['ID', 'Nama'], [
    [1, 'Gaji'],
    [2, 'Makan'],
    [3, 'Transport'],
    [4, 'Belanja'],
    [5, 'Hiburan'],
    [6, 'Kesehatan'],
    [7, 'Tabungan'],
    [8, 'Lainnya'],
    [9, 'Transfer Antar Bank']
  ]);

  // 3. Tabel Transaksi / Mutasi
  init(SH.T, ['ID', 'Tanggal', 'Jenis', 'DompetAsal', 'DompetTujuan', 'Kategori', 'Keterangan', 'Jumlah'], []);

  // 4. Konfigurasi Sistem
  init(SH.S, ['Key', 'Value'], [
    ['BudgetPercent', 15]
  ]);

  // 5. Master Portofolio Tabungan & Bunga (Savings & Investments)
  init(SH.SI, ['ID', 'Nama', 'Kategori', 'NominalPokok', 'BungaPertahun', 'Catatan'], [
    [1, 'Deposito Seabank', 'Deposito Bank', 10000000, 6.0, 'Bunga cair bulanan'],
    [2, 'Kantong Terkunci Jago', 'Tabungan Digital', 5000000, 5.0, 'Dana darurat'],
    [3, 'SBN ORI026', 'Surat Berharga Negara (SBN)', 10000000, 6.4, 'Kupon tetap per bulan'],
    [4, 'RDPU Sucorinvest', 'Reksadana Pasar Uang', 3500000, 4.8, 'Likuid sewaktu-waktu']
  ]);

  // Bersihkan Sheet1 default jika kosong
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  try {
    SpreadsheetApp.getUi().alert('Setup Selesai! Seluruh struktur tabel dan portofolio tabungan telah siap.');
  } catch (e) {
    Logger.log('Setup Selesai via web/service context.');
  }
}

function rows_(name) {
  const s = sheet_(name);
  if (!s || s.getLastRow() <= 1) return [];
  const d = s.getDataRange().getValues();
  d.shift(); // Buang header baris 1
  return d.filter(r => r[0] !== '' && r[0] !== null);
}

function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  const s = String(d).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
}

const getWallets_ = () => rows_(SH.W).map(r => ({ id: r[0], name: String(r[1]).trim(), type: r[2] || 'Harian' }));

const getCategories_ = () => rows_(SH.C).map(r => ({ id: r[0], name: String(r[1]).trim() }));

const getTransactions_ = () => rows_(SH.T).map(r => ({
  id: Number(r[0]),
  date: fmtDate_(r[1]),
  type: String(r[2]).toLowerCase(),
  walletFrom: String(r[3] || '').trim(),
  walletTo: String(r[4] || '').trim(),
  category: String(r[5] || '').trim(),
  description: String(r[6] || '').trim(),
  amount: Number(r[7] || 0)
}));

const getSavingsInstruments_ = () => rows_(SH.SI).map(r => ({
  id: Number(r[0]),
  name: String(r[1] || '').trim(),
  category: String(r[2] || 'Lainnya').trim(),
  amount: Number(r[3] || 0),
  rate: Number(r[4] || 0),
  notes: String(r[5] || '').trim()
}));

function getBudgetPercent_() {
  const s = sheet_(SH.S);
  if (!s) return 15;
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === 'BudgetPercent') return Number(d[i][1]) || 15;
  }
  return 15;
}

const ym_ = t => (t && t.date ? String(t.date).slice(0, 7) : '');

const getMonthsList_ = txs => {
  const set = {};
  txs.forEach(t => {
    const m = ym_(t);
    if (m && m.length === 7) set[m] = 1;
  });
  return Object.keys(set).sort().reverse();
};

function filterTx_(txs, month, category, wallet) {
  return txs.filter(t =>
    (!month || month === 'all' || ym_(t) === month) &&
    (!category || category === 'all' || t.category === category) &&
    (!wallet || wallet === 'all' || t.walletFrom === wallet || t.walletTo === wallet)
  );
}

const sum_ = (txs, type) => txs.filter(t => t.type === type).reduce((s, t) => s + t.amount, 0);

function groupKeluar_(txs, key) {
  const o = {};
  txs.filter(t => t.type === 'keluar').forEach(t => {
    const k = t[key] || 'Lainnya';
    o[k] = (o[k] || 0) + t.amount;
  });
  return o;
}

/**
 * Menghitung estimasi ketahanan saldo operasional harian (Cash Runway).
 * Mengabaikan pos tabungan dan belanja besar (outlier > Rp350.000) agar rata-rata murni belanja rutin.
 */
function calcCashRunway_(txs, totalHarian) {
  const expenseTxs = txs.filter(t => t.type === 'keluar' && t.category !== 'Tabungan');
  if (!expenseTxs.length || totalHarian <= 0) {
    return {
      avgDaily: 0,
      daysLeft: totalHarian > 0 ? 30 : 0,
      status: totalHarian > 0 ? 'Aman' : 'Kritis',
      count: 0,
      totalHarian: totalHarian || 0
    };
  }

  // Saring pengeluaran bernominal besar di atas Rp350.000
  const routineTxs = expenseTxs.filter(t => t.amount <= 350000);
  const activeTxs = routineTxs.length ? routineTxs : expenseTxs;

  const dateSet = {};
  activeTxs.forEach(t => { if (t.date) dateSet[t.date] = 1; });
  const uniqueDays = Object.keys(dateSet).length || 1;
  const totalRoutine = activeTxs.reduce((s, t) => s + t.amount, 0);
  const avgDaily = Math.round(totalRoutine / uniqueDays);

  const daysLeft = avgDaily > 0 ? Math.floor(totalHarian / avgDaily) : 99;
  let status = 'Aman';
  if (daysLeft < 7) status = 'Kritis';
  else if (daysLeft <= 14) status = 'Waspada';

  return {
    avgDaily,
    daysLeft,
    status,
    totalHarian,
    count: activeTxs.length
  };
}

function getDashboardDataObject_(filterMonth) {
  const month = filterMonth || 'all';
  const all = getTransactions_();
  const wallets = getWallets_();
  const budgetPercent = getBudgetPercent_();
  const filtered = filterTx_(all, month);
  const savings = getSavingsInstruments_();

  // 1. Total masuk & keluar untuk periode bulan yang difilter
  let totalMasuk = 0;
  let totalKeluar = 0;
  filtered.forEach(t => {
    if (t.type === 'masuk') totalMasuk += t.amount;
    else if (t.type === 'keluar') totalKeluar += t.amount;
  });

  // 2. Saldo dompet riil dihitung dari SELURUH riwayat mutasi kas (kumulatif)
  const bal = {};
  wallets.forEach(w => { bal[w.name] = 0; });
  all.forEach(t => {
    if (t.type === 'masuk') {
      if (t.walletFrom in bal) bal[t.walletFrom] += t.amount;
    } else if (t.type === 'keluar') {
      if (t.walletFrom in bal) bal[t.walletFrom] -= t.amount;
    } else if (t.type === 'transfer') {
      if (t.walletFrom in bal) bal[t.walletFrom] -= t.amount;
      if (t.walletTo in bal) bal[t.walletTo] += t.amount;
    }
  });

  let totalTabungan = 0;
  let totalHarian = 0;
  const wb = wallets.map(w => {
    const b = bal[w.name] || 0;
    if (w.type === 'Tabungan') {
      totalTabungan += b;
    } else {
      totalHarian += b;
    }
    return { id: w.id, name: w.name, type: w.type, balance: b };
  });

  // Urutkan transaksi dari tanggal terbaru ke terlama
  const sortedTx = filtered.slice().sort((a, b) => {
    const dDiff = new Date(b.date) - new Date(a.date);
    if (dDiff !== 0) return dDiff;
    return b.id - a.id;
  });

  // Hitung ketahanan kas dari saldo harian operasional
  const runway = calcCashRunway_(filtered, totalHarian);

  return {
    wallets: wb,
    categories: getCategories_().map(c => c.name),
    transactions: sortedTx,
    months: getMonthsList_(all),
    filterMonth: month,
    totalMasuk,
    totalKeluar,
    totalSaldo: totalHarian,       // Saldo operasional harian yang siap dipakai
    totalTabungan,                // Total dana simpanan
    totalHarian,
    netWorth: totalHarian + totalTabungan,
    budgetPercent,
    maxBudgetNominal: (budgetPercent / 100) * totalMasuk,
    runway,
    savings                       // Portofolio tabungan & bunga
  };
}

function getDashboardData(m) {
  return getDashboardDataObject_(m);
}

function setBudgetPercent(percent, m) {
  const p = parseFloat(percent);
  if (isNaN(p) || p < 1 || p > 100) throw new Error('Persentase budget harus antara 1% - 100%.');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const s = sheet_(SH.S);
    const d = s.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === 'BudgetPercent') {
        s.getRange(i + 1, 2).setValue(p);
        found = true;
        break;
      }
    }
    if (!found) s.appendRow(['BudgetPercent', p]);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function addWallet(name, type, m) {
  name = (name || '').trim();
  if (!name) throw new Error('Nama dompet tidak boleh kosong.');
  if (ACCOUNT_TYPES.indexOf(type) === -1) type = 'Harian';
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ws = getWallets_();
    if (ws.some(w => w.name.toLowerCase() === name.toLowerCase())) {
      throw new Error('Dompet dengan nama "' + name + '" sudah ada!');
    }
    const nextId = ws.length ? Math.max(...ws.map(w => Number(w.id))) + 1 : 1;
    sheet_(SH.W).appendRow([nextId, name, type]);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function deleteRowById_(name, id) {
  const s = sheet_(name);
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(id)) {
      s.deleteRow(i + 1);
      break;
    }
  }
}

function deleteWallet(id, m) {
  const w = getWallets_().find(x => String(x.id) === String(id));
  if (!w) return getDashboardDataObject_(m);
  
  const inUse = getTransactions_().some(t => t.walletFrom === w.name || t.walletTo === w.name);
  if (inUse) {
    throw new Error('Dompet "' + w.name + '" masih memiliki riwayat transaksi dan tidak dapat dihapus.');
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    deleteRowById_(SH.W, id);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function addCategory(name, m) {
  name = (name || '').trim();
  if (!name) throw new Error('Nama kategori tidak boleh kosong.');
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const cs = getCategories_();
    if (cs.some(c => c.name.toLowerCase() === name.toLowerCase())) {
      throw new Error('Kategori "' + name + '" sudah ada!');
    }
    const nextId = cs.length ? Math.max(...cs.map(c => Number(c.id))) + 1 : 1;
    sheet_(SH.C).appendRow([nextId, name]);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function deleteCategory(name, m) {
  name = (name || '').trim();
  const c = getCategories_().find(x => x.name.toLowerCase() === name.toLowerCase());
  if (!c) return getDashboardDataObject_(m);
  
  if (c.name === 'Transfer Antar Bank') {
    throw new Error('Kategori sistem "Transfer Antar Bank" tidak boleh dihapus.');
  }

  const inUse = getTransactions_().some(t => t.category === c.name);
  if (inUse) {
    throw new Error('Kategori "' + c.name + '" masih digunakan dalam riwayat transaksi.');
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    deleteRowById_(SH.C, c.id);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function addTransaction(tx, m) {
  if (!tx || !tx.type || !tx.walletFrom || !tx.amount || !tx.date || !tx.category) {
    throw new Error('Data transaksi tidak lengkap. Silakan periksa kembali formulir.');
  }

  const amt = parseFloat(tx.amount);
  if (isNaN(amt) || amt <= 0) throw new Error('Nominal transaksi harus lebih dari 0.');

  if (tx.type === 'transfer') {
    if (!tx.walletTo || tx.walletFrom === tx.walletTo) {
      throw new Error('Dompet asal dan tujuan transfer tidak boleh sama.');
    }
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const txs = getTransactions_();
    const id = txs.length ? Math.max(...txs.map(t => Number(t.id))) + 1 : 1;
    
    sheet_(SH.T).appendRow([
      id,
      tx.date,
      tx.type,
      tx.walletFrom,
      tx.type === 'transfer' ? tx.walletTo : '',
      tx.category,
      tx.description || '',
      amt
    ]);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function deleteTransaction(id, m) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    deleteRowById_(SH.T, id);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function addSavingsInstrument(item, m) {
  if (!item || !item.name) throw new Error('Nama instrumen tabungan tidak boleh kosong.');
  const amt = parseFloat(item.amount);
  if (isNaN(amt) || amt <= 0) throw new Error('Nominal pokok tabungan harus lebih dari 0.');
  const rate = parseFloat(item.rate) || 0;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const items = getSavingsInstruments_();
    const nextId = items.length ? Math.max(...items.map(x => Number(x.id))) + 1 : 1;
    sheet_(SH.SI).appendRow([
      nextId,
      item.name.trim(),
      (item.category || 'Lainnya').trim(),
      amt,
      rate,
      (item.notes || '').trim()
    ]);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function deleteSavingsInstrument(id, m) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    deleteRowById_(SH.SI, id);
  } finally {
    lock.releaseLock();
  }

  return getDashboardDataObject_(m);
}

function getAnalysisData(month, category, wallet) {
  month = month || 'all';
  category = category || 'all';
  wallet = wallet || 'all';

  const all = getTransactions_();
  const months = getMonthsList_(all);
  const f = filterTx_(all, month, category, wallet);
  const keluarTx = f.filter(t => t.type === 'keluar');
  const totalMasuk = sum_(f, 'masuk');
  const totalKeluar = sum_(f, 'keluar');

  const byMonth = {};
  f.forEach(t => {
    const k = ym_(t) || 'Lainnya';
    byMonth[k] = byMonth[k] || { masuk: 0, keluar: 0 };
    if (t.type === 'masuk') byMonth[k].masuk += t.amount;
    if (t.type === 'keluar') byMonth[k].keluar += t.amount;
  });

  // Komparasi Dinamis vs Bulan Sebelumnya
  let comparison = null;
  const targetMonth = (month && month !== 'all') ? month : (months[0] || null);

  if (targetMonth) {
    const idx = months.indexOf(targetMonth);
    const prevMonth = (idx !== -1 && idx + 1 < months.length) ? months[idx + 1] : null;

    if (prevMonth) {
      const cur = filterTx_(all, targetMonth, category, wallet);
      const prev = filterTx_(all, prevMonth, category, wallet);
      const cIn = sum_(cur, 'masuk');
      const pIn = sum_(prev, 'masuk');
      const cOut = sum_(cur, 'keluar');
      const pOut = sum_(prev, 'keluar');

      comparison = {
        current: targetMonth,
        prev: prevMonth,
        curMasuk: cIn,
        prevMasuk: pIn,
        curKeluar: cOut,
        prevKeluar: pOut,
        curNet: cIn - cOut,
        prevNet: pIn - pOut,
        curByCategory: groupKeluar_(cur, 'category'),
        prevByCategory: groupKeluar_(prev, 'category')
      };
    }
  }

  return {
    months,
    categories: getCategories_().map(c => c.name),
    wallets: getWallets_().map(w => w.name),
    filterMonth: month,
    filterCategory: category,
    filterWallet: wallet,
    totalMasuk,
    totalKeluar,
    net: totalMasuk - totalKeluar,
    count: f.length,
    avgKeluar: keluarTx.length ? (totalKeluar / keluarTx.length) : 0,
    byCategory: groupKeluar_(f, 'category'),
    byWallet: groupKeluar_(f, 'walletFrom'),
    trendByMonth: Object.keys(byMonth).sort().map(k => ({
      month: k,
      masuk: byMonth[k].masuk,
      keluar: byMonth[k].keluar
    })),
    comparison
  };
}

/**
 * Menyusun data Buku Besar - Mutasi Akun (General Ledger) untuk ekspor cetak PDF.
 */
function getLedgerReportData(targetMonth) {
  const all = getTransactions_();
  const wallets = getWallets_();
  const budgetPercent = getBudgetPercent_();
  const months = getMonthsList_(all);

  if (!targetMonth || targetMonth === 'all') {
    targetMonth = months.length ? months[0] : Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM');
  }

  // 1. Hitung Saldo Awal dompet (sebelum bulan target yang dipilih)
  const initialBalances = {};
  wallets.forEach(w => { initialBalances[w.name] = 0; });

  all.forEach(t => {
    const ym = ym_(t);
    if (ym && ym < targetMonth) {
      if (t.type === 'masuk' && t.walletFrom in initialBalances) initialBalances[t.walletFrom] += t.amount;
      else if (t.type === 'keluar' && t.walletFrom in initialBalances) initialBalances[t.walletFrom] -= t.amount;
      else if (t.type === 'transfer') {
        if (t.walletFrom in initialBalances) initialBalances[t.walletFrom] -= t.amount;
        if (t.walletTo in initialBalances) initialBalances[t.walletTo] += t.amount;
      }
    }
  });

  // 2. Transaksi bulan berjalan
  const curTx = all.filter(t => ym_(t) === targetMonth)
                   .sort((a, b) => new Date(a.date) - new Date(b.date) || a.id - b.id);

  let totalMasuk = 0;
  let totalKeluar = 0;
  const ledgerByWallet = {};

  wallets.forEach(w => {
    ledgerByWallet[w.name] = {
      name: w.name,
      type: w.type,
      saldoAwal: initialBalances[w.name] || 0,
      totalDebet: 0,
      totalKredit: 0,
      mutasi: 0,
      saldoAkhir: initialBalances[w.name] || 0,
      entries: []
    };
  });

  curTx.forEach(t => {
    if (t.type === 'masuk') {
      totalMasuk += t.amount;
      if (ledgerByWallet[t.walletFrom]) {
        ledgerByWallet[t.walletFrom].entries.push({
          date: t.date,
          ref: 'CR' + String(t.id).padStart(4, '0'),
          desc: (t.category ? t.category + ' - ' : '') + (t.description || 'Penerimaan Kas'),
          debet: t.amount,
          kredit: 0
        });
      }
    } else if (t.type === 'keluar') {
      totalKeluar += t.amount;
      if (ledgerByWallet[t.walletFrom]) {
        ledgerByWallet[t.walletFrom].entries.push({
          date: t.date,
          ref: 'CD' + String(t.id).padStart(4, '0'),
          desc: (t.category ? t.category + ' - ' : '') + (t.description || 'Pengeluaran Kas'),
          debet: 0,
          kredit: t.amount
        });
      }
    } else if (t.type === 'transfer') {
      if (ledgerByWallet[t.walletFrom]) {
        ledgerByWallet[t.walletFrom].entries.push({
          date: t.date,
          ref: 'TR' + String(t.id).padStart(4, '0'),
          desc: 'Transfer ke ' + t.walletTo + (t.description ? ' (' + t.description + ')' : ''),
          debet: 0,
          kredit: t.amount
        });
      }
      if (ledgerByWallet[t.walletTo]) {
        ledgerByWallet[t.walletTo].entries.push({
          date: t.date,
          ref: 'TR' + String(t.id).padStart(4, '0'),
          desc: 'Transfer dari ' + t.walletFrom + (t.description ? ' (' + t.description + ')' : ''),
          debet: t.amount,
          kredit: 0
        });
      }
    }
  });

  let totalHarianBalance = 0;
  const walletReports = [];
  wallets.forEach(w => {
    const item = ledgerByWallet[w.name];
    let curBal = item.saldoAwal;
    item.entries.forEach(e => {
      curBal += (e.debet - e.kredit);
      e.saldo = curBal;
      item.totalDebet += e.debet;
      item.totalKredit += e.kredit;
    });
    item.saldoAkhir = curBal;
    item.mutasi = item.totalDebet - item.totalKredit;
    if (w.type !== 'Tabungan') {
      totalHarianBalance += curBal;
    }
    walletReports.push(item);
  });

  // 3. Analisa Anggaran
  const maxBudgetNominal = (budgetPercent / 100) * totalMasuk;
  const budgetRatio = maxBudgetNominal > 0 ? (totalKeluar / maxBudgetNominal) * 100 : 0;
  const budgetStatus = totalKeluar > maxBudgetNominal ? 'Overbudget' : (budgetRatio >= 80 ? 'Waspada' : 'Aman');

  // 4. Analisa Komparasi vs Bulan Sebelumnya
  const curIdx = months.indexOf(targetMonth);
  const prevMonth = (curIdx !== -1 && curIdx + 1 < months.length) ? months[curIdx + 1] : null;
  let prevAnalysis = null;

  if (prevMonth) {
    const prevTx = all.filter(t => ym_(t) === prevMonth);
    const prevMasuk = sum_(prevTx, 'masuk');
    const prevKeluar = sum_(prevTx, 'keluar');
    prevAnalysis = {
      prevMonth,
      prevMasuk,
      prevKeluar,
      prevNet: prevMasuk - prevKeluar,
      diffMasuk: totalMasuk - prevMasuk,
      diffKeluar: totalKeluar - prevKeluar,
      diffNet: (totalMasuk - totalKeluar) - (prevMasuk - prevKeluar)
    };
  }

  // 5. Ketahanan Kas (Cash Runway)
  const runway = calcCashRunway_(curTx, totalHarianBalance);
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';

  return {
    targetMonth,
    printDate: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'),
    budgetAnalysis: {
      budgetPercent,
      totalMasuk,
      totalKeluar,
      maxBudgetNominal,
      budgetRatio,
      budgetStatus,
      sisaBudget: maxBudgetNominal - totalKeluar
    },
    prevAnalysis,
    runway,
    walletReports
  };
}


/* ==========================================================================
 * MODUL AI NOTA TRACKER (GEMINI 3 / 2.5 VISION)
 * ========================================================================== */

const GEMINI_MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
const MAX_BASE64_LENGTH = 7 * 1024 * 1024; // 7MB max payload guard

function getApiKey_(clientKey) {
  if (clientKey && typeof clientKey === 'string' && clientKey.trim().length > 10) {
    return clientKey.trim();
  }
  const prop = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (prop && prop.trim().length > 10) return prop.trim();
  return null;
}

function saveGeminiApiKey(key) {
  if (!key || typeof key !== 'string' || key.trim().length < 10) {
    throw new Error('Format Gemini API Key tidak valid.');
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key.trim());
  return { status: 'success', message: 'API Key berhasil disimpan di Script Properties.' };
}

function callGemini_(model, body, apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  let res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  let json = {};
  try { json = JSON.parse(res.getContentText()); } catch (e) {}

  // Fallback otomatis jika model menolak thinkingConfig (Error 400)
  if (res.getResponseCode() === 400 && body.generationConfig && body.generationConfig.thinkingConfig) {
    const retryBody = JSON.parse(JSON.stringify(body));
    delete retryBody.generationConfig.thinkingConfig;
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(retryBody),
      muteHttpExceptions: true
    });
    try { json = JSON.parse(res.getContentText()); } catch (e) {}
  }

  return { code: res.getResponseCode(), json: json, model: model };
}

function executeGeminiFallback_(body, apiKey) {
  let lastRes = null;
  let lastError = '';
  for (let m = 0; m < GEMINI_MODELS.length; m++) {
    const model = GEMINI_MODELS[m];
    try {
      const result = callGemini_(model, body, apiKey);
      if (result.code === 200) return result;
      lastRes = result;
      lastError = (result.json && result.json.error && result.json.error.message) || ('HTTP ' + result.code);
    } catch (e) {
      lastError = String(e);
    }
  }
  return lastRes || { code: 500, json: { error: { message: lastError || 'Semua model Gemini gagal dihubungi.' } } };
}

/**
 * Menganalisa foto nota/struk belanja menggunakan Gemini Vision.
 */
function analyzeReceiptImageServer(base64Data, mimeType, clientKey) {
  const apiKey = getApiKey_(clientKey);
  if (!apiKey) {
    return {
      status: 'no_key',
      message: 'GEMINI_API_KEY belum disetel. Masukkan API Key Anda di form Scan Nota atau simpan di Script Properties.'
    };
  }

  if (!base64Data || base64Data.length > MAX_BASE64_LENGTH) {
    return { status: 'error', message: 'Ukuran foto nota terlalu besar atau kosong. Coba ambil foto ulang.' };
  }

  const categories = getCategories_().map(c => c.name).filter(c => c !== 'Transfer Antar Bank');
  const catListStr = categories.join(', ');

  const prompt =
    'Kamu adalah asisten pencatatan keuangan pribadi dan pembaca struk/nota belanja (OCR) presisi tinggi. ' +
    'Analisis foto nota belanja ini dan ekstrak informasi ke dalam format JSON sesuai skema:\n' +
    '1. is_receipt: true jika dokumen adalah struk/nota/bon belanja/invoice pembayaran. Jika bukan, false.\n' +
    '2. merchant: Nama toko/supermarket/restoran/tempat pembelian.\n' +
    '3. date: Tanggal transaksi nota dalam format YYYY-MM-DD. Jika tahun tidak tertera, gunakan tahun 2026.\n' +
    '4. total_amount: Nominal total pembayaran akhir (Grand Total / Total Tagihan) dalam bentuk angka bulat positif murni tanpa titik/koma.\n' +
    '5. category: Pilih SATU kategori yang paling relevan dari daftar berikut: [' + catListStr + ']. Jika belanja makan/minuman pilih Makan, jika kebutuhan rumah/supermarket pilih Belanja, bensin/parkir pilih Transport, obat/klinik pilih Kesehatan, nonton/games pilih Hiburan, dll.\n' +
    '6. items: Daftar rincian barang yang dibeli (nama barang, jumlah/qty jika ada, harga per barang atau total item).\n' +
    '7. note: Ringkasan singkat 1 kalimat mengenai isi nota.';

  const schema = {
    type: 'OBJECT',
    properties: {
      is_receipt: { type: 'BOOLEAN' },
      merchant: { type: 'STRING' },
      date: { type: 'STRING' },
      category: { type: 'STRING' },
      total_amount: { type: 'NUMBER' },
      items: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            name: { type: 'STRING' },
            qty: { type: 'STRING' },
            price: { type: 'NUMBER' }
          },
          required: ['name', 'price']
        }
      },
      note: { type: 'STRING' }
    },
    required: ['is_receipt', 'total_amount']
  };

  const body = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64Data } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 1024,
      responseMimeType: 'application/json',
      responseSchema: schema,
      thinkingConfig: { thinkingBudget: 0 }
    }
  };

  try {
    const result = executeGeminiFallback_(body, apiKey);
    if (!result || result.code !== 200) {
      const msg = (result && result.json && result.json.error && result.json.error.message) ||
                  ('Gemini API error (HTTP ' + (result ? result.code : '500') + ')');
      return { status: 'error', message: msg };
    }

    const cand = result.json.candidates && result.json.candidates[0];
    if (!cand || !cand.content || !cand.content.parts || !cand.content.parts[0]) {
      return { status: 'error', message: 'AI tidak memberikan respons valid.' };
    }

    const parsed = JSON.parse(cand.content.parts[0].text);
    if (!parsed.is_receipt && (!parsed.total_amount || parsed.total_amount <= 0)) {
      return { status: 'not_receipt', message: 'Foto tidak terdeteksi sebagai nota belanja yang jelas. Pastikan foto fokus dan total nominal terlihat.' };
    }

    return {
      status: 'success',
      merchant: parsed.merchant || 'Nota Belanja',
      date: parsed.date || fmtDate_(new Date()),
      category: parsed.category || 'Belanja',
      total_amount: Math.round(parsed.total_amount || 0),
      items: Array.isArray(parsed.items) ? parsed.items : [],
      note: parsed.note || ''
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal memproses hasil AI: ' + String(err) };
  }
}

function authorizeUrlFetchApp() {
  // Jalankan fungsi ini sekali dari Editor Apps Script jika ingin memberikan izin UrlFetchApp ke backend
  return UrlFetchApp.fetch('https://www.google.com').getResponseCode();
}
