/* ============================================
   Finance Purchasing — PT ALMA ATA LIFE
   App Logic & Google Sheets API Integration
   ============================================ */

// --- Configuration ---
const CONFIG = {
  SPREADSHEET_ID: '1NcCOiL8Hs5xM9QnPRzLaiXj9SajKKYQfrTAwvJ17R6Q',
  SHEETS: {
    TAGIHAN: 'Tagihan Masuk',
    DETAIL_PO: 'Detail PO View',
    FORM: 'Form Pembayaran',
    LOG_PEMBELIAN: 'Log Pembelian',
    LOG_PEMBAYARAN: 'Log Pembayaran',
    REKAP: 'Rekap Hutang Vendor'
  },
  API_KEY: '',
  CLIENT_ID: '1021440664537-9ic5rv15tudiu69e79av1p7g55u9dj33.apps.googleusercontent.com',
  SCOPES: 'https://www.googleapis.com/auth/spreadsheets',
  DISCOVERY_DOC: 'https://sheets.googleapis.com/$discovery/rest?version=v4',
};

// --- State ---
let tokenClient = null;
let gapiInited = false;
let gisInited = false;
let allData = {
  tagihan: [],
  detailPO: [],
  logPembelian: [],
  logPembayaran: [],
  rekap: []
};
let currentPOData = null;

// --- Initialize Google API ---
function gapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({
      apiKey: CONFIG.API_KEY,
      discoveryDocs: [CONFIG.DISCOVERY_DOC],
    });
    gapiInited = true;
    maybeEnableAuth();
  });
}

function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.CLIENT_ID,
    scope: CONFIG.SCOPES,
    callback: '',
  });
  gisInited = true;
  maybeEnableAuth();
}

function maybeEnableAuth() {
  if (gapiInited && gisInited) {
    document.getElementById('loadingSpinner').style.display = 'flex';
    handleAuthClick();
  }
}

function handleAuthClick() {
  tokenClient.callback = async (resp) => {
    if (resp.error) {
      showToast('Autentikasi gagal. Silakan coba lagi.', 'error');
      document.getElementById('loadingSpinner').style.display = 'none';
      return;
    }
    await loadAllData();
  };

  if (gapi.client.getToken() === null) {
    tokenClient.requestAccessToken({ prompt: 'consent' });
  } else {
    tokenClient.requestAccessToken({ prompt: '' });
  }
}

// --- Fetch Data from Sheets ---
async function fetchSheet(sheetName, range) {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: `${sheetName}!${range}`,
    });
    return response.result.values || [];
  } catch (err) {
    console.error(`Error fetching ${sheetName}:`, err);
    return [];
  }
}

async function appendToSheet(sheetName, values) {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: `${sheetName}!A1`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
    }, {
      values: [values]
    });
    return response.result;
  } catch (err) {
    console.error(`Error appending to ${sheetName}:`, err);
    throw err;
  }
}

async function updateSheetCell(sheetName, range, value) {
  try {
    const response = await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: `${sheetName}!${range}`,
      valueInputOption: 'USER_ENTERED',
    }, {
      values: [[value]]
    });
    return response.result;
  } catch (err) {
    console.error(`Error updating ${sheetName}:`, err);
    throw err;
  }
}

async function updateSheetRow(sheetName, rowIndex, headers, values) {
  try {
    const lastCol = String.fromCharCode(64 + headers.length);
    const range = `${sheetName}!A${rowIndex}:${lastCol}${rowIndex}`;
    const response = await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: range,
      valueInputOption: 'USER_ENTERED',
    }, { values: [values] });
    return response.result;
  } catch (err) {
    console.error(`Error updating row in ${sheetName}:`, err);
    throw err;
  }
}

async function deleteSheetRow(sheetName, rowIndex, colCount) {
  try {
    const lastCol = String.fromCharCode(64 + colCount);
    const range = `${sheetName}!A${rowIndex}:${lastCol}${rowIndex}`;
    const emptyValues = Array(colCount).fill('');
    const response = await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range: range,
      valueInputOption: 'USER_ENTERED',
    }, { values: [emptyValues] });
    return response.result;
  } catch (err) {
    console.error(`Error deleting row in ${sheetName}:`, err);
    throw err;
  }
}

async function loadAllData() {
  showLoading(true);
  try {
    const [tagihan, detailPO, logPembelian, logPembayaran, rekap] = await Promise.all([
      fetchSheet(CONFIG.SHEETS.TAGIHAN, 'A1:Z500'),
      fetchSheet(CONFIG.SHEETS.DETAIL_PO, 'A1:I500'),
      fetchSheet(CONFIG.SHEETS.LOG_PEMBELIAN, 'A1:M500'),
      fetchSheet(CONFIG.SHEETS.LOG_PEMBAYARAN, 'A1:L500'),
      fetchSheet(CONFIG.SHEETS.REKAP, 'A1:J500'),
    ]);

    allData.tagihan = parseRows(tagihan);
    allData.detailPO = parseRows(detailPO);
    allData.logPembelian = parseRows(logPembelian);
    allData.logPembayaran = parseRows(logPembayaran);
    allData.rekap = parseRows(rekap);

    renderDashboard();
    renderTagihan();
    renderDetailPO();
    renderLogPembelian();
    renderLogPembayaran();
    renderRekap();

    showToast('Data berhasil dimuat', 'success');
  } catch (err) {
    console.error('Error loading data:', err);
    showToast('Gagal memuat data', 'error');
  } finally {
    showLoading(false);
  }
}

function parseRows(rawData) {
  if (rawData.length < 2) return [];
  const headers = rawData[0];
  return rawData.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => {
      obj[h] = row[j] || '';
    });
    obj._rowIndex = i + 2; // sheet row (1-based, +1 for header)
    obj._headers = headers;
    return obj;
  });
}

// --- Refresh ---
async function refreshData() {
  await loadAllData();
}

// --- Render Dashboard ---
function renderDashboard() {
  const rekap = allData.rekap;
  const tagihan = allData.tagihan;

  const totalPO = rekap.length;
  const lunas = rekap.filter(r => r['Status Pembayaran'] === 'Lunas').length;
  const belumLunas = totalPO - lunas;

  const totalHutang = rekap.reduce((sum, r) => {
    const sisa = parseCurrency(r['Sisa Tagihan']);
    return sum + (sisa > 0 ? sisa : 0);
  }, 0);

  document.getElementById('stat-total-po').textContent = totalPO;
  document.getElementById('stat-menunggu').textContent = belumLunas;
  document.getElementById('stat-lunas').textContent = lunas;
  document.getElementById('stat-hutang').textContent = formatCurrency(totalHutang);

  const recent = tagihan.slice(-10).reverse();
  const tbody = document.getElementById('dashboard-recent');

  if (recent.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = recent.map(r => `
    <tr>
      <td>${r['Tanggal'] || '—'}</td>
      <td><strong>${r['PO Number'] || '—'}</strong></td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Total Harga'] || '—'}</td>
      <td>${statusBadge(r['Status Pembayaran'])}</td>
    </tr>
  `).join('');
}

// --- Render Tagihan Masuk ---
function renderTagihan() {
  const tbody = document.getElementById('tagihan-body');
  const data = allData.tagihan;

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td>${r['Tanggal'] || '—'}</td>
      <td><strong>${r['PO Number'] || '—'}</strong></td>
      <td>${r['Unit'] || '—'}</td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Total Harga'] || '—'}</td>
      <td>${r['Total Dibayar'] || '—'}</td>
      <td>${r['Sisa Tagihan'] || '—'}</td>
      <td>${statusBadge(r['Status Pembayaran'])}</td>
    </tr>
  `).join('');
}

// --- Render Detail PO ---
function renderDetailPO() {
  const tbody = document.getElementById('detail-po-body');
  const data = allData.detailPO;

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td><strong>${r['PO Number'] || '—'}</strong></td>
      <td>${r['Tanggal'] || '—'}</td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Kode Barang'] || '—'}</td>
      <td>${r['Deskripsi'] || '—'}</td>
      <td>${r['Qty'] || '—'}</td>
      <td>${r['Satuan'] || '—'}</td>
      <td>${r['Harga Satuan'] || '—'}</td>
      <td><strong>${r['Total Item'] || r['Total'] || '—'}</strong></td>
    </tr>
  `).join('');
}

// --- Render Log Pembelian ---
function renderLogPembelian() {
  const tbody = document.getElementById('log-pembelian-body');
  const data = allData.logPembelian;

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td>${r['Tanggal Pembelian'] || '—'}</td>
      <td><strong>${r['Nomor PO'] || '—'}</strong></td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Metode'] || '—'}</td>
      <td>${r['Total PO'] || '—'}</td>
      <td>${r['Jumlah'] || '—'}</td>
      <td>${r['Sisa Tagihan'] || '—'}</td>
      <td>${r['Keterangan'] || '—'}</td>
      <td>${statusBadge(r['Status'])}</td>
      <td class="action-cell">
        <button class="btn-icon btn-edit" onclick="openEditModal('logPembelian', ${r._rowIndex})" title="Edit">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        </button>
        <button class="btn-icon btn-delete" onclick="openDeleteModal('logPembelian', ${r._rowIndex}, '${r['Nomor PO']}')" title="Hapus">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        </button>
      </td>
    </tr>
  `).join('');
}

// --- Render Log Pembayaran ---
function renderLogPembayaran() {
  const tbody = document.getElementById('log-pembayaran-body');
  const data = allData.logPembayaran;

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td>${r['Tanggal Bayar'] || '—'}</td>
      <td><strong>${r['Nomor PO'] || '—'}</strong></td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Metode'] || '—'}</td>
      <td>${r['No Bukti'] || '—'}</td>
      <td>${r['Total PO'] || '—'}</td>
      <td>${r['Jumlah Bayar'] || '—'}</td>
      <td>${r['Sisa Tagihan'] || '—'}</td>
      <td>${r['Keterangan'] || '—'}</td>
      <td class="action-cell">
        <button class="btn-icon btn-edit" onclick="openEditModal('logPembayaran', ${r._rowIndex})" title="Edit">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
        </button>
        <button class="btn-icon btn-delete" onclick="openDeleteModal('logPembayaran', ${r._rowIndex}, '${r['Nomor PO']}')" title="Hapus">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        </button>
      </td>
    </tr>
  `).join('');
}

// --- Render Rekap Hutang ---
function renderRekap() {
  const tbody = document.getElementById('rekap-body');
  const data = allData.rekap;

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">Belum ada data</td></tr>';
    return;
  }

  tbody.innerHTML = data.map(r => `
    <tr>
      <td>${r['Tanggal'] || '—'}</td>
      <td><strong>${r['PO Number'] || '—'}</strong></td>
      <td>${r['Unit'] || '—'}</td>
      <td>${r['Vendor'] || '—'}</td>
      <td>${r['Total Harga'] || '—'}</td>
      <td>${r['Total Dibayar'] || '—'}</td>
      <td>${r['Sisa Tagihan'] || '—'}</td>
      <td>${statusBadge(r['Status Pembayaran'])}</td>
      <td class="deskripsi-cell">${(r['Deskripsi Barang'] || '—').replace(/\n/g, '<br>')}</td>
    </tr>
  `).join('');
}

// --- Generic Table Filter ---
function filterTable(bodyId, searchId) {
  const search = document.getElementById(searchId).value.toLowerCase();
  const rows = document.querySelectorAll(`#${bodyId} tr`);
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(search) ? '' : 'none';
  });
}

// --- Form Pembayaran Logic ---
let currentPOStatus = null; // 'lunas' or 'belum'

async function loadPOData() {
  let poNumber = document.getElementById('input-po').value.trim();
  if (!poNumber) {
    showToast('Masukkan Nomor PO terlebih dahulu', 'error');
    return;
  }

  showLoading(true);

  const tagihan = allData.tagihan.find(r => String(r['PO Number']).trim() === poNumber);
  if (!tagihan) {
    showToast(`PO ${poNumber} tidak ditemukan di Tagihan Masuk`, 'error');
    showLoading(false);
    return;
  }

  currentPOData = tagihan;

  // Fill form - left side (info card)
  document.getElementById('form-tanggal').textContent = tagihan['Tanggal'] || '—';
  document.getElementById('form-unit').textContent = tagihan['Unit'] || '—';
  document.getElementById('form-vendor').textContent = tagihan['Vendor'] || '—';
  document.getElementById('form-alamat').textContent = tagihan['Alamat Vendor'] || '—';
  document.getElementById('form-total').textContent = tagihan['Total Harga'] || '—';

  // Load detail items - trim PO number for matching
  const details = allData.detailPO.filter(r => String(r['PO Number']).trim() === poNumber);
  renderDetailItems(details);

  // Load payment history from Log Pembayaran - trim for matching
  const history = allData.logPembayaran.filter(r => String(r['Nomor PO']).trim() === poNumber);
  renderPaymentHistory(history);

  // Calculate summary
  let totalDibayar = 0;
  history.forEach(r => {
    totalDibayar += parseCurrency(r['Jumlah Bayar']);
  });

  const totalPO = parseCurrency(tagihan['Total Harga']);
  const sisa = totalPO - totalDibayar;

  document.getElementById('summary-total-po').textContent = formatCurrency(totalPO);
  document.getElementById('summary-dibayar').textContent = formatCurrency(totalDibayar);
  document.getElementById('summary-sisa').textContent = formatCurrency(sisa > 0 ? sisa : 0);

  // Check status
  const statusBanner = document.getElementById('status-banner');
  const btnSubmit = document.getElementById('btn-submit');
  const btnRealize = document.getElementById('btn-realize');
  const btnPDF = document.getElementById('btn-pdf');

  if (sisa <= 0 || tagihan['Status Pembayaran'] === 'Lunas') {
    // LUNAS
    currentPOStatus = 'lunas';
    statusBanner.className = 'status-banner lunas';
    statusBanner.innerHTML = `
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
      PO ini sudah LUNAS — Tidak bisa generate pembayaran lagi
    `;
    statusBanner.style.display = 'flex';

    // Disable buttons
    btnSubmit.disabled = true;
    btnSubmit.style.opacity = '0.5';
    btnSubmit.style.cursor = 'not-allowed';
    btnRealize.disabled = true;
    btnRealize.style.opacity = '0.5';
    btnRealize.style.cursor = 'not-allowed';

    // Jumlah = 0
    document.getElementById('form-jumlah').value = 'Rp 0';
  } else {
    // BELUM LUNAS
    currentPOStatus = 'belum';
    statusBanner.className = 'status-banner belum';
    statusBanner.innerHTML = `
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      Sisa tagihan: ${formatCurrency(sisa)} — Silakan input pembayaran
    `;
    statusBanner.style.display = 'flex';

    // Enable buttons
    btnSubmit.disabled = false;
    btnSubmit.style.opacity = '1';
    btnSubmit.style.cursor = 'pointer';
    btnRealize.disabled = false;
    btnRealize.style.opacity = '1';
    btnRealize.style.cursor = 'pointer';

    document.getElementById('form-jumlah').value = formatCurrency(sisa > 0 ? sisa : 0);
  }

  document.getElementById('form-summary').style.display = 'block';
  document.getElementById('form-content-loaded').style.display = 'block';
  
  // Update steps indicator
  document.getElementById('step-1').classList.add('completed');
  document.getElementById('step-1').classList.remove('active');
  document.getElementById('step-2').classList.add('active');
  document.getElementById('step-3').classList.add('active');

  showLoading(false);
  showToast(`PO ${poNumber} berhasil dimuat`, 'success');
}

function renderDetailItems(details) {
  const card = document.getElementById('detail-items-card');
  const tbody = document.getElementById('detail-items-body');
  const tfoot = document.getElementById('detail-items-footer');

  if (details.length === 0) {
    card.style.display = 'none';
    return;
  }

  card.style.display = 'block';

  let totalAll = 0;
  tbody.innerHTML = details.map(r => {
    const total = parseCurrency(r['Total Item'] || r['Total'] || '0');
    totalAll += total;
    return `
      <tr>
        <td>${r['Kode Barang'] || ''}</td>
        <td>${r['Deskripsi'] || ''}</td>
        <td class="text-center">${r['Qty'] || ''}</td>
        <td>${r['Satuan'] || ''}</td>
        <td class="text-right">${r['Harga Satuan'] || ''}</td>
        <td class="text-right">${r['Total Item'] || r['Total'] || ''}</td>
      </tr>
    `;
  }).join('');

  tfoot.innerHTML = `
    <tr>
      <td colspan="5" style="text-align:right; font-weight:600; color:var(--text-secondary);">Total</td>
      <td><strong>${formatCurrency(totalAll)}</strong></td>
    </tr>
  `;
}

function renderPaymentHistory(history) {
  const card = document.getElementById('history-card');
  const tbody = document.getElementById('history-body');
  const tfoot = document.getElementById('history-footer');

  if (history.length === 0) {
    card.style.display = 'none';
    return;
  }

  card.style.display = 'block';

  let totalBayar = 0;
  tbody.innerHTML = history.map(r => {
    const jumlah = parseCurrency(r['Jumlah Bayar']);
    totalBayar += jumlah;
    return `
      <tr>
        <td>${r['Tanggal Bayar'] || '—'}</td>
        <td>${r['Keterangan'] || '—'}</td>
        <td>${r['No Bukti'] || '—'}</td>
        <td><strong>${r['Jumlah Bayar'] || '—'}</strong></td>
      </tr>
    `;
  }).join('');

  tfoot.innerHTML = `
    <tr>
      <td colspan="3" style="text-align:right; font-weight:600; color:var(--text-secondary);">Total Pembayaran</td>
      <td><strong>${formatCurrency(totalBayar)}</strong></td>
    </tr>
  `;
}

// --- Submit to Log Pembelian ---
async function submitToLogPembelian() {
  // Check if LUNAS
  if (currentPOStatus === 'lunas') {
    showToast('PO ini sudah LUNAS — Tidak bisa generate pembayaran lagi', 'error');
    return;
  }

  const poNumber = document.getElementById('input-po').value.trim();
  const metode = document.getElementById('form-metode').value;
  const rekening = document.getElementById('form-rekening').value;
  const bukti = document.getElementById('form-bukti').value;
  const jumlah = document.getElementById('form-jumlah').value;
  const keterangan = document.getElementById('form-keterangan').value;
  const tglBayar = document.getElementById('form-tgl-bayar').value;

  if (!poNumber) {
    showToast('Masukkan Nomor PO terlebih dahulu', 'error');
    return;
  }

  if (!jumlah || parseCurrency(jumlah) <= 0) {
    showToast('Jumlah pembayaran harus diisi', 'error');
    return;
  }

  if (!metode) {
    showToast('Pilih metode pembayaran', 'error');
    return;
  }

  showLoading(true);

  try {
    const totalPO = currentPOData ? parseCurrency(currentPOData['Total Harga']) : 0;
    const jumlahNum = parseCurrency(jumlah);
    const sisa = totalPO - jumlahNum;

    const rowData = [
      tglBayar || new Date().toLocaleDateString('en-GB'),
      poNumber,
      currentPOData ? currentPOData['Vendor'] : '',
      metode,
      rekening,
      bukti,
      currentPOData ? currentPOData['Total Harga'] : '',
      formatCurrency(jumlahNum),
      formatCurrency(sisa > 0 ? sisa : 0),
      keterangan,
      'Draft'
    ];

    await appendToSheet(CONFIG.SHEETS.LOG_PEMBELIAN, rowData);
    showToast('Berhasil disimpan ke Log Pembelian (Draft)', 'success');

    // Clear form
    clearForm();
    await loadAllData();
  } catch (err) {
    showToast('Gagal menyimpan: ' + err.message, 'error');
  } finally {
    showLoading(false);
  }
}

// --- Realize Payment ---
async function realizePayment() {
  // Check if LUNAS
  if (currentPOStatus === 'lunas') {
    showToast('PO ini sudah LUNAS — Tidak bisa generate pembayaran lagi', 'error');
    return;
  }

  const poNumber = document.getElementById('input-po').value.trim();
  const metode = document.getElementById('form-metode').value;
  const rekening = document.getElementById('form-rekening').value;
  const bukti = document.getElementById('form-bukti').value;
  const jumlah = document.getElementById('form-jumlah').value;
  const keterangan = document.getElementById('form-keterangan').value;
  const tglBayar = document.getElementById('form-tgl-bayar').value;

  if (!poNumber) {
    showToast('Masukkan Nomor PO terlebih dahulu', 'error');
    return;
  }

  if (!jumlah || parseCurrency(jumlah) <= 0) {
    showToast('Jumlah pembayaran harus diisi', 'error');
    return;
  }

  if (!metode) {
    showToast('Pilih metode pembayaran', 'error');
    return;
  }

  showLoading(true);

  try {
    const totalPO = currentPOData ? parseCurrency(currentPOData['Total Harga']) : 0;
    const jumlahNum = parseCurrency(jumlah);
    const sisa = totalPO - jumlahNum;

    // 1. Add to Log Pembayaran
    const paymentRow = [
      tglBayar || new Date().toLocaleDateString('en-GB'),
      poNumber,
      currentPOData ? currentPOData['Vendor'] : '',
      metode,
      rekening || 'KAS',
      bukti,
      currentPOData ? currentPOData['Total Harga'] : '',
      formatCurrency(jumlahNum),
      formatCurrency(sisa > 0 ? sisa : 0),
      keterangan
    ];

    await appendToSheet(CONFIG.SHEETS.LOG_PEMBAYARAN, paymentRow);

    // 2. Add to Log Pembelian as Realisasi
    const purchaseRow = [
      tglBayar || new Date().toLocaleDateString('en-GB'),
      poNumber,
      currentPOData ? currentPOData['Vendor'] : '',
      metode,
      rekening,
      bukti,
      currentPOData ? currentPOData['Total Harga'] : '',
      formatCurrency(jumlahNum),
      formatCurrency(sisa > 0 ? sisa : 0),
      keterangan,
      'Realisasi'
    ];

    await appendToSheet(CONFIG.SHEETS.LOG_PEMBELIAN, purchaseRow);

    showToast('Pembayaran berhasil direalisasi!', 'success');

    clearForm();
    await loadAllData();
  } catch (err) {
    showToast('Gagal merealisasi: ' + err.message, 'error');
  } finally {
    showLoading(false);
  }
}

// --- Clear Form ---
function clearForm() {
  document.getElementById('input-po').value = '';
  document.getElementById('form-tanggal').textContent = '—';
  document.getElementById('form-unit').textContent = '—';
  document.getElementById('form-vendor').textContent = '—';
  document.getElementById('form-alamat').textContent = '—';
  document.getElementById('form-total').textContent = '—';
  document.getElementById('form-tgl-bayar').value = '';
  document.getElementById('form-metode').value = '';
  document.getElementById('form-rekening').value = '';
  document.getElementById('form-bukti').value = '';
  document.getElementById('form-jumlah').value = '';
  document.getElementById('form-keterangan').value = '';
  document.getElementById('form-summary').style.display = 'none';
  document.getElementById('form-content-loaded').style.display = 'none';
  document.getElementById('detail-items-card').style.display = 'none';
  document.getElementById('history-card').style.display = 'none';
  
  // Reset steps
  document.getElementById('step-1').classList.remove('completed');
  document.getElementById('step-1').classList.add('active');
  document.getElementById('step-2').classList.remove('active');
  document.getElementById('step-3').classList.remove('active');
  
  currentPOData = null;
}

// --- PDF Generation ---
function generatePDF() {
  const poNumber = document.getElementById('input-po').value.trim();
  if (!poNumber) {
    showToast('Masukkan Nomor PO terlebih dahulu', 'error');
    return;
  }

  const vendor = document.getElementById('form-vendor').value;
  const unit = document.getElementById('form-unit').value;
  const tanggal = document.getElementById('form-tanggal').value;
  const alamat = document.getElementById('form-alamat').value;
  const totalTagihan = document.getElementById('form-total').value;
  const tglBayar = document.getElementById('form-tgl-bayar').value;
  const metode = document.getElementById('form-metode').value;
  const rekening = document.getElementById('form-rekening').value;
  const bukti = document.getElementById('form-bukti').value;
  const jumlah = document.getElementById('form-jumlah').value;
  const keterangan = document.getElementById('form-keterangan').value;

  const isLunas = currentPOStatus === 'lunas';

  const details = allData.detailPO.filter(r => String(r['PO Number']).trim() === poNumber);
  let totalAll = 0;
  let detailRows = details.map(r => {
    const total = parseCurrency(r['Total Item'] || r['Total'] || '0');
    totalAll += total;
    return `
      <tr>
        <td>${r['Kode Barang'] || ''}</td>
        <td>${r['Deskripsi'] || ''}</td>
        <td class="text-center">${r['Qty'] || ''}</td>
        <td>${r['Satuan'] || ''}</td>
        <td class="text-right">${r['Harga Satuan'] || ''}</td>
        <td class="text-right">${r['Total Item'] || r['Total'] || ''}</td>
      </tr>
    `;
  }).join('');
  const history = allData.logPembayaran.filter(r => String(r['Nomor PO']).trim() === poNumber);
  let historyRows = history.map(r => `
    <tr>
      <td>${r['Tanggal Bayar'] || ''}</td>
      <td>${r['Keterangan'] || ''}</td>
      <td class="text-right">${r['Jumlah Bayar'] || ''}</td>
    </tr>
  `).join('');

  const printContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Form Pembayaran - ${poNumber}</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #000; padding: 30px 40px; line-height: 1.5; }
        
        /* Header */
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #000; }
        .header-left { font-size: 16px; font-weight: bold; }
        .header-right { font-size: 14px; font-weight: bold; }
        .lunas-badge { font-size: 11px; font-weight: 900; color: rgba(180,50,50,0.8); border: 2px solid rgba(180,50,50,0.6); padding: 2px 8px; display: inline-block; margin-left: 10px; }
        
        /* Info Boxes */
        .info-boxes { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 16px; }
        .info-box { border: 1px solid #000; padding: 10px; }
        .info-box-title { font-weight: bold; font-size: 10px; text-transform: uppercase; margin-bottom: 6px; border-bottom: 1px solid #ccc; padding-bottom: 4px; }
        .info-row { display: flex; margin-bottom: 3px; font-size: 10.5px; }
        .info-label { width: 110px; flex-shrink: 0; }
        .info-value { flex: 1; }
        .po-highlight { background: #fff3cd; padding: 1px 4px; font-weight: bold; }
        .info-bold { font-weight: bold; }
        
        /* Tables */
        table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 10px; }
        th, td { border: 1px solid #000; padding: 5px 6px; text-align: left; }
        th { background: #f5f5f5; font-weight: bold; font-size: 9px; text-transform: uppercase; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .total-row { font-weight: bold; background: #f9f9f9; }
        
        /* Signatures */
        .signatures { display: flex; justify-content: space-between; margin-top: 50px; }
        .sig-block { text-align: center; width: 45%; }
        .sig-title { font-weight: bold; margin-bottom: 60px; font-size: 10px; }
        .sig-name { font-weight: bold; font-size: 10px; border-top: 1px solid #000; padding-top: 4px; display: inline-block; min-width: 180px; }
        .sig-role { font-size: 9px; color: #333; margin-top: 2px; }
        
        .no-data { font-style: italic; color: #666; font-size: 10px; padding: 10px; }
        
        @media print {
          @page { size: A4 portrait; margin: 15mm; }
          body { padding: 0; }
        }
      </style>
    </head>
    <body>
      <!-- Header -->
      <div class="header">
        <div class="header-left">
          FORM PAYMENT
          ${isLunas ? '<span class="lunas-badge">✓ SUDAH LUNAS</span>' : ''}
        </div>
        <div class="header-right">PT ALMA ATA LIFE</div>
      </div>

      <!-- Info Boxes -->
      <div class="info-boxes">
        <div class="info-box">
          <div class="info-row"><div class="info-label">NO PO</div><div class="info-value">: <span class="po-highlight">${poNumber}</span></div></div>
          <div class="info-row"><div class="info-label">Tanggal</div><div class="info-value">: ${tanggal}</div></div>
          <div class="info-row"><div class="info-label">Vendor</div><div class="info-value">: ${vendor}</div></div>
          <div class="info-row"><div class="info-label">Alamat</div><div class="info-value">: ${alamat}</div></div>
          <div class="info-row"><div class="info-label">Unit</div><div class="info-value">: ${unit}</div></div>
          <div class="info-row"><div class="info-label">Tagihan</div><div class="info-value info-bold">: ${totalTagihan}</div></div>
        </div>
        <div class="info-box">
          <div class="info-row"><div class="info-label">Tanggal Pembayaran</div><div class="info-value">: ${isLunas ? '-' : (tglBayar || '-')}</div></div>
          <div class="info-row"><div class="info-label">Metode</div><div class="info-value">: ${isLunas ? '-' : (metode || '-')}</div></div>
          <div class="info-row"><div class="info-label">Rek. BCA</div><div class="info-value">: ${isLunas ? '-' : (rekening || '-')}</div></div>
          <div class="info-row"><div class="info-label">No Bukti</div><div class="info-value">: ${isLunas ? '-' : (bukti || '-')}</div></div>
          <div class="info-row"><div class="info-label">Jumlah Pembayaran</div><div class="info-value info-bold">: ${isLunas ? 'SUDAH LUNAS' : jumlah}</div></div>
          <div class="info-row"><div class="info-label">Keterangan</div><div class="info-value">: ${isLunas ? '-' : (keterangan || '-')}</div></div>
        </div>
      </div>

      <!-- Detail Barang -->
      <div style="font-weight:bold;font-size:10px;margin-bottom:4px;">DETAIL BARANG</div>
      ${details.length > 0 ? `
      <table>
        <thead>
          <tr>
            <th style="width:50px">Kode Barang</th>
            <th>Deskripsi Barang</th>
            <th style="width:40px" class="text-center">Qty</th>
            <th style="width:50px">Satuan</th>
            <th style="width:90px" class="text-right">Harga Satuan</th>
            <th style="width:90px" class="text-right">Total</th>
          </tr>
        </thead>
        <tbody>${detailRows}</tbody>
      </table>
      ` : '<p class="no-data">Detail barang tidak tersedia</p>'}

      <!-- Riwayat Pembayaran -->
      <div style="font-weight:bold;font-size:10px;margin:10px 0 4px;">RIWAYAT PEMBAYARAN</div>
      ${history.length > 0 ? `
      <table>
        <thead>
          <tr>
            <th style="width:80px">History Bayar</th>
            <th style="width:80px">Keterangan</th>
            <th class="text-right">Jumlah Bayar</th>
          </tr>
        </thead>
        <tbody>
          ${historyRows}
          <tr class="total-row">
            <td colspan="2" class="text-right">TOTAL PEMBAYARAN</td>
            <td class="text-right">${document.getElementById('summary-dibayar').textContent}</td>
          </tr>
          ${!isLunas && jumlah && parseCurrency(jumlah) > 0 ? `
          <tr>
            <td colspan="2" class="text-right" style="font-style:italic">Pelunasan Pembayaran (${keterangan || '-'})</td>
            <td class="text-right info-bold">${jumlah}</td>
          </tr>
          ` : ''}
        </tbody>
      </table>
      ` : '<p class="no-data">Belum ada riwayat pembayaran</p>'}

      <!-- Tanda Tangan -->
      <div class="signatures">
        <div class="sig-block">
          <div class="sig-title">Mengetahui dan Menyetujui</div>
          <div class="sig-name">Prof. Dr. H. Hamam Hadi, M.S., Sc.D., Sp.GK</div>
          <div class="sig-role">Pimpinan PT AA LIFE</div>
        </div>
        <div class="sig-block">
          <div class="sig-title">&nbsp;</div>
          <div class="sig-name">Hasrul Eko Marsanto</div>
          <div class="sig-role">Finance PT AA LIFE</div>
        </div>
      </div>
    </body>
    </html>
  `;

  const printWindow = window.open('', '_blank');
  printWindow.document.write(printContent);
  printWindow.document.close();
  printWindow.focus();
  
  setTimeout(() => {
    printWindow.print();
  }, 500);
}

// --- Edit & Delete ---
const EDIT_FIELDS = {
  logPembelian: [
    { key: 'Tanggal Pembelian', label: 'Tanggal', type: 'date' },
    { key: 'Nomor PO', label: 'Nomor PO', type: 'text' },
    { key: 'Vendor', label: 'Vendor', type: 'text' },
    { key: 'Metode', label: 'Metode', type: 'select', options: ['TRANSFER', 'CASH', 'TRANSFER BANK'] },
    { key: 'Total PO', label: 'Total PO', type: 'currency' },
    { key: 'Jumlah', label: 'Jumlah', type: 'currency' },
    { key: 'Sisa Tagihan', label: 'Sisa Tagihan', type: 'currency' },
    { key: 'Keterangan', label: 'Keterangan', type: 'select', options: ['DP', 'TR', 'FL'] },
    { key: 'Status', label: 'Status', type: 'select', options: ['Draft', 'Realisasi'] },
  ],
  logPembayaran: [
    { key: 'Tanggal Bayar', label: 'Tanggal Bayar', type: 'date' },
    { key: 'Nomor PO', label: 'Nomor PO', type: 'text' },
    { key: 'Vendor', label: 'Vendor', type: 'text' },
    { key: 'Metode', label: 'Metode', type: 'select', options: ['TRANSFER', 'CASH', 'TRANSFER BANK'] },
    { key: 'No Bukti', label: 'No Bukti', type: 'text' },
    { key: 'Total PO', label: 'Total PO', type: 'currency' },
    { key: 'Jumlah Bayar', label: 'Jumlah Bayar', type: 'currency' },
    { key: 'Sisa Tagihan', label: 'Sisa Tagihan', type: 'currency' },
    { key: 'Keterangan', label: 'Keterangan', type: 'select', options: ['DP', 'TR', 'FL'] },
  ],
};

let editingSheet = null;
let editingRow = null;
let deletingSheet = null;
let deletingRow = null;

function openEditModal(sheetKey, rowIndex) {
  editingSheet = sheetKey;
  editingRow = rowIndex;

  const data = allData[sheetKey === 'logPembelian' ? 'logPembelian' : 'logPembayaran'];
  const row = data.find(r => r._rowIndex === rowIndex);
  if (!row) { showToast('Data tidak ditemukan', 'error'); return; }

  const fields = EDIT_FIELDS[sheetKey];
  const body = document.getElementById('modal-body');
  const sheetLabel = sheetKey === 'logPembelian' ? 'Log Pembelian' : 'Log Pembayaran';
  document.getElementById('modal-title').textContent = `Edit — ${sheetLabel}`;

  body.innerHTML = '<div class="form-grid">' + fields.map(f => {
    const val = row[f.key] || '';
    if (f.type === 'select') {
      return `<div class="form-group">
        <label class="form-label">${f.label}</label>
        <select class="form-input" data-key="${f.key}">
          <option value="">Pilih...</option>
          ${f.options.map(o => `<option value="${o}" ${val === o ? 'selected' : ''}>${o}</option>`).join('')}
        </select>
      </div>`;
    }
    return `<div class="form-group">
      <label class="form-label">${f.label}</label>
      <input type="${f.type === 'currency' ? 'text' : f.type}" class="form-input" data-key="${f.key}" value="${val}">
    </div>`;
  }).join('') + '</div>';

  document.getElementById('editModal').style.display = 'flex';
}

function closeEditModal() {
  document.getElementById('editModal').style.display = 'none';
  editingSheet = null;
  editingRow = null;
}

async function saveEdit() {
  if (!editingSheet || !editingRow) return;
  showLoading(true);

  try {
    const fields = EDIT_FIELDS[editingSheet];
    const sheetName = editingSheet === 'logPembelian' ? CONFIG.SHEETS.LOG_PEMBELIAN : CONFIG.SHEETS.LOG_PEMBAYARAN;
    const dataKey = editingSheet === 'logPembelian' ? 'logPembelian' : 'logPembayaran';
    const row = allData[dataKey].find(r => r._rowIndex === editingRow);
    if (!row) throw new Error('Data tidak ditemukan');

    const headers = row._headers;
    const values = headers.map(h => {
      const el = document.querySelector(`[data-key="${h}"]`);
      return el ? el.value : (row[h] || '');
    });

    await updateSheetRow(sheetName, editingRow, headers, values);
    showToast('Data berhasil diupdate', 'success');
    closeEditModal();
    await loadAllData();
  } catch (err) {
    showToast('Gagal update: ' + err.message, 'error');
  } finally {
    showLoading(false);
  }
}

function openDeleteModal(sheetKey, rowIndex, poNumber) {
  deletingSheet = sheetKey;
  deletingRow = rowIndex;
  const sheetLabel = sheetKey === 'logPembelian' ? 'Log Pembelian' : 'Log Pembayaran';
  document.getElementById('delete-message').textContent =
    `Yakin ingin menghapus data PO ${poNumber} dari ${sheetLabel}?`;
  document.getElementById('deleteModal').style.display = 'flex';
}

function closeDeleteModal() {
  document.getElementById('deleteModal').style.display = 'none';
  deletingSheet = null;
  deletingRow = null;
}

async function confirmDelete() {
  if (!deletingSheet || !deletingRow) return;
  showLoading(true);

  try {
    const sheetName = deletingSheet === 'logPembelian' ? CONFIG.SHEETS.LOG_PEMBELIAN : CONFIG.SHEETS.LOG_PEMBAYARAN;
    const dataKey = deletingSheet === 'logPembelian' ? 'logPembelian' : 'logPembayaran';
    const row = allData[dataKey].find(r => r._rowIndex === deletingRow);
    const colCount = row ? row._headers.length : 11;

    await deleteSheetRow(sheetName, deletingRow, colCount);
    showToast('Data berhasil dihapus', 'success');
    closeDeleteModal();
    await loadAllData();
  } catch (err) {
    showToast('Gagal hapus: ' + err.message, 'error');
  } finally {
    showLoading(false);
  }
}

// --- Navigation ---
document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', (e) => {
    e.preventDefault();
    const page = item.dataset.page;

    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    item.classList.add('active');

    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById(`page-${page}`).classList.add('active');

    closeSidebar();
  });
});

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeEditModal();
    closeDeleteModal();
  }
});

// --- Mobile Sidebar ---
const menuToggle = document.getElementById('menuToggle');
const sidebar = document.getElementById('sidebar');
const overlay = document.getElementById('overlay');

menuToggle.addEventListener('click', toggleSidebar);

function toggleSidebar() {
  sidebar.classList.toggle('open');
  overlay.classList.toggle('active');
}

function closeSidebar() {
  sidebar.classList.remove('open');
  overlay.classList.remove('active');
}

// --- Utilities ---
function formatCurrency(amount) {
  const num = typeof amount === 'string' ? parseCurrency(amount) : amount;
  if (isNaN(num) || num === 0) return 'Rp 0';
  const formatted = 'Rp ' + Math.round(Math.abs(num)).toLocaleString('id-ID');
  return num < 0 ? `(${formatted})` : formatted;
}

function parseCurrency(str) {
  if (!str) return 0;
  const cleaned = String(str).replace(/[^0-9,.-]/g, '').replace(/,/g, '.');
  return parseFloat(cleaned) || 0;
}

function statusBadge(status) {
  if (!status) return '<span class="badge badge-draft">—</span>';
  const s = status.toLowerCase();
  if (s === 'lunas') return `<span class="badge badge-lunas">${status}</span>`;
  if (s === 'belum lunas' || s === 'belum dibayar' || s === 'menunggu pembayaran') return `<span class="badge badge-belum">${status}</span>`;
  if (s === 'draft') return `<span class="badge badge-draft">${status}</span>`;
  if (s === 'realisasi') return `<span class="badge badge-realisasi">${status}</span>`;
  return `<span class="badge badge-draft">${status}</span>`;
}

function showLoading(show) {
  document.getElementById('loadingSpinner').style.display = show ? 'flex' : 'none';
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3500);
}

// --- Dark Mode ---
function toggleTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  document.documentElement.setAttribute('data-theme', isDark ? 'light' : 'dark');
  localStorage.setItem('fp-theme', isDark ? 'light' : 'dark');
}

function loadTheme() {
  const saved = localStorage.getItem('fp-theme');
  if (saved === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
  }
}

// --- Currency Input Mask ---
function setupCurrencyMask() {
  const input = document.getElementById('form-jumlah');
  if (!input) return;
  input.addEventListener('input', (e) => {
    let val = e.target.value.replace(/[^0-9]/g, '');
    if (!val) { e.target.value = ''; return; }
    e.target.value = 'Rp ' + parseInt(val, 10).toLocaleString('id-ID');
  });
  input.addEventListener('focus', (e) => {
    if (e.target.value === 'Rp 0') e.target.value = '';
  });
}

// --- Init on Load ---
window.addEventListener('load', () => {
  setupCurrencyMask();
  loadTheme();
  const gapiScript = document.createElement('script');
  gapiScript.src = 'https://apis.google.com/js/api.js';
  gapiScript.onload = gapiLoaded;
  document.head.appendChild(gapiScript);

  const gisScript = document.createElement('script');
  gisScript.src = 'https://accounts.google.com/gsi/client';
  gisScript.onload = gisLoaded;
  document.head.appendChild(gisScript);
});
