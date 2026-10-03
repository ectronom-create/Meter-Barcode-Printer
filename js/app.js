/**
 * Meter Label Printer — Application Logic
 * Supports: Device_pallet_number, Device_box_number, Serial_number, Module_ICCID
 */

// =============================================
//  STATE
// =============================================
let metersData   = [];  // [{meterSN, iccid, palletNum, boxNum, boxSuffix}]
let currentRecord = null;

// Default Sample Data (legacy 2-column format)
const sampleData = [
  { meterSN: '062621000035277', iccid: '8996803620041204244F', palletNum: 'PLT-001', boxNum: 'BOX-001', boxSuffix: '001' },
  { meterSN: '062621000035276', iccid: '8996803620041418935F', palletNum: 'PLT-001', boxNum: 'BOX-001', boxSuffix: '001' },
  { meterSN: '062621000035275', iccid: '8996803620041166344F', palletNum: 'PLT-001', boxNum: 'BOX-002', boxSuffix: '002' }
];

// =============================================
//  INIT
// =============================================
document.addEventListener('DOMContentLoaded', () => {
  initStorage();
  initNav();
  initEventListeners();
  updateUI();
  checkZebraService();
  setInterval(checkZebraService, 4000);

  if (metersData.length > 0) selectRecord(metersData[0]);
});

// =============================================
//  STORAGE
// =============================================
function initStorage() {
  try {
    const saved = localStorage.getItem('meter_records_v2');
    if (saved) {
      metersData = JSON.parse(saved);
    } else {
      // try old key
      const old = localStorage.getItem('meter_records_v1');
      if (old) {
        const parsed = JSON.parse(old);
        metersData = parsed.map(r => ({
          meterSN: r.meterSN || '',
          iccid:   r.iccid   || '',
          palletNum: r.palletNum || '',
          boxNum:    r.boxNum    || '',
          boxSuffix: r.boxSuffix || ''
        }));
      } else {
        metersData = [...sampleData];
      }
      saveStorage();
    }
  } catch (e) {
    console.error('Storage error:', e);
    metersData = [...sampleData];
  }
}

function saveStorage() {
  try {
    localStorage.setItem('meter_records_v2', JSON.stringify(metersData));
  } catch (e) {
    console.error('Storage write error:', e);
  }
}

// =============================================
//  NAVIGATION
// =============================================
function initNav() {
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const targetId = btn.getAttribute('data-tab');
      document.querySelectorAll('.page').forEach(page => {
        page.style.display = page.id === targetId ? 'flex' : 'none';
      });

      if (targetId === 'tab-print') {
        setTimeout(() => document.getElementById('searchMeterSN')?.focus(), 80);
      }
      if (targetId === 'tab-pallets') {
        renderPallets();
      }
    });
  });
}

// =============================================
//  EVENT LISTENERS
// =============================================
function initEventListeners() {
  // File Upload
  const fileInput = document.getElementById('excelFileInput');
  const dropzone  = document.getElementById('uploadDropzone');
  if (dropzone && fileInput) {
    dropzone.addEventListener('click', (e) => {
      if (!e.target.closest('.btn-upload-browse')) fileInput.click();
    });
    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault(); dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault(); dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) handleExcelFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleExcelFile(e.target.files[0]);
    });
  }

  // Load Sample
  document.getElementById('btnLoadSample')?.addEventListener('click', () => {
    metersData = [...sampleData];
    saveStorage();
    updateUI();
    selectRecord(metersData[0]);
    showToast('Sample data loaded!', 'success');
  });

  // Clear Data
  document.getElementById('btnClearData')?.addEventListener('click', () => {
    if (confirm('Are you sure you want to clear all saved data?')) {
      metersData = [];
      currentRecord = null;
      saveStorage();
      updateUI();
      clearPreview();
      showToast('All data cleared.', 'warning');
    }
  });

  // Scan Input
  const si = document.getElementById('searchMeterSN');
  if (si) {
    si.focus();
    si.addEventListener('input', (e) => {
      const q = e.target.value.trim();
      if (q) searchAndPreview(q, false);
    });
    si.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const q = si.value.trim();
        if (q) {
          const autoPrint = document.getElementById('chkAutoPrint')?.checked;
          const found = searchAndPreview(q, true);
          if (found && autoPrint) printCurrentSticker();
          si.select();
        }
      }
    });
  }

  document.getElementById('btnClearScan')?.addEventListener('click', () => {
    const si = document.getElementById('searchMeterSN');
    if (si) { si.value = ''; si.focus(); }
    hideSearchError();
  });

  // Print Buttons
  document.getElementById('btnPrintCurrent')?.addEventListener('click', () => printCurrentSticker(false));
  document.getElementById('btnPrintBrowser')?.addEventListener('click', () => printCurrentSticker(true));
  document.getElementById('btnBatchPrint')?.addEventListener('click', () => batchPrint(metersData));

  // Label Size
  document.getElementById('labelSizeSelect')?.addEventListener('change', () => {
    updateStickerSizeStyles();
    if (currentRecord) renderSticker(currentRecord);
  });
  document.getElementById('customWidth')?.addEventListener('input', () => {
    updateStickerSizeStyles();
    if (currentRecord) renderSticker(currentRecord);
  });
  document.getElementById('customHeight')?.addEventListener('input', () => {
    updateStickerSizeStyles();
    if (currentRecord) renderSticker(currentRecord);
  });

  // Barcode height
  document.getElementById('barcodeHeight')?.addEventListener('input', () => {
    if (currentRecord) renderSticker(currentRecord);
  });

  // Table Search
  document.getElementById('tableSearch')?.addEventListener('input', (e) => {
    filterTable(e.target.value.trim());
  });
}

// =============================================
//  EXCEL IMPORT
// =============================================
function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data     = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
      const ws       = workbook.Sheets[workbook.SheetNames[0]];
      const rows     = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });

      if (!rows || rows.length === 0) {
        showToast('File is empty or has no valid data rows!', 'error');
        return;
      }

      const headers = Object.keys(rows[0]);
      const mapping = detectColumns(headers);

      if (!mapping.serial || !mapping.iccid) {
        showToast(`Cannot detect required columns. Found: [${headers.join(', ')}]`, 'error');
        return;
      }

      const imported = [];
      rows.forEach(row => {
        const meterSN  = String(row[mapping.serial]  || '').trim();
        const iccid    = String(row[mapping.iccid]   || '').trim();
        const palletNum = mapping.pallet ? String(row[mapping.pallet] || '').trim() : '';
        const boxNum    = mapping.box    ? String(row[mapping.box]    || '').trim() : '';
        // Last 3 digits of box number
        const boxSuffix = boxNum ? boxNum.replace(/\D/g, '').slice(-3) : '';

        if (meterSN && iccid) {
          imported.push({ meterSN, iccid, palletNum, boxNum, boxSuffix });
        }
      });

      if (imported.length === 0) {
        showToast('No valid records found in the file.', 'error');
        return;
      }

      metersData = imported;
      saveStorage();
      updateUI();
      selectRecord(metersData[0]);
      showToast(`✅ Imported ${metersData.length} records successfully!`, 'success');

      // Switch to Print tab
      document.querySelector('[data-tab="tab-print"]')?.click();

    } catch (err) {
      console.error('Parse error:', err);
      showToast('Error reading file. Check the format and try again.', 'error');
    }
  };
  reader.onerror = () => showToast('Could not read the file.', 'error');
  reader.readAsArrayBuffer(file);
}

function detectColumns(headers) {
  let serial = null, iccid = null, pallet = null, box = null;

  for (const h of headers) {
    const c = h.trim().toLowerCase().replace(/[\s_\-]/g, '');
    if (!serial  && (c.includes('serialnumber') || c.includes('serial') || c.includes('metersn') || c.includes('meter'))) serial  = h;
    else if (!iccid   && (c.includes('moduleiccid') || c.includes('iccid') || c.includes('sim')))  iccid   = h;
    else if (!pallet  && (c.includes('palletnumber') || c.includes('pallet'))) pallet  = h;
    else if (!box     && (c.includes('boxnumber')    || c.includes('box')))    box     = h;
  }

  // fallback
  if (!serial) serial = headers.find(h => /serial|meter/i.test(h)) || headers[0];
  if (!iccid)  iccid  = headers.find(h => /icc/i.test(h))          || headers[1];
  if (!pallet) pallet = headers.find(h => /pallet/i.test(h))        || null;
  if (!box)    box    = headers.find(h => /box/i.test(h))            || null;

  return { serial, iccid, pallet, box };
}

// =============================================
//  SEARCH / SELECT
// =============================================
function searchAndPreview(snQuery, showErrorOnNotFound = false) {
  const q = snQuery.trim().toLowerCase();
  if (!q) return false;

  let found = metersData.find(r => r.meterSN.toLowerCase() === q);
  if (!found) {
    found = metersData.find(r =>
      r.meterSN.toLowerCase().includes(q) || q.includes(r.meterSN.toLowerCase())
    );
  }

  if (found) {
    selectRecord(found);
    hideSearchError();
    return true;
  } else {
    if (showErrorOnNotFound) showSearchError(`Meter not found: "${snQuery}"`);
    return false;
  }
}

function selectRecord(record) {
  currentRecord = record;

  // Info card fields
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val || '—'; };
  set('detailPallet',  record.palletNum);
  set('detailBox',     record.boxNum);
  set('detailMeterSN', record.meterSN);
  set('detailICCID',   record.iccid);

  // Show data section
  const infoEmpty = document.getElementById('infoEmpty');
  const infoData  = document.getElementById('infoData');
  if (infoEmpty) infoEmpty.style.display = 'none';
  if (infoData)  infoData.style.display  = 'flex';

  renderSticker(record);

  // Highlight row
  document.querySelectorAll('#metersTableBody tr').forEach(tr => {
    tr.classList.toggle('row-selected', tr.dataset.sn === record.meterSN);
  });
}

// =============================================
//  STICKER RENDER
// Build label line: SN: 062621000035277  (277)
// =============================================
function buildSnLine(record) {
  if (record.boxSuffix) {
    return `SN: ${record.meterSN} (${record.boxSuffix})`;
  }
  return `SN: ${record.meterSN}`;
}

function renderSticker(record) {
  const container = document.getElementById('stickerPreview');
  if (!container || !record) return;

  const bHeight = parseInt(document.getElementById('barcodeHeight')?.value || '32', 10);
  const snLine  = buildSnLine(record);

  container.innerHTML = `
    <div class="sticker-iccid">ICCID: ${escapeHtml(record.iccid)}</div>
    <div class="sticker-barcode-wrapper">
      <svg id="previewBarcodeSvg"></svg>
    </div>
    <div class="sticker-metersn">${escapeHtml(snLine)}</div>
  `;

  try {
    JsBarcode('#previewBarcodeSvg', record.iccid, {
      format: 'CODE128',
      width: 1.25,
      height: bHeight,
      displayValue: false,
      margin: 0,
      background: 'transparent',
      lineColor: '#000000'
    });
  } catch (err) {
    console.error('Barcode error:', err);
  }
}

// =============================================
//  PALLET VIEW
// =============================================
function renderPallets() {
  const grid = document.getElementById('palletGrid');
  if (!grid) return;

  if (!metersData.length) {
    grid.innerHTML = '<div class="pallet-empty">No data loaded yet. Import an Excel file first.</div>';
    return;
  }

  // Group by pallet
  const palletMap = new Map();
  metersData.forEach(r => {
    const key = r.palletNum || '(No Pallet)';
    if (!palletMap.has(key)) palletMap.set(key, []);
    palletMap.get(key).push(r);
  });

  grid.innerHTML = '';
  palletMap.forEach((records, palletName) => {
    // Unique boxes
    const boxes = [...new Set(records.map(r => r.boxNum).filter(Boolean))];

    const card = document.createElement('div');
    card.className = 'pallet-card';

    const boxPills = boxes.map(b => {
      const suffix = b.replace(/\D/g, '').slice(-3);
      return `<span class="box-pill" title="${escapeHtml(b)}">${escapeHtml(b)}</span>`;
    }).join('');

    card.innerHTML = `
      <div class="pallet-card-header">
        <div class="pallet-card-title">
          📦 ${escapeHtml(palletName)}
          <small>Pallet</small>
        </div>
        <span class="pallet-badge">${records.length} labels</span>
      </div>
      <div class="pallet-card-body">
        <div class="pallet-stat-row">
          <span>Total Meters</span>
          <strong>${records.length}</strong>
        </div>
        <div class="pallet-stat-row">
          <span>Boxes</span>
          <strong>${boxes.length}</strong>
        </div>
        ${boxes.length > 0 ? `
          <div class="pallet-boxes-label">Boxes in this pallet</div>
          <div class="pallet-box-pills">${boxPills}</div>
        ` : ''}
      </div>
      <div class="pallet-card-footer">
        <button class="btn-pallet-print" onclick="printPallet('${escapeHtml(palletName)}')">
          ⚡ Print All ${records.length} Labels
        </button>
      </div>
    `;

    grid.appendChild(card);
  });
}

window.printPallet = async function(palletName) {
  const meterRecords = metersData.filter(r => (r.palletNum || '(No Pallet)') === palletName);
  if (!meterRecords.length) return;

  if (!confirm(`Print ${meterRecords.length} labels for pallet "${palletName}"?\n\nA pallet cover sticker will be printed first.`)) return;

  // Prepend the pallet header record
  const headerRecord = {
    type:        'pallet_header',
    palletNum:   palletName,
    totalLabels: meterRecords.length
  };
  const allRecords = [headerRecord, ...meterRecords];

  const ok = await tryDirectZebraPrint(allRecords);
  if (!ok) printRecords(allRecords);
};

// =============================================
//  ZEBRA SERVICE
// =============================================
let isZebraDirectAvailable = false;

async function checkZebraService() {
  try {
    const res = await fetch('http://localhost:5000/api/print-zebra', { method: 'OPTIONS' });
    if (res.ok || res.status === 200) {
      isZebraDirectAvailable = true;
      updatePrinterStatus(true);
      return;
    }
  } catch (e) { /* not running */ }
  isZebraDirectAvailable = false;
  updatePrinterStatus(false);
}

function updatePrinterStatus(online) {
  const dot  = document.getElementById('statusDot');
  const text = document.getElementById('statusText');
  if (dot)  dot.className  = 'status-dot ' + (online ? 'online' : 'offline');
  if (text) text.textContent = online ? 'Zebra ZT411 — Ready' : 'Zebra not connected';
}

// =============================================
//  PRINT
// =============================================
async function printCurrentSticker(forceBrowser = false) {
  if (!currentRecord) {
    showToast('Scan or select a meter first!', 'warning');
    return;
  }
  if (!forceBrowser) {
    const ok = await tryDirectZebraPrint([currentRecord]);
    if (ok) return;
  }
  printRecords([currentRecord]);
}

async function batchPrint(records) {
  if (!records || records.length === 0) {
    showToast('No data to print. Import an Excel file first.', 'warning');
    return;
  }
  if (!confirm(`Print all ${records.length} labels?`)) return;
  const ok = await tryDirectZebraPrint(records);
  if (!ok) printRecords(records);
}

async function tryDirectZebraPrint(records) {
  try {
    const res = await fetch('http://localhost:5000/api/print-zebra', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ records })
    });
    if (res.ok) {
      showToast(`⚡ Sent ${records.length} label(s) to Zebra ZT411!`, 'success');
      return true;
    }
  } catch (e) { /* fallback */ }
  return false;
}

// Isolated iframe browser print
function printRecords(recordsList) {
  const size = getSelectedDimensions();

  let iframe = document.getElementById('thermal-print-iframe');
  if (iframe) iframe.remove();

  iframe = document.createElement('iframe');
  iframe.id = 'thermal-print-iframe';
  Object.assign(iframe.style, {
    position: 'fixed', right: '0', bottom: '0',
    width: '1px', height: '1px',
    border: 'none', opacity: '0.01'
  });
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;
  doc.open();
  doc.write(`<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><title>Print</title>
<style>
@page { size: ${size.width}mm ${size.height}mm; margin: 0mm !important; }
* { box-sizing: border-box !important; margin: 0 !important; padding: 0 !important; }
html, body {
  width: ${size.width}mm !important;
  background: #fff !important;
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, Courier, monospace !important;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}
/* ---- Regular label page ---- */
.label-page {
  width: ${size.width}mm !important; height: ${size.height}mm !important;
  max-height: ${size.height}mm !important;
  page-break-after: always !important; break-after: page !important;
  page-break-inside: avoid !important;
  display: flex !important; flex-direction: column !important;
  align-items: center !important; justify-content: space-between !important;
  padding: 1mm 1.5mm !important; overflow: hidden !important; text-align: center !important;
}
.label-page:last-child { page-break-after: auto !important; break-after: auto !important; }
.label-iccid {
  font-size: 7.5pt !important; font-weight: 800 !important;
  line-height: 1.1 !important; direction: ltr !important;
  color: #000 !important; width: 100% !important; white-space: nowrap !important;
}
.label-barcode {
  width: 100% !important; display: flex !important;
  align-items: center !important; justify-content: center !important; margin: 0.3mm 0 !important;
}
.label-barcode svg { max-width: 98% !important; height: ${size.height <= 25 ? '11mm' : '15mm'} !important; display: block !important; }
.label-sn {
  font-size: 8pt !important; font-weight: 900 !important;
  line-height: 1.1 !important; direction: ltr !important;
  color: #000 !important; width: 100% !important; white-space: nowrap !important;
  padding-top: 0.5mm !important; margin-top: 0.5mm !important;
}
/* ---- Pallet header page ---- */
.pallet-header-page {
  width: ${size.width}mm !important; height: ${size.height}mm !important;
  max-height: ${size.height}mm !important;
  page-break-after: always !important; break-after: page !important;
  page-break-inside: avoid !important;
  display: flex !important; flex-direction: column !important;
  align-items: center !important; justify-content: space-between !important;
  padding: 1mm !important; overflow: hidden !important;
  text-align: center !important;
  border: 1.5pt solid #000 !important;
}
.ph-top {
  font-size: 7pt !important; font-weight: 700 !important;
  letter-spacing: 0.08em !important; color: #000 !important;
  text-transform: uppercase !important; width: 100% !important;
}
.ph-num {
  font-size: 11.5pt !important; font-weight: 900 !important;
  color: #000 !important; line-height: 1.2 !important;
  width: 100% !important; overflow: hidden !important;
  white-space: nowrap !important; text-overflow: ellipsis !important;
}
.ph-divider {
  width: 85% !important; height: 1pt !important;
  background: #000 !important; border: none !important;
  margin: 1mm auto !important;
}
.ph-total {
  font-size: 7pt !important; font-weight: 700 !important;
  color: #000 !important; width: 100% !important;
}
</style></head><body><div id="lc"></div></body></html>`);
  doc.close();

  const container = doc.getElementById('lc');
  recordsList.forEach(rec => {
    const recType = rec.type || 'label';

    if (recType === 'pallet_header') {
      // ---- Pallet Cover Sticker ----
      const page = doc.createElement('div');
      page.className = 'pallet-header-page';

      const topDiv = doc.createElement('div');
      topDiv.className = 'ph-top';
      topDiv.textContent = '★ PALLET ★';

      const numDiv = doc.createElement('div');
      numDiv.className = 'ph-num';
      numDiv.textContent = rec.palletNum || '';

      const hr = doc.createElement('div');
      hr.className = 'ph-divider';

      const totalDiv = doc.createElement('div');
      totalDiv.className = 'ph-total';
      totalDiv.textContent = `Total: ${rec.totalLabels} labels`;

      page.appendChild(topDiv);
      page.appendChild(numDiv);
      page.appendChild(hr);
      page.appendChild(totalDiv);
      container.appendChild(page);

    } else {
      // ---- Regular Meter Label ----
      const snLine = buildSnLine(rec);

      const page = doc.createElement('div');
      page.className = 'label-page';

      const iccidDiv = doc.createElement('div');
      iccidDiv.className = 'label-iccid';
      iccidDiv.textContent = 'ICCID: ' + rec.iccid;

      const barcodeDiv = doc.createElement('div');
      barcodeDiv.className = 'label-barcode';
      const svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
      barcodeDiv.appendChild(svg);

      const snDiv = doc.createElement('div');
      snDiv.className = 'label-sn';
      snDiv.textContent = snLine;

      page.appendChild(iccidDiv);
      page.appendChild(barcodeDiv);
      page.appendChild(snDiv);
      container.appendChild(page);

      try {
        JsBarcode(svg, rec.iccid, {
          format: 'CODE128',
          width: 1.25,
          height: size.height <= 25 ? 28 : (size.height > 30 ? 44 : 34),
          displayValue: false,
          margin: 0,
          background: '#ffffff',
          lineColor: '#000000'
        });
      } catch (err) {
        console.error('Barcode render error:', err);
      }
    }
  });

  setTimeout(() => {
    try { iframe.contentWindow.focus(); iframe.contentWindow.print(); }
    catch (e) { window.print(); }
  }, 150);
}


// =============================================
//  DIMENSIONS
// =============================================
function getSelectedDimensions() {
  const val = document.getElementById('labelSizeSelect')?.value || '50x25';
  if (val === 'custom') {
    return {
      width:  parseFloat(document.getElementById('customWidth')?.value  || '50'),
      height: parseFloat(document.getElementById('customHeight')?.value || '25')
    };
  }
  const [w, h] = val.split('x');
  return { width: parseFloat(w), height: parseFloat(h) };
}

function updateStickerSizeStyles() {
  const val      = document.getElementById('labelSizeSelect')?.value || '50x25';
  const customBox = document.getElementById('customDimensionsBox');
  const sticker   = document.getElementById('stickerPreview');
  if (customBox) customBox.style.display = (val === 'custom') ? 'flex' : 'none';
  const dims = getSelectedDimensions();
  if (sticker) {
    ['width','height','minWidth','minHeight','maxWidth','maxHeight'].forEach((p, i) => {
      sticker.style[p] = (i % 2 === 0 ? dims.width : dims.height) + 'mm';
    });
  }
}

// =============================================
//  UPDATE UI
// =============================================
function updateUI() {
  const badge = document.getElementById('totalRecordsBadge');
  if (badge) badge.textContent = metersData.length;

  // Count unique pallets
  const palletBadge = document.getElementById('totalPalletsBadge');
  if (palletBadge) {
    const pallets = new Set(metersData.map(r => r.palletNum).filter(Boolean));
    palletBadge.textContent = pallets.size;
  }

  renderTable(metersData);
  updateStickerSizeStyles();
}

// =============================================
//  TABLE
// =============================================
function renderTable(data) {
  const tbody    = document.getElementById('metersTableBody');
  const countLbl = document.getElementById('tableCountLabel');
  if (countLbl) countLbl.textContent = `${data.length} record${data.length !== 1 ? 's' : ''}`;
  if (!tbody) return;

  if (data.length === 0) {
    tbody.innerHTML = `
      <tr><td colspan="6" style="text-align:center; padding:2rem; color:#9ca3af;">
        No data. Import an Excel file or load sample data.
      </td></tr>`;
    return;
  }

  tbody.innerHTML = data.map((item, idx) => `
    <tr data-sn="${escapeHtml(item.meterSN)}"
        class="${currentRecord && currentRecord.meterSN === item.meterSN ? 'row-selected' : ''}">
      <td>${idx + 1}</td>
      <td>${escapeHtml(item.palletNum || '—')}</td>
      <td>${escapeHtml(item.boxNum    || '—')}</td>
      <td>${escapeHtml(item.meterSN)}</td>
      <td>${escapeHtml(item.iccid)}</td>
      <td>
        <button class="btn-row-print" onclick="selectAndShowSticker(${metersData.indexOf(item)})">
          Print
        </button>
      </td>
    </tr>
  `).join('');
}

function filterTable(keyword) {
  if (!keyword) { renderTable(metersData); return; }
  const k = keyword.toLowerCase();
  renderTable(metersData.filter(r =>
    r.meterSN.toLowerCase().includes(k)  ||
    r.iccid.toLowerCase().includes(k)    ||
    (r.palletNum || '').toLowerCase().includes(k) ||
    (r.boxNum    || '').toLowerCase().includes(k)
  ));
}

window.selectAndShowSticker = function(index) {
  const record = metersData[index];
  if (!record) return;
  selectRecord(record);
  const si = document.getElementById('searchMeterSN');
  if (si) si.value = record.meterSN;
  document.querySelector('[data-tab="tab-print"]')?.click();
};

// =============================================
//  CLEAR / RESET
// =============================================
function clearPreview() {
  const sticker = document.getElementById('stickerPreview');
  if (sticker) sticker.innerHTML = '';
  const infoEmpty = document.getElementById('infoEmpty');
  const infoData  = document.getElementById('infoData');
  if (infoEmpty) infoEmpty.style.display = 'flex';
  if (infoData)  infoData.style.display  = 'none';
}

// =============================================
//  NOTIFICATIONS
// =============================================
let toastTimer = null;
function showToast(message, type = 'info') {
  const bar = document.getElementById('toastBar');
  if (!bar) return;
  const icons = { success: '✅', error: '❌', warning: '⚠️', info: 'ℹ️' };
  bar.className = `toast-bar ${type}`;
  bar.innerHTML = `<span>${icons[type] || 'ℹ️'}</span> ${message}`;
  bar.style.display = 'flex';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { bar.style.display = 'none'; }, 4000);
}

function showAlert(msg, type) { showToast(msg, type); }

function showSearchError(message) {
  const el = document.getElementById('searchErrorBox');
  if (el) { el.textContent = message; el.style.display = 'block'; }
}

function hideSearchError() {
  const el = document.getElementById('searchErrorBox');
  if (el) el.style.display = 'none';
}

// =============================================
//  UTILS
// =============================================
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
