const VERSION = 'v2.9.0';

// ─── State ───────────────────────────────────────────────────────
let masterData = null;   // { circuitName, serialNumber }[]
let newData = null;      // raw rows from new file (all columns kept)
let newHeaders = null;
let validationResults = null;
let renderedTabs = new Set();

// ─── Helpers ──────────────────────────────────────────────────────
function formatCellValue(val, isDateCol = false) {
  if (val instanceof Date) {
    // Use UTC methods — SheetJS creates dates at UTC midnight, local methods shift the day in negative-offset timezones
    const mm = String(val.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(val.getUTCDate()).padStart(2, '0');
    const yyyy = val.getUTCFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  // Some date cells aren't formatted as dates in Excel so SheetJS returns a raw serial number.
  // Only convert for columns whose header contains "date" — safe for serial number columns.
  if (isDateCol && typeof val === 'number' && val > 25569 && val < 62091) {
    const d = new Date(Math.round((val - 25569) * 86400000));
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    const yyyy = d.getUTCFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  return val;
}

// ─── File Loading ─────────────────────────────────────────────────
function readFileAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

async function loadMaster(file) {
  showError('');
  const ab = await readFileAsArrayBuffer(file);
  const wb = XLSX.read(ab, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let headerRow1Idx = -1;
  for (let i = 0; i < Math.min(raw.length, 10); i++) {
    const r = raw[i].map(c => String(c).trim().toUpperCase());
    if (r.includes('CIRCUIT') && r.includes('SERIAL')) {
      headerRow1Idx = i;
      break;
    }
  }

  let circuitCol = -1, serialCol = -1;

  if (headerRow1Idx >= 0) {
    const row1 = raw[headerRow1Idx].map(c => String(c).trim().toUpperCase());
    const row2 = headerRow1Idx + 1 < raw.length ? raw[headerRow1Idx + 1].map(c => String(c).trim().toUpperCase()) : [];

    for (let c = 0; c < row1.length; c++) {
      if (row1[c] === 'CIRCUIT' && row2[c] === 'NAME') { circuitCol = c; }
      if (row1[c] === 'SERIAL' && row2[c] === 'NUMBER') { serialCol = c; }
    }
    if (circuitCol === -1) {
      for (let c = 0; c < row1.length; c++) {
        if (row1[c] === 'CIRCUIT' || row1[c] === 'NOMENCLATURE') circuitCol = c;
        if (row1[c] === 'SERIAL') serialCol = c;
      }
    }

    const dataStart = headerRow1Idx + 2;
    const data = [];
    for (let i = dataStart; i < raw.length; i++) {
      const row = raw[i];
      const cn = String(row[circuitCol] ?? '').trim();
      const sn = String(row[serialCol] ?? '').trim();
      if (cn || sn) data.push({ circuitName: cn, serialNumber: sn });
    }
    return data;
  }

  const headerRow = raw[0].map(c => String(c).trim().toUpperCase());
  for (let c = 0; c < headerRow.length; c++) {
    if (headerRow[c].includes('CIRCUIT') || headerRow[c] === 'NOMENCLATURE') circuitCol = c;
    if (headerRow[c].includes('SERIAL')) serialCol = c;
  }

  if (circuitCol === -1 || serialCol === -1) {
    throw new Error('Could not find CIRCUIT NAME and SERIAL NUMBER columns in master file.');
  }

  const data = [];
  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    const cn = String(row[circuitCol] ?? '').trim();
    const sn = String(row[serialCol] ?? '').trim();
    if (cn || sn) data.push({ circuitName: cn, serialNumber: sn });
  }
  return data;
}

async function loadNew(file) {
  showError('');
  const ab = await readFileAsArrayBuffer(file);
  const isCSV = file.name.toLowerCase().endsWith('.csv');
  let rawRows;

  if (isCSV) {
    const text = new TextDecoder().decode(ab);
    const wb = XLSX.read(text, { type: 'string' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  } else {
    const wb = XLSX.read(ab, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  }

  let headerIdx = -1;
  for (let i = 0; i < Math.min(rawRows.length, 5); i++) {
    const r = rawRows[i].map(c => String(c).trim().toLowerCase());
    if (r.includes('nomenclature') && r.some(v => v.includes('serial'))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) {
    throw new Error('Could not find Nomenclature and Serial Number columns in new file.');
  }

  const headers = rawRows[headerIdx].map(c => String(c).trim());
  const nomenclatureIdx = headers.findIndex(h => h.toLowerCase() === 'nomenclature');
  const serialIdx = headers.findIndex(h => h.toLowerCase() === 'serial number');

  if (nomenclatureIdx === -1 || serialIdx === -1) {
    throw new Error('New file must have "Nomenclature" and "Serial Number" columns.');
  }

  const dateColIndices = new Set(
    headers.map((h, i) => h.toLowerCase().includes('date') ? i : -1).filter(i => i >= 0)
  );

  const rows = [];
  for (let i = headerIdx + 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = formatCellValue(row[idx] ?? '', dateColIndices.has(idx)); });
    rows.push(obj);
  }

  return { headers, rows, nomenclatureIdx, serialIdx };
}

// ─── Validation ───────────────────────────────────────────────────
function runValidation(master, newFileData) {
  const masterSet = new Set();
  const masterCircuitSet = new Set();
  const masterSerialSet = new Set();

  master.forEach(({ circuitName, serialNumber }) => {
    masterSet.add(`${circuitName.toUpperCase()}|||${serialNumber.toUpperCase()}`);
    masterCircuitSet.add(circuitName.toUpperCase());
    masterSerialSet.add(serialNumber.toUpperCase());
  });

  const results = newFileData.rows.map(row => {
    const cn = String(row['Nomenclature'] ?? '').trim();
    const sn = String(row['Serial Number'] ?? '').trim();
    const key = `${cn.toUpperCase()}|||${sn.toUpperCase()}`;

    const pairMatch = masterSet.has(key);
    const circuitMatch = masterCircuitSet.has(cn.toUpperCase());
    const serialMatch = masterSerialSet.has(sn.toUpperCase());

    let status, issue;
    if (pairMatch) {
      status = 'PASS';
      issue = '';
    } else if (!circuitMatch && !serialMatch) {
      status = 'FAIL';
      issue = 'Circuit Name + Serial Number not in master';
    } else if (!circuitMatch) {
      status = 'FAIL';
      issue = 'Circuit Name not in master';
    } else if (!serialMatch) {
      status = 'FAIL';
      issue = 'Serial Number not in master';
    } else {
      status = 'FAIL';
      issue = 'Pair mismatch — each exists in master but not together';
    }

    return { ...row, _status: status, _issue: issue, _circuitName: cn, _serialNumber: sn };
  });

  return results;
}

// ─── UI Rendering ─────────────────────────────────────────────────
function renderStats(results) {
  const total = results.length;
  const fails = results.filter(r => r._status === 'FAIL').length;
  const passes = total - fails;
  const pct = total ? Math.round((passes / total) * 100) : 0;

  document.getElementById('stats-row').innerHTML = `
    <div class="stat-card ${fails > 0 ? 'danger' : 'success'}">
      <div class="stat-num">${fails}</div>
      <div class="stat-label">Mismatches</div>
    </div>
    <div class="stat-card success">
      <div class="stat-num">${passes.toLocaleString()}</div>
      <div class="stat-label">Validated OK</div>
    </div>
    <div class="stat-card">
      <div class="stat-num">${total.toLocaleString()}</div>
      <div class="stat-label">Total Rows</div>
    </div>
    <div class="stat-card ${pct < 100 ? 'warn' : 'success'}">
      <div class="stat-num">${pct}%</div>
      <div class="stat-label">Match Rate</div>
    </div>
  `;
}

function renderTable(containerId, rows, allHeaders, showIssueCol, description) {
  const container = document.getElementById(containerId);
  const descHtml = description ? `<p class="tab-info">${description}</p>` : '';
  if (!rows.length) {
    container.innerHTML = descHtml + '<div class="empty-state">No rows to display.</div>';
    return;
  }

  const displayHeaders = showIssueCol
    ? ['_status', '_issue', 'Nomenclature', 'Serial Number', ...allHeaders.filter(h => h !== 'Nomenclature' && h !== 'Serial Number')]
    : ['_status', 'Nomenclature', 'Serial Number', ...allHeaders.filter(h => h !== 'Nomenclature' && h !== 'Serial Number')];

  const labelMap = { _status: 'Status', _issue: 'Issue' };

  const thead = displayHeaders.map(h => `<th>${labelMap[h] || h}</th>`).join('');

  const tbody = rows.map(row => {
    const isFail = row._status === 'FAIL';
    const cells = displayHeaders.map(h => {
      const val = row[h] ?? '';
      if (h === '_status') {
        return `<td><span class="badge ${isFail ? 'badge-fail' : 'badge-pass'}">${val}</span></td>`;
      }
      if (h === '_issue') return `<td class="${isFail ? 'highlight-bad' : ''}">${val}</td>`;
      if ((h === 'Nomenclature' || h === 'Serial Number') && isFail) {
        return `<td class="highlight-bad">${val}</td>`;
      }
      return `<td>${val}</td>`;
    }).join('');
    return `<tr class="${isFail ? 'mismatch-row' : 'pass-row'}">${cells}</tr>`;
  }).join('');

  container.innerHTML = descHtml + `<div class="table-wrap"><table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></div>`;
}

function renderNotInMaster(master, newResults) {
  const newSerials = new Set(newResults.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = master.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));

  const container = document.getElementById('tab-notinmaster');
  const descHtml = '<p class="tab-info">Master records whose serial number does not appear anywhere in the new file — these circuits were expected but are completely absent from the new results.</p>';
  if (!missing.length) {
    container.innerHTML = descHtml + '<div class="empty-state">All master serials are present in the new file.</div>';
    return;
  }

  const rows = missing.map(m => `
    <tr>
      <td class="highlight-bad">${m.circuitName}</td>
      <td class="highlight-bad">${m.serialNumber}</td>
    </tr>
  `).join('');

  container.innerHTML = descHtml + `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Circuit Name (Master)</th><th>Serial Number (Master)</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

function renderUniqueFailures() {
  const container = document.getElementById('tab-uniqueerrors');
  const desc = '<p class="tab-info">Each unique failing relay shown once — deduplicated by Nomenclature + Serial Number + Issue. The Count column shows how many test rows had that exact failure. Use this to get a clean list of distinct problems without repeat noise.</p>';

  const fails = validationResults.filter(r => r._status === 'FAIL');
  if (!fails.length) {
    container.innerHTML = desc + '<div class="empty-state">No failures found.</div>';
    return;
  }

  // Deduplicate by Nomenclature + Serial Number + Issue
  const seen = new Map();
  const order = masterSortOrders();
  fails.forEach(r => {
    const key = `${r._circuitName}|||${r._serialNumber}|||${r._issue}`;
    if (seen.has(key)) {
      seen.get(key).count++;
    } else {
      seen.set(key, { ...r, count: 1 });
    }
  });

  const unique = sortByMaster([...seen.values()], order);

  const thead = '<tr><th>Count</th><th>Status</th><th>Issue</th><th>Nomenclature</th><th>Serial Number</th></tr>';
  const tbody = unique.map(r => `
    <tr class="mismatch-row">
      <td><span class="badge badge-fail">${r.count}</span></td>
      <td><span class="badge badge-fail">FAIL</span></td>
      <td class="highlight-bad">${r._issue}</td>
      <td class="highlight-bad">${r._circuitName}</td>
      <td class="highlight-bad">${r._serialNumber}</td>
    </tr>`).join('');

  container.innerHTML = desc + `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
}

function masterSortOrders() {
  const byCircuit = new Map();
  const bySerial  = new Map();
  masterData.forEach((m, i) => {
    const cn = m.circuitName.toUpperCase();
    const sn = m.serialNumber.toUpperCase();
    if (!byCircuit.has(cn)) byCircuit.set(cn, i);
    if (!bySerial.has(sn))  bySerial.set(sn, i);
  });
  return { byCircuit, bySerial };
}

function sortByMaster(rows, orders) {
  const { byCircuit, bySerial } = orders;
  const getIdx = row => {
    const cn = (row._circuitName  || '').toUpperCase();
    const sn = (row._serialNumber || '').toUpperCase();
    if (byCircuit.has(cn)) return byCircuit.get(cn);
    if (bySerial.has(sn))  return bySerial.get(sn);
    return Infinity;
  };
  // Find comment value — try common column name variants
  const getComment = row =>
    String(row['Comments'] || row['Comment'] || row['COMMENTS'] || row['comment'] || '');

  return [...rows].sort((a, b) => {
    const ai = getIdx(a);
    const bi = getIdx(b);
    if (ai !== bi) return ai - bi;
    // Tie (same master position, or both unresolved) — sort by Comments with
    // numeric awareness so R1-2A sorts before R1-10A correctly
    return getComment(a).localeCompare(getComment(b), undefined, { numeric: true, sensitivity: 'base' });
  });
}

function renderForTab(tabId) {
  const order = masterSortOrders();
  if (tabId === 'uniqueerrors') {
    renderUniqueFailures();
  } else if (tabId === 'exceptions') {
    const fails = sortByMaster(validationResults.filter(r => r._status === 'FAIL'), order);
    renderTable('tab-exceptions', fails, newData.headers, true, 'Rows from the new file where the circuit name, serial number, or the combination was not found in the master. Each row shows the specific reason it failed.');
  } else if (tabId === 'notinmaster') {
    renderNotInMaster(masterData, validationResults);
  } else if (tabId === 'fulldata') {
    const sorted = sortByMaster(validationResults, order);
    renderTable('tab-fulldata', sorted, newData.headers, true, 'Every row from the new file in master order. Green = matched the master (PASS). Red = did not match (FAIL).');
  }
}

// ─── Excel Export ─────────────────────────────────────────────────
function exportExcel(results, masterData, allHeaders) {
  const wb = XLSX.utils.book_new();
  const order = masterSortOrders();
  const sorted = sortByMaster(results, order);
  const sortedFails = sorted.filter(r => r._status === 'FAIL');

  const fullHeaders = ['Status', 'Issue', ...allHeaders];
  const fullRows = sorted.map(r => {
    const row = [r._status, r._issue];
    allHeaders.forEach(h => row.push(r[h] ?? ''));
    return row;
  });
  const fullSheet = XLSX.utils.aoa_to_sheet([fullHeaders, ...fullRows]);

  sorted.forEach((r, i) => {
    const rowIdx = i + 1;
    const cellAddr = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
    if (!fullSheet[cellAddr]) return;
    fullSheet[cellAddr].s = r._status === 'FAIL'
      ? { fill: { fgColor: { rgb: 'FFDDDD' } }, font: { bold: true, color: { rgb: 'CC0000' } } }
      : { font: { color: { rgb: '007744' } } };
  });

  XLSX.utils.book_append_sheet(wb, fullSheet, 'Full Data');

  const failRows = sortedFails;
  if (failRows.length) {
    const excRows = failRows.map(r => {
      const row = [r._status, r._issue];
      allHeaders.forEach(h => row.push(r[h] ?? ''));
      return row;
    });
    const excSheet = XLSX.utils.aoa_to_sheet([fullHeaders, ...excRows]);
    XLSX.utils.book_append_sheet(wb, excSheet, 'Failures');
  }

  // Unique Failures sheet
  const seenU = new Map();
  sortedFails.forEach(r => {
    const key = `${r._circuitName}|||${r._serialNumber}|||${r._issue}`;
    if (seenU.has(key)) { seenU.get(key).count++; }
    else { seenU.set(key, { ...r, count: 1 }); }
  });
  const uniqueErrRows = [...seenU.values()];
  if (uniqueErrRows.length) {
    const ueData = uniqueErrRows.map(r => [r.count, r._status, r._issue, r._circuitName, r._serialNumber]);
    const ueSheet = XLSX.utils.aoa_to_sheet([['Count', 'Status', 'Issue', 'Nomenclature', 'Serial Number'], ...ueData]);
    XLSX.utils.book_append_sheet(wb, ueSheet, 'Unique Failures');
  }

  const newSerials = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = masterData.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));
  if (missing.length) {
    const missingRows = missing.map(m => [m.circuitName, m.serialNumber]);
    const missingSheet = XLSX.utils.aoa_to_sheet([['Circuit Name', 'Serial Number'], ...missingRows]);
    XLSX.utils.book_append_sheet(wb, missingSheet, 'Missing From New File');
  }

  const total = results.length;
  const fails = results.filter(r => r._status === 'FAIL').length;
  const summaryData = [
    ['Relay Data Checker — Validation Report'],
    ['Generated', new Date().toLocaleString()],
    ['Version', VERSION],
    [],
    ['Total Rows', total],
    ['Passed', total - fails],
    ['Failed', fails],
    ['Match Rate', total ? `${Math.round(((total - fails) / total) * 100)}%` : 'N/A'],
    ['Master Serials Missing from New File', missing.length],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

  XLSX.writeFile(wb, `relay-data-checker_validation_${new Date().toISOString().slice(0,19).replace('T','_').replace(/:/g,'-')}.xlsx`);
}

function exportCSV(results, masterData, allHeaders) {
  const escape = v => {
    const s = String(v ?? '');
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
  };

  const order = masterSortOrders();
  const sorted = sortByMaster(results, order);

  const lines = [];
  const fullHeaders = ['Status', 'Issue', ...allHeaders];

  lines.push('FULL DATA');
  lines.push(fullHeaders.map(escape).join(','));
  sorted.forEach(r => {
    lines.push([r._status, r._issue, ...allHeaders.map(h => r[h] ?? '')].map(escape).join(','));
  });
  lines.push('');

  const fails = sortByMaster(results.filter(r => r._status === 'FAIL'), order);
  lines.push('FAILURES');
  lines.push(fullHeaders.map(escape).join(','));
  fails.forEach(r => {
    lines.push([r._status, r._issue, ...allHeaders.map(h => r[h] ?? '')].map(escape).join(','));
  });
  lines.push('');

  const seenC = new Map();
  fails.forEach(r => {
    const key = `${r._circuitName}|||${r._serialNumber}|||${r._issue}`;
    if (seenC.has(key)) { seenC.get(key).count++; }
    else { seenC.set(key, { ...r, count: 1 }); }
  });
  lines.push('UNIQUE FAILURES');
  lines.push('Count,Status,Issue,Nomenclature,Serial Number');
  [...seenC.values()].forEach(r => {
    lines.push([r.count, r._status, r._issue, r._circuitName, r._serialNumber].map(escape).join(','));
  });
  lines.push('');

  const newSerials = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = masterData.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));
  lines.push('MISSING FROM NEW FILE');
  lines.push('Circuit Name (Master),Serial Number (Master)');
  missing.forEach(m => lines.push([escape(m.circuitName), escape(m.serialNumber)].join(',')));

  const csv = lines.join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `relay-data-checker_report_${new Date().toISOString().slice(0,19).replace('T','_').replace(/:/g,'-')}.csv`;
  a.click();
}

// ─── Progress helper ──────────────────────────────────────────────
function setProgress(pct) {
  const bar = document.getElementById('progress-bar');
  const fill = document.getElementById('progress-fill');
  if (pct === null) { bar.classList.remove('visible'); return; }
  bar.classList.add('visible');
  fill.style.width = pct + '%';
}

function showError(msg) {
  const el = document.getElementById('error-banner');
  el.textContent = msg;
  el.classList.toggle('visible', !!msg);
}

function checkReady() {
  document.getElementById('run-btn').disabled = !(masterData && newData);
}

// ─── Event Wiring ─────────────────────────────────────────────────
document.getElementById('version-label').textContent = VERSION;
document.getElementById('theme-toggle').textContent = '☀ Light Mode';

document.getElementById('file-master').addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('status-master');
  const zoneEl = document.getElementById('zone-master');
  try {
    statusEl.textContent = 'Loading...';
    statusEl.className = 'zone-status';
    masterData = await loadMaster(file);
    statusEl.textContent = `✓ ${file.name} — ${masterData.length} records`;
    statusEl.className = 'zone-status ok';
    zoneEl.classList.add('loaded');
  } catch (err) {
    statusEl.textContent = '✗ ' + err.message;
    statusEl.className = 'zone-status err';
    masterData = null;
    showError(err.message);
  }
  checkReady();
});

document.getElementById('file-new').addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  const statusEl = document.getElementById('status-new');
  const zoneEl = document.getElementById('zone-new');
  try {
    statusEl.textContent = 'Loading...';
    statusEl.className = 'zone-status';
    newData = await loadNew(file);
    statusEl.textContent = `✓ ${file.name} — ${newData.rows.length} rows`;
    statusEl.className = 'zone-status ok';
    zoneEl.classList.add('loaded');
  } catch (err) {
    statusEl.textContent = '✗ ' + err.message;
    statusEl.className = 'zone-status err';
    newData = null;
    showError(err.message);
  }
  checkReady();
});

['zone-master', 'zone-new'].forEach(zoneId => {
  const zone = document.getElementById(zoneId);
  const inputId = zoneId === 'zone-master' ? 'file-master' : 'file-new';
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) {
      const input = document.getElementById(inputId);
      const dt = new DataTransfer();
      dt.items.add(file);
      input.files = dt.files;
      input.dispatchEvent(new Event('change'));
    }
  });
});

document.getElementById('run-btn').addEventListener('click', async () => {
  showError('');
  setProgress(10);
  document.getElementById('run-btn').disabled = true;
  document.getElementById('results-panel').classList.remove('visible');
  document.getElementById('export-row').style.display = 'none';

  await new Promise(r => setTimeout(r, 50));
  setProgress(40);

  try {
    validationResults = runValidation(masterData, newData);
    setProgress(80);
    await new Promise(r => setTimeout(r, 50));

    renderedTabs = new Set();
    renderStats(validationResults);

    // Reset to first tab and set loading placeholders in all panels
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelector('[data-tab="exceptions"]').classList.add('active');
    const exceptionsPanel = document.getElementById('tab-exceptions');
    exceptionsPanel.classList.add('active');
    exceptionsPanel.innerHTML = '<div class="loading-state"><span class="loading-dot"></span>Rendering rows\u2026</div>';
    document.getElementById('tab-uniqueerrors').innerHTML = '<div class="empty-state">Click tab to load.</div>';
    document.getElementById('tab-notinmaster').innerHTML = '<div class="empty-state">Click tab to load.</div>';
    document.getElementById('tab-fulldata').innerHTML = '<div class="empty-state">Click tab to load.</div>';

    document.getElementById('results-panel').classList.add('visible');
    document.getElementById('export-row').style.display = 'flex';
    setProgress(null);

    // Lazy-render the active (exceptions) tab after paint
    requestAnimationFrame(() => requestAnimationFrame(() => {
      renderForTab('exceptions');
      renderedTabs.add('exceptions');
    }));
  } catch (err) {
    showError('Validation error: ' + err.message);
    setProgress(null);
  }

  document.getElementById('run-btn').disabled = false;
});

document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    const tabId = tab.dataset.tab;
    const panel = document.getElementById('tab-' + tabId);
    panel.classList.add('active');

    if (validationResults && !renderedTabs.has(tabId)) {
      panel.innerHTML = '<div class="loading-state"><span class="loading-dot"></span>Rendering rows\u2026</div>';
      requestAnimationFrame(() => requestAnimationFrame(() => {
        renderForTab(tabId);
        renderedTabs.add(tabId);
      }));
    }
  });
});

document.getElementById('theme-toggle').addEventListener('click', () => {
  const isLight = document.body.classList.toggle('light');
  document.getElementById('theme-toggle').textContent = isLight ? '☾ Dark Mode' : '☀ Light Mode';
});

document.getElementById('dl-excel').addEventListener('click', () => {
  if (validationResults) exportExcel(validationResults, masterData, newData.headers);
});

document.getElementById('dl-csv').addEventListener('click', () => {
  if (validationResults) exportCSV(validationResults, masterData, newData.headers);
});
