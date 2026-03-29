// ─── State ───────────────────────────────────────────────────────
let masterData = null;   // { circuitName, serialNumber }[]
let newData = null;      // raw rows from new file (all columns kept)
let newHeaders = null;
let validationResults = null;
let fullDataRendered = false;

// ─── Helpers ──────────────────────────────────────────────────────
function formatCellValue(val) {
  if (val instanceof Date) {
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    const dd = String(val.getDate()).padStart(2, '0');
    const yyyy = val.getFullYear();
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
        if (row1[c] === 'CIRCUIT') circuitCol = c;
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
    if (headerRow[c].includes('CIRCUIT')) circuitCol = c;
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

  const rows = [];
  for (let i = headerIdx + 1; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (row.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = formatCellValue(row[idx] ?? ''); });
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

// ─── Excel Export ─────────────────────────────────────────────────
function exportExcel(results, masterData, allHeaders) {
  const wb = XLSX.utils.book_new();

  const fullHeaders = ['Status', 'Issue', ...allHeaders];
  const fullRows = results.map(r => {
    const row = [r._status, r._issue];
    allHeaders.forEach(h => row.push(r[h] ?? ''));
    return row;
  });
  const fullSheet = XLSX.utils.aoa_to_sheet([fullHeaders, ...fullRows]);

  results.forEach((r, i) => {
    const rowIdx = i + 1;
    const cellAddr = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
    if (!fullSheet[cellAddr]) return;
    fullSheet[cellAddr].s = r._status === 'FAIL'
      ? { fill: { fgColor: { rgb: 'FFDDDD' } }, font: { bold: true, color: { rgb: 'CC0000' } } }
      : { font: { color: { rgb: '007744' } } };
  });

  XLSX.utils.book_append_sheet(wb, fullSheet, 'Full Data');

  const failRows = results.filter(r => r._status === 'FAIL');
  if (failRows.length) {
    const excRows = failRows.map(r => {
      const row = [r._status, r._issue];
      allHeaders.forEach(h => row.push(r[h] ?? ''));
      return row;
    });
    const excSheet = XLSX.utils.aoa_to_sheet([fullHeaders, ...excRows]);
    XLSX.utils.book_append_sheet(wb, excSheet, 'Failures');
  }

  const newSerials = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = masterData.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));
  if (missing.length) {
    const missingRows = missing.map(m => [m.circuitName, m.serialNumber]);
    const missingSheet = XLSX.utils.aoa_to_sheet([['Circuit Name', 'Serial Number'], ...missingRows]);
    XLSX.utils.book_append_sheet(wb, missingSheet, 'Not In Master');
  }

  const total = results.length;
  const fails = results.filter(r => r._status === 'FAIL').length;
  const summaryData = [
    ['Relay Data Checker — Validation Report'],
    ['Generated', new Date().toLocaleString()],
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

function exportCSV(results, allHeaders) {
  const fails = results.filter(r => r._status === 'FAIL');
  if (!fails.length) { alert('No failures to export.'); return; }
  const headers = ['Status', 'Issue', ...allHeaders];
  const rows = fails.map(r => {
    return [r._status, r._issue, ...allHeaders.map(h => {
      const v = String(r[h] ?? '');
      return v.includes(',') ? `"${v}"` : v;
    })].join(',');
  });
  const csv = [headers.join(','), ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `relay-data-checker_failures_${new Date().toISOString().slice(0,19).replace('T','_').replace(/:/g,'-')}.csv`;
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

  await new Promise(r => setTimeout(r, 50));
  setProgress(40);

  try {
    validationResults = runValidation(masterData, newData);
    setProgress(80);
    await new Promise(r => setTimeout(r, 50));

    fullDataRendered = false;
    renderStats(validationResults);

    const fails = validationResults.filter(r => r._status === 'FAIL');
    renderTable('tab-exceptions', fails, newData.headers, true, 'Rows from the new file where the circuit name, serial number, or the combination was not found in the master. Each row shows the specific reason it failed.');
    document.getElementById('tab-fulldata').innerHTML = '<div class="empty-state">Click the Full Data tab to load all rows.</div>';
    renderNotInMaster(masterData, validationResults);

    document.getElementById('results-panel').classList.add('visible');
    setProgress(null);
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
    const panel = document.getElementById('tab-' + tab.dataset.tab);
    panel.classList.add('active');

    if (tab.dataset.tab === 'fulldata' && !fullDataRendered && validationResults) {
      panel.innerHTML = '<div class="loading-state"><span class="loading-dot"></span>Rendering rows…</div>';
      // Two rAF calls: first lets the browser paint the loading state,
      // second runs the heavy render in the next frame after paint.
      requestAnimationFrame(() => requestAnimationFrame(() => {
        renderTable('tab-fulldata', validationResults, newData.headers, true, 'Every row from the new file. Green = matched the master (PASS). Red = did not match (FAIL).');
        fullDataRendered = true;
      }));
    }
  });
});

document.getElementById('theme-toggle').addEventListener('click', () => {
  const isLight = document.body.classList.toggle('light');
  document.getElementById('theme-toggle').textContent = isLight ? 'Dark Mode' : 'Light Mode';
});

document.getElementById('dl-excel').addEventListener('click', () => {
  if (validationResults) exportExcel(validationResults, masterData, newData.headers);
});

document.getElementById('dl-csv').addEventListener('click', () => {
  if (validationResults) exportCSV(validationResults, newData.headers);
});
