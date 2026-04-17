const VERSION = 'v2.20.0';

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
        if (row1[c].includes('SERIAL')) {
          if (serialCol === -1 || row1[c] === 'SERIAL NUMBER') serialCol = c;
        }
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
    // Prefer exact 'SERIAL NUMBER' — prevents 'DEVICE SERIAL' or other
    // columns that contain the word SERIAL from overriding the right column
    if (headerRow[c].includes('SERIAL')) {
      if (serialCol === -1 || headerRow[c] === 'SERIAL NUMBER') serialCol = c;
    }
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

function renderSummary(results, master) {
  const total     = results.length;
  const fails     = results.filter(r => r._status === 'FAIL');
  const failRate  = total ? fails.length / total : 0;

  const byType = {};
  fails.forEach(r => { byType[r._issue] = (byType[r._issue] || 0) + 1; });
  const sorted     = Object.entries(byType).sort((a, b) => b[1] - a[1]);
  const dominant   = sorted[0];
  const domPct     = dominant ? Math.round((dominant[1] / fails.length) * 100) : 0;

  const uniqueFails = new Set(fails.map(r => `${r._circuitName}|||${r._serialNumber}|||${r._issue}`)).size;
  const repeatedTesting = fails.length > 0 && fails.length > uniqueFails * 1.5;

  const newSerials    = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missingCount  = master.filter(m => !newSerials.has(m.serialNumber.toUpperCase())).length;
  const missingPct    = master.length ? Math.round((missingCount / master.length) * 100) : 0;

  let analysis = '';

  if (fails.length === 0) {
    analysis = `<strong>Clean run.</strong> Every relay matched the master on both circuit name and serial number. No action required.`;

  } else if (failRate > 0.9) {
    if (dominant && dominant[0].includes('Serial Number not in master') && domPct > 60) {
      analysis = `<strong>Likely column mismatch in the master file.</strong> Over ${Math.round(failRate * 100)}% of records failed on serial number alone. A failure rate this high on a single field almost always points to the master reading the wrong column — for example, a "Device Serial" or secondary serial field instead of the relay serial number. Verify which serial column was picked up in Step 1.`;
    } else if (dominant && dominant[0].includes('Circuit Name + Serial Number')) {
      analysis = `<strong>Possible wrong master file.</strong> Nearly every record failed with neither field found. This usually means the master file covers a different location or system than the new results file — the two datasets don't appear to be from the same installation.`;
    } else {
      analysis = `<strong>Systemic issue — not individual relay failures.</strong> A ${Math.round(failRate * 100)}% failure rate is too high to be isolated hardware problems. Most likely cause: a column mapping issue, outdated master file, or a major reconfiguration that hasn't been captured in the master yet.`;
    }

  } else if (failRate > 0.3) {
    if (dominant && dominant[0].includes('Pair mismatch') && domPct > 50) {
      analysis = `<strong>Suggests hardware has been reorganized.</strong> The dominant failure is pair mismatch — circuit names and serial numbers both exist in the master, just not paired together. This pattern is consistent with relays being physically moved, swapped between rack positions, or reinstalled after maintenance without a corresponding master update.`;
    } else if (dominant && dominant[0].includes('Serial Number') && domPct > 60) {
      analysis = `<strong>Serial number population has changed.</strong> High failure rate driven by serial number mismatches with circuit names largely intact. This suggests a significant portion of hardware has been replaced since the master was last updated — the rack positions are correct but the units in them are different.`;
    } else {
      analysis = `<strong>Widespread discrepancies across multiple failure types.</strong> The mix of failure reasons at this rate points to either an outdated master file or a large-scale change (new equipment batch, section reconfiguration) that hasn't been fully registered.`;
    }

  } else if (failRate > 0.05) {
    if (dominant && dominant[0].includes('Pair mismatch') && domPct > 50) {
      analysis = `<strong>Isolated swaps or moves detected.</strong> Most failures are pair mismatches on specific units — both values are known to the master, just not together. This is typical after spot maintenance where individual relays were replaced or repositioned. Field verification of those specific units is recommended.`;
    } else if (dominant && dominant[0].includes('Serial Number') && domPct > 60) {
      analysis = `<strong>Small number of serial replacements.</strong> A targeted set of serial number mismatches with correct circuit names suggests these positions had hardware replaced recently. Likely just needs the master updated for those specific units.`;
    } else {
      analysis = `<strong>Minor discrepancies — likely recent changes.</strong> The ${Math.round(failRate * 100)}% failure rate is low enough that this appears to be normal drift from installations or replacements since the last master update rather than a systemic problem.`;
    }

  } else {
    analysis = `<strong>Data is in good shape.</strong> ${Math.round(failRate * 100)}% failure rate. The small number of failures likely represent recently installed or replaced components not yet registered in the master — routine follow-up items rather than a data quality concern.`;
  }

  if (repeatedTesting) {
    analysis += ` <em>Note: ${uniqueFails} unique failures across ${fails.length} failed rows — some units appear to have been tested more than once.</em>`;
  }

  if (missingPct > 20) {
    analysis += ` <em>Coverage gap: ${missingCount} master circuits (${missingPct}%) were not tested — this may be an incomplete run covering only part of the installation.</em>`;
  } else if (missingCount > 0) {
    analysis += ` <em>${missingCount} master circuit${missingCount > 1 ? 's were' : ' was'} not present in the new file.</em>`;
  }

  document.getElementById('summary-box').innerHTML = analysis;
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
  const totalCount = unique.reduce((sum, r) => sum + r.count, 0);

  const thead = '<tr><th>Count</th><th>Status</th><th>Issue</th><th>Nomenclature</th><th>Serial Number</th><th>Report Number</th></tr>';
  const tbody = unique.map(r => {
    const reportNum = r['Report Number'] || r['Report No'] || r['Report#'] || r['Report No.'] || '';
    return `
    <tr class="mismatch-row">
      <td><span class="badge badge-fail">${r.count}</span></td>
      <td><span class="badge badge-fail">FAIL</span></td>
      <td class="highlight-bad">${r._issue}</td>
      <td class="highlight-bad">${r._circuitName}</td>
      <td class="highlight-bad">${r._serialNumber}</td>
      <td>${reportNum}</td>
    </tr>`;
  }).join('');

  const tfoot = `<tfoot><tr><td><strong>${totalCount}</strong></td><td colspan="5">total failure instances — matches the Mismatches count above</td></tr></tfoot>`;
  container.innerHTML = desc + `<div class="table-wrap"><table><thead>${thead}</thead><tbody>${tbody}</tbody>${tfoot}</table></div>`;
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
async function exportExcel(results, masterData, allHeaders) {
  const wb = new ExcelJS.Workbook();
  const order = masterSortOrders();
  const sorted = sortByMaster(results, order);
  const sortedFails = sorted.filter(r => r._status === 'FAIL');

  const fullHeaders = ['Status', 'Issue', ...allHeaders];

  // ExcelJS columns are 1-based; +1 converts findIndex result
  const trCol = fullHeaders.findIndex(h =>
    typeof h === 'string' && h.toUpperCase().includes('TEST RESULT')
  ) + 1;

  const FILL_RED   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDDD' } };
  const FILL_GREEN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DDFFDD' } };
  const FONT_RED   = { bold: true, color: { argb: 'CC0000' } };
  const FONT_GREEN = { color: { argb: '007744' } };

  function addFullSheet(name, rows) {
    const ws = wb.addWorksheet(name);
    ws.addRow(fullHeaders).font = { bold: true };
    rows.forEach(r => {
      const rowData = [r._status, r._issue, ...allHeaders.map(h => r[h] ?? '')];
      const row = ws.addRow(rowData);
      const isFail = r._status === 'FAIL';
      const fill = isFail ? FILL_RED : FILL_GREEN;
      const font = isFail ? FONT_RED : FONT_GREEN;
      const statusCell = row.getCell(1);
      statusCell.fill = fill;
      statusCell.font = font;
      if (trCol > 0) {
        const trCell = row.getCell(trCol);
        trCell.fill = fill;
        trCell.font = font;
      }
    });
  }

  addFullSheet('Full Data', sorted);
  if (sortedFails.length) addFullSheet('Failures', sortedFails);

  // Unique Failures sheet
  const seenU = new Map();
  sortedFails.forEach(r => {
    const key = `${r._circuitName}|||${r._serialNumber}|||${r._issue}`;
    if (seenU.has(key)) seenU.get(key).count++;
    else seenU.set(key, { ...r, count: 1 });
  });
  const uniqueErrRows = [...seenU.values()];
  if (uniqueErrRows.length) {
    const wsUE = wb.addWorksheet('Unique Failures');
    wsUE.addRow(['Count', 'Status', 'Issue', 'Nomenclature', 'Serial Number', 'Report Number']).font = { bold: true };
    const getReport = r => r['Report Number'] || r['Report No'] || r['Report#'] || r['Report No.'] || '';
    uniqueErrRows.forEach(r => wsUE.addRow([r.count, r._status, r._issue, r._circuitName, r._serialNumber, getReport(r)]));
    const totalCount = uniqueErrRows.reduce((sum, r) => sum + r.count, 0);
    const totalRow = wsUE.addRow([totalCount, '', 'TOTAL failure instances', '', '', '']);
    totalRow.font = { bold: true };
  }

  const newSerials = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = masterData.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));
  if (missing.length) {
    const wsMissing = wb.addWorksheet('Missing From New File');
    wsMissing.addRow(['Circuit Name', 'Serial Number']).font = { bold: true };
    missing.forEach(m => wsMissing.addRow([m.circuitName, m.serialNumber]));
  }

  const total = results.length;
  const fails = results.filter(r => r._status === 'FAIL').length;
  const wsSummary = wb.addWorksheet('Summary');
  [
    ['Relay Data Checker — Validation Report'],
    ['Generated', new Date().toLocaleString()],
    ['Version', VERSION],
    [],
    ['Total Rows', total],
    ['Passed', total - fails],
    ['Failed', fails],
    ['Match Rate', total ? `${Math.round(((total - fails) / total) * 100)}%` : 'N/A'],
    ['Master Serials Missing from New File', missing.length],
  ].forEach(row => wsSummary.addRow(row));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `relay-data-checker_validation_${new Date().toISOString().slice(0,19).replace('T','_').replace(/:/g,'-')}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
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
  lines.push('Count,Status,Issue,Nomenclature,Serial Number,Report Number');
  [...seenC.values()].forEach(r => {
    const rpt = r['Report Number'] || r['Report No'] || r['Report#'] || r['Report No.'] || '';
    lines.push([r.count, r._status, r._issue, r._circuitName, r._serialNumber, rpt].map(escape).join(','));
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
    renderSummary(validationResults, masterData);

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

document.getElementById('dl-excel').addEventListener('click', async () => {
  if (validationResults) await exportExcel(validationResults, masterData, newData.headers);
});

document.getElementById('dl-csv').addEventListener('click', () => {
  if (validationResults) exportCSV(validationResults, masterData, newData.headers);
});

document.getElementById('reset-btn').addEventListener('click', () => {
  masterData = null;
  newData = null;
  validationResults = null;
  renderedTabs = new Set();

  ['master', 'new'].forEach(side => {
    document.getElementById(`status-${side}`).textContent = 'No file loaded';
    document.getElementById(`status-${side}`).className = 'zone-status';
    document.getElementById(`zone-${side}`).classList.remove('loaded');
    document.getElementById(`file-${side}`).value = '';
  });

  document.getElementById('run-btn').disabled = true;
  document.getElementById('results-panel').classList.remove('visible');
  document.getElementById('export-row').style.display = 'none';
  document.getElementById('progress-bar').classList.remove('visible');
  showError('');
  ['exceptions','uniqueerrors','notinmaster','fulldata'].forEach(id => {
    document.getElementById(`tab-${id}`).innerHTML = '';
  });
  document.getElementById('stats-row').innerHTML = '';
  document.getElementById('summary-box').innerHTML = '';
});


// ─── Border Rail Scene Animation ──────────────────────────────────
(function () {
  const canvas = document.getElementById('rail-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const PAD  = 14;   // px from viewport edge to track centre
  const GAP  = 8;    // gap between the two rails
  const TL   = 68;   // train length (along track)
  const TW   = 22;   // train width (perpendicular to track)
  const MAX_V      = 3.2;
  const BRAKE_DIST = 190;
  const STOP_GAP   = 55;

  // 4 signals evenly spread around the perimeter
  const signals = [
    { frac: 0.10, state: 'lunar', timer: rnd() },
    { frac: 0.35, state: 'red',   timer: rnd() + 70 },
    { frac: 0.60, state: 'lunar', timer: rnd() + 140 },
    { frac: 0.85, state: 'red',   timer: rnd() + 30 },
  ];

  function rnd() { return 220 + Math.random() * 430; }

  let trainFrac = 0, trainV = MAX_V;

  // ── Geometry helpers ───────────────────────────────────────────
  function perim() {
    return 2 * (canvas.width - 2 * PAD + canvas.height - 2 * PAD);
  }

  function wrap(t) { return ((t % 1) + 1) % 1; }

  // Forward pixel distance from fraction a to fraction b
  function fwdDist(a, b) {
    let d = b - a;
    if (d < 0) d += 1;
    return d * perim();
  }

  // x, y, angle (radians) at fraction t around the rectangle
  function trackPoint(t) {
    const W = canvas.width, H = canvas.height;
    const w = W - 2 * PAD, h = H - 2 * PAD;
    const p = perim();
    const d = wrap(t) * p;
    if (d < w)         return { x: PAD + d,           y: PAD,           angle: 0 };
    if (d < w + h)     return { x: PAD + w,            y: PAD + (d-w),   angle: Math.PI / 2 };
    if (d < 2*w + h)   return { x: PAD + w - (d-w-h), y: PAD + h,       angle: Math.PI };
                       return { x: PAD,                y: PAD+h-(d-2*w-h), angle: -Math.PI / 2 };
  }

  // ── Update ─────────────────────────────────────────────────────
  function update() {
    signals.forEach(s => {
      if (--s.timer <= 0) {
        s.state = s.state === 'red' ? 'lunar' : 'red';
        s.timer = rnd();
      }
    });

    const nextRed = signals
      .filter(s => s.state === 'red')
      .map(s => ({ s, d: fwdDist(trainFrac, s.frac) }))
      .filter(({ d }) => d > 8 && d < BRAKE_DIST + 100)
      .sort((a, b) => a.d - b.d)[0];

    if (nextRed) {
      const d = nextRed.d;
      if (d < BRAKE_DIST) trainV = Math.max(0, trainV - 0.09);
      if (d <= STOP_GAP)  trainV = 0;
    } else {
      trainV = Math.min(MAX_V, trainV + 0.055);
    }

    trainFrac = wrap(trainFrac + trainV / perim());
  }

  // ── Rounded-rect path (no native roundRect dependency) ─────────
  function rr(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.arcTo(x + w, y,   x + w, y + r,   r);
    ctx.lineTo(x + w, y + h - r);
    ctx.arcTo(x + w, y+h, x+w-r, y + h,   r);
    ctx.lineTo(x + r, y + h);
    ctx.arcTo(x,   y+h, x,   y+h-r,       r);
    ctx.lineTo(x,   y + r);
    ctx.arcTo(x,   y,   x+r, y,            r);
    ctx.closePath();
  }

  // ── Draw track ─────────────────────────────────────────────────
  function drawTrack() {
    const W = canvas.width, H = canvas.height;
    const p = perim();
    const tieStep = 22;

    // Ties (short perpendicular rectangles)
    ctx.fillStyle = 'rgba(20,35,55,0.85)';
    for (let d = 0; d < p; d += tieStep) {
      const pt = trackPoint(d / p);
      ctx.save();
      ctx.translate(pt.x, pt.y);
      ctx.rotate(pt.angle);
      ctx.fillRect(-GAP * 1.6, -2, GAP * 3.2, 4);
      ctx.restore();
    }

    // Two rails (inner & outer rectangles)
    for (const offset of [-GAP / 2, GAP / 2]) {
      const r = PAD + offset;
      ctx.strokeStyle = offset < 0 ? 'rgba(155,180,200,0.75)' : 'rgba(100,130,155,0.65)';
      ctx.lineWidth = 2.2;
      ctx.strokeRect(r, r, W - 2 * r, H - 2 * r);
    }
  }

  // ── Draw one signal ────────────────────────────────────────────
  function drawSignal(s) {
    const pt = trackPoint(s.frac);
    // Inward perpendicular (right-hand side of direction of travel = toward page centre)
    const ipx = -Math.sin(pt.angle);
    const ipy =  Math.cos(pt.angle);

    const base = { x: pt.x + ipx * (GAP + 4),  y: pt.y + ipy * (GAP + 4) };
    const tip  = { x: pt.x + ipx * (GAP + 22), y: pt.y + ipy * (GAP + 22) };

    // Mast
    ctx.strokeStyle = 'rgba(50,70,90,0.9)'; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(base.x, base.y); ctx.lineTo(tip.x, tip.y); ctx.stroke();

    const now = Date.now();
    const lunarOn = Math.floor(now / 545) % 2 === 0; // 55 flashes/min

    if (s.state === 'lunar') {
      if (lunarOn) {
        // Outer glow
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 8, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(180,215,255,0.15)'; ctx.fill();
        // Light
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 4.5, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(220,238,255,0.95)'; ctx.fill();
      } else {
        ctx.beginPath(); ctx.arc(tip.x, tip.y, 4.5, 0, Math.PI * 2);
        ctx.fillStyle = 'rgba(220,238,255,0.06)'; ctx.fill();
      }
    } else {
      // Red glow
      ctx.beginPath(); ctx.arc(tip.x, tip.y, 9, 0, Math.PI * 2);
      ctx.fillStyle = 'rgba(255,30,30,0.12)'; ctx.fill();
      // Red light
      ctx.beginPath(); ctx.arc(tip.x, tip.y, 4.5, 0, Math.PI * 2);
      ctx.fillStyle = '#ff2020'; ctx.fill();
    }
  }

  // ── Draw train ─────────────────────────────────────────────────
  function drawTrain() {
    const pt = trackPoint(trainFrac);
    ctx.save();
    ctx.translate(pt.x, pt.y);
    ctx.rotate(pt.angle);

    // Body centred on track point, oriented along angle
    const bx = -TL / 2, by = -TW / 2;

    // Shadow / depth
    ctx.fillStyle = 'rgba(0,0,0,0.25)';
    rr(bx + 2, by + 2, TL, TW, 3); ctx.fill();

    // Body gradient (silver)
    const g = ctx.createLinearGradient(bx, by, bx, by + TW);
    g.addColorStop(0,   '#ccd8e4');
    g.addColorStop(0.4, '#b0c2d0');
    g.addColorStop(1,   '#808fa0');
    ctx.fillStyle = g; rr(bx, by, TL, TW, 3); ctx.fill();

    // RTA red stripe
    ctx.fillStyle = '#c8102e';
    ctx.fillRect(bx, by + Math.round(TW * 0.60), TL, 4);

    // Roof shine
    const shine = ctx.createLinearGradient(bx, by, bx, by + TW * 0.35);
    shine.addColorStop(0, 'rgba(255,255,255,0.35)');
    shine.addColorStop(1, 'rgba(255,255,255,0)');
    ctx.fillStyle = shine; ctx.fillRect(bx + 3, by, TL - 6, TW * 0.35);

    // Windows
    ctx.fillStyle = 'rgba(10,28,58,0.9)';
    for (let i = 0; 5 + i * 11 + 8 <= TL - 5; i++) {
      rr(bx + 5 + i * 11, by + 3, 8, 6, 1); ctx.fill();
    }

    // Bogies
    ctx.fillStyle = '#111a24';
    rr(bx + 5,      by + TW, 22, 4, 1); ctx.fill();
    rr(bx + TL - 27, by + TW, 22, 4, 1); ctx.fill();

    // Wheels
    [[bx + 12], [bx + 22], [bx + TL - 22], [bx + TL - 12]].forEach(([wx]) => {
      const wy = by + TW + 6;
      ctx.beginPath(); ctx.arc(wx, wy, 4.5, 0, Math.PI * 2);
      ctx.fillStyle = '#111a24'; ctx.fill();
      ctx.strokeStyle = '#334455'; ctx.lineWidth = 1; ctx.stroke();
      ctx.beginPath(); ctx.arc(wx, wy, 1.5, 0, Math.PI * 2);
      ctx.fillStyle = '#445566'; ctx.fill();
    });

    ctx.restore();
  }

  // ── Draw frame ─────────────────────────────────────────────────
  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    drawTrack();
    signals.forEach(drawSignal);
    drawTrain();
  }

  // ── Resize & loop ──────────────────────────────────────────────
  function resize() {
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
  }

  window.addEventListener('resize', resize);
  resize();

  function frame() { update(); draw(); requestAnimationFrame(frame); }
  requestAnimationFrame(frame);
}());
