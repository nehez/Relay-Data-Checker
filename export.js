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

  function autoFitColumns(ws) {
    const widths = {};
    ws.eachRow(row => {
      row.eachCell({ includeEmpty: false }, (cell, colIdx) => {
        const len = cell.value != null ? String(cell.value).length : 0;
        widths[colIdx] = Math.max(widths[colIdx] || 0, len);
      });
    });
    Object.entries(widths).forEach(([colIdx, len]) => {
      ws.getColumn(Number(colIdx)).width = Math.min(len + 3, 60);
    });
  }

  function fitSerialNomenclature(ws, headers) {
    headers.forEach((h, i) => {
      const hl = String(h).toLowerCase();
      if (hl === 'nomenclature' || hl === 'serial number') {
        const colIdx = i + 1;
        let maxLen = String(h).length;
        ws.getColumn(colIdx).eachCell({ includeEmpty: false }, cell => {
          const len = cell.value != null ? String(cell.value).length : 0;
          if (len > maxLen) maxLen = len;
        });
        ws.getColumn(colIdx).width = Math.min(maxLen + 3, 60);
      }
    });
  }

  function alignRight(ws) {
    ws.eachRow(row => {
      row.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = { horizontal: 'right' };
      });
    });
  }

  function addFullSheet(name, rows) {
    const ws = wb.addWorksheet(name);
    ws.addRow(fullHeaders).font = { bold: true };
    const trHeaderName = trCol > 0 ? fullHeaders[trCol - 1] : null;
    rows.forEach(r => {
      const rowData = [r._status, r._issue, ...allHeaders.map(h => r[h] ?? '')];
      const row = ws.addRow(rowData);
      const isFail = r._status === 'FAIL';
      const statusCell = row.getCell(1);
      statusCell.fill = isFail ? FILL_RED : FILL_GREEN;
      statusCell.font = isFail ? FONT_RED : FONT_GREEN;
      if (trHeaderName) {
        const trCell = row.getCell(trCol);
        const trVal = String(r[trHeaderName] ?? '').toUpperCase().trim();
        if (trVal.includes('FAIL')) {
          trCell.fill = FILL_RED; trCell.font = FONT_RED;
        } else if (trVal.includes('PASS')) {
          trCell.fill = FILL_GREEN; trCell.font = FONT_GREEN;
        }
      }
    });
    fitSerialNomenclature(ws, fullHeaders);
    alignRight(ws);
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
    wsUE.addRow(['Count', 'Status', 'Issue', 'Nomenclature', 'Serial Number', 'Report Number', 'Comments']).font = { bold: true };
    const getReport  = r => r['Report Number'] || r['Report No'] || r['Report#'] || r['Report No.'] || '';
    const getComment = r => r['Comments'] || r['Comment'] || r['COMMENTS'] || r['comment'] || '';
    uniqueErrRows.forEach(r => wsUE.addRow([r.count, r._status, r._issue, r._circuitName, r._serialNumber, getReport(r), getComment(r)]));
    const totalCount = uniqueErrRows.reduce((sum, r) => sum + r.count, 0);
    const totalRow = wsUE.addRow([totalCount, '', 'TOTAL failure instances', '', '', '', '']);
    totalRow.font = { bold: true };
    autoFitColumns(wsUE);
    alignRight(wsUE);
  }

  const newSerials = new Set(results.map(r => String(r['Serial Number'] ?? '').trim().toUpperCase()));
  const missing = masterData.filter(m => !newSerials.has(m.serialNumber.toUpperCase()));
  if (missing.length) {
    const wsMissing = wb.addWorksheet('Missing From New File');
    wsMissing.addRow(['Circuit Name', 'Serial Number']).font = { bold: true };
    missing.forEach(m => wsMissing.addRow([m.circuitName, m.serialNumber]));
    autoFitColumns(wsMissing);
    alignRight(wsMissing);
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
  autoFitColumns(wsSummary);
  alignRight(wsSummary);

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
