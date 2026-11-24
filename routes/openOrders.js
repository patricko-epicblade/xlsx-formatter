const ExcelJS = require('exceljs');
const { DATE_FMT, isoToDate, styleHeaderRow, safeFileName, makeRouter } = require('./common');

const MONEY = '$#,##0.00;[Red]($#,##0.00)';
const NUM = '#,##0.##;[Red](#,##0.##)';
const BAND_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const SUB_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const WARN_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
const TYPES = ['text', 'money', 'number', 'date'];

// Body: { title, asOfLabel, outputFileName, columns: [{header, type}], rows: [{customer, status, warn?, cells: []}], notes?: [string] }
function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.columns) || body.columns.length === 0) return 'columns must be a non-empty array';
  for (let i = 0; i < body.columns.length; i++) {
    const c = body.columns[i];
    if (!c || typeof c.header !== 'string') return `columns[${i}].header must be a string`;
    if (c.type != null && !TYPES.includes(c.type)) return `columns[${i}].type must be one of ${TYPES.join(', ')}`;
  }
  if (!Array.isArray(body.rows)) return 'rows must be an array';
  for (let i = 0; i < body.rows.length; i++) {
    const r = body.rows[i];
    if (!r || typeof r !== 'object' || Array.isArray(r)) return `rows[${i}] must be an object`;
    if (typeof r.customer !== 'string' || !r.customer.trim()) return `rows[${i}].customer must be a non-empty string`;
    if (r.status != null && typeof r.status !== 'string') return `rows[${i}].status must be a string`;
    if (!Array.isArray(r.cells) || r.cells.length !== body.columns.length) return `rows[${i}].cells must have ${body.columns.length} entries`;
    for (const v of r.cells) if (v != null && typeof v !== 'string' && typeof v !== 'number') return `rows[${i}].cells may only hold strings, numbers or null`;
  }
  return null;
}

async function build(body) {
  const cols = [{ header: 'Customer', type: 'text', w: 28 }, { header: 'Estimate status', type: 'text', w: 16 },
    ...body.columns.map((c) => ({ header: c.header, type: c.type || 'text', w: c.type === 'money' ? 15 : c.type === 'date' ? 13 : c.type === 'number' ? 11 : 20 }))];
  const N = cols.length;
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Open Orders', {
    pageSetup: { orientation: 'landscape', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }, printTitlesRow: '1:3' },
  });
  const groups = new Map();
  for (const r of body.rows) { const k = r.customer.trim(); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); }
  const names = [...groups.keys()].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
  const byStatus = {};
  for (const r of body.rows) { const s = r.status || 'Unknown'; byStatus[s] = (byStatus[s] || 0) + 1; }

  sheet.addRow([body.title || 'Open Orders - All Customers']);
  sheet.mergeCells(1, 1, 1, N);
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getRow(1).height = 24;
  sheet.addRow([`${body.asOfLabel ? `As of ${body.asOfLabel} | ` : ''}${body.rows.length} open lines | ${names.length} customers | ${Object.entries(byStatus).map(([s, n]) => `${s} ${n}`).join(', ')}`]);
  sheet.mergeCells(2, 1, 2, N);
  sheet.getCell(2, 1).font = { italic: true, size: 11 };
  const header = sheet.addRow(cols.map((c) => c.header));
  styleHeaderRow(header);
  header.height = 32;
  for (let n = 1; n <= N; n++) header.getCell(n).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cols.forEach((c, i) => { sheet.getColumn(i + 1).width = c.w; });
  if (!body.rows.length) sheet.addRow(['No open orders found']).font = { italic: true };

  const fmtCell = (cell, type) => {
    if (type === 'money') { cell.numFmt = MONEY; cell.alignment = { horizontal: 'right', vertical: 'top' }; }
    else if (type === 'number') { cell.numFmt = NUM; cell.alignment = { horizontal: 'right', vertical: 'top' }; }
    else if (type === 'date') { cell.numFmt = DATE_FMT; cell.alignment = { horizontal: 'center', vertical: 'top' }; }
    else cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  };
  const hasMoney = cols.some((c) => c.type === 'money');
  const sums = (rows) => cols.map((c, i) => (c.type === 'money' && i >= 2 ? Math.round(rows.reduce((t, r) => t + (typeof r.cells[i - 2] === 'number' ? r.cells[i - 2] : 0), 0) * 100) / 100 : null));
  const totalRow = (label, rows, fill, border) => {
    const row = sheet.addRow(sums(rows));
    row.getCell(1).value = label;
    row.getCell(2).value = `${rows.length} ${rows.length === 1 ? 'line' : 'lines'}`;
    for (let n = 1; n <= N; n++) {
      const c = row.getCell(n);
      c.font = { bold: true };
      if (fill) c.fill = SUB_FILL;
      c.border = border;
      if (cols[n - 1].type === 'money') { c.numFmt = MONEY; c.alignment = { horizontal: 'right' }; }
    }
  };

  for (const name of names) {
    const rows = groups.get(name);
    const band = sheet.addRow([`${name} (${rows.length} ${rows.length === 1 ? 'line' : 'lines'})`]);
    sheet.mergeCells(band.number, 1, band.number, N);
    band.getCell(1).font = { bold: true, size: 12 };
    band.getCell(1).fill = BAND_FILL;
    for (const r of rows) {
      const vals = [name, r.status || '', ...cols.slice(2).map((c, i) => {
        const v = r.cells[i];
        if (v == null || v === '') return null;
        if (c.type === 'date' && typeof v === 'string') return isoToDate(v) || v;
        return v;
      })];
      const row = sheet.addRow(vals);
      cols.forEach((c, i) => fmtCell(row.getCell(i + 1), c.type));
      if (r.warn) for (let n = 1; n <= N; n++) row.getCell(n).fill = WARN_FILL;
    }
    if (hasMoney) totalRow(`Total ${name}`, rows, true, { top: { style: 'thin' } });
  }
  if (body.rows.length) {
    if (hasMoney) totalRow('GRAND TOTAL', body.rows, false, { top: { style: 'medium' } });
    for (const t of body.notes || []) {
      const note = sheet.addRow([t]);
      sheet.mergeCells(note.number, 1, note.number, N);
      note.getCell(1).font = { italic: true, size: 10 };
      note.getCell(1).alignment = { wrapText: true, vertical: 'top' };
    }
  }
  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  return { buffer: await wb.xlsx.writeBuffer(), fileName: safeFileName(body.outputFileName, 'Open_Orders_All_Customers.xlsx') };
}

module.exports = makeRouter({ validate, build });
