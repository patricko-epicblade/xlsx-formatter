const ExcelJS = require('exceljs');
const { styleHeaderRow, safeFileName, makeRouter } = require('./common');

// Body: { title, instructions?, outputFileName, names: [string], markHeader?, choices?: [string] }
function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.names)) return 'names must be an array';
  for (let i = 0; i < body.names.length; i++) {
    if (typeof body.names[i] !== 'string' || !body.names[i].trim()) return `names[${i}] must be a non-empty string`;
  }
  if (body.choices != null && (!Array.isArray(body.choices) || body.choices.some((c) => typeof c !== 'string' || /[,"]/.test(c)))) {
    return 'choices must be an array of plain strings';
  }
  return null;
}

async function build(body) {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Customers', { pageSetup: { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '1:3' } });
  sheet.addRow([body.title || 'Customer list']);
  sheet.mergeCells(1, 1, 1, 3);
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getRow(1).height = 24;
  sheet.addRow([body.instructions || '']);
  sheet.mergeCells(2, 1, 2, 3);
  sheet.getCell(2, 1).font = { italic: true };
  sheet.getCell(2, 1).alignment = { wrapText: true, vertical: 'top' };
  sheet.getRow(2).height = 32;
  styleHeaderRow(sheet.addRow(['#', 'Customer', body.markHeader || 'Whose customer?']));
  [6, 48, 26].forEach((w, i) => { sheet.getColumn(i + 1).width = w; });
  const seen = new Map();
  for (const n of body.names) { const t = n.trim(); if (!seen.has(t.toLowerCase())) seen.set(t.toLowerCase(), t); }
  const names = [...seen.values()].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
  names.forEach((n, i) => {
    const r = sheet.addRow([i + 1, n, null]);
    r.getCell(1).alignment = { horizontal: 'center' };
    if (body.choices && body.choices.length) {
      r.getCell(3).dataValidation = { type: 'list', allowBlank: true, formulae: [`"${body.choices.join(',')}"`] };
    }
    r.getCell(3).border = { bottom: { style: 'hair' }, left: { style: 'thin' }, right: { style: 'thin' } };
  });
  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  sheet.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3, column: 3 } };
  return { buffer: await wb.xlsx.writeBuffer(), fileName: safeFileName(body.outputFileName, 'Customer_List.xlsx') };
}

module.exports = makeRouter({ validate, build });
