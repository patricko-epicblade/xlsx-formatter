const ExcelJS = require('exceljs');
const {
  MONEY_FMT, DATE_FMT, isoToDate, styleHeaderRow, safeFileName, validateLines, makeRouter,
} = require('./common');

const HEADERS = [
  'Customer', 'Invoice #', 'Invoice Date', 'Due Date', 'Days Past Due',
  'Open Balance', 'PO #', 'Terms', 'Lot #',
];
const WIDTHS = [36, 14, 14, 14, 15, 16, 16, 14, 14];

function ageFill(days) {
  if (typeof days !== 'number' || days < 1) return null;
  const argb = days <= 30 ? 'FFFFF2CC' : days <= 60 ? 'FFF8CBAD' : 'FFFFC7CE';
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.customers)) return 'customers must be an array';
  for (let i = 0; i < body.customers.length; i++) {
    const c = body.customers[i];
    if (!c || typeof c !== 'object') return `customers[${i}] must be an object`;
    if (typeof c.customer !== 'string' || !c.customer) return `customers[${i}].customer must be a non-empty string`;
    if (typeof c.netPastDue !== 'number' || !isFinite(c.netPastDue)) return `customers[${i}].netPastDue must be a number`;
    const e = validateLines(c.lines, `customers[${i}]`);
    if (e) return e;
  }
  return null;
}

async function build(body) {
  const companyName = body.companyName || 'Customer';
  const asOfLabel = body.asOfLabel || '';
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Overdue Invoices');

  sheet.addRow([`Overdue Invoices - ${companyName}`]);
  sheet.addRow([`As of ${asOfLabel}`]);
  sheet.mergeCells(1, 1, 1, HEADERS.length);
  sheet.mergeCells(2, 1, 2, HEADERS.length);
  sheet.getRow(1).height = 24;
  sheet.getRow(2).height = 20;
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 1).font = { italic: true, size: 12 };
  sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };

  const headerRow = sheet.addRow(HEADERS);
  styleHeaderRow(headerRow);
  WIDTHS.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });

  if (body.customers.length === 0) {
    const r = sheet.addRow(['No overdue invoices']);
    r.font = { italic: true };
  }

  let grand = 0;
  let invoiceCount = 0;
  for (const c of body.customers) {
    for (const l of c.lines) {
      const hasDays = typeof l.daysPastDue === 'number';
      if (hasDays) invoiceCount++;
      const r = sheet.addRow([
        c.customer,
        l.num == null ? '' : String(l.num),
        isoToDate(l.date),
        isoToDate(l.dueDate),
        hasDays ? l.daysPastDue : null,
        l.openBalance,
        l.poNumber || '',
        l.terms || '',
        l.lotNumber == null ? '' : String(l.lotNumber),
      ]);
      [2, 3, 4, 5, 7, 8, 9].forEach((n) => {
        r.getCell(n).alignment = { horizontal: 'center', vertical: 'middle' };
      });
      r.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      r.getCell(3).numFmt = DATE_FMT;
      r.getCell(4).numFmt = DATE_FMT;
      r.getCell(6).numFmt = MONEY_FMT;
      r.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
      const fill = ageFill(l.daysPastDue);
      if (fill) r.getCell(5).fill = fill;
    }
    const sub = sheet.addRow([`Total for ${c.customer}`, null, null, null, null, c.netPastDue, c.note || '']);
    sub.font = { bold: true };
    sub.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    sub.getCell(6).numFmt = MONEY_FMT;
    sub.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
    sub.getCell(7).font = { bold: true, italic: true };
    sub.getCell(7).alignment = { horizontal: 'left', vertical: 'middle' };
    for (let n = 1; n <= HEADERS.length; n++) {
      sub.getCell(n).border = { top: { style: 'thin' } };
    }
    grand += c.netPastDue;
  }

  if (body.customers.length > 0) {
    const g = sheet.addRow([
      'GRAND TOTAL', null, null, null, `${invoiceCount} invoices`,
      Math.round(grand * 100) / 100,
    ]);
    g.font = { bold: true };
    g.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    g.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    g.getCell(6).numFmt = MONEY_FMT;
    g.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
    for (let n = 1; n <= HEADERS.length; n++) {
      g.getCell(n).border = { top: { style: 'thin' }, bottom: { style: 'double' } };
    }
  }

  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  sheet.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: HEADERS.length },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Overdue_Invoices.xlsx') };
}

module.exports = makeRouter({ validate, build });
