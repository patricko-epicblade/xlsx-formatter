const ExcelJS = require('exceljs');
const {
  DATE_FMT, isoToDate, styleHeaderRow, safeFileName, makeRouter,
} = require('./common');

const HEADERS = [
  'Customer', 'FA #', 'Lot #', 'Description', 'Amount', 'Ordered', 'Shipped', 'Status', 'Outcome',
];
const WIDTHS = [30, 12, 16, 50, 14, 13, 28, 12, 22];
const AMOUNT_FMT = '$#,##0.00';
const REMAKE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
const OPTIONAL_STRINGS = ['estimateNum', 'lot', 'description', 'ordered', 'shipped', 'status', 'outcome'];

function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.rows)) return 'rows must be an array';
  for (let i = 0; i < body.rows.length; i++) {
    const r = body.rows[i];
    if (!r || typeof r !== 'object' || Array.isArray(r)) return `rows[${i}] must be an object`;
    if (typeof r.customer !== 'string' || !r.customer.trim()) return `rows[${i}].customer must be a non-empty string`;
    if (typeof r.faNumber !== 'string' || !r.faNumber.trim()) return `rows[${i}].faNumber must be a non-empty string`;
    if (typeof r.amount !== 'number' || !isFinite(r.amount)) return `rows[${i}].amount must be a number`;
    if (typeof r.remake !== 'boolean') return `rows[${i}].remake must be a boolean`;
    for (const k of OPTIONAL_STRINGS) {
      if (r[k] != null && typeof r[k] !== 'string') return `rows[${i}].${k} must be a string`;
    }
  }
  return null;
}

async function build(body) {
  const companyName = body.companyName || 'Customer';
  const asOfLabel = body.asOfLabel || '';
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Closed FA', {
    pageSetup: {
      orientation: 'landscape',
      paperSize: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
      printTitlesRow: '1:3',
    },
  });

  sheet.addRow([`Closed FA (First Article) Report - ${companyName}`]);
  sheet.addRow([`As of ${asOfLabel}`]);
  sheet.mergeCells(1, 1, 1, HEADERS.length);
  sheet.mergeCells(2, 1, 2, HEADERS.length);
  sheet.getRow(1).height = 24;
  sheet.getRow(2).height = 20;
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 1).font = { italic: true, size: 12 };
  sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };

  styleHeaderRow(sheet.addRow(HEADERS));
  WIDTHS.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });

  if (body.rows.length === 0) {
    const r = sheet.addRow(['No FA trials found']);
    r.font = { italic: true };
  }

  const str = (v) => (v == null ? '' : String(v));
  let prevCustomer = null;
  for (const l of body.rows) {
    const r = sheet.addRow([
      l.customer, l.faNumber, str(l.lot), str(l.description), l.amount,
      isoToDate(l.ordered), str(l.shipped), str(l.status), str(l.outcome),
    ]);
    for (let n = 1; n <= HEADERS.length; n++) {
      const c = r.getCell(n);
      c.alignment = { horizontal: 'left', vertical: 'middle' };
      if (l.remake) c.fill = REMAKE_FILL;
      if (l.customer !== prevCustomer) c.border = { top: { style: 'thin' } };
    }
    prevCustomer = l.customer;
    r.getCell(4).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.getCell(7).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.getCell(5).numFmt = AMOUNT_FMT;
    r.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
    r.getCell(6).numFmt = DATE_FMT;
    r.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  sheet.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: HEADERS.length },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Closed_FA_Report.xlsx') };
}

module.exports = makeRouter({ validate, build });
