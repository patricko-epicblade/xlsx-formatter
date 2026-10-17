const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const {
  MONEY_FMT, DATE_FMT, isoToDate, styleHeaderRow, safeFileName, validateLines, makeRouter,
} = require('./common');

const LOGO_PATH = path.join(__dirname, '..', 'assets', 'logo.png');
const HEADERS = [
  'Invoice #', 'Invoice Date', 'Due Date', 'Days Past Due', 'Open Balance', 'PO #', 'Terms', 'Lot #',
];
const WIDTHS = [14, 14, 14, 15, 16, 14, 14, 12];
const GREY_ITALIC = { italic: true, color: { argb: 'FF808080' } };

// Merge a text line across columns 1..lastCol, wrap it, and size the row to fit.
// Column width units are roughly one character, so lines = ceil(chars / merged width).
function wrappedLine(sheet, row, lastCol, text, font) {
  const width = WIDTHS.slice(0, lastCol).reduce((a, b) => a + b, 0);
  const str = String(text);
  const lines = Math.max(1, Math.ceil(str.length / (width * 0.95)));
  sheet.mergeCells(row, 1, row, lastCol);
  const cell = sheet.getCell(row, 1);
  cell.value = str;
  if (font) cell.font = font;
  cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  const r = sheet.getRow(row);
  r.height = Math.max(r.height || 0, lines * 15 + (lines > 1 ? 2 : 0));
}

function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  const h = body.header;
  if (!h || typeof h !== 'object') return 'header must be an object';
  for (const k of ['addressLines', 'remitToLines']) {
    if (h[k] != null && !Array.isArray(h[k])) return `header.${k} must be an array`;
  }
  const c = body.customer;
  if (!c || typeof c !== 'object') return 'customer must be an object';
  if (typeof c.customer !== 'string' || !c.customer) return 'customer.customer must be a non-empty string';
  if (typeof c.netPastDue !== 'number' || !isFinite(c.netPastDue)) return 'customer.netPastDue must be a number';
  return validateLines(c.lines, 'customer');
}

async function build(body) {
  const header = body.header;
  const cust = body.customer;
  const nonBlank = (arr) => (arr || []).filter((s) => s != null && String(s).trim() !== '');
  const addressLines = nonBlank(header.addressLines);
  const remitToLines = nonBlank(header.remitToLines);
  const ncol = HEADERS.length;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Statement', {
    pageSetup: {
      orientation: 'portrait',
      paperSize: 1,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
    },
  });
  WIDTHS.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });

  // Top-left: logo if the file exists, else company name
  const hasLogo = fs.existsSync(LOGO_PATH);
  sheet.mergeCells(1, 1, 1, 4);
  if (hasLogo) {
    const imageId = workbook.addImage({ filename: LOGO_PATH, extension: 'png' });
    // Epic Blade logo is ~2.3:1 (assets/logo.png, 640x278); keep that ratio so it isn't squashed.
    sheet.addImage(imageId, { tl: { col: 0.1, row: 0.1 }, ext: { width: 172, height: 75 } });
    sheet.getRow(1).height = 60;
  } else {
    const c = sheet.getCell(1, 1);
    c.value = header.companyName || '';
    c.font = { bold: true, size: 18 };
    c.alignment = { horizontal: 'left', vertical: 'middle' };
    sheet.getRow(1).height = 30;
  }
  // Top-right: STATEMENT + date
  sheet.mergeCells(1, 5, 1, ncol);
  const t = sheet.getCell(1, 5);
  t.value = 'STATEMENT';
  t.font = { bold: true, size: 20 };
  t.alignment = { horizontal: 'right', vertical: 'middle' };
  sheet.mergeCells(2, 5, 2, ncol);
  const d = sheet.getCell(2, 5);
  d.value = `Date: ${body.asOfLabel || ''}`;
  d.alignment = { horizontal: 'right', vertical: 'middle' };

  // Address lines below the logo / name
  let row = 2;
  if (addressLines.length === 0) {
    sheet.mergeCells(row, 1, row, 4);
    const a = sheet.getCell(row, 1);
    a.value = '[address to come]';
    a.font = GREY_ITALIC;
    row++;
  } else {
    for (const line of addressLines) {
      // Row 2 shares the line with the "Date:" cell in E:H, so it can only span A:D.
      wrappedLine(sheet, row, row === 2 ? 4 : ncol, line);
      row++;
    }
  }

  row++; // blank
  sheet.getCell(row, 1).value = 'Bill to:';
  row++;
  sheet.getCell(row, 1).value = cust.customer;
  sheet.getCell(row, 1).font = { bold: true };
  row += 2; // blank row between

  const headerRow = sheet.getRow(row);
  HEADERS.forEach((h, i) => { headerRow.getCell(i + 1).value = h; });
  styleHeaderRow(headerRow);
  row++;

  for (const l of cust.lines) {
    const r = sheet.getRow(row++);
    const hasDays = typeof l.daysPastDue === 'number';
    const vals = [
      l.num == null ? '' : String(l.num),
      isoToDate(l.date),
      isoToDate(l.dueDate),
      hasDays ? l.daysPastDue : null,
      l.openBalance,
      l.poNumber || '',
      l.terms || '',
      l.lotNumber == null ? '' : String(l.lotNumber),
    ];
    vals.forEach((v, i) => { r.getCell(i + 1).value = v; });
    r.getCell(2).numFmt = DATE_FMT;
    r.getCell(3).numFmt = DATE_FMT;
    r.getCell(5).numFmt = MONEY_FMT;
    [1, 2, 3, 4, 6, 7, 8].forEach((n) => {
      r.getCell(n).alignment = { horizontal: 'center', vertical: 'middle' };
    });
    r.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
  }

  const tot = sheet.getRow(row++);
  tot.getCell(1).value = 'Total past due';
  tot.getCell(5).value = cust.netPastDue;
  tot.getCell(5).numFmt = MONEY_FMT;
  tot.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
  for (let n = 1; n <= ncol; n++) {
    tot.getCell(n).font = { bold: true };
    tot.getCell(n).border = { top: { style: 'thin' } };
  }

  row++; // blank
  sheet.getCell(row, 1).value = 'Remit to / how to pay:';
  sheet.getCell(row, 1).font = { bold: true };
  row++;
  if (remitToLines.length === 0) {
    sheet.getCell(row, 1).value = '[payment instructions to come]';
    sheet.getCell(row, 1).font = GREY_ITALIC;
  } else {
    for (const line of remitToLines) {
      wrappedLine(sheet, row++, ncol, line);
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Statement.xlsx') };
}

module.exports = makeRouter({ validate, build });
