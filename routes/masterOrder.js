const ExcelJS = require('exceljs');
const {
  DATE_FMT, isoToDate, styleHeaderRow, safeFileName, makeRouter,
} = require('./common');

const AMOUNT_FMT = '$#,##0.00;[Red]($#,##0.00)';
const QTY_FMT = '#,##0.##;[Red](#,##0.##)';
const SUBTOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const OVER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8CBAD' } };
const NUMBER_FIELDS = ['orderedAmount', 'shippedAmount', 'remainingAmount'];
const OPTIONAL_STRINGS = ['orderNum', 'lot', 'description', 'status', 'ordered', 'lastShipped', 'note', 'section'];

function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.rows)) return 'rows must be an array';
  for (let i = 0; i < body.rows.length; i++) {
    const r = body.rows[i];
    if (!r || typeof r !== 'object' || Array.isArray(r)) return `rows[${i}] must be an object`;
    if (typeof r.customer !== 'string' || !r.customer.trim()) return `rows[${i}].customer must be a non-empty string`;
    for (const k of NUMBER_FIELDS) {
      if (typeof r[k] !== 'number' || !isFinite(r[k])) return `rows[${i}].${k} must be a number`;
    }
    for (const k of OPTIONAL_STRINGS) {
      if (r[k] != null && typeof r[k] !== 'string') return `rows[${i}].${k} must be a string`;
    }
    for (const k of ['percentUsed', 'orderedQty', 'shippedQty', 'remainingQty', 'shippedCount', 'daysSinceShipment']) {
      if (r[k] != null && (typeof r[k] !== 'number' || !isFinite(r[k]))) return `rows[${i}].${k} must be a number`;
    }
    if (r.sharedLot != null && typeof r.sharedLot !== 'boolean') return `rows[${i}].sharedLot must be a boolean`;
    if (r.overShipped != null && typeof r.overShipped !== 'boolean') return `rows[${i}].overShipped must be a boolean`;
  }
  return null;
}

const sumOf = (rows, k) => Math.round(rows.reduce((t, r) => t + (typeof r[k] === 'number' ? r[k] : 0), 0) * 100) / 100;

async function build(body) {
  const companyName = body.companyName || 'Customer';
  const asOfLabel = body.asOfLabel || '';
  const hasQty = body.hasQty === true;
  const workbook = new ExcelJS.Workbook();

  // Column list; the quantity columns exist only when the QuickBooks report had a Qty column.
  const cols = [
    { h: 'Customer', w: 28 }, { h: 'Order #', w: 12 }, { h: 'Lot #', w: 16 }, { h: 'Description', w: 50 },
    { h: 'Status', w: 11 }, { h: 'Order date', w: 13 },
    { h: 'Ordered $', w: 15, k: 'orderedAmount', fmt: AMOUNT_FMT }, { h: 'Shipped $', w: 15, k: 'shippedAmount', fmt: AMOUNT_FMT },
    { h: 'Remaining $', w: 15, k: 'remainingAmount', fmt: AMOUNT_FMT }, { h: '% used', w: 9, k: 'percentUsed', fmt: '0.0"%"' },
  ];
  if (hasQty) {
    cols.push({ h: 'Ordered qty', w: 12, k: 'orderedQty', fmt: QTY_FMT }, { h: 'Shipped qty', w: 12, k: 'shippedQty', fmt: QTY_FMT },
      { h: 'Remaining qty', w: 13, k: 'remainingQty', fmt: QTY_FMT });
  }
  cols.push({ h: 'Shipments', w: 10, k: 'shippedCount', fmt: '0' }, { h: 'Last shipped', w: 13 },
    { h: 'Days since shipment', w: 12, k: 'daysSinceShipment', fmt: '0' }, { h: 'Note', w: 45 });
  const N = cols.length;
  const at = (h) => cols.findIndex((c) => c.h === h) + 1;

  const sheet = workbook.addWorksheet('Master Orders', {
    pageSetup: {
      orientation: 'landscape', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }, printTitlesRow: '1:3',
    },
  });
  sheet.addRow([`Open Master Orders Report - ${companyName}`]);
  sheet.addRow([`As of ${asOfLabel}`]);
  sheet.mergeCells(1, 1, 1, N);
  sheet.mergeCells(2, 1, 2, 2);
  sheet.mergeCells(2, 3, 2, N);
  sheet.getRow(1).height = 24;
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 1).font = { italic: true, size: 12 };
  sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  const overCount = body.rows.filter((r) => r.overShipped).length;
  const summary = [`${body.rows.length} open master ${body.rows.length === 1 ? 'order' : 'orders'}`,
    `${sumOf(body.rows, 'remainingAmount').toLocaleString('en-US', { style: 'currency', currency: 'USD' })} remaining`];
  const stat = (st) => { const l = body.rows.filter((r) => String(r.status || '').toLowerCase() === st.toLowerCase()); return l.length ? `${st} ${l.length}: ${sumOf(l, 'remainingAmount').toLocaleString('en-US', { style: 'currency', currency: 'USD' })} remaining` : null; };
  for (const st of ['Accepted', 'Pending']) { const t = stat(st); if (t) summary.push(t); }
  if (overCount) summary.push(`${overCount} over-shipped (red)`);
  if (!hasQty) summary.push(body.qtyNote || 'No Qty column in the QuickBooks report, so quantities are left out.');
  sheet.getCell(2, 3).value = summary.join(' | ');
  sheet.getCell(2, 3).font = { size: 11 };
  sheet.getCell(2, 3).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  sheet.getRow(2).height = Math.max(20, 16 * Math.ceil(sheet.getCell(2, 3).value.length / Math.max(60, cols.slice(2).reduce((t, c) => t + c.w, 0) * 0.9)));

  const headerRow = sheet.addRow(cols.map((c) => c.h));
  styleHeaderRow(headerRow);
  headerRow.height = 32;
  for (let n = 1; n <= N; n++) headerRow.getCell(n).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cols.forEach((c, i) => { sheet.getColumn(i + 1).width = c.w; });

  if (body.rows.length === 0) sheet.addRow(['No open master orders found']).font = { italic: true };

  const str = (v) => (v == null ? '' : String(v));
  const totalsRow = (label, count, list, fill, border) => {
    const row = sheet.addRow([]);
    row.getCell(1).value = label;
    row.getCell(2).value = `${count} ${count === 1 ? 'order' : 'orders'}`;
    for (const c of cols) {
      if (!c.k || !['orderedAmount', 'shippedAmount', 'remainingAmount', 'orderedQty', 'shippedQty', 'remainingQty'].includes(c.k)) continue;
      row.getCell(at(c.h)).value = sumOf(list, c.k);
      row.getCell(at(c.h)).numFmt = c.fmt;
    }
    const ord = sumOf(list, 'orderedAmount');
    if (ord > 0) { row.getCell(at('% used')).value = Math.round((sumOf(list, 'shippedAmount') / ord) * 1000) / 10; row.getCell(at('% used')).numFmt = '0.0"%"'; }
    for (let n = 1; n <= N; n++) {
      const c = row.getCell(n);
      c.font = { bold: true };
      if (fill) c.fill = SUBTOTAL_FILL;
      c.border = border;
      c.alignment = { horizontal: n <= 2 ? 'left' : 'right', vertical: 'middle' };
    }
  };

  for (const l of body.rows) {
    const vals = cols.map((c) => {
      if (c.k) return typeof l[c.k] === 'number' ? l[c.k] : null;
      switch (c.h) {
        case 'Customer': return l.customer;
        case 'Order #': return str(l.orderNum);
        case 'Lot #': return str(l.lot);
        case 'Description': return str(l.description);
        case 'Status': return str(l.status);
        case 'Order date': return isoToDate(l.ordered);
        case 'Last shipped': return isoToDate(l.lastShipped);
        default: return str(l.note);
      }
    });
    const r = sheet.addRow(vals);
    cols.forEach((c, i) => {
      const cell = r.getCell(i + 1);
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      if (c.fmt) { cell.numFmt = c.fmt; cell.alignment = { horizontal: 'right', vertical: 'middle' }; }
      if (l.overShipped) cell.fill = OVER_FILL;
    });
    for (const h of ['Order date', 'Last shipped']) { r.getCell(at(h)).numFmt = DATE_FMT; r.getCell(at(h)).alignment = { horizontal: 'center', vertical: 'middle' }; }
    r.getCell(at('Description')).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.getCell(at('Note')).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
  }
  if (body.rows.length) {
    totalsRow('TOTAL', body.rows.length, body.rows, false, { top: { style: 'medium' } });
    for (const st of ['Accepted', 'Pending']) {
      const l = body.rows.filter((r) => String(r.status || '').toLowerCase() === st.toLowerCase());
      if (l.length) totalsRow(`  of which ${st}`, l.length, l, false, {});
    }
    const note = sheet.addRow(['Shipped $ is the invoice Amount, so freight and tax on an invoice are counted as shipped. Orders on a shared lot are marked "check by hand".']);
    sheet.mergeCells(note.number, 1, note.number, N);
    note.getCell(1).font = { italic: true, size: 10 };
    note.getCell(1).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  }

  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  sheet.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3, column: N } };

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Master_Order_Report.xlsx') };
}

module.exports = makeRouter({ validate, build });
