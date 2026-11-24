const ExcelJS = require('exceljs');
const {
  DATE_FMT, isoToDate, styleHeaderRow, safeFileName, makeRouter,
} = require('./common');

// Suggested Outcome values: the dropdown, the header hint, the off-list highlight and the summary line all read this one list.
// Other text is still allowed (see the data validation below).
const OUTCOME_VALUES = ['CEB', 'Cancelled', 'Passed', 'Failed', 'Pending'];
const OUTCOME_COL = 10;
const HEADERS = [
  'Customer', 'FA #', 'Lot #', 'Description', 'Amount', 'Ordered', 'Shipped', 'Days to ship', 'Status',
  `Outcome (e.g. ${OUTCOME_VALUES.join(', ')})`, 'Note',
];
const WIDTHS = [30, 12, 16, 60, 14, 13, 28, 9, 12, 26, 28];
const RED_FONT = { color: { argb: 'FFC00000' } };
const OFF_LIST_FILL = { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFF2CC' }, fgColor: { argb: 'FFFFF2CC' } };
const SUBTOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const AMOUNT_FMT = '$#,##0.00';
const REMAKE_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
const OPTIONAL_STRINGS = ['estimateNum', 'lot', 'description', 'ordered', 'shipped', 'status', 'outcome', 'note', 'section'];

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
    if (r.shippedFromOriginal != null && typeof r.shippedFromOriginal !== 'boolean') return `rows[${i}].shippedFromOriginal must be a boolean`;
    if (r.daysToShip != null && (typeof r.daysToShip !== 'number' || !isFinite(r.daysToShip))) return `rows[${i}].daysToShip must be a number`;
    if (r.shippedDates != null && (!Array.isArray(r.shippedDates) || r.shippedDates.some((d) => typeof d !== 'string'))) {
      return `rows[${i}].shippedDates must be an array of strings`;
    }
  }
  if (body.notCarried != null) {
    if (!Array.isArray(body.notCarried)) return 'notCarried must be an array';
    for (let i = 0; i < body.notCarried.length; i++) {
      const n = body.notCarried[i];
      if (!n || typeof n !== 'object' || Array.isArray(n)) return `notCarried[${i}] must be an object`;
      for (const k of ['faNumber', 'customer', 'outcome', 'note']) {
        if (n[k] != null && typeof n[k] !== 'string') return `notCarried[${i}].${k} must be a string`;
      }
    }
  }
  return null;
}

const median = (a) => { const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };

/** One-line summary for row 2, e.g. "126 trials | 111 shipped | 15 not shipped | 5 outcomes entered (4 CEB, 1 Cancelled) | 121 open | median 34 days to ship". */
function summaryLine(rows, notCarried) {
  const shipped = rows.filter((r) => r.shipped || (Array.isArray(r.shippedDates) && r.shippedDates.length)).length;
  const parts = [`${rows.length} ${rows.length === 1 ? 'trial' : 'trials'}`, `${shipped} shipped`, `${rows.length - shipped} not shipped`];
  // An Outcome counts as entered when it is not blank and not the generated "Remake" marker.
  const entered = rows.filter((r) => { const o = String(r.outcome || '').trim(); return o && o !== 'Remake'; });
  const detail = [];
  for (const v of OUTCOME_VALUES) {
    const n = entered.filter((r) => String(r.outcome).trim().toLowerCase() === v.toLowerCase()).length;
    if (n) detail.push(`${n} ${v}`);
  }
  const other = entered.length - detail.reduce((t, d) => t + parseInt(d, 10), 0);
  if (other) detail.push(`${other} other`);
  parts.push(`${entered.length} ${entered.length === 1 ? 'outcome' : 'outcomes'} entered${detail.length ? ` (${detail.join(', ')})` : ''}`);
  parts.push(`${rows.length - entered.length} open`);
  const days = rows.map((r) => r.daysToShip).filter((d) => typeof d === 'number');
  if (days.length) { const md = median(days); parts.push(`median ${Number.isInteger(md) ? md : md.toFixed(1)} days to ship`); }
  if (notCarried && notCarried.length) parts.push(`${notCarried.length} typed ${notCarried.length === 1 ? 'edit' : 'edits'} not carried over (see sheet)`);
  return parts.join(' | ');
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
  sheet.mergeCells(2, 1, 2, 2);
  sheet.mergeCells(2, 3, 2, HEADERS.length);
  sheet.getRow(1).height = 24;
  sheet.getRow(2).height = 20;
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 1).font = { italic: true, size: 12 };
  sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 3).value = summaryLine(body.rows, body.notCarried);
  sheet.getCell(2, 3).font = { size: 11 };
  sheet.getCell(2, 3).alignment = { horizontal: 'left', vertical: 'middle' };

  const headerRow = sheet.addRow(HEADERS);
  styleHeaderRow(headerRow);
  headerRow.height = 32;
  for (let n = 1; n <= HEADERS.length; n++) headerRow.getCell(n).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  headerRow.getCell(7).note = 'Red = the remake has no ship date of its own; this is the ship date of the original order.';
  WIDTHS.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });

  if (body.rows.length === 0) {
    const r = sheet.addRow(['No FA trials found']);
    r.font = { italic: true };
  }

  const str = (v) => (v == null ? '' : String(v));
  // Customer sections: the customer's divider row (name, trial count, total amount) comes first, then its trials.
  // A remake keeps its own customer name but sits in its original's section (row.section).
  const sectionOf = (l) => l.section || l.customer;
  const totals = new Map();
  for (const l of body.rows) {
    const t = totals.get(sectionOf(l)) || { n: 0, sum: 0 };
    t.n += 1;
    t.sum += l.amount;
    totals.set(sectionOf(l), t);
  }
  let prevSection = null;
  const dataRowNums = [];
  for (const l of body.rows) {
    if (sectionOf(l) !== prevSection) {
      prevSection = sectionOf(l);
      const t = totals.get(prevSection);
      const d = sheet.addRow([prevSection, `${t.n} ${t.n === 1 ? 'trial' : 'trials'}`, '', '', Math.round(t.sum * 100) / 100]);
      for (let n = 1; n <= HEADERS.length; n++) {
        const c = d.getCell(n);
        c.font = { bold: true };
        c.fill = SUBTOTAL_FILL;
        c.border = { top: { style: 'thin' } };
        c.alignment = { horizontal: 'left', vertical: 'middle' };
      }
      d.getCell(5).numFmt = AMOUNT_FMT;
      d.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
    }
    // Exactly one ship date -> a real Excel date; several dates (or none) stay text.
    let dates = Array.isArray(l.shippedDates) ? l.shippedDates : null;
    if (!dates) dates = /^\d{2}\/\d{2}\/\d{4}$/.test(str(l.shipped)) ? [l.shipped.slice(6) + '-' + l.shipped.slice(0, 2) + '-' + l.shipped.slice(3, 5)] : [];
    const shippedCell = dates.length === 1 && isoToDate(dates[0]) ? isoToDate(dates[0]) : str(l.shipped);
    const r = sheet.addRow([
      l.customer, l.faNumber, str(l.lot), str(l.description), l.amount,
      isoToDate(l.ordered), shippedCell, l.daysToShip == null ? null : l.daysToShip, str(l.status), str(l.outcome), str(l.note),
    ]);
    dataRowNums.push(r.number);
    for (let n = 1; n <= HEADERS.length; n++) {
      const c = r.getCell(n);
      c.alignment = { horizontal: 'left', vertical: 'middle' };
      if (l.remake) c.fill = REMAKE_FILL;
    }
    r.getCell(4).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.getCell(7).alignment = shippedCell instanceof Date
      ? { horizontal: 'center', vertical: 'middle' }
      : { horizontal: 'left', vertical: 'top', wrapText: true };
    if (shippedCell instanceof Date) r.getCell(7).numFmt = DATE_FMT;
    if (l.shippedFromOriginal && shippedCell !== '') r.getCell(7).font = RED_FONT;
    r.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(11).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.getCell(5).numFmt = AMOUNT_FMT;
    r.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
    r.getCell(6).numFmt = DATE_FMT;
    r.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // Outcome: dropdown of suggestions, but any other text is allowed (no error alert); off-list text gets a light yellow fill.
  // Generated "Remake..." outcomes are not highlighted.
  // Data rows only (not customer divider rows). Contiguous runs of data rows (between dividers) each get one validation range
  // and one conditional format, so no two ranges ever overlap.
  const colL = String.fromCharCode(64 + OUTCOME_COL);
  const blocks = [];
  for (const rn of dataRowNums) {
    const last = blocks[blocks.length - 1];
    if (last && last[1] === rn - 1) last[1] = rn; else blocks.push([rn, rn]);
  }
  const validation = {
    type: 'list', allowBlank: true, formulae: [`"${OUTCOME_VALUES.join(',')}"`],
    showInputMessage: true, promptTitle: 'Outcome', prompt: `Pick one (${OUTCOME_VALUES.join(', ')}) or type your own.`,
  };
  // showErrorMessage is deliberately never set: typing any other text must work without an alert.
  // One validation per block, added as a range. Never set cell.dataValidation on column J: exceljs merges per-cell rules into overlapping ranges.
  for (const [from, to] of blocks) sheet.dataValidations.add(`${colL}${from}:${colL}${to}`, { ...validation });
  blocks.forEach(([from, to], i) => {
    sheet.addConditionalFormatting({
      ref: `${colL}${from}:${colL}${to}`,
      rules: [{
        type: 'expression', priority: i + 1, style: { fill: OFF_LIST_FILL },
        formulae: [`AND(TRIM(${colL}${from})<>"",LEFT(${colL}${from},6)<>"Remake",NOT(OR(${OUTCOME_VALUES.map((v) => `${colL}${from}="${v}"`).join(',')})))`],
      }],
    });
  });

  sheet.views = [{ state: 'frozen', ySplit: 3 }];
  sheet.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: HEADERS.length },
  };

  // Typed Outcome / Note from the previous report whose FA # is not in this report: listed so nothing typed is lost silently.
  if (Array.isArray(body.notCarried) && body.notCarried.length) {
    const nc = workbook.addWorksheet('Not carried over');
    nc.addRow(['Typed in the previous report, but the FA # is not in this report. Copy anything you still need.']).font = { italic: true };
    styleHeaderRow(nc.addRow(['FA #', 'Customer', 'Outcome', 'Note']));
    [14, 30, 22, 50].forEach((w, i) => { nc.getColumn(i + 1).width = w; });
    for (const n of body.notCarried) nc.addRow([str(n.faNumber), str(n.customer), str(n.outcome), str(n.note)]);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Closed_FA_Report.xlsx') };
}

module.exports = makeRouter({ validate, build });
module.exports.OUTCOME_VALUES = OUTCOME_VALUES;
module.exports.summaryLine = summaryLine;
