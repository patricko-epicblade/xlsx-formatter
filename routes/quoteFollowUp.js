const ExcelJS = require('exceljs');
const {
  DATE_FMT, isoToDate, styleHeaderRow, safeFileName, makeRouter,
} = require('./common');

const AMOUNT_FMT = '$#,##0.00;[Red]($#,##0.00)';
const QTY_FMT = '#,##0.##';
const STATUSES = ['Won', 'Lost', 'Pending'];
const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } };
const ALT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
const money = (n) => Number(n).toLocaleString('en-US', { style: 'currency', currency: 'USD' });

const QUOTE_STRINGS = ['rfq', 'company', 'contact', 'date', 'salesPerson', 'accountManager', 'status', 'ageBand'];
const QUOTE_NUMBERS = ['parts', 'annualOpportunity', 'annualProfit'];
const LINE_STRINGS = ['rfq', 'part', 'company', 'status', 'vendorChoice'];
const LINE_NUMBERS = ['quantity', 'price', 'sellPrice', 'vendorCost', 'profitEach'];

function validate(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return 'Body must be a JSON object';
  if (!Array.isArray(body.quotes)) return 'quotes must be an array';
  if (!Array.isArray(body.lines)) return 'lines must be an array';
  if (!body.summary || typeof body.summary !== 'object') return 'summary must be an object';
  for (const st of STATUSES) {
    const s = body.summary[st];
    if (!s || typeof s !== 'object') return `summary.${st} must be an object`;
    for (const k of ['count', 'annualOpportunity', 'annualProfit']) {
      if (typeof s[k] !== 'number' || !isFinite(s[k])) return `summary.${st}.${k} must be a number`;
    }
  }
  if (body.leftOut != null && typeof body.leftOut !== 'number') return 'leftOut must be a number';
  if (body.futureDated != null && typeof body.futureDated !== 'number') return 'futureDated must be a number';
  if (body.winRate != null && !Array.isArray(body.winRate)) return 'winRate must be an array';
  if (body.checks != null && (!Array.isArray(body.checks) || body.checks.some((c) => typeof c !== 'string'))) return 'checks must be an array of strings';
  for (let i = 0; i < body.quotes.length; i++) {
    const q = body.quotes[i];
    if (!q || typeof q !== 'object' || Array.isArray(q)) return `quotes[${i}] must be an object`;
    if (typeof q.rfq !== 'string' || !q.rfq.trim()) return `quotes[${i}].rfq must be a non-empty string`;
    if (!STATUSES.includes(q.status)) return `quotes[${i}].status must be Won, Lost or Pending`;
    for (const k of QUOTE_STRINGS) if (q[k] != null && typeof q[k] !== 'string') return `quotes[${i}].${k} must be a string`;
    for (const k of QUOTE_NUMBERS) if (typeof q[k] !== 'number' || !isFinite(q[k])) return `quotes[${i}].${k} must be a number`;
    if (q.daysPending != null && (typeof q.daysPending !== 'number' || !isFinite(q.daysPending))) return `quotes[${i}].daysPending must be a number`;
  }
  for (let i = 0; i < body.lines.length; i++) {
    const l = body.lines[i];
    if (!l || typeof l !== 'object' || Array.isArray(l)) return `lines[${i}] must be an object`;
    if (typeof l.rfq !== 'string' || !l.rfq.trim()) return `lines[${i}].rfq must be a non-empty string`;
    for (const k of LINE_STRINGS) if (l[k] != null && typeof l[k] !== 'string') return `lines[${i}].${k} must be a string`;
    for (const k of LINE_NUMBERS) if (l[k] != null && (typeof l[k] !== 'number' || !isFinite(l[k]))) return `lines[${i}].${k} must be a number`;
    if (l.counted != null && typeof l.counted !== 'boolean') return `lines[${i}].counted must be a boolean`;
  }
  return null;
}

function title(sheet, text, sub, width) {
  sheet.addRow([text]);
  sheet.addRow([sub]);
  sheet.mergeCells(1, 1, 1, width);
  sheet.mergeCells(2, 1, 2, width);
  sheet.getRow(1).height = 24;
  sheet.getCell(1, 1).font = { bold: true, size: 16 };
  sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  sheet.getCell(2, 1).font = { italic: true, size: 12 };
  sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };
}

function table(sheet, cols, headerAt) {
  const headerRow = sheet.getRow(headerAt);
  headerRow.values = cols.map((c) => c.h);
  styleHeaderRow(headerRow);
  headerRow.height = 32;
  for (let n = 1; n <= cols.length; n++) headerRow.getCell(n).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cols.forEach((c, i) => { sheet.getColumn(i + 1).width = c.w; });
  sheet.views = [{ state: 'frozen', ySplit: headerAt }];
  sheet.autoFilter = { from: { row: headerAt, column: 1 }, to: { row: headerAt, column: cols.length } };
}

function pageSetup(headerAt) {
  return {
    orientation: 'landscape', paperSize: 1, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }, printTitlesRow: `1:${headerAt}`,
  };
}

async function build(body) {
  const companyName = body.companyName || 'Customer';
  const asOfLabel = body.asOfLabel || '';
  const workbook = new ExcelJS.Workbook();
  const sub = `As of ${asOfLabel}`;
  const str = (v) => (v == null ? '' : String(v));
  const num = (v) => (typeof v === 'number' ? v : null);

  // ---- Quotes tab (first): status summary above the table, then one row per RFQ #
  const qCols = [
    { h: 'RFQ #', w: 16 }, { h: 'Company', w: 28 }, { h: 'Customer', w: 24 }, { h: 'Date', w: 16 }, { h: 'Sales Person', w: 20 },
    { h: 'Account Manager', w: 20 }, { h: 'Status', w: 11 }, { h: 'Days pending', w: 11, fmt: '#,##0' }, { h: 'Age', w: 11 }, { h: 'Parts', w: 8, fmt: '0' },
    { h: 'Annual opportunity $', w: 20, fmt: AMOUNT_FMT }, { h: 'Annual profit $', w: 18, fmt: AMOUNT_FMT },
  ];
  const qs = workbook.addWorksheet('Quotes', { pageSetup: pageSetup(9) });
  const leftOut = typeof body.leftOut === 'number' && body.leftOut > 0
    ? ` | ${body.leftOut} spreadsheet ${body.leftOut === 1 ? 'line has' : 'lines have'} no RFQ # and ${body.leftOut === 1 ? 'is' : 'are'} not listed (see Notes)` : '';
  const ahead = typeof body.futureDated === 'number' && body.futureDated > 0
    ? ` | ${body.futureDated} pending ${body.futureDated === 1 ? 'quote is' : 'quotes are'} dated in the future ("Date ahead"; probably a date typo)` : '';
  title(qs, `Quote Follow-Up Report - ${companyName}`, `${sub} | ${body.quotes.length} ${body.quotes.length === 1 ? 'quote' : 'quotes'}, biggest annual opportunity first${leftOut}${ahead}`, qCols.length);

  // Summary block (rows 3-7), same grey header and total styling as the table below.
  const sumHeader = qs.getRow(3);
  sumHeader.values = ['Status', 'Quotes', 'Annual opportunity $', 'Annual profit $', 'Share of opportunity'];
  for (let n = 1; n <= 5; n++) {
    sumHeader.getCell(n).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
    sumHeader.getCell(n).font = { bold: true };
    sumHeader.getCell(n).alignment = { horizontal: n === 1 ? 'left' : 'right', vertical: 'middle' };
  }
  const totalOpp = STATUSES.reduce((t, st) => t + body.summary[st].annualOpportunity, 0);
  const sumRow = (at, label, s, bold) => {
    const share = totalOpp ? s.annualOpportunity / totalOpp : null;
    const r = qs.getRow(at);
    r.values = [label, s.count, s.annualOpportunity, s.annualProfit, share];
    r.getCell(2).numFmt = '#,##0';
    r.getCell(3).numFmt = AMOUNT_FMT;
    r.getCell(4).numFmt = AMOUNT_FMT;
    r.getCell(5).numFmt = '0.0%';
    for (let n = 1; n <= 5; n++) {
      r.getCell(n).alignment = { horizontal: n === 1 ? 'left' : 'right', vertical: 'middle' };
      if (bold) { r.getCell(n).font = { bold: true }; r.getCell(n).border = { top: { style: 'medium' } }; }
    }
  };
  STATUSES.forEach((st, i) => sumRow(4 + i, st, body.summary[st]));
  const total = { count: 0, annualOpportunity: 0, annualProfit: 0 };
  for (const st of STATUSES) {
    total.count += body.summary[st].count;
    total.annualOpportunity += body.summary[st].annualOpportunity;
    total.annualProfit += body.summary[st].annualProfit;
  }
  total.annualOpportunity = Math.round(total.annualOpportunity * 100) / 100;
  total.annualProfit = Math.round(total.annualProfit * 100) / 100;
  sumRow(7, 'All quotes', total, true);
  const pend = body.summary.Pending;
  qs.getCell(8, 1).value = `Follow up on the ${pend.count} pending ${pend.count === 1 ? 'quote' : 'quotes'}: ${money(pend.annualOpportunity)} of annual opportunity, ${money(pend.annualProfit)} annual profit. Filter Status = Pending and sort by Sales Person or Company.`;
  qs.mergeCells(8, 1, 8, qCols.length);
  qs.getCell(8, 1).alignment = { horizontal: 'left', vertical: 'middle' };
  qs.getCell(8, 1).font = { bold: true };

  table(qs, qCols, 9);
  if (body.quotes.length === 0) qs.addRow(['No quotes found']).font = { italic: true };
  for (const q of body.quotes) {
    const r = qs.addRow([q.rfq, str(q.company), str(q.contact), isoToDate(q.date), str(q.salesPerson), str(q.accountManager), q.status,
      num(q.daysPending), str(q.ageBand), q.parts, q.annualOpportunity, q.annualProfit]);
    qCols.forEach((c, i) => {
      const cell = r.getCell(i + 1);
      cell.alignment = { horizontal: c.fmt ? 'right' : 'left', vertical: 'middle' };
      if (c.fmt) cell.numFmt = c.fmt;
    });
    r.getCell(4).numFmt = DATE_FMT;
    r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
  }
  if (body.quotes.length) {
    const first = 10;
    const last = 9 + body.quotes.length;
    const r = qs.addRow([]);
    r.getCell(1).value = 'TOTAL (visible rows)';
    // SUBTOTAL(109) follows the filter, so the totals match what is on screen.
    r.getCell(7).value = { formula: `SUBTOTAL(103,A${first}:A${last})`, result: body.quotes.length };
    r.getCell(10).value = { formula: `SUBTOTAL(109,J${first}:J${last})`, result: body.quotes.reduce((t, q) => t + q.parts, 0) };
    r.getCell(11).value = { formula: `SUBTOTAL(109,K${first}:K${last})`, result: total.annualOpportunity };
    r.getCell(12).value = { formula: `SUBTOTAL(109,L${first}:L${last})`, result: total.annualProfit };
    r.getCell(11).numFmt = AMOUNT_FMT;
    r.getCell(12).numFmt = AMOUNT_FMT;
    for (let n = 1; n <= qCols.length; n++) {
      r.getCell(n).font = { bold: true };
      r.getCell(n).fill = TOTAL_FILL;
      r.getCell(n).border = { top: { style: 'medium' } };
    }
  }

  // ---- Win Rate tab: by the customer's sales person
  const wCols = [
    { h: "Customer's sales person", w: 28 }, { h: 'Company', w: 30 }, { h: 'Quotes', w: 10, fmt: '#,##0' }, { h: 'Won', w: 9, fmt: '#,##0' },
    { h: 'Lost', w: 9, fmt: '#,##0' }, { h: 'Pending', w: 10, fmt: '#,##0' }, { h: 'Win %', w: 10, fmt: '0%' }, { h: 'Pending annual opportunity $', w: 26, fmt: AMOUNT_FMT },
  ];
  const wr = workbook.addWorksheet('Win Rate', { pageSetup: pageSetup(3) });
  title(wr, `Win Rate - ${companyName}`, `${sub} | by the customer's sales person; Win % = Won / (Won + Lost), blank when nothing is decided yet`, wCols.length);
  table(wr, wCols, 3);
  for (const p of body.winRate || []) {
    const r = wr.addRow([str(p.contact), str(p.company), p.quotes, p.won, p.lost, p.pending, num(p.winPct), p.pendingOpportunity]);
    wCols.forEach((c, i) => {
      const cell = r.getCell(i + 1);
      cell.alignment = { horizontal: c.fmt ? 'right' : 'left', vertical: 'middle' };
      if (c.fmt) cell.numFmt = c.fmt;
    });
  }

  // ---- Part Detail tab
  const dCols = [
    { h: 'RFQ #', w: 16 }, { h: 'Part #', w: 40 }, { h: 'Company', w: 28 }, { h: 'Status', w: 11 }, { h: 'Annual quantity', w: 13, fmt: QTY_FMT },
    { h: 'Price (landed cost)', w: 16, fmt: AMOUNT_FMT }, { h: 'Sell price', w: 14, fmt: AMOUNT_FMT }, { h: 'Vendor cost', w: 14, fmt: AMOUNT_FMT },
    { h: 'Vendor choice', w: 26 }, { h: 'Profit each', w: 14, fmt: AMOUNT_FMT }, { h: 'Counted in totals', w: 12 },
  ];
  const ds = workbook.addWorksheet('Part Detail', { pageSetup: pageSetup(3) });
  title(ds, `Part Detail - ${companyName}`, `${sub} | every line behind each quote; "No" = another price tier or vendor quote for a part already counted (the Won line, else the largest quantity; not added to the totals)`, dCols.length);
  table(ds, dCols, 3);
  for (const l of body.lines) {
    const counted = l.counted !== false;
    const r = ds.addRow([l.rfq, str(l.part), str(l.company), str(l.status), num(l.quantity), num(l.price), num(l.sellPrice), num(l.vendorCost),
      str(l.vendorChoice), num(l.profitEach), counted ? 'Yes' : 'No']);
    dCols.forEach((c, i) => {
      const cell = r.getCell(i + 1);
      cell.alignment = { horizontal: c.fmt ? 'right' : 'left', vertical: 'middle' };
      if (c.fmt) cell.numFmt = c.fmt;
      if (!counted) cell.fill = ALT_FILL;
    });
    r.getCell(11).alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ---- Notes tab
  const ns = workbook.addWorksheet('Notes', { pageSetup: pageSetup(2) });
  title(ns, `Notes - ${companyName}`, sub, 1);
  ns.getColumn(1).width = 130;
  const notes = [
    'Status comes from the Status column of the quote spreadsheet and applies to the whole quote (every line with the same RFQ #). If any line is Won, the quote is Won. Otherwise, if any line is Lost, the quote is Lost. If every line is blank, it is Pending, counted from the day it was sent.',
    'Each part is counted once per quote: the Won line if there is one, otherwise the line with the largest annual quantity. The other price tiers and vendor quotes are listed in Part Detail as not counted.',
    "Annual opportunity is the sell price x annual quantity for each counted part (or the spreadsheet's own Annual Opportunity Size when filled in).",
    'RFQ numbers that differ by a suffix (-1, -2, -3) are separate quotes.',
  ];
  for (const t of [...notes, ...(body.checks || []).map((c) => `Check: ${c}`)]) {
    const r = ns.addRow([t]);
    r.getCell(1).alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    r.height = Math.max(18, 15 * Math.ceil(t.length / 120));
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fileName: safeFileName(body.outputFileName, 'Quote_Follow_Up_Report.xlsx') };
}

module.exports = makeRouter({ validate, build });
