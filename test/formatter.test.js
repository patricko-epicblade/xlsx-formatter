const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const app = require('../index');

const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture.json'), 'utf8'));
// Human-readable example outputs: written only when FORMATTER_EXAMPLE_DIR is set to an existing folder
const EXAMPLE_DIR = process.env.FORMATTER_EXAMPLE_DIR || '';

let server;
let base;

before(async () => {
  await new Promise((resolve) => {
    server = app.listen(0, '127.0.0.1', resolve);
  });
  base = `http://127.0.0.1:${server.address().port}`;
});

after(() => new Promise((resolve) => server.close(resolve)));

async function postJson(url, body, raw) {
  return fetch(base + url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: raw !== undefined ? raw : JSON.stringify(body),
  });
}

async function loadSheet(res) {
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return { wb, buf, sheet: wb.worksheets[0] };
}

function saveExample(name, buf) {
  if (EXAMPLE_DIR && fs.existsSync(EXAMPLE_DIR)) fs.writeFileSync(path.join(EXAMPLE_DIR, name), buf);
}

const fillOf = (cell) => cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb;
const MONEY = '$#,##0.00;($#,##0.00)';

test('POST /format/ar-overdue builds the internal report', async () => {
  const res = await postJson('/format/ar-overdue', {
    companyName: fixture.companyName,
    asOfLabel: fixture.asOfLabel,
    outputFileName: fixture.internalFileName,
    customers: fixture.customers,
  });
  assert.equal(res.status, 200);
  assert.match(res.headers.get('content-disposition'), /filename="20261012_Overdue_Invoices\.xlsx"/);
  const { sheet, buf } = await loadSheet(res);
  saveExample('ar-overdue-example.xlsx', buf);

  assert.equal(sheet.name, 'Overdue Invoices');
  assert.equal(sheet.getCell('A1').value, 'Overdue Invoices - Epic Blade & Industrial Solutions, LLC');
  assert.equal(sheet.getCell('A1').font.bold, true);
  assert.equal(sheet.getCell('A1').font.size, 16);
  assert.equal(sheet.getCell('A2').value, 'As of October 12, 2026');
  assert.equal(sheet.getCell('A2').font.italic, true);

  const headers = ['Customer', 'Invoice #', 'Invoice Date', 'Due Date', 'Days Past Due',
    'Open Balance', 'PO #', 'Terms', 'Lot #'];
  headers.forEach((h, i) => {
    const c = sheet.getCell(3, i + 1);
    assert.equal(c.value, h);
    assert.equal(fillOf(c), 'FFBFBFBF');
    assert.equal(c.font.bold, true);
  });

  // 3 header rows + (3+1) + (1+1) + (2+1) + grand = 13 rows
  assert.equal(sheet.rowCount, 13);
  assert.equal(sheet.getCell('A4').value, 'Acme Corp');
  assert.equal(sheet.getCell('A6').value, 'Acme Corp'); // customer name on credit line too
  assert.equal(sheet.getCell('B4').value, '90001');

  // Subtotals
  assert.equal(sheet.getCell('A7').value, 'Total for Acme Corp');
  assert.equal(sheet.getCell('F7').value, 3261.05);
  assert.equal(sheet.getCell('A7').font.bold, true);
  assert.equal(sheet.getCell('A9').value, 'Total for Beta Industrial LTDA');
  assert.equal(sheet.getCell('F9').value, 800);
  assert.equal(sheet.getCell('A12').value, 'Total for Zeta Co');
  assert.equal(sheet.getCell('F12').value, -150);
  assert.equal(sheet.getCell('G12').value, 'Credits cover overdue balance - no statement');
  assert.equal(sheet.getCell('G12').font.italic, true);

  // Grand total + invoice count
  assert.equal(sheet.getCell('A13').value, 'GRAND TOTAL');
  assert.equal(sheet.getCell('A13').font.bold, true);
  assert.equal(sheet.getCell('F13').value, 3911.05);
  assert.equal(sheet.getCell('E13').value, '4 invoices');

  // Days past due colors: 5 yellow, 45 orange, 90 red, credit none
  assert.equal(fillOf(sheet.getCell('E4')), 'FFFFF2CC');
  assert.equal(fillOf(sheet.getCell('E5')), 'FFF8CBAD');
  assert.equal(fillOf(sheet.getCell('E8')), 'FFFFC7CE');
  assert.equal(fillOf(sheet.getCell('E10')), 'FFFFF2CC');
  assert.ok(!fillOf(sheet.getCell('E6')), 'credit line has no fill');
  assert.equal(sheet.getCell('E6').value, null);
  assert.equal(sheet.getCell('E4').value, 5);

  // Formats and real dates
  assert.equal(sheet.getCell('F4').numFmt, MONEY);
  assert.equal(sheet.getCell('F6').value, -239.45);
  assert.equal(sheet.getCell('F7').numFmt, MONEY);
  assert.equal(sheet.getCell('F13').numFmt, MONEY);
  assert.ok(sheet.getCell('C4').value instanceof Date);
  assert.equal(sheet.getCell('C4').value.toISOString().slice(0, 10), '2026-08-05');
  assert.equal(sheet.getCell('D4').value.toISOString().slice(0, 10), '2026-10-07');
  assert.equal(sheet.getCell('C4').numFmt, 'mm/dd/yyyy');
  assert.equal(sheet.getCell('D4').numFmt, 'mm/dd/yyyy');

  // Freeze + autofilter
  assert.equal(sheet.views[0].state, 'frozen');
  assert.equal(sheet.views[0].ySplit, 3);
  assert.ok(sheet.autoFilter);

  // Print setup
  assert.equal(sheet.pageSetup.orientation, 'landscape');
  assert.equal(sheet.pageSetup.fitToPage, true);
  assert.equal(sheet.pageSetup.fitToWidth, 1);
  assert.equal(sheet.pageSetup.fitToHeight, 0);
  assert.equal(sheet.pageSetup.printTitlesRow, '1:3');
  const m = sheet.pageSetup.margins;
  assert.equal(m.left, 0.5);
  assert.equal(m.right, 0.5);
  assert.equal(m.top, 0.5);
  assert.equal(m.bottom, 0.5);
});

test('POST /format/ar-overdue with no customers gives a single placeholder line', async () => {
  const res = await postJson('/format/ar-overdue', {
    companyName: 'X', asOfLabel: 'October 12, 2026', outputFileName: 'e.xlsx', customers: [],
  });
  assert.equal(res.status, 200);
  const { sheet } = await loadSheet(res);
  assert.equal(sheet.getCell('A4').value, 'No overdue invoices');
  assert.equal(sheet.rowCount, 4);
});

test('POST /format/statement with empty address/remit shows placeholders', async () => {
  const cust = fixture.customers[0];
  const res = await postJson('/format/statement', {
    asOfLabel: fixture.asOfLabel,
    outputFileName: cust.statementFileName,
    header: { companyName: fixture.companyName, addressLines: [], remitToLines: [] },
    customer: cust,
  });
  assert.equal(res.status, 200);
  assert.match(res.headers.get('content-disposition'), /filename="20261012_Acme_Corp_Statement\.xlsx"/);
  const { sheet, buf } = await loadSheet(res);
  saveExample('statement-example.xlsx', buf);

  assert.equal(sheet.name, 'Statement');
  assert.equal(sheet.pageSetup.orientation, 'portrait');
  assert.equal(sheet.pageSetup.fitToWidth, 1);
  assert.equal(sheet.pageSetup.fitToPage, true);

  // assets/logo.png exists, so the logo image is placed top-left (keeping its ~2.3:1 ratio);
  // without it, the company name is written in bold 18pt instead.
  const images = sheet.getImages();
  if (fs.existsSync(path.join(__dirname, '..', 'assets', 'logo.png'))) {
    assert.equal(images.length, 1);
    assert.equal(images[0].range.tl.nativeCol, 0);
    assert.equal(images[0].range.tl.nativeRow, 0);
  } else {
    assert.equal(images.length, 0);
    assert.equal(sheet.getCell('A1').value, fixture.companyName);
    assert.equal(sheet.getCell('A1').font.size, 18);
    assert.equal(sheet.getCell('A1').font.bold, true);
  }
  assert.equal(sheet.getCell('E1').value, 'STATEMENT');
  assert.equal(sheet.getCell('E1').font.size, 20);
  assert.equal(sheet.getCell('E2').value, 'Date: October 12, 2026');

  assert.equal(sheet.getCell('A2').value, '[address to come]');
  assert.equal(sheet.getCell('A2').font.italic, true);
  assert.equal(sheet.getCell('A4').value, 'Bill to:');
  assert.equal(sheet.getCell('A5').value, 'Acme Corp');
  assert.equal(sheet.getCell('A5').font.bold, true);

  ['Invoice #', 'Invoice Date', 'Due Date', 'Days Past Due', 'Open Balance', 'PO #', 'Terms', 'Lot #']
    .forEach((h, i) => {
      const c = sheet.getCell(7, i + 1);
      assert.equal(c.value, h);
      assert.equal(fillOf(c), 'FFBFBFBF');
    });

  // 2 invoice lines, a "Credits" subheading, then the labelled credit; no age colors
  assert.equal(sheet.getCell('A8').value, '90001');
  assert.equal(sheet.getCell('E8').value, 1500.5);
  assert.equal(sheet.getCell('A10').value, 'Credits');
  assert.equal(sheet.getCell('A11').value, 'Credit Memo CM-101');
  assert.equal(sheet.getCell('E11').value, -239.45);
  assert.equal(sheet.getCell('E11').numFmt, MONEY);
  assert.ok(!fillOf(sheet.getCell('D8')), 'no age colors on statement');
  assert.ok(sheet.getCell('B8').value instanceof Date);
  assert.equal(sheet.getCell('B8').numFmt, 'mm/dd/yyyy');

  assert.equal(sheet.getCell('A12').value, 'Total past due as of October 12, 2026');
  assert.equal(sheet.getCell('A12').font.bold, true);
  assert.equal(sheet.getCell('E12').value, 3261.05);
  assert.equal(sheet.getCell('E12').numFmt, MONEY);
  assert.equal(sheet.getCell('E12').border.top.style, 'thin');

  // No remit-to text: the section is left off, with no placeholder
  const texts = [];
  sheet.eachRow((r) => texts.push(String(r.getCell(1).value)));
  assert.ok(!texts.includes('Remit to / how to pay:'));
  assert.ok(!texts.includes('[payment instructions to come]'));
});

test('POST /format/statement renders provided address and remit lines', async () => {
  const res = await postJson('/format/statement', {
    asOfLabel: fixture.asOfLabel,
    outputFileName: 's.xlsx',
    header: {
      companyName: 'Co',
      addressLines: ['1 Main St', 'Town, ST 00000'],
      remitToLines: ['Check payable to Co', 'ACH: see invoice'],
    },
    customer: fixture.customers[1],
  });
  assert.equal(res.status, 200);
  const { sheet } = await loadSheet(res);
  assert.equal(sheet.getCell('A2').value, '1 Main St');
  assert.equal(sheet.getCell('A3').value, 'Town, ST 00000');
  const texts = [];
  sheet.eachRow((r) => texts.push(String(r.getCell(1).value)));
  assert.ok(texts.includes('Check payable to Co'));
  assert.ok(texts.includes('ACH: see invoice'));
  assert.ok(!texts.includes('[address to come]'));
  assert.ok(!texts.includes('[payment instructions to come]'));
});

test('invalid bodies return 400', async () => {
  let res = await postJson('/format/ar-overdue', { companyName: 'x' });
  assert.equal(res.status, 400);
  res = await postJson('/format/ar-overdue', null, '{not json');
  assert.equal(res.status, 400);
  res = await postJson('/format/ar-overdue', {
    customers: [{ customer: 'A', netPastDue: 1, lines: [{ openBalance: 'x' }] }],
  });
  assert.equal(res.status, 400);
  res = await postJson('/format/statement', { customer: {} });
  assert.equal(res.status, 400);
  res = await postJson('/format/statement', null, '[]');
  assert.equal(res.status, 400);
});

test('regression: POST /format still formats an uploaded xlsx', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(['Part', 'Company', 'Description', '2026-01', '2026-02', 'TOTAL']);
  ws.addRow(['P-1', 'ACME', 'Widget', 5, 0, 5]);
  const input = Buffer.from(await wb.xlsx.writeBuffer());

  const form = new FormData();
  form.append('file', new Blob([input]), 'in.xlsx');
  form.append('companyName', 'Acme');
  form.append('reportStart', '2026-01');
  form.append('reportEnd', '2026-02');
  form.append('outputFileName', 'out.xlsx');
  const res = await fetch(base + '/format', { method: 'POST', body: form });
  assert.equal(res.status, 200);
  assert.match(res.headers.get('content-disposition'), /filename="out\.xlsx"/);
  const { sheet } = await loadSheet(res);
  assert.equal(sheet.name, 'Usage Chart');
  assert.equal(sheet.getCell('A1').value, 'Customer Usage Report - Acme');
  assert.equal(sheet.getCell('A2').value, 'Period: 2026-01 through 2026-02');
  assert.equal(sheet.getCell('A3').value, 'Part');
  assert.equal(fillOf(sheet.getCell('A3')), 'FFBFBFBF');
  assert.equal(sheet.getCell('D4').value, 5);
  assert.equal(fillOf(sheet.getCell('D4')), 'FFE2F0D9');

  // Missing file still 400
  const bad = await fetch(base + '/format', { method: 'POST', body: new FormData() });
  assert.equal(bad.status, 400);
});

test('long remit-to and address lines wrap across A:H instead of clipping', async () => {
  const longRemit = 'Remit by ACH to Example Bank, routing 000000000, account 000000000, ' + 'reference the invoice number. '.repeat(5);
  const longAddr = '1234 Very Long Street Name Boulevard, Building 7, Suite 900, Some Extremely Long City Name, ST 00000-0000, Country';
  const res = await postJson('/format/statement', {
    asOfLabel: fixture.asOfLabel,
    outputFileName: 's.xlsx',
    header: { companyName: 'Co', addressLines: ['Short line', longAddr], remitToLines: [longRemit.slice(0, 200)] },
    customer: fixture.customers[1],
  });
  assert.equal(res.status, 200);
  const { sheet } = await loadSheet(res);
  assert.equal(longRemit.slice(0, 200).length, 200);

  let remitRow;
  sheet.eachRow((r, n) => { if (String(r.getCell(1).value).startsWith('Remit by ACH')) remitRow = n; });
  assert.ok(remitRow, 'remit line present');
  const remit = sheet.getCell(remitRow, 1);
  assert.equal(remit.value.length, 200);
  assert.ok(remit.alignment.wrapText);
  assert.equal(sheet.getCell(remitRow, 8).master.address, `A${remitRow}`, 'merged through H');
  assert.ok(sheet.getRow(remitRow).height >= 30, "row tall enough for 2+ lines");

  const addr = sheet.getCell('A3');
  assert.equal(addr.value, longAddr);
  assert.ok(addr.alignment.wrapText);
  assert.equal(sheet.getCell('H3').master.address, 'A3');
  assert.ok(sheet.getRow(3).height >= 30);
});

test('Content-Disposition is safe for non-Latin-1 file names', async () => {
  const names = ['Café Ltda Statement.xlsx', 'O’Neil.xlsx', '日本語の請求書.xlsx', '😀 Statement.xlsx'];
  for (const name of names) {
    const res = await postJson('/format/statement', {
      asOfLabel: fixture.asOfLabel,
      outputFileName: name,
      header: { companyName: 'Co', addressLines: [], remitToLines: [] },
      customer: fixture.customers[1],
    });
    assert.equal(res.status, 200, name);
    const cd = res.headers.get('content-disposition');
    const m = /^attachment; filename="([^"\/]*)"; filename\*=UTF-8''([A-Za-z0-9%._~-]+)$/.exec(cd);
    assert.ok(m, `header parses for ${name}: ${cd}`);
    assert.match(m[1], /^[ -~]*$/);
    assert.equal(decodeURIComponent(m[2]), name);
  }
  const res = await postJson('/format/ar-overdue', {
    companyName: 'X', asOfLabel: 'a', outputFileName: '日本.xlsx', customers: [],
  });
  assert.equal(res.status, 200);
  assert.ok(res.headers.get('content-disposition').endsWith("filename*=UTF-8''%E6%97%A5%E6%9C%AC.xlsx"));
});

// ---- /format/closed-fa ----
const closedFa = JSON.parse(fs.readFileSync(path.join(__dirname, 'closed-fa-fixture.json'), 'utf8'));

test('POST /format/closed-fa builds the report', async () => {
  const res = await postJson('/format/closed-fa', closedFa);
  assert.equal(res.status, 200);
  assert.match(res.headers.get('content-disposition'), /filename="20261022_Closed_FA_Report\.xlsx"/);
  const { sheet, buf } = await loadSheet(res);
  saveExample('closed-fa-example.xlsx', buf);

  assert.equal(sheet.name, 'Closed FA');
  assert.equal(sheet.getCell('A1').value, 'Closed FA (First Article) Report - Test Company, LLC');
  assert.equal(sheet.getCell('A1').font.bold, true);
  assert.equal(sheet.getCell('A1').font.size, 16);
  assert.equal(sheet.getCell('A2').value, 'As of October 22, 2026');
  assert.equal(sheet.getCell('A2').font.italic, true);
  const headers = ['Customer', 'FA #', 'Lot #', 'Description', 'Amount', 'Ordered', 'Shipped', 'Days to ship', 'Status', 'Outcome (e.g. CEB, Cancelled, Passed, Failed, Pending)', 'Note'];
  headers.forEach((h, i) => {
    const c = sheet.getCell(3, i + 1);
    assert.equal(c.value, h);
    assert.equal(c.font.bold, true);
    assert.equal(fillOf(c), 'FFBFBFBF');
  });
  // 4 trials + one divider row per customer (Acme, Beta)
  assert.equal(sheet.rowCount, 3 + closedFa.rows.length + 2);

  // Customer divider rows carry the subtotals: Acme row 4 (3 trials), Beta row 8 (1 trial)
  assert.equal(sheet.getCell('A4').value, 'Acme Corp');
  assert.equal(sheet.getCell('B4').value, '3 trials');
  assert.equal(sheet.getCell('E4').value, 13845.6);
  assert.equal(sheet.getCell('E4').numFmt, '$#,##0.00');
  assert.equal(sheet.getCell('A4').font.bold, true);
  assert.equal(sheet.getCell('A4').border.top.style, 'thin');
  assert.equal(sheet.getCell('K4').border.top.style, 'thin');
  assert.equal(sheet.getCell('A8').value, 'Beta Industrial LTDA');
  assert.equal(sheet.getCell('B8').value, '1 trial');
  assert.equal(sheet.getCell('E8').value, 2500.5);

  // Amount + real date (first trial is row 5)
  const amt = sheet.getCell('E5');
  assert.equal(amt.value, 12345.6);
  assert.equal(amt.numFmt, '$#,##0.00');
  const ordered = sheet.getCell('F5');
  assert.ok(ordered.value instanceof Date);
  assert.equal(ordered.value.toISOString().slice(0, 10), '2026-03-15');
  assert.equal(ordered.numFmt, 'mm/dd/yyyy');

  // Description wrap, multi-line, top aligned; shipped: several dates stay text, one date is a real date, none is blank
  const desc = sheet.getCell('D5');
  assert.equal(desc.value, 'Test Widget Trial\nSecond line of notes\nThird line');
  assert.equal(desc.alignment.wrapText, true);
  assert.equal(desc.alignment.vertical, 'top');
  assert.equal(sheet.getCell('G5').value, '04/01/2026, 04/15/2026, 05/02/2026');
  assert.equal(sheet.getCell('G5').alignment.wrapText, true);
  assert.ok(sheet.getCell('G6').value instanceof Date);
  assert.equal(sheet.getCell('G6').value.toISOString().slice(0, 10), '2026-06-01');
  assert.equal(sheet.getCell('G6').numFmt, 'mm/dd/yyyy');
  assert.ok(sheet.getCell('G9').value instanceof Date);
  assert.ok(!sheet.getCell('G7').value);
  assert.ok(sheet.getColumn(4).width >= 45);

  // Days to ship (number, blank when not shipped), Note, Outcome, Status
  assert.equal(sheet.getCell('H5').value, 17);
  assert.equal(sheet.getCell('H6').value, 12);
  assert.ok(sheet.getCell('H7').value == null);
  assert.equal(sheet.getCell('K7').value, 'Not shipped; call Bob');
  assert.equal(sheet.getCell('K9').value, 'New since last report');
  assert.equal(sheet.getCell('J5').value, 'WON');
  assert.equal(sheet.getCell('I5').value, 'Converted');

  // Remake fill across the row, and only there (notes are not highlighted)
  for (let n = 1; n <= 11; n++) assert.equal(fillOf(sheet.getCell(6, n)), 'FFF2F2F2');
  assert.notEqual(fillOf(sheet.getCell('A5')), 'FFF2F2F2');
  assert.notEqual(fillOf(sheet.getCell('A7')), 'FFF2F2F2');
  assert.equal(fillOf(sheet.getCell('K7')), fillOf(sheet.getCell('A7')));

  // Freeze / autofilter / print setup
  assert.equal(sheet.views[0].state, 'frozen');
  assert.equal(sheet.views[0].ySplit, 3);
  assert.equal(String(sheet.autoFilter), 'A3:K3');
  assert.equal(sheet.pageSetup.orientation, 'landscape');
  assert.equal(sheet.pageSetup.fitToWidth, 1);
  assert.equal(sheet.pageSetup.fitToHeight, 0);
  assert.equal(sheet.pageSetup.printTitlesRow, '1:3');
  assert.equal(sheet.pageSetup.margins.left, 0.5);
});

test('POST /format/closed-fa: a remake of another customer sits in the original section and counts in its subtotal', async () => {
  const rows = [
    { ...closedFa.rows[0], shippedDates: [], shipped: '', daysToShip: null },
    { ...closedFa.rows[1], customer: 'Zed Ltd', section: 'Acme Corp', amount: 100 },
  ];
  const { sheet } = await loadSheet(await postJson('/format/closed-fa', { ...closedFa, rows }));
  assert.equal(sheet.rowCount, 3 + 1 + 2);
  assert.equal(sheet.getCell('B4').value, '2 trials');
  assert.equal(sheet.getCell('E4').value, 12445.6);
  assert.equal(sheet.getCell('A6').value, 'Zed Ltd');
  assert.ok(!sheet.getCell('G5').value);
});

test('POST /format/closed-fa works without the round-3 fields (old rows: shipped text, no note)', async () => {
  const old = closedFa.rows.map(({ shippedDates, daysToShip, note, section, ...rest }) => rest);
  const { sheet } = await loadSheet(await postJson('/format/closed-fa', { ...closedFa, rows: old }));
  assert.equal(sheet.rowCount, 3 + 4 + 2);
  assert.ok(sheet.getCell('G6').value instanceof Date); // a single "MM/DD/YYYY" text still becomes a date
  assert.ok(sheet.getCell('H5').value == null);
});

const readPrev = (buf) => fetch(base + '/read/closed-fa', { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: buf });

test('POST /read/closed-fa returns Customer, FA #, Outcome and Note of every trial, skipping divider rows', async () => {
  const gen = await postJson('/format/closed-fa', closedFa);
  const res = await readPrev(Buffer.from(await gen.arrayBuffer()));
  assert.equal(res.status, 200);
  const { rows } = await res.json();
  assert.deepEqual(rows.map((r) => [r.customer, r.faNumber, r.outcome, r.note]), [
    ['Acme Corp', '90003FA', 'WON', ''],
    ['Acme Corp', '90003FAR', 'Remake', ''],
    ['Acme Corp', '90001FA', '', 'Not shipped; call Bob'],
    ['Beta Industrial LTDA', '90002FA', 'Cancelled', 'New since last report'],
  ]);
});

test('POST /read/closed-fa reads a hand-edited report and an older report without a Note column', async () => {
  const gen = await postJson('/format/closed-fa', closedFa);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await gen.arrayBuffer()));
  const ws = wb.worksheets[0];
  ws.getCell('J7').value = 'Won - reorder';
  ws.getCell('K7').value = 'Not shipped; Bob says Friday';
  let res = await readPrev(Buffer.from(await wb.xlsx.writeBuffer()));
  let { rows } = await res.json();
  assert.equal(rows[2].outcome, 'Won - reorder');
  assert.equal(rows[2].note, 'Not shipped; Bob says Friday');
  ws.spliceColumns(11, 1); // an old report has no Note column
  res = await readPrev(Buffer.from(await wb.xlsx.writeBuffer()));
  ({ rows } = await res.json());
  assert.equal(rows.length, 4);
  assert.equal(rows[2].note, '');
  assert.equal(rows[2].outcome, 'Won - reorder');
});

test('POST /read/closed-fa rejects an empty body and a file that is not a report with 400', async () => {
  assert.equal((await readPrev(Buffer.alloc(0))).status, 400);
  assert.equal((await readPrev(Buffer.from('not an xlsx'))).status, 400);
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Other').addRow(['a', 'b']);
  assert.equal((await readPrev(Buffer.from(await wb.xlsx.writeBuffer()))).status, 400);
});

test('POST /format/closed-fa with no rows shows a placeholder', async () => {
  const res = await postJson('/format/closed-fa', { companyName: 'X', asOfLabel: 'a', rows: [] });
  assert.equal(res.status, 200);
  const { sheet } = await loadSheet(res);
  assert.equal(sheet.getCell('A3').value, 'Customer');
  assert.equal(sheet.getCell('A4').value, 'No FA trials found');
  assert.equal(sheet.rowCount, 4);
});

test('POST /format/closed-fa rejects bad bodies with 400', async () => {
  const good = closedFa.rows[0];
  const bads = [
    '{not json',
    JSON.stringify([]),
    JSON.stringify({}),
    JSON.stringify({ rows: 'x' }),
    JSON.stringify({ rows: [null] }),
    JSON.stringify({ rows: [{ ...good, customer: '' }] }),
    JSON.stringify({ rows: [{ ...good, faNumber: undefined }] }),
    JSON.stringify({ rows: [{ ...good, amount: '5' }] }),
    JSON.stringify({ rows: [{ ...good, remake: 'no' }] }),
    JSON.stringify({ rows: [{ ...good, description: 5 }] }),
    JSON.stringify({ rows: [{ ...good, daysToShip: '5' }] }),
    JSON.stringify({ rows: [{ ...good, shippedDates: '2026-01-01' }] }),
    JSON.stringify({ rows: [{ ...good, note: 5 }] }),
  ];
  for (const raw of bads) {
    const res = await postJson('/format/closed-fa', null, raw);
    assert.equal(res.status, 400, raw);
  }
});

test('POST /format/statement with totalDue lists not-yet-due lines as Current and shows two totals', async () => {
  const res = await postJson('/format/statement', {
    asOfLabel: 'November 1, 2026',
    outputFileName: 'full.xlsx',
    header: { companyName: 'Test Co', addressLines: ['1 Main St'], remitToLines: ['Pay by check'] },
    customer: {
      customer: 'Acme Corp',
      totalDue: 250,
      pastDue: 100,
      lines: [
        { type: 'Invoice', num: '1', date: '2026-08-01', dueDate: '2026-10-01', daysPastDue: 31, openBalance: 100 },
        { type: 'Invoice', num: '2', date: '2026-10-01', dueDate: '2026-12-01', daysPastDue: null, current: true, openBalance: 300 },
        { type: 'Credit Memo', num: 'C1', date: '2026-09-01', dueDate: '2026-09-01', daysPastDue: null, openBalance: -150 },
      ],
    },
  });
  assert.equal(res.status, 200);
  const { sheet, buf } = await loadSheet(res);
  saveExample('statement-full-example.xlsx', buf);
  // header row 7; overdue line 8, "Not yet due" 9, current line 10, "Credits" 11, credit 12
  assert.equal(sheet.getCell('D8').value, 31);
  assert.equal(sheet.getCell('A9').value, 'Not yet due');
  assert.equal(sheet.getCell('A9').font.italic, true);
  assert.equal(sheet.getCell('D10').value, 'Current');
  assert.equal(sheet.getCell('A11').value, 'Credits');
  assert.equal(sheet.getCell('A12').value, 'Credit Memo C1');
  assert.equal(sheet.getCell('D12').value, null); // credits have no age
  assert.equal(sheet.getCell('A13').value, 'Total balance due as of November 1, 2026');
  assert.equal(sheet.getCell('E13').value, 250);
  assert.equal(sheet.getCell('E13').border.top.style, 'thin');
  assert.equal(sheet.getCell('A14').value, 'Past due');
  assert.equal(sheet.getCell('E14').value, 100);
  assert.equal(sheet.getCell('E14').numFmt, MONEY);
  assert.equal(sheet.getCell('A14').font.bold, true);
  // 100 past due less the 150 credit floors at zero
  assert.equal(sheet.getCell('A15').value, 'Past due after credits');
  assert.equal(sheet.getCell('E15').value, 0);
  assert.equal(sheet.getCell('A17').value, 'Remit to / how to pay:');
});

test('POST /format/statement: no "after credits" line without credits; a credit with no type or number says "Credit"', async () => {
  const base = {
    asOfLabel: 'November 1, 2026', outputFileName: 'x.xlsx', header: { companyName: 'Co' },
  };
  const lines = [
    { type: 'Invoice', num: '1', date: '2026-08-01', dueDate: '2026-10-01', daysPastDue: 31, openBalance: 100 },
  ];
  let res = await postJson('/format/statement', { ...base, customer: { customer: 'A', totalDue: 100, pastDue: 100, lines } });
  let { sheet } = await loadSheet(res);
  let texts = [];
  sheet.eachRow((r) => texts.push(String(r.getCell(1).value)));
  assert.ok(!texts.includes('Past due after credits'));
  assert.ok(!texts.includes('Not yet due') && !texts.includes('Credits'));

  res = await postJson('/format/statement', {
    ...base,
    customer: { customer: 'A', totalDue: 60, pastDue: 60, lines: lines.concat([
      { num: '', date: '2026-09-01', dueDate: '2026-09-01', daysPastDue: null, openBalance: -40 },
    ]) },
  });
  ({ sheet } = await loadSheet(res));
  texts = [];
  sheet.eachRow((r) => texts.push(String(r.getCell(1).value)));
  assert.ok(texts.includes('Credit'));
  let afterRow;
  sheet.eachRow((r, n) => { if (r.getCell(1).value === 'Past due after credits') afterRow = n; });
  assert.equal(sheet.getCell(afterRow, 5).value, 20);
});

test('POST /format/statement with totalDue needs pastDue too', async () => {
  const res = await postJson('/format/statement', {
    header: {}, customer: { customer: 'A', totalDue: 5, lines: [] },
  });
  assert.equal(res.status, 400);
});

// ---- Closed FA, round 4 ----------------------------------------------------------------------
test('closed-fa: summary line in row 2, Outcome dropdown that allows other text, off-list highlight rule', async () => {
  const res = await postJson('/format/closed-fa', closedFa);
  assert.equal(res.status, 200);
  const { sheet } = await loadSheet(res);
  const n = closedFa.rows.length;
  assert.match(sheet.getCell('C2').value, new RegExp(`^${n} trials \| \d+ shipped \| \d+ not shipped`));
  assert.equal(sheet.getCell('A2').value, 'As of October 22, 2026');
  const dv = sheet.getCell('J5').dataValidation;
  assert.equal(dv.type, 'list');
  assert.equal(dv.formulae[0], '"CEB,Cancelled,Passed,Failed,Pending"');
  assert.ok(!dv.showErrorMessage, 'typing other text must not be blocked');
  const cf = sheet.conditionalFormattings.find((c) => /^J\d+:J\d+$/.test(c.ref));
  assert.ok(cf, 'conditional formatting on the Outcome column');
  assert.match(cf.rules[0].formulae[0], /J\d+="CEB".*J\d+="Pending"/);
  assert.equal(cf.rules[0].style.fill.bgColor.argb, 'FFFFF2CC');
});

test('closed-fa: summary counts outcomes and median days; Not carried over sheet and count', async () => {
  const { summaryLine } = require('../routes/closedFa');
  const rows = [
    { shipped: '01/02/2026', outcome: 'CEB', daysToShip: 10 }, { shipped: '01/02/2026', outcome: 'ceb', daysToShip: 30 },
    { shipped: '', outcome: 'Cancelled', daysToShip: null }, { shipped: '', outcome: 'Whatever', daysToShip: null },
  ];
  assert.equal(summaryLine(rows, []), '4 trials | 2 shipped | 2 not shipped | 4 outcomes entered (2 CEB, 1 Cancelled, 1 other) | 0 open | median 20 days to ship');
  assert.match(summaryLine(rows, [{}]), /\| 1 typed edit not carried over \(see sheet\)$/);
  const res = await postJson('/format/closed-fa', { ...closedFa, notCarried: [{ faNumber: '50009FA', customer: 'Acme', outcome: 'Won', note: 'typed' }] });
  assert.equal(res.status, 200);
  const { wb, sheet } = await loadSheet(res);
  assert.match(sheet.getCell('C2').value, /1 typed edit not carried over/);
  const nc = wb.getWorksheet('Not carried over');
  assert.deepEqual([1, 2, 3, 4].map((c) => nc.getCell(3, c).value), ['50009FA', 'Acme', 'Won', 'typed']);
  const plain = await loadSheet(await postJson('/format/closed-fa', closedFa));
  assert.equal(plain.wb.getWorksheet('Not carried over'), undefined);
  assert.equal((await postJson('/format/closed-fa', { ...closedFa, notCarried: 'x' })).status, 400);
});

test('closed-fa: a remake showing the original ship date is red; the new header still reads back', async () => {
  const rows = [
    { customer: 'Acme', faNumber: '1FA', amount: 1, remake: false, ordered: '2026-03-01', shipped: '03/20/2026', shippedDates: ['2026-03-20'], daysToShip: 19 },
    { customer: 'Acme', faNumber: '2FAR', amount: 1, remake: true, ordered: '2026-04-10', shipped: '03/20/2026', shippedDates: ['2026-03-20'], daysToShip: null, shippedFromOriginal: true },
  ];
  const res = await postJson('/format/closed-fa', { companyName: 'T', asOfLabel: 'x', rows });
  const { sheet, buf } = await loadSheet(res);
  assert.notEqual(sheet.getCell('G5').font.color && sheet.getCell('G5').font.color.argb, 'FFC00000');
  assert.equal(sheet.getCell('G6').font.color.argb, 'FFC00000');
  assert.equal(sheet.getCell('H6').value, null);
  const { readPreviousReport } = require('../routes/closedFaPrevious');
  assert.deepEqual((await readPreviousReport(buf)).map((r) => r.faNumber), ['1FA', '2FAR']);
});

// ---- Closed FA, round 5 ----------------------------------------------------------------------
const cellsOf = (ref) => {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
  return { c1: m[1], r1: +m[2], r2: +(m[4] || m[2]) };
};
const overlap = (a, b) => a.c1 === b.c1 && a.r1 <= b.r2 && b.r1 <= a.r2;

test('closed-fa: Outcome validation and yellow rule cover data rows only, no overlaps, never an error alert', async () => {
  const JSZip = require('jszip');
  const res = await postJson('/format/closed-fa', closedFa);
  const buf = Buffer.from(await res.arrayBuffer());
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const { sheet } = await loadSheet(await postJson('/format/closed-fa', closedFa));
  const dividers = new Set();
  sheet.eachRow((r, n) => { if (n > 3 && r.getCell(2).value && /^\d+ trials?$/.test(String(r.getCell(2).value))) dividers.add(n); });
  assert.ok(dividers.size > 0);
  const dvs = [...xml.matchAll(/<dataValidation [^>]*>/g)].map((m) => m[0]);
  assert.ok(dvs.length > 0);
  const ranges = [];
  for (const tag of dvs) {
    assert.ok(!/showErrorMessage="(1|true)"/.test(tag), `error alert must never be on: ${tag}`);
    for (const ref of /sqref="([^"]+)"/.exec(tag)[1].split(' ')) ranges.push(cellsOf(ref));
  }
  const cfs = [...xml.matchAll(/<conditionalFormatting sqref="([^"]+)"/g)].map((m) => cellsOf(m[1]));
  assert.ok(cfs.length > 0);
  for (const list of [ranges, cfs]) {
    list.forEach((a, i) => {
      assert.equal(a.c1, 'J');
      for (let r = a.r1; r <= a.r2; r++) assert.ok(!dividers.has(r), `row ${r} is a divider row`);
      list.forEach((b, j) => { if (i < j) assert.ok(!overlap(a, b), `overlapping ranges ${JSON.stringify(a)} ${JSON.stringify(b)}`); });
    });
  }
  // every data row is covered by a validation
  sheet.eachRow((r, n) => { if (n > 3 && !dividers.has(n) && r.getCell(1).value !== 'No FA trials found') assert.ok(ranges.some((a) => n >= a.r1 && n <= a.r2), `row ${n} lacks the dropdown`); });
});

test('closed-fa: many customers of varied sizes give exactly one non-overlapping Outcome validation per block', async () => {
  const JSZip = require('jszip');
  const sizes = [11, 1, 2, 45, 3, 1, 7];
  const rows = [];
  sizes.forEach((n, ci) => { for (let i = 0; i < n; i++) rows.push({ customer: `Cust ${ci}`, faNumber: `${ci}${i}FA`, amount: 1, remake: false, ordered: '2026-03-01', shipped: '', shippedDates: [] }); });
  const res = await postJson('/format/closed-fa', { companyName: 'T', asOfLabel: 'x', rows });
  const zip = await JSZip.loadAsync(Buffer.from(await res.arrayBuffer()));
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const refs = [...xml.matchAll(/<dataValidation [^>]*sqref="([^"]+)"/g)].map((m) => m[1]);
  assert.equal(refs.length, sizes.length);
  const rs = refs.map((r) => cellsOf(r));
  rs.forEach((a, i) => rs.forEach((b, j) => { if (i < j) assert.ok(!overlap(a, b), `overlap ${refs[i]} ${refs[j]}`); }));
  assert.ok(!/showErrorMessage/.test(xml));
});

test('closed-fa: summary shows entered vs open outcomes', () => {
  const { summaryLine } = require('../routes/closedFa');
  const rows = [
    { shipped: '01/02/2026', outcome: 'CEB', daysToShip: 10 }, { shipped: '', outcome: 'Remake' }, { shipped: '', outcome: '' },
    { shipped: '', outcome: 'Cancelled' },
  ];
  assert.equal(summaryLine(rows, []), '4 trials | 1 shipped | 3 not shipped | 2 outcomes entered (1 CEB, 1 Cancelled) | 2 open | median 10 days to ship');
});
