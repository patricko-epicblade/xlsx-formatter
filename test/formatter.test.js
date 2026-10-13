const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const app = require('../index');

const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixture.json'), 'utf8'));
// Human-readable example outputs (written only if this folder exists on the machine)
const EXAMPLE_DIR = 'C:\\Users\\aweso\\Documents\\Epic Blade\\automations\\02-account-statement\\04-build\\formatter-test';

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
  if (fs.existsSync(EXAMPLE_DIR)) fs.writeFileSync(path.join(EXAMPLE_DIR, name), buf);
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

  // No logo file in the repo, so the company name is used
  assert.equal(sheet.getCell('A1').value, fixture.companyName);
  assert.equal(sheet.getCell('A1').font.size, 18);
  assert.equal(sheet.getCell('A1').font.bold, true);
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

  // 3 lines, no age colors
  assert.equal(sheet.getCell('A8').value, '90001');
  assert.equal(sheet.getCell('E8').value, 1500.5);
  assert.equal(sheet.getCell('E10').value, -239.45);
  assert.equal(sheet.getCell('E10').numFmt, MONEY);
  assert.ok(!fillOf(sheet.getCell('D8')), 'no age colors on statement');
  assert.ok(sheet.getCell('B8').value instanceof Date);
  assert.equal(sheet.getCell('B8').numFmt, 'mm/dd/yyyy');

  assert.equal(sheet.getCell('A11').value, 'Total past due');
  assert.equal(sheet.getCell('A11').font.bold, true);
  assert.equal(sheet.getCell('E11').value, 3261.05);
  assert.equal(sheet.getCell('E11').numFmt, MONEY);
  assert.equal(sheet.getCell('E11').border.top.style, 'thin');

  assert.equal(sheet.getCell('A13').value, 'Remit to / how to pay:');
  assert.equal(sheet.getCell('A13').font.bold, true);
  assert.equal(sheet.getCell('A14').value, '[payment instructions to come]');
  assert.equal(sheet.getCell('A14').font.italic, true);
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
