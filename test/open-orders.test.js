const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const app = require('../index');

let server;
let base;
before(async () => { await new Promise((r) => { server = app.listen(0, '127.0.0.1', r); }); base = `http://127.0.0.1:${server.address().port}`; });
after(() => new Promise((r) => server.close(r)));
const post = (path, body, raw) => fetch(`${base}${path}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: raw !== undefined ? raw : JSON.stringify(body) });
const load = async (res) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(Buffer.from(await res.arrayBuffer())); return wb.worksheets[0]; };

const sample = () => ({
  title: 'Open Orders - All Customers', asOfLabel: 'November 23, 2026', outputFileName: 'oo.xlsx',
  columns: [{ header: 'PO #', type: 'text' }, { header: 'Ship date', type: 'date' }, { header: 'Amount', type: 'money' }],
  rows: [
    { customer: 'Zeta Co', status: 'Pending', cells: ['P2', '2026-12-01', 50] },
    { customer: 'Acme', status: 'Accepted', cells: ['P1', '2026-11-30', 100] },
    { customer: 'Acme', status: 'Accepted', cells: ['P3', null, 25.5] },
    { customer: 'Unknown', status: 'No estimate found', warn: true, cells: ['P4', null, 5] },
  ],
});

test('open-orders: grouped by customer A-Z, status column, group and grand totals', async () => {
  const res = await post('/format/open-orders', sample());
  assert.equal(res.status, 200);
  const ws = await load(res);
  assert.equal(ws.name, 'Open Orders');
  assert.deepEqual(ws.getRow(3).values.slice(1), ['Customer', 'Estimate status', 'PO #', 'Ship date', 'Amount']);
  const col1 = [];
  ws.eachRow((r) => col1.push(String(r.getCell(1).value)));
  assert.ok(col1.indexOf('Acme (2 lines)') < col1.indexOf('Unknown (1 line)') && col1.indexOf('Unknown (1 line)') < col1.indexOf('Zeta Co (1 line)'));
  assert.ok(col1.includes('Total Acme'));
  assert.equal(ws.getRow(col1.indexOf('GRAND TOTAL') + 1).getCell(5).value, 180.5);
  const acme = ws.getRow(col1.indexOf('Acme (2 lines)') + 2);
  assert.equal(acme.getCell(2).value, 'Accepted');
  assert.ok(acme.getCell(4).value instanceof Date);
  assert.equal(ws.getRow(col1.indexOf('Unknown (1 line)') + 2).getCell(1).fill.fgColor.argb, 'FFFFE699');
  assert.equal(ws.views[0].ySplit, 3);
});

test('open-orders: empty rows still produce a file; bad bodies give 400', async () => {
  const b = sample();
  b.rows = [];
  assert.equal((await post('/format/open-orders', b)).status, 200);
  const bad = sample();
  bad.rows[0].cells = ['only one'];
  assert.equal((await post('/format/open-orders', bad)).status, 400);
  assert.equal((await post('/format/open-orders', null, '{nope')).status, 400);
});

test('customer-list: title, sorted unique names, blank mark column with Dena/Erin dropdown', async () => {
  const res = await post('/format/customer-list', { title: 'Customers for Dena to sort (Dena / Erin)', outputFileName: 'l.xlsx', markHeader: 'Dena or Erin?', choices: ['Dena', 'Erin'], names: ['Zed', 'Alpha', 'alpha ', 'Beta'] });
  assert.equal(res.status, 200);
  const ws = await load(res);
  assert.equal(ws.getCell('A1').value, 'Customers for Dena to sort (Dena / Erin)');
  assert.deepEqual([ws.getCell('B4').value, ws.getCell('B5').value, ws.getCell('B6').value], ['Alpha', 'Beta', 'Zed']);
  assert.equal(ws.getCell('C4').value, null);
  assert.equal(ws.getCell('C4').dataValidation.formulae[0], '"Dena,Erin"');
  assert.equal(ws.rowCount, 6);
  assert.equal((await post('/format/customer-list', { names: [''] })).status, 400);
});
