const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const app = require('../index');

let server;
let base;
before(async () => {
  await new Promise((resolve) => { server = app.listen(0, '127.0.0.1', resolve); });
  base = `http://127.0.0.1:${server.address().port}`;
});
after(() => new Promise((resolve) => server.close(resolve)));

const row = (o = {}) => ({
  customer: 'Acme', orderNum: '11111M', lot: '70001M', description: 'Knives', status: 'Accepted', ordered: '2026-01-05',
  orderedAmount: 1000, shippedAmount: 250, remainingAmount: 750, percentUsed: 25, shippedCount: 1,
  lastShipped: '2026-03-01', daysSinceShipment: 266, note: '', ...o,
});
const post = (body, raw) => fetch(`${base}/format/master-order`, {
  method: 'POST', headers: { 'Content-Type': 'application/json' }, body: raw !== undefined ? raw : JSON.stringify(body),
});
async function load(res) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await res.arrayBuffer()));
  return wb.worksheets[0];
}

test('master-order: house style, headers, dollars only when hasQty is false, no subtotal rows, TOTAL and status rows', async () => {
  const res = await post({ companyName: 'Co', asOfLabel: 'November 22, 2026', outputFileName: 'x.xlsx', hasQty: false, qtyNote: 'No Qty', rows: [row(), row({ orderNum: '22222M' })] });
  assert.equal(res.status, 200);
  assert.match(res.headers.get('content-disposition'), /x\.xlsx/);
  const ws = await load(res);
  assert.equal(ws.name, 'Master Orders');
  assert.equal(ws.getCell('A1').value, 'Open Master Orders Report - Co');
  assert.match(String(ws.getCell('C2').value), /^2 open master orders \| \$1,500\.00 remaining \| Accepted 2: \$1,500\.00 remaining \| No Qty/);
  const headers = ws.getRow(3).values.slice(1);
  assert.deepEqual(headers, ['Customer', 'Order #', 'Lot #', 'Description', 'Status', 'Order date', 'Ordered $', 'Shipped $', 'Remaining $', '% used', 'Shipments', 'Last shipped', 'Days since shipment', 'Note']);
  assert.equal(ws.getCell(3, 1).fill.fgColor.argb, 'FFBFBFBF');
  assert.equal(ws.views[0].ySplit, 3);
  assert.ok(ws.autoFilter);
  assert.equal(ws.getCell(4, 4).value, 'Knives'); // first row is an order, no customer subtotal
  assert.ok(ws.getCell(4, 6).value instanceof Date);
  assert.notEqual(ws.getCell(5, 1).fill?.fgColor?.argb, 'FFFFF2CC', 'no idle (yellow) row');
  assert.equal(ws.getCell(6, 1).value, 'TOTAL');
  assert.equal(ws.getCell(6, 9).value, 1500);
  assert.match(String(ws.getCell(8, 1).value), /freight and tax/);
});

test('master-order: quantity columns appear with hasQty and totals add up', async () => {
  const res = await post({ hasQty: true, rows: [row({ orderedQty: 100, shippedQty: 40, remainingQty: 60 }), row({ customer: 'Zed', orderedQty: 10, shippedQty: 12, remainingQty: -2 })] });
  const ws = await load(res);
  const headers = ws.getRow(3).values.slice(1);
  assert.deepEqual(headers.slice(10, 13), ['Ordered qty', 'Shipped qty', 'Remaining qty']);
  const total = ws.getRow(6); // 2 orders from row 4, then TOTAL
  assert.equal(total.getCell(13).value, 58);
});

test('master-order: empty rows give a single "No open master orders found" line', async () => {
  const ws = await load(await post({ rows: [] }));
  assert.equal(ws.getCell(4, 1).value, 'No open master orders found');
  assert.equal(ws.rowCount, 4);
});

test('master-order: 400 on a bad body, 400 on invalid JSON', async () => {
  assert.equal((await post({ rows: 'x' })).status, 400);
  assert.equal((await post({ rows: [{ customer: 'A' }] })).status, 400);
  assert.equal((await post(null, '{nope')).status, 400);
});

test('master-order: over-shipped rows are red, no customer subtotal rows', async () => {
  const res = await post({ rows: [row({ customer: 'Solo', overShipped: true, remainingAmount: -5 }), row({ customer: 'Two' }), row({ customer: 'Two', orderNum: '33333M' })] });
  const ws = await load(res);
  assert.match(String(ws.getCell('C2').value), /1 over-shipped \(red\)/);
  assert.equal(ws.getCell(4, 4).value, 'Knives');
  assert.equal(ws.getCell(4, 1).fill.fgColor.argb, 'FFF8CBAD');
  assert.equal(ws.getCell(5, 1).value, 'Two'); // an order, not a subtotal
  assert.equal(ws.getCell(7, 1).value, 'TOTAL');
});

test('master-order: Accepted and Pending are totalled separately', async () => {
  const res = await post({ rows: [row(), row({ status: 'Pending', remainingAmount: 100, orderedAmount: 200, shippedAmount: 100 })] });
  const ws = await load(res);
  assert.match(String(ws.getCell('C2').value), /Accepted 1: \$750\.00 remaining \| Pending 1: \$100\.00 remaining/);
  assert.equal(ws.getCell(6, 1).value, 'TOTAL');
  assert.match(String(ws.getCell(7, 1).value), /of which Accepted/);
  assert.equal(ws.getCell(7, 9).value, 750);
  assert.match(String(ws.getCell(8, 1).value), /of which Pending/);
  assert.equal(ws.getCell(8, 9).value, 100);
});
