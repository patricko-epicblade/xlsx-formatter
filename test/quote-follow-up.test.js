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

const quote = (o = {}) => ({
  rfq: 'R-1', company: 'Acme', contact: 'Cy', date: '2026-06-01', salesPerson: 'Sam', accountManager: 'Ann', status: 'Pending',
  daysPending: 177, parts: 2, annualOpportunity: 1000, annualProfit: 400, ...o,
});
const line = (o = {}) => ({ rfq: 'R-1', part: 'P1', company: 'Acme', status: 'Pending', quantity: 10, price: 5, sellPrice: 10, vendorCost: 4, vendorChoice: 'V1', profitEach: 5, counted: true, ...o });
const summary = () => ({
  Won: { count: 0, annualOpportunity: 0, annualProfit: 0 }, Lost: { count: 0, annualOpportunity: 0, annualProfit: 0 },
  Pending: { count: 1, annualOpportunity: 1000, annualProfit: 400 },
});
const post = (body, raw) => fetch(`${base}/format/quote-follow-up`, {
  method: 'POST', headers: { 'Content-Type': 'application/json' }, body: raw !== undefined ? raw : JSON.stringify(body),
});
async function load(res) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await res.arrayBuffer()));
  return wb;
}

test('quote-follow-up: four tabs, summary above the table, house style, filter and frozen header, uncounted lines marked, win rate', async () => {
  const res = await post({
    companyName: 'Co', asOfLabel: 'November 25, 2026', outputFileName: 'x.xlsx', quotes: [quote()], summary: summary(), checks: ['one check'],
    lines: [line(), line({ counted: false })],
    winRate: [{ contact: 'Pat', company: 'Acme', quotes: 3, won: 1, lost: 1, pending: 1, winPct: 0.5, pendingOpportunity: 1000 }],
  });
  assert.equal(res.status, 200);
  const wb = await load(res);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ['Quotes', 'Win Rate', 'Part Detail', 'Notes']);
  const q = wb.getWorksheet('Quotes');
  assert.equal(q.getCell(9, 1).fill.fgColor.argb, 'FFBFBFBF');
  assert.equal(q.views[0].ySplit, 9);
  assert.ok(q.autoFilter);
  assert.equal(q.getCell(10, 7).value, 'Pending');
  assert.equal(q.getCell(10, 11).value, 1000);
  assert.equal(q.getCell(9, 9).value, 'Age');
  // summary block above the table
  assert.equal(q.getCell(6, 1).value, 'Pending');
  assert.equal(q.getCell(6, 3).value, 1000);
  assert.equal(q.getCell(6, 2).value, 1);
  assert.equal(q.getCell(7, 1).value, 'All quotes');
  const w = wb.getWorksheet('Win Rate');
  assert.equal(w.getCell(4, 1).value, 'Pat');
  assert.equal(w.getCell(4, 7).value, 0.5);
  const d = wb.getWorksheet('Part Detail');
  assert.equal(d.getCell(4, 11).value, 'Yes');
  assert.equal(d.getCell(5, 11).value, 'No');
});

test('quote-follow-up: left-out lines and date-ahead quotes are noted on the Quotes tab; date-ahead shows blank days', async () => {
  const res = await post({
    asOfLabel: 'x', quotes: [quote({ daysPending: null, ageBand: 'Date ahead' })], lines: [line()], summary: summary(), leftOut: 65, futureDated: 1,
  });
  assert.equal(res.status, 200);
  const q = (await load(res)).getWorksheet('Quotes');
  assert.match(String(q.getCell(2, 1).value), /65 spreadsheet lines have no RFQ #/);
  assert.match(String(q.getCell(2, 1).value), /dated in the future/);
  assert.equal(q.getCell(10, 8).value, null);
  assert.equal(q.getCell(10, 9).value, 'Date ahead');
});

test('quote-follow-up: empty quotes still produce a workbook', async () => {
  const res = await post({ asOfLabel: 'x', quotes: [], lines: [], summary: summary() });
  assert.equal(res.status, 200);
  assert.equal((await load(res)).getWorksheet('Quotes').getCell(10, 1).value, 'No quotes found');
});

test('quote-follow-up: bad input gets a 400 with a message', async () => {
  assert.equal((await post({}, '{bad')).status, 400);
  assert.equal((await post({ quotes: 'x' })).status, 400);
  const r = await post({ quotes: [quote({ status: 'Maybe' })], lines: [], summary: summary() });
  assert.equal(r.status, 400);
  assert.match(await r.text(), /status must be Won, Lost or Pending/);
  const r2 = await post({ quotes: [quote({ annualOpportunity: 'a' })], lines: [], summary: summary() });
  assert.equal(r2.status, 400);
});
