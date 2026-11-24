const express = require('express');
const ExcelJS = require('exceljs');

// Reads a Closed FA report (.xlsx) written by /format/closed-fa and returns the cells a person may have edited, so the
// next run can carry them over (fa-rules.js carryOver). Columns are found by header name, so an older report without
// the Note column still reads. Customer divider rows ("<n> trials" in the FA # column) are skipped.
const SUBTOTAL_RE = /^\d[\d,]* trials?$/i;

const cellText = (cell) => {
  const v = cell && cell.value;
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') return String(cell.text || '').trim();
  return String(v).trim();
};

/** buffer (xlsx) -> [{ faNumber, customer, outcome, note }] in sheet order. Throws if it is not a Closed FA report. */
async function readPreviousReport(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.getWorksheet('Closed FA') || wb.worksheets[0];
  if (!sheet) throw new Error('The previous report has no worksheet.');

  let headerRow = 0;
  const col = {};
  for (let r = 1; r <= Math.min(sheet.rowCount, 10) && !headerRow; r++) {
    const found = {};
    // "Outcome (e.g. CEB, ...)" is the header since round 4; older reports say just "Outcome".
    sheet.getRow(r).eachCell((c, n) => { found[cellText(c).toLowerCase().replace(/^(outcome)\s*\(.*\)$/, '$1')] = n; });
    if (found['fa #'] && found.customer && found.outcome) { headerRow = r; Object.assign(col, found); }
  }
  if (!headerRow) throw new Error('The previous report does not look like a Closed FA report (no Customer / FA # / Outcome header row).');

  const rows = [];
  for (let r = headerRow + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const faNumber = cellText(row.getCell(col['fa #']));
    if (SUBTOTAL_RE.test(faNumber)) continue;
    const customer = cellText(row.getCell(col.customer));
    if (!customer || customer === 'No FA trials found') continue;
    rows.push({
      faNumber, customer,
      outcome: cellText(row.getCell(col.outcome)),
      note: col.note ? cellText(row.getCell(col.note)) : '',
    });
  }
  return rows;
}

// POST /read/closed-fa  (body = the .xlsx bytes)  ->  { rows: [...] }
const router = express.Router();
router.post('/', express.raw({ type: () => true, limit: '10mb' }), async (req, res) => {
  try {
    if (!Buffer.isBuffer(req.body) || !req.body.length) return res.status(400).send('Body must be the .xlsx file');
    let rows;
    try { rows = await readPreviousReport(req.body); } catch (e) { return res.status(400).send(`Cannot read the previous report: ${e.message}`); }
    res.json({ rows });
  } catch (error) {
    console.error('Read error:', error);
    res.status(500).send(`Read error: ${error.message}`);
  }
});

module.exports = router;
module.exports.readPreviousReport = readPreviousReport;
