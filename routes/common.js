const express = require('express');

const GREY_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFBFBFBF' },
};
const MONEY_FMT = '$#,##0.00;($#,##0.00)';
const DATE_FMT = 'mm/dd/yyyy';
const XLSX_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// ISO yyyy-mm-dd -> real Excel date (UTC midnight so no timezone shift); null if missing/invalid.
function isoToDate(iso) {
  if (typeof iso !== 'string' || !/^\d{4}-\d{2}-\d{2}/.test(iso)) return null;
  const d = new Date(iso.slice(0, 10) + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

function styleHeaderRow(row) {
  row.font = { bold: true };
  row.eachCell((cell) => {
    cell.fill = GREY_FILL;
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
}

function safeFileName(name, fallback) {
  const n = typeof name === 'string' && name.trim() ? name.trim() : fallback;
  return n.replace(/["\r\n\\/]/g, '_');
}

// RFC 6266 / 5987 header: ASCII fallback plus UTF-8 encoded filename*.
function contentDisposition(name) {
  let n = String(name);
  if (typeof n.toWellFormed === 'function') n = n.toWellFormed(); // lone surrogates would make encodeURIComponent throw
  const ascii = n
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\x20-\x7e]|["\\\/]/g, '_');
  const encoded = encodeURIComponent(n).replace(
    /['()*]/g, (ch) => '%' + ch.charCodeAt(0).toString(16).toUpperCase());
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encoded}`;
}

function sendXlsx(res, buffer, fileName) {
  res.setHeader('Content-Type', XLSX_TYPE);
  res.setHeader('Content-Disposition', contentDisposition(fileName));
  res.setHeader('Content-Length', buffer.length);
  res.end(buffer);
}

// Returns an error string or null.
function validateLines(lines, where) {
  if (!Array.isArray(lines)) return `${where}.lines must be an array`;
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (!l || typeof l !== 'object') return `${where}.lines[${i}] must be an object`;
    if (typeof l.openBalance !== 'number' || !isFinite(l.openBalance))
      return `${where}.lines[${i}].openBalance must be a number`;
  }
  return null;
}

// Router with JSON body parsing scoped to it only (not global), and 400 on bad JSON.
function makeRouter(handler) {
  const router = express.Router();
  router.use(express.json({ limit: '5mb' }));
  router.post('/', async (req, res) => {
    try {
      const err = handler.validate(req.body);
      if (err) return res.status(400).send(err);
      const { buffer, fileName } = await handler.build(req.body);
      sendXlsx(res, buffer, fileName);
    } catch (error) {
      console.error('Formatting error:', error);
      res.status(500).send(`Formatting error: ${error.message}`);
    }
  });
  // eslint-disable-next-line no-unused-vars
  router.use((err, req, res, next) => {
    const status = err && (err.status || err.statusCode);
    if (err && err.type === 'entity.parse.failed') {
      return res.status(400).send(`Invalid JSON body: ${err.message}`);
    }
    if (status && status < 500) return res.status(status).send(err.message);
    console.error('Formatting error:', err);
    res.status(500).send(`Formatting error: ${err && err.message}`);
  });
  return router;
}

module.exports = {
  GREY_FILL, MONEY_FMT, DATE_FMT, isoToDate, styleHeaderRow, safeFileName,
  validateLines, makeRouter, contentDisposition,
};
