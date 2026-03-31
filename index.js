const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const upload = multer();

// Health check route
app.get('/', (req, res) => {
  res.status(200).send('XLSX formatter is running');
});

// Optional GET route for testing
app.get('/format', (req, res) => {
  res.status(200).send('Formatter endpoint is live. Use POST to send a file.');
});

app.post('/format', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded');
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const sheet = workbook.worksheets[0];

    const companyName = req.body.companyName || 'Customer';
    const reportStart = req.body.reportStart || '';
    const reportEnd = req.body.reportEnd || '';

    sheet.spliceRows(
      1,
      0,
      [`Customer Usage Report - ${companyName}`],
      [`Period: ${reportStart} through ${reportEnd}`]
    );

    const lastCol = Math.max(sheet.columnCount, 1);

    sheet.mergeCells(1, 1, 1, lastCol);
    sheet.mergeCells(2, 1, 2, lastCol);

    sheet.getRow(1).font = { bold: true, size: 16 };
    sheet.getRow(1).alignment = { horizontal: 'center' };

    sheet.getRow(2).font = { italic: true };
    sheet.getRow(2).alignment = { horizontal: 'center' };

    const headerRowIndex = 3;
    const headerRow = sheet.getRow(headerRowIndex);

    headerRow.font = { bold: true };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFBFBFBF' },
      };
      cell.border = {
        bottom: { style: 'thin' },
      };
      cell.alignment = { horizontal: 'center' };
    });

    sheet.views = [{ state: 'frozen', ySplit: headerRowIndex }];

    sheet.columns.forEach((col) => {
      let maxLength = 10;

      col.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
        if (cell.value === 0) {
          cell.value = '';
        }

        if (rowNumber > headerRowIndex && typeof cell.value === 'number') {
          cell.alignment = { horizontal: 'right' };
        }

        const value = cell.value == null ? '' : String(cell.value);
        maxLength = Math.max(maxLength, value.length);
      });

      col.width = maxLength + 2;
    });

    const firstCol = sheet.getColumn(1);
    firstCol.eachCell((cell, rowNumber) => {
      if (rowNumber > headerRowIndex) {
        cell.font = { bold: true };
      }
    });

    sheet.autoFilter = {
      from: { row: headerRowIndex, column: 1 },
      to: { row: headerRowIndex, column: sheet.columnCount },
    };

    sheet.name = 'Usage Chart';

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.send(buffer);
  } catch (error) {
    console.error('Formatting error:', error);
    res.status(500).send(`Formatting error: ${error.message}`);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Formatter running on port ${PORT}`);
});