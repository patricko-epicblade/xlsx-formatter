const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');

const upload = multer();
const app = express();

app.post('/format', upload.single('file'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const sheet = workbook.worksheets[0];

    const companyName = req.body.companyName || 'Customer';
    const reportStart = req.body.reportStart || '';
    const reportEnd = req.body.reportEnd || '';

    // -------------------------
    // 1. Insert title rows
    // -------------------------
    sheet.spliceRows(1, 0,
      [`Customer Usage Report - ${companyName}`],
      [`Period: ${reportStart} through ${reportEnd}`]
    );

    const lastCol = sheet.columnCount;

    sheet.mergeCells(1, 1, 1, lastCol);
    sheet.mergeCells(2, 1, 2, lastCol);

    sheet.getRow(1).font = { bold: true, size: 16 };
    sheet.getRow(1).alignment = { horizontal: 'center' };

    sheet.getRow(2).font = { italic: true };
    sheet.getRow(2).alignment = { horizontal: 'center' };

    // -------------------------
    // 2. Header styling
    // -------------------------
    const headerRowIndex = 3;
    const headerRow = sheet.getRow(headerRowIndex);

    headerRow.font = { bold: true };

    headerRow.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFBFBFBF' }
      };
      cell.border = {
        bottom: { style: 'thin' }
      };
      cell.alignment = { horizontal: 'center' };
    });

    // -------------------------
    // 3. Freeze header
    // -------------------------
    sheet.views = [{ state: 'frozen', ySplit: headerRowIndex }];

    // -------------------------
    // 4. Column formatting
    // -------------------------
    sheet.columns.forEach((col, colIndex) => {
      let maxLength = 10;

      col.eachCell({ includeEmpty: true }, (cell, rowNumber) => {

        // Replace 0 with blank
        if (cell.value === 0) {
          cell.value = '';
        }

        // Right-align numeric cells (data rows only)
        if (rowNumber > headerRowIndex && typeof cell.value === 'number') {
          cell.alignment = { horizontal: 'right' };
        }

        // Track width
        if (cell.value) {
          maxLength = Math.max(maxLength, cell.value.toString().length);
        }
      });

      col.width = maxLength + 2;
    });

    // -------------------------
    // 5. Bold OEM / first column
    // -------------------------
    const oemColumn = sheet.getColumn(1);
    oemColumn.eachCell((cell, rowNumber) => {
      if (rowNumber > headerRowIndex) {
        cell.font = { bold: true };
      }
    });

    // -------------------------
    // 6. Add filter row
    // -------------------------
    sheet.autoFilter = {
      from: {
        row: headerRowIndex,
        column: 1
      },
      to: {
        row: headerRowIndex,
        column: sheet.columnCount
      }
    };

    // -------------------------
    // 7. Clean sheet name
    // -------------------------
    sheet.name = "Usage Chart";

    // -------------------------
    // 8. Output file
    // -------------------------
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send('Formatting error');
  }
});

app.listen(3000, () => console.log('Formatter running'));