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
    const outputFileName = req.body.outputFileName || 'report.xlsx';

    // Insert top rows
    sheet.spliceRows(
      1,
      0,
      [`Customer Usage Report - ${companyName}`],
      [`Period: ${reportStart} through ${reportEnd}`]
    );

    const headerRowIndex = 3;
    const dataStartRow = headerRowIndex + 1;
    const lastCol = Math.max(sheet.columnCount, 1);
    const lastRow = sheet.rowCount;

    // Merge title rows
    sheet.mergeCells(1, 1, 1, lastCol);
    sheet.mergeCells(2, 1, 2, lastCol);

    // Row 1 / 2: keep visible, left aligned
    sheet.getRow(1).height = 24;
    sheet.getRow(2).height = 20;

    sheet.getCell(1, 1).font = { bold: true, size: 16 };
    sheet.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };

    sheet.getCell(2, 1).font = { italic: true, size: 12 };
    sheet.getCell(2, 1).alignment = { horizontal: 'left', vertical: 'middle' };

    // Header styling
    const headerRow = sheet.getRow(headerRowIndex);
    headerRow.font = { bold: true };
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFBFBFBF' },
      };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
      };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
      };
    });

    // Freeze top rows
    sheet.views = [{ state: 'frozen', ySplit: headerRowIndex }];

    // Identify column types from headers
    const columnInfo = [];
    for (let col = 1; col <= lastCol; col++) {
      const rawHeader = sheet.getCell(headerRowIndex, col).value;
      const headerText = rawHeader == null ? '' : String(rawHeader).trim();
      const isPartColumn = col === 1;
      const isTotalColumn = /TOTAL/i.test(headerText);
      const isDateColumn = /^\d{4}-\d{2}$/.test(headerText);
      const isCompanyColumn = /^company$/i.test(headerText);
      const yearMatch = headerText.match(/^(\d{4})/);
      const year = yearMatch ? yearMatch[1] : null;

      columnInfo[col] = {
        headerText,
        isPartColumn,
        isTotalColumn,
        isDateColumn,
        isCompanyColumn,
        year,
      };
    }

    // Group date columns by year only (exclude TOTAL columns)
    for (let col = 1; col <= lastCol; col++) {
      const info = columnInfo[col];
      if (info && info.isDateColumn) {
        sheet.getColumn(col).outlineLevel = 1;
      }
    }
    sheet.properties.outlineLevelCol = 1;

    // Make Company header invisible: white text on white background
    for (let col = 1; col <= lastCol; col++) {
      const info = columnInfo[col];
      if (info.isCompanyColumn) {
        const headerCell = sheet.getCell(headerRowIndex, col);
        headerCell.font = {
          ...(headerCell.font || {}),
          bold: true,
          color: { argb: 'FFFFFFFF' },
        };
        headerCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' },
        };
        headerCell.border = {};
      }
    }

    // Table/body formatting
    for (let rowNumber = dataStartRow; rowNumber <= lastRow; rowNumber++) {
      for (let col = 1; col <= lastCol; col++) {
        const cell = sheet.getCell(rowNumber, col);
        const info = columnInfo[col];
        const value = cell.value;

        // Replace numeric 0 with blank
        if (value === 0) {
          cell.value = '';
        }

        // Default alignment
        if (info.isPartColumn || info.isCompanyColumn) {
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else if (typeof cell.value === 'number') {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }

        // Column A bold
        if (info.isPartColumn) {
          cell.font = { ...(cell.font || {}), bold: true };
        }

        // Company column body cells invisible: white text on white background
        if (info.isCompanyColumn) {
          cell.font = { ...(cell.font || {}), color: { argb: 'FFFFFFFF' } };
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          continue;
        }

        // TOTAL columns below header: light blue fill + bold text
        if (info.isTotalColumn) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFDDEBF7' },
          };
          cell.font = { ...(cell.font || {}), bold: true };
        }

        // Non-TOTAL numeric data cells (excluding column A): light green fill
        if (
          !info.isPartColumn &&
          !info.isTotalColumn &&
          typeof cell.value === 'number' &&
          cell.value !== 0
        ) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE2F0D9' },
          };
        }
      }
    }

    // Column widths
    for (let col = 1; col <= lastCol; col++) {
      const info = columnInfo[col];
      const column = sheet.getColumn(col);

      // Part column: wide enough so text stays inside column A
      if (info.isPartColumn) {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          const text = cell.value == null ? '' : String(cell.value);
          maxLength = Math.max(maxLength, text.length);
        });
        column.width = Math.max(35, maxLength + 2);
        continue;
      }

      // Company column: keep it available but visually hidden
      if (info.isCompanyColumn) {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
          const text = cell.value == null ? '' : String(cell.value);
          maxLength = Math.max(maxLength, text.length);
        });
        column.width = Math.max(18, Math.min(maxLength + 2, 30));
        continue;
      }

      // TOTAL columns: fixed width to show full text
      if (info.isTotalColumn) {
        column.width = 17;
        continue;
      }

      // All other columns: autosize
      let maxLength = 10;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const text = cell.value == null ? '' : String(cell.value);
        maxLength = Math.max(maxLength, text.length);
      });
      column.width = Math.min(Math.max(maxLength + 2, 10), 16);
    }

    // Auto filter on header row
    sheet.autoFilter = {
      from: { row: headerRowIndex, column: 1 },
      to: { row: headerRowIndex, column: lastCol },
    };

    // Clean sheet name
    sheet.name = 'Usage Chart';

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${outputFileName}"`
    );
    res.setHeader('Content-Length', buffer.length);

    res.end(buffer);
  } catch (error) {
    console.error('Formatting error:', error);
    res.status(500).send(`Formatting error: ${error.message}`);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Formatter running on port ${PORT}`);
});