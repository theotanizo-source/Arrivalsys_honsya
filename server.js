const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');

const app    = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static(path.join(__dirname)));

app.post('/upload', upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, {
      type: 'buffer',
      cellStyles: true,
      cellNF: true,
      sheetStubs: true
    });

    const results = {};
    wb.SheetNames.forEach(name => {
      results[name] = parseSheet(wb, name);
    });

    res.json({ sheetNames: wb.SheetNames, sheets: results });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

function parseSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return { rows: [], colWidths: [], rowHeights: [], merges: [] };

  const range    = XLSX.utils.decode_range(ws['!ref']);
  const colWidths  = [];
  const rowHeights = [];

  if (ws['!cols']) {
    ws['!cols'].forEach((col, i) => {
      colWidths[i] = col && col.wch ? Math.round(col.wch * 7) : 64;
    });
  }
  if (ws['!rows']) {
    ws['!rows'].forEach((row, i) => {
      rowHeights[i] = row && row.hpx ? row.hpx
                    : row && row.hpt ? Math.round(row.hpt * 1.33)
                    : 20;
    });
  }

  const merges = (ws['!merges'] || []).map(m => ({
    sr: m.s.r, sc: m.s.c, er: m.e.r, ec: m.e.c
  }));

  const rows = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const cells = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      cells.push({
        v: cell ? (cell.w !== undefined ? cell.w : (cell.v !== undefined ? String(cell.v) : '')) : '',
        s: cell ? (cell.s || null) : null
      });
    }
    rows.push(cells);
  }

  return { rows, colWidths, rowHeights, merges,
           minR: range.s.r, minC: range.s.c };
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
