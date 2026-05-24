const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure data directory exists on startup
const UPLOAD_DIR = path.join(__dirname, '../data');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    cb(null, 'lottery-data' + path.extname(file.originalname));
  }
});
const upload = multer({ storage, fileFilter: (req, file, cb) => {
  const ext = path.extname(file.originalname).toLowerCase();
  if (['.xlsx', '.xls', '.csv'].includes(ext)) cb(null, true);
  else cb(new Error('Apenas arquivos .xlsx, .xls ou .csv são aceitos'));
}});

app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Nenhum arquivo enviado' });
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: null });
    res.json({ success: true, rows: data.length, columns: Object.keys(data[0] || {}), data });
  } catch (err) {
    res.status(500).json({ error: 'Erro ao processar arquivo: ' + err.message });
  }
});

app.get('/api/data', (req, res) => {
  const dataPath = path.join(__dirname, '../data/lottery-data.xlsx');
  const csvPath = path.join(__dirname, '../data/lottery-data.csv');
  const filePath = fs.existsSync(dataPath) ? dataPath : fs.existsSync(csvPath) ? csvPath : null;
  if (!filePath) return res.json({ data: [], columns: [] });
  try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: null });
    res.json({ data, columns: Object.keys(data[0] || {}) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`\n🎰 Lottery Dashboard rodando em http://localhost:${PORT}\n`);
});
