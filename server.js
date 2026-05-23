const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const XLSX_PATH = path.join(__dirname, 'data', 'MegaSena_Dashboard.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Load and parse xlsx ─────────────────────────────────────────────────────
async function loadDraws() {
  if (!fs.existsSync(XLSX_PATH)) {
    throw new Error(`Arquivo não encontrado: ${XLSX_PATH}`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(XLSX_PATH);

  const sheet = workbook.getWorksheet('📥 DADOS') || workbook.worksheets[0];
  const draws = [];
  let headerRow = -1;

  sheet.eachRow((row, rowNum) => {
    const vals = row.values; // 1-indexed
    if (headerRow === -1) {
      const str = vals.slice(1).join('|').toLowerCase();
      if (str.includes('concurso') && str.includes('bola')) {
        headerRow = rowNum;
      }
      return;
    }
    const concurso = vals[1];
    const data = vals[2];
    const bolas = [vals[3], vals[4], vals[5], vals[6], vals[7], vals[8]];
    if (!concurso || typeof concurso !== 'number') return;
    if (bolas.some(b => !b || typeof b !== 'number')) return;

    let dataStr = '';
    if (data instanceof Date) {
      dataStr = data.toLocaleDateString('pt-BR');
    } else if (typeof data === 'string') {
      dataStr = data;
    } else {
      dataStr = String(data);
    }

    draws.push({
      concurso: Math.round(concurso),
      data: dataStr,
      bolas: bolas.map(b => Math.round(b)).sort((a, b) => a - b),
    });
  });

  draws.sort((a, b) => a.concurso - b.concurso);
  return draws;
}

// ─── Frequency analysis ───────────────────────────────────────────────────────
function calcFrequency(draws) {
  const freq = {};
  for (let i = 1; i <= 60; i++) freq[i] = 0;

  for (const d of draws) {
    for (const b of d.bolas) freq[b]++;
  }

  const total = draws.length;
  const values = Object.values(freq);
  const max = Math.max(...values);
  const min = Math.min(...values);
  const threshold_hot = max - (max - min) * 0.2;
  const threshold_cold = min + (max - min) * 0.2;

  const lastSeen = {};
  for (let i = 1; i <= 60; i++) lastSeen[i] = -1;
  for (let i = 0; i < draws.length; i++) {
    for (const b of draws[i].bolas) lastSeen[b] = i;
  }

  const result = [];
  for (let i = 1; i <= 60; i++) {
    const f = freq[i];
    const atraso = total - 1 - lastSeen[i];
    let categoria = 'Médio';
    if (f >= threshold_hot) categoria = 'Quente';
    else if (f <= threshold_cold) categoria = 'Frio';

    const last100 = draws.slice(-100);
    const freqLast100 = last100.filter(d => d.bolas.includes(i)).length;
    const expectedLast100 = (f / total) * 100;
    const tendencia = ((freqLast100 - expectedLast100) / expectedLast100) * 100;

    result.push({ numero: i, frequencia: f, taxa: f / total, categoria, atraso, tendencia });
  }
  return result;
}

// ─── Stats ────────────────────────────────────────────────────────────────────
function calcStats(draws) {
  const somaDist = {};
  const pariDist = { 0:0,1:0,2:0,3:0,4:0,5:0,6:0 };

  for (const d of draws) {
    const soma = d.bolas.reduce((a, b) => a + b, 0);
    const faixa = getSomaFaixa(soma);
    somaDist[faixa] = (somaDist[faixa] || 0) + 1;
    const pares = d.bolas.filter(b => b % 2 === 0).length;
    pariDist[pares]++;
  }

  return { somaDist, pariDist, total: draws.length };
}

function getSomaFaixa(soma) {
  if (soma < 100) return 'Abaixo de 100';
  if (soma <= 129) return '100 – 129';
  if (soma <= 159) return '130 – 159';
  if (soma <= 189) return '160 – 189';
  if (soma <= 219) return '190 – 219';
  if (soma <= 249) return '220 – 249';
  if (soma <= 279) return '250 – 279';
  return '280 ou mais';
}

// ─── Game generator ───────────────────────────────────────────────────────────
function generateGames(draws, params = {}) {
  const {
    somaMin = 130, somaMax = 230,
    minPares = 2, maxPares = 4,
    count = 10,
  } = params;

  const freq = calcFrequency(draws);
  const quentes = freq.filter(f => f.categoria === 'Quente').map(f => f.numero);
  const frios = freq.filter(f => f.categoria === 'Frio').map(f => f.numero);
  const atrasados = freq.sort((a, b) => b.atraso - a.atraso).slice(0, 15).map(f => f.numero);
  const tendAlta = freq.filter(f => f.tendencia > 20).map(f => f.numero);
  const tendBaixa = freq.filter(f => f.tendencia < -20).map(f => f.numero);

  const strategies = ['Balanceado', 'Quentes', 'Frios', 'Atrasados', 'Tendência alta', 'Tendência baixa'];
  const games = [];
  let attempts = 0;

  while (games.length < count && attempts < 50000) {
    attempts++;
    const strategy = strategies[games.length % strategies.length];
    let pool;

    if (strategy === 'Quentes' && quentes.length >= 6) pool = quentes;
    else if (strategy === 'Frios' && frios.length >= 6) pool = frios;
    else if (strategy === 'Atrasados' && atrasados.length >= 6) pool = atrasados;
    else if (strategy === 'Tendência alta' && tendAlta.length >= 6) pool = tendAlta;
    else if (strategy === 'Tendência baixa' && tendBaixa.length >= 6) pool = tendBaixa;
    else pool = Array.from({length: 60}, (_, i) => i + 1);

    const shuffled = [...pool].sort(() => Math.random() - 0.5);
    const picked = shuffled.slice(0, 6).sort((a, b) => a - b);

    const soma = picked.reduce((a, b) => a + b, 0);
    const pares = picked.filter(b => b % 2 === 0).length;

    if (soma >= somaMin && soma <= somaMax && pares >= minPares && pares <= maxPares) {
      games.push({ numeros: picked, soma, pares, estrategia: strategy, valido: true });
    }
  }

  return games;
}

// ─── API Routes ───────────────────────────────────────────────────────────────
let cachedDraws = null;
let lastLoaded = 0;

async function getDraws() {
  if (cachedDraws && Date.now() - lastLoaded < 60000) return cachedDraws;
  cachedDraws = await loadDraws();
  lastLoaded = Date.now();
  return cachedDraws;
}

app.get('/api/draws', async (req, res) => {
  try {
    const draws = await getDraws();
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 20;
    const start = (page - 1) * limit;
    const paginated = draws.slice().reverse().slice(start, start + limit);
    res.json({ draws: paginated, total: draws.length, page, limit });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/frequency', async (req, res) => {
  try {
    const draws = await getDraws();
    res.json(calcFrequency(draws));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/stats', async (req, res) => {
  try {
    const draws = await getDraws();
    const freq = calcFrequency(draws);
    const stats = calcStats(draws);
    const latest = draws.slice(-1)[0];
    const mostFreq = [...freq].sort((a, b) => b.frequencia - a.frequencia)[0];
    const leastFreq = [...freq].sort((a, b) => a.frequencia - b.frequencia)[0];
    const mostDelayed = [...freq].sort((a, b) => b.atraso - a.atraso)[0];
    res.json({ ...stats, latest, mostFreq, leastFreq, mostDelayed });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/generate', async (req, res) => {
  try {
    const draws = await getDraws();
    const games = generateGames(draws, req.body);
    res.json(games);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/reload', async (req, res) => {
  cachedDraws = null;
  try {
    const draws = await getDraws();
    res.json({ ok: true, total: draws.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('*', (_, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, () => {
  console.log(`🎲 Lottery Dashboard rodando em http://localhost:${PORT}`);
  if (!fs.existsSync(XLSX_PATH)) {
    console.warn(`⚠️  Arquivo Excel não encontrado. Coloque o arquivo em: ${XLSX_PATH}`);
  }
});
