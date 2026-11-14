const fs = require('fs');
const path = require('path');

// Carrega variáveis do arquivo .env caso ele exista localmente
try {
  const envPath = path.join(__dirname, '.env');
  if (fs.existsSync(envPath)) {
    const envConfig = fs.readFileSync(envPath, 'utf8');
    envConfig.split('\n').forEach(line => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) return;
      const [key, ...valueParts] = trimmed.split('=');
      if (key && valueParts.length > 0) {
        const val = valueParts.join('=').trim().replace(/(^['"]|['"]$)/g, '');
        process.env[key.trim()] = val;
      }
    });
  }
} catch (e) {
  console.warn('Aviso: Não foi possível carregar o arquivo .env:', e.message);
}

const express = require('express');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');

const JWT_SECRET = process.env.JWT_SECRET || 'supersecretkey_megasena_dashboard';
const USERS_PATH = path.join(__dirname, 'data', 'users.json');

function loadUsers() {
  if (!fs.existsSync(USERS_PATH)) {
    fs.writeFileSync(USERS_PATH, JSON.stringify([]));
    return [];
  }
  try {
    const data = fs.readFileSync(USERS_PATH, 'utf8');
    return JSON.parse(data || '[]');
  } catch (e) {
    return [];
  }
}

function saveUsers(users) {
  fs.writeFileSync(USERS_PATH, JSON.stringify(users, null, 2));
}

// Middleware to authenticate token
function authenticateToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Acesso negado. Faça login para gerar jogos.' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Token inválido ou expirado.' });
    }
    req.user = user;
    next();
  });
}

const app = express();
const PORT = process.env.PORT || 3000;
const XLSX_PATH = path.join(__dirname, 'data', 'MegaSena_Dashboard.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Load and parse xlsx ─────────────────────────────────────────────────────
async function parseCaixaRawXlsx(filePath) {
  const buf = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buf);
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

  const draws = [];
  const rowRegex = /<(?:x:)?row[^>]*>(.*?)<\/(?:x:)?row>/gs;
  let match;
  let isFirst = true;

  while ((match = rowRegex.exec(sheetXml)) !== null) {
    if (isFirst) {
      isFirst = false;
      continue;
    }
    const rowContent = match[1];
    const cellRegex = /<(?:x:)?c[^>]*?(?:t="([^"]*)")?[^>]*>.*?(?:<(?:x:)?v>(.*?)<\/(?:x:)?v>)?<\/(?:x:)?c>/gs;
    const cells = [];
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
      cells.push(cellMatch[2] !== undefined ? cellMatch[2] : '');
    }

    if (cells.length < 8) continue;
    const concurso = parseInt(cells[0], 10);
    const dataStr = cells[1];
    const b1 = parseInt(cells[2], 10);
    const b2 = parseInt(cells[3], 10);
    const b3 = parseInt(cells[4], 10);
    const b4 = parseInt(cells[5], 10);
    const b5 = parseInt(cells[6], 10);
    const b6 = parseInt(cells[7], 10);

    if (concurso > 0 && !isNaN(concurso) && !isNaN(b1) && !isNaN(b2) && !isNaN(b3) && !isNaN(b4) && !isNaN(b5) && !isNaN(b6)) {
      draws.push({
        concurso,
        data: dataStr,
        bolas: [b1, b2, b3, b4, b5, b6].sort((a, b) => a - b),
      });
    }
  }

  draws.sort((a, b) => a.concurso - b.concurso);
  return draws;
}

async function loadDraws() {
  const candidates = [
    path.join(__dirname, 'data', 'Mega-Sena (3).xlsx'),
    path.join(__dirname, 'data', 'Mega-Sena (1).xlsx'),
    path.join(__dirname, 'data', 'Mega-Sena.xlsx'),
    path.join(__dirname, 'data', 'MegaSena_Dashboard.xlsx'),
  ];

  const foundPath = candidates.find(p => fs.existsSync(p));
  if (!foundPath) {
    throw new Error(`Nenhuma planilha encontrada em data/ (procurado por: Mega-Sena.xlsx, MegaSena_Dashboard.xlsx)`);
  }

  // Tenta primeiro via parser direto Caixa (suporta o formato original baixado da Caixa)
  try {
    const draws = await parseCaixaRawXlsx(foundPath);
    if (draws.length > 0) {
      console.log(`Carregados ${draws.length} concursos via parser Caixa de: ${path.basename(foundPath)}`);
      return draws;
    }
  } catch (err) {
    // Se não for formato Caixa raw, cai no ExcelJS abaixo
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(foundPath);

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
  console.log(`Carregados ${draws.length} concursos via ExcelJS de: ${path.basename(foundPath)}`);
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

  const last100 = draws.slice(-100);
  const result = [];
  for (let i = 1; i <= 60; i++) {
    const f = freq[i];
    const atraso = total - 1 - lastSeen[i];
    let categoria = 'Médio';
    if (f >= threshold_hot) categoria = 'Quente';
    else if (f <= threshold_cold) categoria = 'Frio';

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

// ─── Teoria dos Conjuntos (Coocorrência, Probabilidade Condicional e Lift) ──
function calcSetTheory(draws, skipTrios = false) {
  const N = 60;
  const k = 6;
  const T = draws.length;
  const expectedPairProb = (k * (k - 1)) / (N * (N - 1)); // 30 / 3540 = 1 / 118
  const expectedPairCount = T * expectedPairProb;
  const expectedTrioProb = (k * (k - 1) * (k - 2)) / (N * (N - 1) * (N - 2)); // 120 / 205320 = 1 / 1711
  const expectedTrioCount = T * expectedTrioProb;

  const numCounts = new Array(N + 1).fill(0);
  const pairCounts = Array.from({ length: N + 1 }, () => new Array(N + 1).fill(0));
  const trioMap = skipTrios ? null : new Map();

  for (const d of draws) {
    const b = d.bolas;
    for (let i = 0; i < b.length; i++) {
      numCounts[b[i]]++;
      for (let j = i + 1; j < b.length; j++) {
        pairCounts[b[i]][b[j]]++;
        pairCounts[b[j]][b[i]]++;
        if (!skipTrios) {
          for (let m = j + 1; m < b.length; m++) {
            const key = `${b[i]}-${b[j]}-${b[m]}`;
            trioMap.set(key, (trioMap.get(key) || 0) + 1);
          }
        }
      }
    }
  }

  const allPairs = [];
  const partnersByNumber = skipTrios ? null : {};
  if (!skipTrios) {
    for (let i = 1; i <= N; i++) partnersByNumber[i] = [];
  }

  for (let i = 1; i <= N; i++) {
    for (let j = i + 1; j <= N; j++) {
      const c = pairCounts[i][j];
      const pA = numCounts[i] / T;
      const pB = numCounts[j] / T;
      const pAB = c / T;
      const lift = pA * pB > 0 ? pAB / (pA * pB) : 0;
      const vsExpected = expectedPairCount > 0 ? c / expectedPairCount : 0;
      const confAtoB = numCounts[i] > 0 ? c / numCounts[i] : 0;
      const confBtoA = numCounts[j] > 0 ? c / numCounts[j] : 0;

      const pairObj = {
        a: i,
        b: j,
        count: c,
        support: pAB,
        confAtoB,
        confBtoA,
        lift,
        vsExpected,
      };
      allPairs.push(pairObj);

      if (!skipTrios) {
        partnersByNumber[i].push({
          parceiro: j,
          count: c,
          conf: confAtoB,
          lift,
          vsExpected,
        });
        partnersByNumber[j].push({
          parceiro: i,
          count: c,
          conf: confBtoA,
          lift,
          vsExpected,
        });
      }
    }
  }

  if (!skipTrios) {
    for (let i = 1; i <= N; i++) {
      partnersByNumber[i].sort((x, y) => y.count - x.count || y.lift - x.lift);
      partnersByNumber[i] = partnersByNumber[i].slice(0, 10);
    }
  }

  const topPairs = [...allPairs]
    .sort((x, y) => y.count - x.count || y.lift - x.lift)
    .slice(0, 20);

  if (skipTrios) {
    return {
      totalDraws: T,
      expectedPairProb,
      expectedPairCount,
      topPairs,
      pairCounts,
    };
  }

  const topLiftPairs = [...allPairs]
    .sort((x, y) => y.lift - x.lift || y.count - x.count)
    .slice(0, 20);

  const topTrios = Array.from(trioMap.entries())
    .map(([key, count]) => ({
      numeros: key.split('-').map(Number),
      count,
      support: count / T,
      vsExpected: expectedTrioCount > 0 ? count / expectedTrioCount : 0,
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 15);

  return {
    totalDraws: T,
    totalPossiblePairs: (N * (N - 1)) / 2,
    expectedPairProb,
    expectedPairCount,
    expectedTrioProb,
    expectedTrioCount,
    topPairs,
    topLiftPairs,
    topTrios,
    partnersByNumber,
    pairCounts,
  };
}

// ─── Game generator ───────────────────────────────────────────────────────────
function shuffle(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function checkPastWins(picked, draws) {
  const wins = { sena: 0, quina: 0, quadra: 0 };
  for (const d of draws) {
    const matches = picked.filter(n => d.bolas.includes(n)).length;
    if (matches === 6) wins.sena++;
    else if (matches === 5) wins.quina++;
    else if (matches === 4) wins.quadra++;
  }
  return wins;
}

function hasConsecutive(picked) {
  let count = 1;
  for (let i = 1; i < picked.length; i++) {
    if (picked[i] === picked[i - 1] + 1) {
      count++;
      if (count >= 3) return true;
    } else {
      count = 1;
    }
  }
  return false;
}

function createSeededRng(seed) {
  let a = (seed + 0x6D2B79F5) | 0;
  return function () {
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffleWithRng(array, rng = Math.random) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function generateGames(draws, params = {}) {
  if (!params._raw && (params.mode === 'regressivo_elite' || params.regressiveFilter === true || params.regressiveFilter === 'true')) {
    return selectRegressiveGames(draws, params);
  }

  const somaMin = Number(params.somaMin) || 130;
  const somaMax = Number(params.somaMax) || 230;
  const count = Math.min(Math.max(Number(params.count) || 10, 1), 30);
  const mode = ['vetor_feixe', 'vetor_nucleo'].includes(params.mode) ? params.mode : 'vetor_paginado';
  const pageOffset = Math.max(Number(params.pageOffset) || 0, 0);
  const rng = params._beamSeed !== undefined ? createSeededRng(Number(params._beamSeed)) : Math.random;

  const PRECO_UNITARIO = 5.00;
  const N = 60;
  const T = draws.length || 1;

  // 1. Frequência individual P(A) + Matriz de Coocorrência P(A ∩ B) e Condicional P(B | A)
  const freqList = calcFrequency(draws);
  const { pairCounts, expectedPairCount, topPairs } = calcSetTheory(draws, true);
  const expectedCondProb = 5 / (N - 1); // ~8.47%

  const freqMap = {};
  let minFreq = Infinity;
  let maxFreq = -Infinity;
  for (const f of freqList) {
    freqMap[f.numero] = f.frequencia;
    if (f.frequencia < minFreq) minFreq = f.frequencia;
    if (f.frequencia > maxFreq) maxFreq = f.frequencia;
  }

  // Vetor de peso marginal individual P(A) normalizado (1.0 a 1.35)
  const baseWeights = {};
  for (let i = 1; i <= N; i++) {
    const f = freqMap[i] || minFreq;
    baseWeights[i] = 1.0 + 0.35 * ((f - minFreq) / (maxFreq - minFreq || 1));
  }

  // Matriz de Sinergia Combinada (Probabilidade Conjunta P(A∩B) + Probabilidade Condicional P(B|A))
  const pairSynergy = Array.from({ length: N + 1 }, () => new Array(N + 1).fill(1));
  const allSeedPairs = [];

  for (let a = 1; a <= N; a++) {
    for (let b = a + 1; b <= N; b++) {
      const c = pairCounts[a][b] || 0;
      const ratioJoint = expectedPairCount > 0 ? c / expectedPairCount : 1;
      const confAB = freqMap[a] > 0 ? c / freqMap[a] : 0;
      const confBA = freqMap[b] > 0 ? c / freqMap[b] : 0;
      const ratioCond = expectedCondProb > 0 ? ((confAB + confBA) / 2) / expectedCondProb : 1;

      // Score combina 50% Prob. Conjunta P(A∩B) + 50% Prob. Condicional P(B|A)
      const score = 0.5 * ratioJoint + 0.5 * ratioCond;
      pairSynergy[a][b] = score;
      pairSynergy[b][a] = score;

      allSeedPairs.push({ a, b, count: c, score, ratioJoint });
    }
  }

  // Ordena o vetor mestre de pares do maior para o menor score
  allSeedPairs.sort((x, y) => y.score - x.score || y.count - x.count);

  const batchUsage = {};
  for (let i = 1; i <= N; i++) batchUsage[i] = 0;
  const usedPairsInBatch = new Set();
  const games = [];

  // Função auxiliar: pontua um candidato x em relação ao subconjunto parcial J já escolhido
  function scoreCandidateWithSet(x, partialSet) {
    let sumSynergy = 0;
    for (const y of partialSet) {
      sumSynergy += pairSynergy[x][y];
    }
    const avgSynergy = sumSynergy / partialSet.length;
    // Aplica penalidade moderada de repetição no lote para equilibrar cobertura e sinergia alta
    const usagePenalty = Math.pow(1 + 0.22 * batchUsage[x], 1.5);
    return (avgSynergy * baseWeights[x]) / usagePenalty;
  }

  // Verifica se o candidato x pode entrar no conjunto parcial sem violar |J_i ∩ J_k| <= 1
  function canAddCandidate(x, partialSet, strictDisjoint = true) {
    if (partialSet.includes(x)) return false;
    if (strictDisjoint) {
      for (const y of partialSet) {
        const key = x < y ? `${x}-${y}` : `${y}-${x}`;
        if (usedPairsInBatch.has(key)) return false;
      }
    }
    const testArr = [...partialSet, x].sort((a, b) => a - b);
    if (hasConsecutive(testArr)) return false;
    return true;
  }

  // ─── Modo C: Fechamento Condensado (5 Canais de Fusão Estrutural de Pontes) ──
  if (mode === 'vetor_nucleo') {
    const base20 = generateGames(draws, { mode: 'vetor_paginado', count: 20, pageOffset: 0, _raw: true });
    const g1 = base20[0].numeros;
    const a = topPairs[0].a;
    const b = topPairs[0].b;
    const gA = (base20.find(g => g.numeros.includes(a) && !g.numeros.includes(b))?.numeros || []).filter(n => n !== a && n !== b);
    const gB = (base20.find(g => g.numeros.includes(b) && !g.numeros.includes(a))?.numeros || []).filter(n => n !== a && n !== b);
    const g2 = base20[1].numeros;
    const g3 = base20[2].numeros;
    const g4 = base20[3].numeros;
    const g5 = base20[4].numeros;

    // Atraso de cada dezena (0 = saiu no último concurso, >=1 = janela ativa)
    const lastSeen = {};
    for (let i = 1; i <= N; i++) lastSeen[i] = T;
    for (let dIdx = T - 1; dIdx >= 0; dIdx--) {
      const balls = draws[dIdx].bolas || draws[dIdx].numeros || draws[dIdx].dezenas || [];
      for (const ball of balls) {
        const d = T - 1 - dIdx;
        if (d < lastSeen[ball]) lastSeen[ball] = d;
      }
    }

    const getActiveWeight = (x) => {
      const d = lastSeen[x] || 0;
      if (d >= 1 && d <= 12) return 1.18;
      if (d >= 13 && d <= 20) return 1.08;
      if (d === 0) return 0.92;
      return 0.96;
    };

    const getLiftPure = (x, y) => ((pairCounts[x][y] || 0) * T) / ((freqMap[x] || 1) * (freqMap[y] || 1));
    const getHybridLift = (x, y) => {
      const c = pairCounts[x][y] || 0;
      return 0.5 * getLiftPure(x, y) + 0.5 * (c / (expectedPairCount || 1));
    };

    const combinations = (arr, k) => {
      if (k === 0) return [[]];
      if (arr.length < k) return [];
      const res = [];
      for (let i = 0; i <= arr.length - k; i++) {
        for (const t of combinations(arr.slice(i + 1), k - 1)) res.push([arr[i], ...t]);
      }
      return res;
    };

    // ─── CANAL 1: Espinha Dorsal de 4 Nós (G3 \ G1 + Pass2(G3) + GB) ───────────
    const ch1_G3Spine = [];
    {
      let g3Core = g3.filter(n => !g1.includes(n));
      if (g3Core.filter(x => lastSeen[x] >= 2 && lastSeen[x] <= 22).length < 3) {
        for (const extra of g4) {
          if (!g1.includes(extra) && !g3Core.includes(extra)) g3Core.push(extra);
        }
      }
      const g3Pass2Partners = {};
      for (const u of g3Core) {
        const p2Game = base20.slice(10).find(g => g.numeros.includes(u));
        if (p2Game) {
          const others = p2Game.numeros.filter(x => !g3Core.includes(x) && lastSeen[x] >= 2 && lastSeen[x] <= 22);
          others.sort((x, y) => getLiftPure(u, y) - getLiftPure(u, x));
          g3Pass2Partners[u] = others.slice(0, 2);
        } else {
          g3Pass2Partners[u] = [];
        }
      }
      const gBFiltered = gB.filter(x => lastSeen[x] >= 2 && lastSeen[x] <= 22);
      const g3Filtered = g3Core.filter(x => lastSeen[x] >= 2 && lastSeen[x] <= 22);
      const seen1 = new Map();

      for (const trio of combinations(g3Filtered, 3)) {
        const perms = [
          [trio[0], trio[1], trio[2]],
          [trio[0], trio[2], trio[1]],
          [trio[1], trio[2], trio[0]],
        ];
        for (const [u, v, w] of perms) {
          for (const pU of g3Pass2Partners[u] || []) {
            for (const pV of g3Pass2Partners[v] || []) {
              if (pU === pV) continue;
              const spineEdges = [
                pairCounts[u][v] || 0,
                pairCounts[u][pU] || 0,
                pairCounts[v][pV] || 0,
                pairCounts[pU][pV] || 0,
              ];
              const minSpine = Math.min(...spineEdges);
              const avgSpine = spineEdges.reduce((s, c) => s + c, 0) / 4;

              for (const xB of gBFiltered) {
                const six = [u, v, w, pU, pV, xB].sort((x, y) => x - y);
                if (new Set(six).size !== 6 || hasConsecutive(six)) continue;
                const structEdges = [
                  ...spineEdges,
                  pairCounts[u][w] || 0,
                  pairCounts[v][w] || 0,
                  pairCounts[xB][u] || 0,
                  pairCounts[xB][v] || 0,
                  pairCounts[xB][w] || 0,
                  pairCounts[xB][pU] || 0,
                  pairCounts[xB][pV] || 0,
                ];
                const minStruct = Math.min(...structEdges);
                const wLift = (getLiftPure(w, u) + getLiftPure(w, v) + getLiftPure(w, pU) + getLiftPure(w, pV)) / 4;
                const xBLift = (getLiftPure(xB, u) + getLiftPure(xB, v) + getLiftPure(xB, pU) + getLiftPure(xB, pV) + getLiftPure(xB, w)) / 5;
                const score = minStruct * 100 + minSpine * 25 + avgSpine * 5 + (wLift + xBLift) * 10;
                const key = six.join('-');
                const ex = seen1.get(key);
                if (!ex || score > ex.score) {
                  seen1.set(key, { numeros: six, score, label: 'C1: Espinha G3\\G1 + Pass2 + GB' });
                }
              }
            }
          }
        }
      }
      ch1_G3Spine.push(...[...seen1.values()].sort((x, y) => y.score - x.score));
    }

    // ─── CANAL 2: Fusão de Ramificações Disjuntas (3×gBUnique + 2×g2Unique + 1×gAUnique) ───
    const ch2_DisjointBranches = [];
    {
      const gBUnique = gB.filter(n => !g2.includes(n) && !gA.includes(n) && lastSeen[n] >= 1);
      const gAUnique = gA.filter(n => !g2.includes(n) && !gB.includes(n));
      const g2Unique = g2.filter(n => !gA.includes(n) && !gB.includes(n) && lastSeen[n] >= 1);
      const seen2 = new Set();

      for (const tB of combinations(gBUnique, 3)) {
        for (const p2 of combinations(g2Unique, 2)) {
          for (const xA of gAUnique) {
            const six = [...tB, ...p2, xA].sort((x, y) => x - y);
            if (new Set(six).size !== 6 || hasConsecutive(six)) continue;
            const key = six.join('-');
            if (seen2.has(key)) continue;
            seen2.add(key);

            const structLifts = [];
            const structCounts = [];
            for (let i = 0; i < 6; i++) {
              const inc = [];
              for (let j = 0; j < 6; j++) {
                if (i !== j) inc.push({ c: pairCounts[six[i]][six[j]] || 0, l: getHybridLift(six[i], six[j]) });
              }
              const incByLift = [...inc].sort((x, y) => y.l - x.l);
              const incByCount = [...inc].sort((x, y) => y.c - x.c);
              structLifts.push(incByLift[0].l, incByLift[1].l);
              structCounts.push(incByCount[0].c, incByCount[1].c);
            }
            const avgTop2Lift = structLifts.reduce((s, x) => s + x, 0) / structLifts.length;
            const minTop2Count = Math.min(...structCounts);
            const closesPass1Cycle = base20.slice(0, 10).some(
              g => g.numeros.includes(xA) && tB.some(tb => g.numeros.includes(tb))
            );
            const cycleBonus = closesPass1Cycle ? 12 : 0;
            const score = minTop2Count * 10 + avgTop2Lift * 20 + cycleBonus;
            ch2_DisjointBranches.push({ numeros: six, score, label: 'C2: Fusão Disjunta GB + G2 + GA' });
          }
        }
      }
      ch2_DisjointBranches.sort((x, y) => y.score - x.score);
    }

    // ─── CANAL 3: Ponte Transversal 2×g3 + 2×g2 + 2×gB (Grafo 2-Conectado) ─────
    const scoreBridgeGraph = (nums) => {
      const sorted = [...nums].sort((x, y) => x - y);
      if (sorted.length !== 6 || new Set(sorted).size !== 6 || hasConsecutive(sorted)) return null;
      const soma = sorted.reduce((s, n) => s + n, 0);
      if (soma < 105 || soma > 255) return null;
      const pares = sorted.filter(n => n % 2 === 0).length;
      if (pares < 1 || pares > 5) return null;

      let sumTop2Incident = 0, minTop1Incident = Infinity, minTop2Incident = Infinity, sumLiftTop2 = 0;
      for (let i = 0; i < 6; i++) {
        const v = sorted[i];
        const inc = [];
        for (let j = 0; j < 6; j++) {
          if (i !== j) inc.push({ c: pairCounts[v][sorted[j]] || 0, l: getLiftPure(v, sorted[j]) });
        }
        inc.sort((x, y) => y.c - x.c);
        sumTop2Incident += inc[0].c * 1.4 + inc[1].c * 1.0;
        sumLiftTop2 += inc[0].l * 1.4 + inc[1].l * 1.0;
        if (inc[0].c < minTop1Incident) minTop1Incident = inc[0].c;
        if (inc[1].c < minTop2Incident) minTop2Incident = inc[1].c;
      }
      let twoHopBridgeScore = 0;
      const allEdges = [];
      for (let i = 0; i < 6; i++) {
        for (let j = i + 1; j < 6; j++) {
          const u = sorted[i], v = sorted[j];
          const direct = pairCounts[u][v] || 0;
          allEdges.push(direct);
          let bestHop = 0;
          for (let k = 0; k < 6; k++) {
            if (k !== i && k !== j) {
              const hop = Math.min(pairCounts[u][sorted[k]] || 0, pairCounts[v][sorted[k]] || 0);
              if (hop > bestHop) bestHop = hop;
            }
          }
          if (direct <= 25 && bestHop >= 28) twoHopBridgeScore += (bestHop - 24) * 1.8;
          else if (bestHop >= 30) twoHopBridgeScore += (bestHop - 27) * 0.8;
        }
      }
      allEdges.sort((x, y) => y - x);
      const top6EdgesAvg = allEdges.slice(0, 6).reduce((s, c) => s + c, 0) / 6;
      const activeMult = sorted.reduce((s, x) => s + getActiveWeight(x), 0) / 6;
      const score = (
        (sumTop2Incident / 6) * 2.2 +
        minTop1Incident * 2.5 +
        minTop2Incident * 2.0 +
        top6EdgesAvg * 2.5 +
        sumLiftTop2 * 6.0 +
        twoHopBridgeScore * 1.2
      ) * activeMult;
      return { numeros: sorted, score, label: 'C3: Ponte Transversal G3 + G2 + GB' };
    };

    const ch3_Transversal = [];
    {
      const seen3 = new Set();
      for (const t3 of combinations(g3, 3)) {
        for (const xB of gB) {
          for (const pExtra of combinations([...new Set([...gB, ...g2])], 2)) {
            const r = scoreBridgeGraph([...t3, xB, ...pExtra]);
            if (r && !seen3.has(r.numeros.join('-'))) {
              seen3.add(r.numeros.join('-'));
              ch3_Transversal.push(r);
            }
          }
        }
      }
      for (const p3 of combinations(g3, 2)) {
        for (const p2 of combinations(g2, 2)) {
          for (const pB of combinations(gB, 2)) {
            const r = scoreBridgeGraph([...p3, ...p2, ...pB]);
            if (r && !seen3.has(r.numeros.join('-'))) {
              seen3.add(r.numeros.join('-'));
              ch3_Transversal.push(r);
            }
          }
        }
      }
      ch3_Transversal.sort((x, y) => y.score - x.score);
    }

    // ─── CANAL 4: Fusão Clássica do Pivô [a, b] + g2 + g3 + gA/gB ──────────────
    const ch4_PivotCore = [];
    {
      const ch4A = [];
      const g2Act = g2.filter(x => x !== a && x !== b && lastSeen[x] >= 2 && lastSeen[x] <= 22);
      const g3Mat = [...g3.filter(x => x !== a && x !== b && lastSeen[x] >= 2 && lastSeen[x] <= 22)]
        .sort((x, y) => lastSeen[y] - lastSeen[x]);
      if (g2Act.length >= 2 && g3Mat.length >= 2) {
        for (const p2 of combinations(g2Act.slice(0, 3), 2)) {
          for (const p3 of combinations(g3Mat.slice(0, 3), 2)) {
            const six = [a, b, ...p2, ...p3].sort((x, y) => x - y);
            if (new Set(six).size === 6 && !hasConsecutive(six)) {
              const score = (lastSeen[p3[0]] + lastSeen[p3[1]]) * 5 + (pairCounts[p2[0]][p2[1]] || 0);
              ch4A.push({ numeros: six, score, label: `C4A: Pivô ${String(a).padStart(2, '0')}∩${String(b).padStart(2, '0')} + G2 + G3` });
            }
          }
        }
        ch4A.sort((x, y) => y.score - x.score);
      }

      const ch4B = [];
      const bridgePool = [...new Set([...g2, ...g3])].filter(n => n !== a && n !== b && lastSeen[n] >= 1 && lastSeen[n] <= 22);
      const cores5 = [];
      for (const xA of gA) {
        for (const xB of gB) {
          if (xA === xB) continue;
          for (const y1 of bridgePool) {
            if ([a, b, xA, xB].includes(y1)) continue;
            const edges = [
              pairCounts[xA][xB] || 0,
              pairCounts[y1][xA] || 0,
              pairCounts[y1][xB] || 0,
              pairCounts[y1][b] || 0,
              pairCounts[xA][a] || 0,
              pairCounts[xB][b] || 0,
            ];
            const minE = Math.min(...edges);
            const geo = Math.pow(edges.reduce((p, q) => p * (q || 0.5), 1), 1 / edges.length);
            cores5.push({ core: [a, b, xA, xB, y1], xA, xB, y1, minE, geo });
          }
        }
      }
      cores5.sort((x, y) => y.minE - x.minE || y.geo - x.geo);
      const seen4B = new Set();
      for (const c of cores5.slice(0, 25)) {
        for (const z of bridgePool) {
          if (c.core.includes(z)) continue;
          const nums = [...c.core, z].sort((x, y) => x - y);
          if (hasConsecutive(nums)) continue;
          const nAcima31 = nums.filter(n => n > 31).length;
          if (nAcima31 < 2) continue;
          const key = nums.join('-');
          if (seen4B.has(key)) continue;
          seen4B.add(key);
          const zEdges = c.core.map(m => pairCounts[z][m] || 0);
          const score = c.minE * 20 + c.geo * 3 + Math.min(...zEdges) * 4;
          ch4B.push({ numeros: nums, score, label: `C4B: Pivô ${String(a).padStart(2, '0')}∩${String(b).padStart(2, '0')} + Ramificações` });
        }
      }
      ch4B.sort((x, y) => y.score - x.score);

      const maxLen4 = Math.max(ch4A.length, ch4B.length);
      const seen4 = new Set();
      for (let i = 0; i < maxLen4; i++) {
        if (i < ch4A.length && !seen4.has(ch4A[i].numeros.join('-'))) {
          seen4.add(ch4A[i].numeros.join('-'));
          ch4_PivotCore.push(ch4A[i]);
        }
        if (i < ch4B.length && !seen4.has(ch4B[i].numeros.join('-'))) {
          seen4.add(ch4B[i].numeros.join('-'));
          ch4_PivotCore.push(ch4B[i]);
        }
      }
    }

    // ─── CANAL 5: Fechamento da 2ª Camada (G5 + Pass2(G5) + G4) ─────────────────
    const ch5_G5G4 = [];
    {
      const g5Pass2 = {};
      for (const u of g5) {
        const p2Games = base20.slice(5).filter(g => g.numeros.includes(u));
        const others = [...new Set(p2Games.flatMap(g => g.numeros))].filter(x => !g5.includes(x) && lastSeen[x] >= 1);
        others.sort((x, y) => getLiftPure(u, y) - getLiftPure(u, x));
        g5Pass2[u] = others.slice(0, 2);
      }
      const g5Active = g5.filter(x => lastSeen[x] >= 3);
      const g4Active = g4.filter(x => lastSeen[x] >= 1 && lastSeen[x] <= 25);
      const seen5 = new Map();
      for (const trio of combinations(g5Active, 3)) {
        const [u, v, w] = trio;
        const partners = [...new Set([...(g5Pass2[u] || []), ...(g5Pass2[v] || []), ...(g5Pass2[w] || [])])];
        for (const p of partners) {
          for (const x4 of g4Active) {
            for (const yExtra of [...partners, ...g4Active]) {
              const six = [u, v, w, p, x4, yExtra].sort((x, y) => x - y);
              if (new Set(six).size !== 6 || hasConsecutive(six)) continue;
              const edges = [];
              for (let i = 0; i < 6; i++) for (let j = i + 1; j < 6; j++) edges.push(pairCounts[six[i]][six[j]] || 0);
              edges.sort((x, y) => x - y);
              const score = edges[0] * 100 + edges[1] * 20 + edges.reduce((s, c) => s + c, 0);
              const key = six.join('-');
              if (!seen5.has(key) || score > seen5.get(key).score) {
                seen5.set(key, { numeros: six, score, label: 'C5: Fechamento 2ª Camada G5 + G4' });
              }
            }
          }
        }
      }
      ch5_G5G4.push(...[...seen5.values()].sort((x, y) => y.score - x.score));
    }

    // Escalonamento dos 5 canais de fusão estrutural (3 de C1, 3 de C2, 1 de C3, 2 de C4, 1 de C5 por bloco de 10)
    const schedule = [
      ch1_G3Spine,
      ch1_G3Spine,
      ch1_G3Spine,
      ch2_DisjointBranches,
      ch2_DisjointBranches,
      ch2_DisjointBranches,
      ch3_Transversal,
      ch4_PivotCore,
      ch4_PivotCore,
      ch5_G5G4,
    ];

    const selected = [];
    const seenFinal = new Set();
    const targetTotal = (pageOffset + 1) * count;
    const recent15Draws = draws.slice(-15);

    const pickFromList = (list) => {
      for (const item of list) {
        const key = item.numeros.join('-');
        if (seenFinal.has(key)) continue;
        // Bloqueia combinações que já descarregaram 4+ dezenas nos últimos 15 concursos
        if (recent15Draws.some(d => item.numeros.filter(n => d.bolas.includes(n)).length >= 4)) continue;
        if (selected.some(s => s.numeros.filter(n => item.numeros.includes(n)).length >= 5)) continue;
        seenFinal.add(key);
        selected.push(item);
        return true;
      }
      return false;
    };

    for (let idx = 0; idx < targetTotal; idx++) {
      const ch = schedule[idx % schedule.length];
      if (!pickFromList(ch)) {
        // Fallback em rodízio pelos canais com candidatos restantes
        for (const altCh of [ch1_G3Spine, ch2_DisjointBranches, ch3_Transversal, ch4_PivotCore, ch5_G5G4]) {
          if (pickFromList(altCh)) break;
        }
      }
    }

    const pageSlice = selected.slice(pageOffset * count, (pageOffset + 1) * count);
    while (pageSlice.length < count && base20[pageSlice.length % base20.length]) {
      pageSlice.push({
        numeros: base20[pageSlice.length % base20.length].numeros,
        label: 'Fallback Vetor Paginado',
      });
    }

    for (const item of pageSlice) {
      const picked = item.numeros;
      let totalPairRatio = 0;
      let bestPair = { a: picked[0], b: picked[1], count: 0 };
      let newPairsCount = 0;
      const gamePairs = [];

      for (let i = 0; i < picked.length; i++) {
        for (let j = i + 1; j < picked.length; j++) {
          const pA = picked[i];
          const pB = picked[j];
          const key = `${pA}-${pB}`;
          gamePairs.push(key);
          if (!usedPairsInBatch.has(key)) newPairsCount++;
          const c = pairCounts[pA][pB] || 0;
          totalPairRatio += expectedPairCount > 0 ? c / expectedPairCount : 1;
          if (c > bestPair.count) {
            bestPair = { a: pA, b: pB, count: c };
          }
        }
      }
      gamePairs.forEach(key => usedPairsInBatch.add(key));

      const soma = picked.reduce((p, q) => p + q, 0);
      const pares = picked.filter(x => x % 2 === 0).length;
      const nAcima31 = picked.filter(x => x > 31).length;
      const premios = checkPastWins(picked, draws);
      const afinidadeMedia = Number((totalPairRatio / 15).toFixed(2));

      games.push({
        numeros: picked,
        soma,
        pares,
        afinidadeMedia,
        parMaisForte: `${String(bestPair.a).padStart(2, '0')} e ${String(bestPair.b).padStart(2, '0')} (${bestPair.count}x)`,
        paresIneditos: `${item.label} (${newPairsCount} pares novos)`,
        estrategia: `Modo C (${item.label})`,
        premios,
        custo: PRECO_UNITARIO,
        rateioPotencial: nAcima31 >= 3 ? 'Alto' : 'Normal',
      });
    }

    return games;
  }

  // No Modo A (Paginado) ou quando avançando páginas no Modo B, deslocamos o ponto de partida no vetor de sementes
  const seedStartOffset = (pageOffset * count * 2) % Math.max(allSeedPairs.length - count * 5, 1);

  for (let gIdx = 0; gIdx < count; gIdx++) {
    let bestClique = null;

    // Monta o vetor de sementes disponíveis para abrir este jogo
    const candidateSeeds = [];
    for (let idx = seedStartOffset; idx < allSeedPairs.length; idx++) {
      const sp = allSeedPairs[idx];
      const key = `${sp.a}-${sp.b}`;
      if (usedPairsInBatch.has(key)) continue;

      const penaltyA = Math.pow(1 + 0.25 * batchUsage[sp.a], 1.5);
      const penaltyB = Math.pow(1 + 0.25 * batchUsage[sp.b], 1.5);
      const effectiveSeedScore = (sp.score * baseWeights[sp.a] * baseWeights[sp.b]) / (penaltyA * penaltyB);

      candidateSeeds.push({ ...sp, effectiveSeedScore });
      if (candidateSeeds.length >= 60) break;
    }

    candidateSeeds.sort((x, y) => y.effectiveSeedScore - x.effectiveSeedScore);

    // Define quantas sementes do topo do vetor explorar:
    // Modo A testa as 8 melhores sementes em ordem determinística; Modo B sorteia entre o Top 3 do feixe
    const topSeedsToTry = candidateSeeds.slice(0, 8);
    if (mode === 'vetor_feixe') {
      const top3 = topSeedsToTry.slice(0, 3);
      shuffleWithRng(top3, rng);
      topSeedsToTry.splice(0, top3.length, ...top3);
    }

    const globalCandidateCliques = [];

    for (const seed of topSeedsToTry) {
      const partial2 = [seed.a, seed.b];

      // Vetor ordenado para a 3ª dezena
      const vec3 = [];
      for (let x = 1; x <= N; x++) {
        if (!canAddCandidate(x, partial2, true)) continue;
        vec3.push({ num: x, score: scoreCandidateWithSet(x, partial2) });
      }
      vec3.sort((a, b) => b.score - a.score);

      const branch3 = vec3.slice(0, 3);
      if (mode === 'vetor_feixe') shuffleWithRng(branch3, rng);

      for (const c3 of branch3) {
        const partial3 = [...partial2, c3.num];

        // Vetor ordenado para a 4ª dezena
        const vec4 = [];
        for (let x = 1; x <= N; x++) {
          if (!canAddCandidate(x, partial3, true)) continue;
          vec4.push({ num: x, score: scoreCandidateWithSet(x, partial3) });
        }
        vec4.sort((a, b) => b.score - a.score);

        const branch4 = vec4.slice(0, 3);
        if (mode === 'vetor_feixe') shuffleWithRng(branch4, rng);

        for (const c4 of branch4) {
          const partial4 = [...partial3, c4.num];
          const sum4 = partial4.reduce((a, b) => a + b, 0);

          // Vetor ordenado de candidatos finais para a 5ª e 6ª dezenas
          const vec56 = [];
          for (let x = 1; x <= N; x++) {
            if (!canAddCandidate(x, partial4, true)) continue;
            vec56.push({ num: x, score: scoreCandidateWithSet(x, partial4) });
          }
          vec56.sort((a, b) => b.score - a.score);

          const topPool = vec56.slice(0, 16);

          for (let i = 0; i < topPool.length; i++) {
            for (let j = i + 1; j < topPool.length; j++) {
              const n5 = topPool[i].num;
              const n6 = topPool[j].num;
              const totalSum = sum4 + n5 + n6;
              if (totalSum < somaMin || totalSum > somaMax) continue;

              const key56 = n5 < n6 ? `${n5}-${n6}` : `${n6}-${n5}`;
              if (usedPairsInBatch.has(key56)) continue;

              const fullSet = [...partial4, n5, n6].sort((a, b) => a - b);
              if (hasConsecutive(fullSet)) continue;

              const nAcima31 = fullSet.filter(n => n > 31).length;
              if (nAcima31 < 3) continue;

              const sig = fullSet.join(',');
              if (games.some(g => g.numeros.join(',') === sig)) continue;

              // Calcula a pontuação real dos 15 pares C(6,2) do conjunto completo + penalidade leve de repetição no lote
              let sum15Synergy = 0;
              for (let p1 = 0; p1 < 6; p1++) {
                for (let p2 = p1 + 1; p2 < 6; p2++) {
                  sum15Synergy += pairSynergy[fullSet[p1]][fullSet[p2]];
                }
              }
              let usagePen = 1;
              for (const num of fullSet) {
                usagePen *= (1 + 0.15 * batchUsage[num]);
              }
              const cliqueScore = (sum15Synergy / 15) / Math.pow(usagePen, 0.25);
              globalCandidateCliques.push({ numeros: fullSet, cliqueScore, rawSynergy: sum15Synergy / 15, sig });
            }
          }
        }
      }
      if (mode === 'vetor_paginado' && globalCandidateCliques.length >= 25) break;
      if (mode === 'vetor_feixe' && globalCandidateCliques.length >= 35) break;
    }

    if (globalCandidateCliques.length > 0) {
      // Remove assinaturas duplicadas mantendo o maior score
      const uniqueMap = new Map();
      for (const cl of globalCandidateCliques) {
        if (!uniqueMap.has(cl.sig) || uniqueMap.get(cl.sig).cliqueScore < cl.cliqueScore) {
          uniqueMap.set(cl.sig, cl);
        }
      }
      const uniqueCliques = Array.from(uniqueMap.values()).sort((a, b) => b.cliqueScore - a.cliqueScore);

      if (mode === 'vetor_paginado') {
        bestClique = uniqueCliques[0];
      } else {
        const eliteTop = uniqueCliques.slice(0, Math.min(3, uniqueCliques.length));
        bestClique = eliteTop[Math.floor(rng() * eliteTop.length)];
      }
    }

    // Fallback de segurança caso a janela de soma escolhida pelo usuário seja extremamente estreita
    if (!bestClique) {
      for (const seed of candidateSeeds) {
        const picked = [seed.a, seed.b];
        const rem = [];
        for (let x = 1; x <= N; x++) {
          if (x !== seed.a && x !== seed.b) {
            rem.push({ num: x, score: scoreCandidateWithSet(x, picked) });
          }
        }
        rem.sort((a, b) => b.score - a.score);
        for (const r of rem) {
          if (picked.length < 6 && !hasConsecutive([...picked, r.num].sort((a, b) => a - b))) {
            picked.push(r.num);
          }
        }
        picked.sort((a, b) => a - b);
        const s = picked.reduce((a, b) => a + b, 0);
        if (s >= somaMin && s <= somaMax) {
          bestClique = { numeros: picked };
          break;
        }
      }
      if (!bestClique) {
        bestClique = { numeros: [5, 10, 27, 33, 42, 53] };
      }
    }

    const picked = bestClique.numeros;
    const gamePairs = [];
    let overlapPairs = 0;
    let totalPairRatio = 0;
    let bestPair = { a: picked[0], b: picked[1], count: 0 };

    for (let i = 0; i < picked.length; i++) {
      for (let j = i + 1; j < picked.length; j++) {
        const a = picked[i];
        const b = picked[j];
        const key = `${a}-${b}`;
        gamePairs.push(key);
        if (usedPairsInBatch.has(key)) overlapPairs++;

        const c = pairCounts[a][b] || 0;
        totalPairRatio += expectedPairCount > 0 ? c / expectedPairCount : 1;
        if (c > bestPair.count) {
          bestPair = { a, b, count: c };
        }
      }
    }

    const soma = picked.reduce((a, b) => a + b, 0);
    const pares = picked.filter(b => b % 2 === 0).length;
    const nAcima31 = picked.filter(n => n > 31).length;
    const premios = checkPastWins(picked, draws);
    const afinidadeMedia = (totalPairRatio / 15).toFixed(2);
    const paresIneditos = 15 - overlapPairs;

    games.push({
      numeros: picked,
      soma,
      pares,
      afinidadeMedia: Number(afinidadeMedia),
      parMaisForte: `${String(bestPair.a).padStart(2, '0')} e ${String(bestPair.b).padStart(2, '0')} (${bestPair.count}x)`,
      paresIneditos: `${paresIneditos}/15`,
      estrategia: mode === 'vetor_paginado'
        ? `Modo A: Vetor Determinístico (Pág. ${pageOffset + 1})`
        : `Modo B: Vetor Feixe Top-3 (Feixe ${pageOffset + 1})`,
      premios,
      custo: PRECO_UNITARIO,
      rateioPotencial: nAcima31 >= 4 ? 'Alto' : 'Normal',
    });

    picked.forEach(n => batchUsage[n]++);
    gamePairs.forEach(k => usedPairsInBatch.add(k));
  }

  return games;
}

// ─── Seletor Regressivo Walk-Forward (Descarta 1 e 2 acertos e busca o próximo ≥ 3) ───
const regressiveCache = new Map();

function getRegressiveProfile(draws, { somaMin = 130, somaMax = 230, windowSize = 10 } = {}) {
  const cacheKey = `${draws.length}-${somaMin}-${somaMax}-${windowSize}`;
  if (regressiveCache.has(cacheKey)) return regressiveCache.get(cacheKey);

  const modesConfig = [
    { mode: 'vetor_paginado', label: 'Modo A', pagesToScan: 6 },
    { mode: 'vetor_feixe', label: 'Modo B', pagesToScan: 6 },
    { mode: 'vetor_nucleo', label: 'Modo C', pagesToScan: 4 },
  ];

  const profileByMode = {};
  const startIdx = Math.max(20, draws.length - windowSize);

  for (const cfg of modesConfig) {
    const totalSlots = cfg.pagesToScan * 10;
    const slotStats = Array.from({ length: totalSlots }, (_, sIdx) => ({
      slot: sIdx + 1,
      page: Math.floor(sIdx / 10),
      posInPage: (sIdx % 10) + 1,
      maxHits: 0,
      hits3Plus: 0,
      hits4Plus: 0,
      sumHits: 0,
      bestDraw: '-',
      history: [],
    }));

    for (let idx = startIdx; idx < draws.length; idx++) {
      const target = draws[idx];
      const train = draws.slice(0, idx);
      const cands = [];
      for (let p = 0; p < cfg.pagesToScan; p++) {
        const pageGames = generateGames(train, {
          somaMin,
          somaMax,
          count: 10,
          mode: cfg.mode,
          pageOffset: p,
          _beamSeed: 1000 + p * 37,
          _raw: true,
        });
        cands.push(...pageGames);
      }

      for (let sIdx = 0; sIdx < Math.min(totalSlots, cands.length); sIdx++) {
        const g = cands[sIdx];
        const h = g.numeros.filter(n => target.bolas.includes(n)).length;
        const st = slotStats[sIdx];
        st.sumHits += h;
        st.history.push({ concurso: target.concurso, hits: h });
        if (h >= 3) st.hits3Plus++;
        if (h >= 4) st.hits4Plus++;
        if (h > st.maxHits) {
          st.maxHits = h;
          st.bestDraw = `#${target.concurso} (${h} acertos)`;
        }
      }
    }

    // Gera os candidatos atuais (treinados até o último concurso T) para cada slot
    const currentCands = [];
    for (let p = 0; p < cfg.pagesToScan; p++) {
      const pageGames = generateGames(draws, {
        somaMin,
        somaMax,
        count: 10,
        mode: cfg.mode,
        pageOffset: p,
        _beamSeed: 1000 + p * 37,
        _raw: true,
      });
      currentCands.push(...pageGames);
    }

    const recent15Draws = draws.slice(-15);
    const recentResonanceDraws = draws.slice(-180, -15); // Exclui os últimos 15 sorteios para evitar exaustão de ciclo

    const enrichedCandidates = currentCands.map((g, sIdx) => {
      const st = slotStats[sIdx];
      const exhaustedInLast15 = recent15Draws.some(d => g.numeros.filter(n => d.bolas.includes(n)).length >= 4);
      let maxResonance = 0;
      let bestResDraw = '-';
      for (const d of recentResonanceDraws) {
        const rh = g.numeros.filter(n => d.bolas.includes(n)).length;
        if (rh > maxResonance) {
          maxResonance = rh;
          bestResDraw = `#${d.concurso} (${rh} acertos)`;
        }
      }

      // Critério Regressivo:
      // 1) O slot no Walk-Forward de 10 sorteios atingiu >= 3 acertos (Terno/Quadra/Quina/Sena)
      // 2) E o jogo atual não exauriu 4+ dezenas em nenhum dos últimos 15 concursos
      const cycleExhausted = exhaustedInLast15;
      const wfApproved = st.maxHits >= 3 && !cycleExhausted;
      const resonanceApproved = maxResonance >= 3 && !cycleExhausted;

      return {
        ...g,
        modeCode: cfg.mode,
        modeLabel: cfg.label,
        slotOriginal: st.slot,
        pageOriginal: st.page + 1,
        posInPage: st.posInPage,
        wfMaxHits: st.maxHits,
        wfHits3Plus: st.hits3Plus,
        wfHits4Plus: st.hits4Plus,
        wfBestDraw: st.bestDraw,
        maxResonance,
        bestResDraw,
        cycleExhausted,
        wfApproved,
        resonanceApproved,
        regressiveScore:
          (wfApproved ? 1000 : 0) +
          st.maxHits * 120 +
          st.hits4Plus * 80 +
          st.hits3Plus * 40 +
          maxResonance * 15 +
          g.afinidadeMedia * 10,
      };
    });

    const approvedWFCount = enrichedCandidates.filter(c => c.wfApproved).length;
    const discardedLowHitsCount = enrichedCandidates.filter(c => !c.wfApproved).length;

    profileByMode[cfg.mode] = {
      mode: cfg.mode,
      label: cfg.label,
      totalEvaluated: enrichedCandidates.length,
      approvedWFCount,
      discardedLowHitsCount,
      candidates: enrichedCandidates,
    };
  }

  const result = {
    windowSize,
    fromConcurso: draws[startIdx]?.concurso,
    toConcurso: draws[draws.length - 1]?.concurso,
    modes: profileByMode,
  };
  regressiveCache.set(cacheKey, result);
  return result;
}

function selectRegressiveGames(draws, params = {}) {
  const somaMin = Number(params.somaMin) || 130;
  const somaMax = Number(params.somaMax) || 230;
  const count = Math.min(Math.max(Number(params.count) || 10, 1), 30);
  const pageOffset = Math.max(Number(params.pageOffset) || 0, 0);
  const mode = params.mode || 'vetor_paginado';

  const profile = getRegressiveProfile(draws, { somaMin, somaMax, windowSize: 10 });

  let pool = [];
  if (mode === 'regressivo_elite') {
    // Combina os melhores candidatos aprovados dos Modos C, A e B ordenados por performance regressiva
    const allApproved = [
      ...profile.modes.vetor_nucleo.candidates.filter(c => c.wfApproved),
      ...profile.modes.vetor_paginado.candidates.filter(c => c.wfApproved),
      ...profile.modes.vetor_feixe.candidates.filter(c => c.wfApproved),
    ];
    allApproved.sort((a, b) => b.regressiveScore - a.regressiveScore);
    const fallbackRes = [
      ...profile.modes.vetor_nucleo.candidates.filter(c => !c.wfApproved && c.resonanceApproved),
      ...profile.modes.vetor_paginado.candidates.filter(c => !c.wfApproved && c.resonanceApproved),
      ...profile.modes.vetor_feixe.candidates.filter(c => !c.wfApproved && c.resonanceApproved),
    ].sort((a, b) => b.regressiveScore - a.regressiveScore);
    pool = [...allApproved, ...fallbackRes];
  } else {
    const mProf = profile.modes[mode] || profile.modes.vetor_paginado;
    // Mantém a ordem natural dos slots que atingiram >= 3 acertos no Walk-Forward,
    // descartando os slots de 1 e 2 acertos e buscando o próximo no gerador!
    const wfPass = mProf.candidates.filter(c => c.wfApproved);
    const resPass = mProf.candidates.filter(c => !c.wfApproved && c.resonanceApproved);
    const rest = mProf.candidates.filter(c => !c.wfApproved && !c.resonanceApproved);
    pool = [...wfPass, ...resPass, ...rest];
  }

  const selected = [];
  const seenSigs = new Set();
  const targetTotal = (pageOffset + 1) * count;

  for (const cand of pool) {
    const sig = cand.numeros.join('-');
    if (seenSigs.has(sig)) continue;
    // Evita clones com 5 dezenas idênticas quando puxa de páginas seguintes
    if (selected.some(s => s.numeros.filter(n => cand.numeros.includes(n)).length >= 5)) continue;
    seenSigs.add(sig);
    selected.push(cand);
    if (selected.length >= targetTotal) break;
  }

  const totalDiscarded = mode === 'regressivo_elite'
    ? (profile.modes.vetor_paginado.discardedLowHitsCount +
       profile.modes.vetor_feixe.discardedLowHitsCount +
       profile.modes.vetor_nucleo.discardedLowHitsCount)
    : (profile.modes[mode]?.discardedLowHitsCount || 0);

  const slice = selected.slice(pageOffset * count, (pageOffset + 1) * count);
  return slice.map((c) => {
    const badge = c.wfApproved
      ? `WF ≥3: ${c.wfBestDraw}`
      : `Ressonância ≥3: ${c.bestResDraw}`;
    return {
      ...c,
      descartados1ou2: totalDiscarded,
      paresIneditos: `${badge} | Slot #${c.slotOriginal}`,
      estrategia: `${c.estrategia} • [Aprovado ≥3: ${c.wfApproved ? c.wfBestDraw : c.bestResDraw}]`,
    };
  });
}

// ─── API Routes ───────────────────────────────────────────────────────────────
let loadingPromise = null;
let lastLoaded = 0;

async function getDraws() {
  if (loadingPromise && Date.now() - lastLoaded < 60000) return loadingPromise;
  lastLoaded = Date.now();
  loadingPromise = loadDraws();
  return loadingPromise;
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

app.get('/api/sets', async (req, res) => {
  try {
    const draws = await getDraws();
    const { pairCounts, ...setsData } = calcSetTheory(draws);
    res.json(setsData);
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

// ─── Authentication Endpoints ────────────────────────────────────────────────
app.post('/api/auth/register', async (req, res) => {
  const { username, password, name, email } = req.body;
  if (!username || !password || !name || !email) {
    return res.status(400).json({ error: 'Todos os campos (nome, email, usuário e senha) são obrigatórios.' });
  }

  const users = loadUsers();
  if (users.find(u => u.username.toLowerCase() === username.toLowerCase())) {
    return res.status(400).json({ error: 'Este nome de usuário já existe.' });
  }
  if (users.find(u => u.email.toLowerCase() === email.toLowerCase())) {
    return res.status(400).json({ error: 'Este email já está cadastrado.' });
  }

  try {
    const hashedPassword = await bcrypt.hash(password, 10);
    const newUser = {
      id: Date.now().toString() + Math.random().toString().slice(2, 6),
      username,
      name,
      email: email.toLowerCase(),
      passwordHash: hashedPassword
    };
    users.push(newUser);
    saveUsers(users);
    res.status(201).json({ message: 'Usuário registrado com sucesso!' });
  } catch (e) {
    res.status(500).json({ error: 'Erro ao registrar usuário.' });
  }
});

app.post('/api/auth/login', async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Usuário e senha são obrigatórios.' });
  }

  const users = loadUsers();
  const user = users.find(u => u.username.toLowerCase() === username.toLowerCase());
  if (!user) {
    return res.status(400).json({ error: 'Usuário ou senha inválidos.' });
  }

  try {
    const valid = await bcrypt.compare(password, user.passwordHash);
    if (!valid) {
      return res.status(400).json({ error: 'Usuário ou senha inválidos.' });
    }

    const token = jwt.sign(
      { id: user.id, username: user.username },
      JWT_SECRET,
      { expiresIn: '7d' }
    );
    res.json({ token, username: user.username });
  } catch (e) {
    res.status(500).json({ error: 'Erro ao processar login.' });
  }
});

app.post('/api/auth/google', async (req, res) => {
  const { credential } = req.body;
  if (!credential) {
    return res.status(400).json({ error: 'Credencial do Google ausente.' });
  }

  try {
    // Valida o token com o Google chamando o endpoint de validação
    const googleRes = await fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${credential}`);
    if (!googleRes.ok) {
      return res.status(400).json({ error: 'Token do Google inválido ou expirado.' });
    }

    const payload = await googleRes.json();
    const { email, name, sub } = payload;

    if (!email) {
      return res.status(400).json({ error: 'Não foi possível obter o email do Google.' });
    }

    const users = loadUsers();
    let user = users.find(u => u.email.toLowerCase() === email.toLowerCase());

    // Se o usuário não existe, cria um novo usuário automaticamente
    if (!user) {
      let baseUsername = email.split('@')[0].replace(/[^a-zA-Z0-9]/g, '');
      let username = baseUsername;
      let counter = 1;

      while (users.find(u => u.username.toLowerCase() === username.toLowerCase())) {
        username = baseUsername + counter;
        counter++;
      }

      user = {
        id: Date.now().toString() + Math.random().toString().slice(2, 6),
        username,
        name: name || username,
        email: email.toLowerCase(),
        passwordHash: null, // Usuários do Google não têm senha local
        googleId: sub
      };
      users.push(user);
      saveUsers(users);
    }

    // Gera o token JWT local
    const token = jwt.sign(
      { id: user.id, username: user.username },
      JWT_SECRET,
      { expiresIn: '7d' }
    );

    res.json({ token, username: user.username });
  } catch (e) {
    res.status(500).json({ error: 'Erro ao autenticar com o Google: ' + e.message });
  }
});

app.get('/api/auth/google/client-id', (req, res) => {
  res.json({ clientId: process.env.GOOGLE_CLIENT_ID || '1008719970978-gp2ddch93g5f.apps.googleusercontent.com' });
});

// ─── User CRUD Endpoints (Protected by JWT) ───────────────────────────────────
app.get('/api/users', authenticateToken, (req, res) => {
  try {
    const users = loadUsers();
    const safeUsers = users.map(({ passwordHash, ...rest }) => rest);
    res.json(safeUsers);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/users', authenticateToken, async (req, res) => {
  const { username, password, name, email } = req.body;
  if (!username || !password || !name || !email) {
    return res.status(400).json({ error: 'Todos os campos (nome, email, usuário e senha) são obrigatórios.' });
  }

  const users = loadUsers();
  if (users.find(u => u.username.toLowerCase() === username.toLowerCase())) {
    return res.status(400).json({ error: 'Este nome de usuário já existe.' });
  }
  if (users.find(u => u.email.toLowerCase() === email.toLowerCase())) {
    return res.status(400).json({ error: 'Este email já está cadastrado.' });
  }

  try {
    const hashedPassword = await bcrypt.hash(password, 10);
    const newUser = {
      id: Date.now().toString() + Math.random().toString().slice(2, 6),
      username,
      name,
      email: email.toLowerCase(),
      passwordHash: hashedPassword
    };
    users.push(newUser);
    saveUsers(users);

    const { passwordHash, ...safeUser } = newUser;
    res.status(201).json(safeUser);
  } catch (e) {
    res.status(500).json({ error: 'Erro ao criar usuário.' });
  }
});

app.put('/api/users/:id', authenticateToken, async (req, res) => {
  const { id } = req.params;
  const { username, password, name, email } = req.body;

  if (!username || !name || !email) {
    return res.status(400).json({ error: 'Nome, email e usuário são obrigatórios.' });
  }

  const users = loadUsers();
  const userIndex = users.findIndex(u => u.id === id);
  if (userIndex === -1) {
    return res.status(404).json({ error: 'Usuário não encontrado.' });
  }

  if (users.find(u => u.id !== id && u.username.toLowerCase() === username.toLowerCase())) {
    return res.status(400).json({ error: 'Este nome de usuário já está em uso.' });
  }
  if (users.find(u => u.id !== id && u.email.toLowerCase() === email.toLowerCase())) {
    return res.status(400).json({ error: 'Este email já está em uso.' });
  }

  try {
    const user = users[userIndex];
    user.username = username;
    user.name = name;
    user.email = email.toLowerCase();

    if (password) {
      user.passwordHash = await bcrypt.hash(password, 10);
    }

    users[userIndex] = user;
    saveUsers(users);

    const { passwordHash, ...safeUser } = user;
    res.json(safeUser);
  } catch (e) {
    res.status(500).json({ error: 'Erro ao atualizar usuário.' });
  }
});

app.delete('/api/users/:id', authenticateToken, (req, res) => {
  const { id } = req.params;

  if (req.user.id === id) {
    return res.status(400).json({ error: 'Você não pode excluir sua própria conta enquanto estiver logado.' });
  }

  const users = loadUsers();
  const userIndex = users.findIndex(u => u.id === id);
  if (userIndex === -1) {
    return res.status(404).json({ error: 'Usuário não encontrado.' });
  }

  users.splice(userIndex, 1);
  saveUsers(users);
  res.json({ success: true, message: 'Usuário excluído com sucesso.' });
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

app.get('/api/regression', async (req, res) => {
  try {
    const draws = await getDraws();
    const somaMin = Number(req.query.somaMin) || 130;
    const somaMax = Number(req.query.somaMax) || 230;
    const profile = getRegressiveProfile(draws, { somaMin, somaMax, windowSize: 10 });
    res.json({
      windowSize: profile.windowSize,
      fromConcurso: profile.fromConcurso,
      toConcurso: profile.toConcurso,
      summary: {
        vetor_paginado: {
          label: 'Modo A (Vetor Determinístico)',
          totalEvaluated: profile.modes.vetor_paginado.totalEvaluated,
          approvedWFCount: profile.modes.vetor_paginado.approvedWFCount,
          discardedLowHitsCount: profile.modes.vetor_paginado.discardedLowHitsCount,
          topApproved: profile.modes.vetor_paginado.candidates.filter(c => c.wfApproved).slice(0, 6),
        },
        vetor_feixe: {
          label: 'Modo B (Feixe Top-3)',
          totalEvaluated: profile.modes.vetor_feixe.totalEvaluated,
          approvedWFCount: profile.modes.vetor_feixe.approvedWFCount,
          discardedLowHitsCount: profile.modes.vetor_feixe.discardedLowHitsCount,
          topApproved: profile.modes.vetor_feixe.candidates.filter(c => c.wfApproved).slice(0, 6),
        },
        vetor_nucleo: {
          label: 'Modo C (5 Canais de Fusão)',
          totalEvaluated: profile.modes.vetor_nucleo.totalEvaluated,
          approvedWFCount: profile.modes.vetor_nucleo.approvedWFCount,
          discardedLowHitsCount: profile.modes.vetor_nucleo.discardedLowHitsCount,
          topApproved: profile.modes.vetor_nucleo.candidates.filter(c => c.wfApproved).slice(0, 10),
        },
      },
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/reload', async (req, res) => {
  cachedDraws = null;
  regressiveCache.clear();
  try {
    const draws = await getDraws();
    res.json({ ok: true, total: draws.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('*', (_, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`🎲 Lottery Dashboard rodando em http://localhost:${PORT}`);
    if (!fs.existsSync(XLSX_PATH)) {
      console.warn(`⚠️  Arquivo Excel não encontrado. Coloque o arquivo em: ${XLSX_PATH}`);
    }
  });
}

module.exports = {
  app,
  getDraws,
  generateGames,
  getRegressiveProfile,
  selectRegressiveGames,
};
