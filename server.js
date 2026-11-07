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
function calcSetTheory(draws) {
  const N = 60;
  const k = 6;
  const T = draws.length;
  const expectedPairProb = (k * (k - 1)) / (N * (N - 1)); // 30 / 3540 = 1 / 118
  const expectedPairCount = T * expectedPairProb;
  const expectedTrioProb = (k * (k - 1) * (k - 2)) / (N * (N - 1) * (N - 2)); // 120 / 205320 = 1 / 1711
  const expectedTrioCount = T * expectedTrioProb;

  const numCounts = new Array(N + 1).fill(0);
  const pairCounts = Array.from({ length: N + 1 }, () => new Array(N + 1).fill(0));
  const trioMap = new Map();

  for (const d of draws) {
    const b = d.bolas;
    for (let i = 0; i < b.length; i++) {
      numCounts[b[i]]++;
      for (let j = i + 1; j < b.length; j++) {
        pairCounts[b[i]][b[j]]++;
        pairCounts[b[j]][b[i]]++;
        for (let m = j + 1; m < b.length; m++) {
          const key = `${b[i]}-${b[j]}-${b[m]}`;
          trioMap.set(key, (trioMap.get(key) || 0) + 1);
        }
      }
    }
  }

  const allPairs = [];
  const partnersByNumber = {};
  for (let i = 1; i <= N; i++) partnersByNumber[i] = [];

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

  for (let i = 1; i <= N; i++) {
    partnersByNumber[i].sort((x, y) => y.count - x.count || y.lift - x.lift);
    partnersByNumber[i] = partnersByNumber[i].slice(0, 10);
  }

  const topPairs = [...allPairs]
    .sort((x, y) => y.count - x.count || y.lift - x.lift)
    .slice(0, 20);

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

function generateGames(draws, params = {}) {
  const somaMin = Number(params.somaMin) || 130;
  const somaMax = Number(params.somaMax) || 230;
  const count = Math.min(Math.max(Number(params.count) || 10, 1), 30);
  const mode = params.mode === 'vetor_feixe' ? 'vetor_feixe' : 'vetor_paginado'; // 'vetor_paginado' (Modo A) ou 'vetor_feixe' (Modo B)
  const pageOffset = Math.max(Number(params.pageOffset) || 0, 0);

  const PRECO_UNITARIO = 5.00;
  const N = 60;
  const T = draws.length || 1;

  // 1. Frequência individual P(A) + Matriz de Coocorrência P(A ∩ B) e Condicional P(B | A)
  const freqList = calcFrequency(draws);
  const { pairCounts, expectedPairCount } = calcSetTheory(draws);
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

  // No Modo A (Paginado), deslocamos o ponto de partida no vetor de sementes conforme a página
  const seedStartOffset = mode === 'vetor_paginado' ? (pageOffset * count * 2) % Math.max(allSeedPairs.length - count * 5, 1) : 0;

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
      shuffle(top3);
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
      if (mode === 'vetor_feixe') shuffle(branch3);

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
        if (mode === 'vetor_feixe') shuffle(branch4);

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
        bestClique = eliteTop[Math.floor(Math.random() * eliteTop.length)];
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
        : 'Modo B: Vetor Feixe Top-3',
      premios,
      custo: PRECO_UNITARIO,
      rateioPotencial: nAcima31 >= 4 ? 'Alto' : 'Normal',
    });

    picked.forEach(n => batchUsage[n]++);
    gamePairs.forEach(k => usedPairsInBatch.add(k));
  }

  return games;
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
