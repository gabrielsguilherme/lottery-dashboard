const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
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
  const {
    somaMin = 130, somaMax = 230,
    minPares = 2, maxPares = 4,
    count = 10,
  } = params;

  const PRECO_UNITARIO = 5.00;
  const games = [];
  let attempts = 0;

  // Calcula a frequência e isola os números medianos (nem muito nem pouco sorteados)
  const freq = calcFrequency(draws);
  const medios = freq.filter(f => f.categoria === 'Médio').map(f => f.numero);

  // Fallback caso a categoria médio não retorne o suficiente (improvável)
  const pool = medios.length >= 10 ? medios : Array.from({length: 60}, (_, i) => i + 1);

  while (games.length < count && attempts < 500000) {
    attempts++;
    
    // Sorteia 6 números EXCLUSIVAMENTE do pool de números medianos
    const pickedSet = new Set();
    while (pickedSet.size < 6) {
      pickedSet.add(pool[Math.floor(Math.random() * pool.length)]);
    }
    const picked = Array.from(pickedSet).sort((a, b) => a - b);

    // Cálculos para os filtros
    const soma = picked.reduce((a, b) => a + b, 0);
    const pares = picked.filter(b => b % 2 === 0).length;
    const nAcima31 = picked.filter(n => n > 31).length;

    // Filtros de validação
    if (
      soma >= somaMin && 
      soma <= somaMax && 
      nAcima31 >= 3 && 
      !hasConsecutive(picked)
    ) {
      const premios = checkPastWins(picked, draws);
      
      games.push({ 
        numeros: picked, 
        soma, 
        pares, 
        estrategia: 'Frequência Média', 
        premios,
        custo: PRECO_UNITARIO,
        rateioPotencial: nAcima31 >= 4 ? 'Alto' : 'Normal'
      });
    }
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

app.post('/api/generate', authenticateToken, async (req, res) => {
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
