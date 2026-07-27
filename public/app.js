const state = { stats: null, draws: [], frequency: [], charts: {}, token: null, username: null, isLoggedIn: false };

function updateAuthState() {
  const token = localStorage.getItem('token');
  const username = localStorage.getItem('username');

  if (token && username) {
    state.token = token;
    state.username = username;
    state.isLoggedIn = true;
    document.getElementById('authContainer').style.display = 'none';
    document.getElementById('generatorContainer').style.display = 'block';
    document.getElementById('userGreeting').innerText = `Olá, ${username}!`;
  } else {
    state.token = null;
    state.username = null;
    state.isLoggedIn = false;
    document.getElementById('authContainer').style.display = 'block';
    document.getElementById('generatorContainer').style.display = 'none';
  }
}

function setupAuthListeners() {
  document.getElementById('toRegisterBtn').addEventListener('click', () => {
    document.getElementById('loginForm').style.display = 'none';
    document.getElementById('registerForm').style.display = 'block';
    document.getElementById('authTitle').innerText = 'Crie sua conta';
    document.getElementById('authHint').innerText = 'Escolha um usuário e senha para cadastrar-se no painel.';
  });

  document.getElementById('toLoginBtn').addEventListener('click', () => {
    document.getElementById('registerForm').style.display = 'none';
    document.getElementById('loginForm').style.display = 'block';
    document.getElementById('authTitle').innerText = 'Acesse sua conta para gerar jogos';
    document.getElementById('authHint').innerText = 'É necessário fazer login para ter acesso ao motor estatístico de geração de apostas.';
  });

  document.getElementById('loginBtn').addEventListener('click', async () => {
    const username = document.getElementById('loginUsername').value.trim();
    const password = document.getElementById('loginPassword').value;

    if (!username || !password) {
      alert('Por favor, digite o usuário e a senha.');
      return;
    }

    try {
      const data = await api('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });

      localStorage.setItem('token', data.token);
      localStorage.setItem('username', data.username);
      updateAuthState();
      await generateGames();
    } catch (e) {
      alert(e.message || 'Erro ao fazer login.');
    }
  });

  document.getElementById('registerBtn').addEventListener('click', async () => {
    const username = document.getElementById('registerUsername').value.trim();
    const password = document.getElementById('registerPassword').value;

    if (!username || !password) {
      alert('Por favor, defina um usuário e uma senha.');
      return;
    }

    try {
      const data = await api('/api/auth/register', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });

      alert(data.message || 'Usuário cadastrado com sucesso! Faça o login agora.');
      document.getElementById('toLoginBtn').click();
    } catch (e) {
      alert(e.message || 'Erro ao realizar cadastro.');
    }
  });

  document.getElementById('logoutBtn').addEventListener('click', () => {
    localStorage.removeItem('token');
    localStorage.removeItem('username');
    updateAuthState();
  });
}

async function api(url, options) {
  const res = await fetch(url, options);
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

function setTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`panel-${tab.dataset.tab}`).classList.add('active');
    });
  });
}

function renderStats(stats) {
  const grid = document.getElementById('statsGrid');
  grid.innerHTML = `
    <div class="card"><div class="stat-label">Total de sorteios</div><div class="stat-value">${stats.total}</div><div class="stat-extra">Base lida da planilha</div></div>
    <div class="card"><div class="stat-label">Último concurso</div><div class="stat-value">${stats.latest.concurso}</div><div class="stat-extra">${stats.latest.data}</div></div>
    <div class="card"><div class="stat-label">Mais sorteado</div><div class="stat-value">${stats.mostFreq.numero}</div><div class="stat-extra">${stats.mostFreq.frequencia} vezes</div></div>
    <div class="card"><div class="stat-label">Menos sorteado</div><div class="stat-value">${stats.leastFreq.numero}</div><div class="stat-extra">${stats.leastFreq.frequencia} vezes</div></div>
    <div class="card"><div class="stat-label">Maior atraso</div><div class="stat-value">${stats.mostDelayed.numero}</div><div class="stat-extra">${stats.mostDelayed.atraso} concursos</div></div>
  `;
  document.getElementById('latestSummary').innerHTML = `
    <p class="muted">Último resultado carregado:</p>
    <h3>Concurso ${stats.latest.concurso}</h3>
    <p>${stats.latest.data}</p>
    <div class="balls">${stats.latest.bolas.map(n => `<span class="ball">${n}</span>`).join('')}</div>
  `;
}

function renderDraws(draws) {
  document.getElementById('resultsTable').innerHTML = draws.map(d => `
    <tr>
      <td>#${d.concurso}</td>
      <td>${d.data}</td>
      <td><div class="balls">${d.bolas.map(n => `<span class="ball">${n}</span>`).join('')}</div></td>
    </tr>
  `).join('');
}

function categoriaTag(cat) {
  const cls = cat === 'Quente' ? 'hot' : cat === 'Frio' ? 'cold' : 'mid';
  return `<span class="tag ${cls}">${cat}</span>`;
}

function renderFrequency(freq) {
  document.getElementById('frequencyTable').innerHTML = freq.map(f => `
    <tr>
      <td>${f.numero}</td><td>${f.frequencia}</td>
      <td>${(f.taxa * 100).toFixed(2)}%</td>
      <td>${categoriaTag(f.categoria)}</td>
      <td>${f.atraso}</td>
      <td>${f.tendencia.toFixed(1)}%</td>
    </tr>
  `).join('');

  destroyChart('freqChart');
  state.charts.freqChart = new Chart(document.getElementById('freqChart'), {
    type: 'bar',
    data: { labels: freq.map(f => f.numero), datasets: [{ label: 'Frequência', data: freq.map(f => f.frequencia), backgroundColor: freq.map(f => f.categoria === 'Quente' ? '#209869' : f.categoria === 'Frio' ? '#ef4444' : '#005ca9') }] },
    options: { responsive:true, plugins:{legend:{labels:{color:'#475569'}}}, scales:{x:{ticks:{color:'#64748b'}},y:{ticks:{color:'#64748b'}}} }
  });

  const topDelay = [...freq].sort((a,b) => b.atraso - a.atraso).slice(0, 10);
  destroyChart('delayChart');
  state.charts.delayChart = new Chart(document.getElementById('delayChart'), {
    type: 'bar',
    data: { labels: topDelay.map(f => f.numero), datasets: [{ label: 'Atraso', data: topDelay.map(f => f.atraso), backgroundColor: '#f29100' }] },
    options: { indexAxis:'y', responsive:true, plugins:{legend:{labels:{color:'#475569'}}}, scales:{x:{ticks:{color:'#64748b'}},y:{ticks:{color:'#64748b'}}} }
  });
}

function renderStatsCharts(stats) {
  destroyChart('sumChart');
  state.charts.sumChart = new Chart(document.getElementById('sumChart'), {
    type: 'bar',
    data: { labels: Object.keys(stats.somaDist), datasets: [{ label: 'Sorteios', data: Object.values(stats.somaDist), backgroundColor: '#005ca9' }] },
    options: { plugins:{legend:{labels:{color:'#475569'}}}, scales:{x:{ticks:{color:'#64748b'}},y:{ticks:{color:'#64748b'}}} }
  });
  destroyChart('parityChart');
  state.charts.parityChart = new Chart(document.getElementById('parityChart'), {
    type: 'doughnut',
    data: { labels: Object.keys(stats.pariDist).map(k => `${k} pares`), datasets: [{ data: Object.values(stats.pariDist), backgroundColor: ['#ef4444','#f97316','#f29100','#209869','#14b8a6','#3b82f6','#8b5cf6'] }] },
    options: { plugins:{legend:{labels:{color:'#475569'}}} }
  });
}

function destroyChart(name) { if (state.charts[name]) state.charts[name].destroy(); }

async function generateGames() {
  if (!state.isLoggedIn) return;

  const body = {
    somaMin: Number(document.getElementById('somaMin').value),
    somaMax: Number(document.getElementById('somaMax').value),
    count: Number(document.getElementById('countGames').value),
  };

  try {
    const games = await api('/api/generate', { 
      method: 'POST', 
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${state.token}`
      }, 
      body: JSON.stringify(body) 
    });

    document.getElementById('gamesTable').innerHTML = games.map((g, i) => `
      <tr>
        <td>Jogo ${i + 1}</td>
        <td><div class="balls">${g.numeros.map(n => `<span class="ball">${n}</span>`).join('')}</div></td>
        <td>${g.soma}</td>
        <td>${g.pares}</td>
        <td><span class="tag mid">${g.estrategia}</span></td>
        <td>R$ ${g.custo.toFixed(2)}</td>
        <td>${g.rateioPotencial === 'Alto' ? '<span class="tag hot">Alto</span>' : '<span class="tag mid">Normal</span>'}</td>
        <td>${g.premios.sena > 0 ? '<strong>Sena!</strong>' : `${g.premios.quina}Q / ${g.premios.quadra}q`}</td>
      </tr>
    `).join('');
  } catch (e) {
    if (e.message.includes('401') || e.message.includes('403')) {
      localStorage.removeItem('token');
      localStorage.removeItem('username');
      updateAuthState();
    } else {
      alert('Erro ao gerar jogos: ' + e.message);
    }
  }
}

async function loadAll() {
  const [stats, draws, frequency] = await Promise.all([
    api('/api/stats'), api('/api/draws?limit=20'), api('/api/frequency'),
  ]);
  state.stats = stats; state.draws = draws.draws; state.frequency = frequency;
  renderStats(stats); renderDraws(draws.draws); renderFrequency(frequency); renderStatsCharts(stats);
  updateAuthState();
  if (state.isLoggedIn) {
    await generateGames();
  }
}

document.getElementById('generateBtn').addEventListener('click', generateGames);
document.getElementById('generateBtnTop').addEventListener('click', () => {
  document.querySelector('[data-tab="generator"]').click();
  generateGames();
});
document.getElementById('reloadBtn').addEventListener('click', async () => {
  await api('/api/reload');
  await loadAll();
  alert('Planilha recarregada com sucesso.');
});

setTabs();
setupAuthListeners();
loadAll().catch(err => {
  document.body.innerHTML = `<div style="padding:24px;color:white;font-family:Inter,sans-serif"><h1>Erro ao carregar dashboard</h1><pre>${err.message}</pre><p>Coloque o arquivo <strong>MegaSena_Dashboard.xlsx</strong> dentro da pasta <strong>data/</strong>.</p></div>`;
});
