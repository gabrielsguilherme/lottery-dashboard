const state = { stats: null, draws: [], frequency: [], charts: {}, token: null, username: null, isLoggedIn: false };

function updateAuthState() {
  const token = localStorage.getItem('token');
  const username = localStorage.getItem('username');

  if (token && username) {
    state.token = token;
    state.username = username;
    state.isLoggedIn = true;
    document.getElementById('authContainer').style.display = 'none';
    document.getElementById('appContent').style.display = 'block';
    document.getElementById('headerActions').style.display = 'flex';
    document.getElementById('userGreeting').innerText = `Olá, ${username}!`;
    document.getElementById('usersTab').style.display = 'inline-block';
  } else {
    state.token = null;
    state.username = null;
    state.isLoggedIn = false;
    document.getElementById('authContainer').style.display = 'block';
    document.getElementById('appContent').style.display = 'none';
    document.getElementById('headerActions').style.display = 'none';
    document.getElementById('usersTab').style.display = 'none';
    
    const activeTab = document.querySelector('.tab.active');
    if (activeTab && activeTab.dataset.tab === 'users') {
      document.querySelector('[data-tab="results"]').click();
    }
  }
}

async function setupGoogleAuth() {
  try {
    const { clientId } = await api('/api/auth/google/client-id');
    if (clientId && typeof google !== 'undefined') {
      google.accounts.id.initialize({
        client_id: clientId,
        callback: handleGoogleLogin
      });
      google.accounts.id.renderButton(
        document.getElementById('googleBtn'),
        { theme: 'outline', size: 'large', width: '280' }
      );
    }
  } catch (e) {
    console.error('Erro ao inicializar Google Sign-In:', e);
  }
}

async function handleGoogleLogin(response) {
  try {
    const data = await api('/api/auth/google', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ credential: response.credential })
    });

    localStorage.setItem('token', data.token);
    localStorage.setItem('username', data.username);
    updateAuthState();
    await generateGames();
  } catch (e) {
    alert('Erro no login do Google: ' + e.message);
  }
}

async function loadUsersList() {
  if (!state.isLoggedIn) return;
  try {
    const users = await api('/api/users', {
      headers: { 'Authorization': `Bearer ${state.token}` }
    });
    
    document.getElementById('usersTableBody').innerHTML = users.map(u => `
      <tr>
        <td>${u.name}</td>
        <td>${u.email}</td>
        <td>${u.username}</td>
        <td>
          <button class="btn-secondary" style="padding: 6px 12px; font-size: 0.85rem;" onclick="openEditUserModal('${u.id}', '${u.name}', '${u.email}', '${u.username}')">Editar</button>
          <button class="btn-secondary" style="padding: 6px 12px; font-size: 0.85rem; border-color: var(--danger); color: var(--danger);" onclick="deleteUser('${u.id}')">Excluir</button>
        </td>
      </tr>
    `).join('');
  } catch (e) {
    alert('Erro ao carregar usuários: ' + e.message);
  }
}

window.openEditUserModal = function(id, name, email, username) {
  document.getElementById('modalUserId').value = id;
  document.getElementById('modalName').value = name;
  document.getElementById('modalEmail').value = email;
  document.getElementById('modalUsername').value = username;
  document.getElementById('modalPassword').value = '';
  
  document.getElementById('modalTitle').innerText = 'Editar Usuário';
  document.getElementById('passwordHint').style.display = 'inline';
  document.getElementById('userModal').style.display = 'grid';
};

window.deleteUser = async function(id) {
  if (!confirm('Tem certeza que deseja excluir este usuário?')) return;
  try {
    await api(`/api/users/${id}`, {
      method: 'DELETE',
      headers: { 'Authorization': `Bearer ${state.token}` }
    });
    alert('Usuário excluído com sucesso!');
    loadUsersList();
  } catch (e) {
    alert(e.message || 'Erro ao excluir usuário.');
  }
};

function setupUserCRUD() {
  document.getElementById('createUserBtn').addEventListener('click', () => {
    document.getElementById('modalUserId').value = '';
    document.getElementById('modalName').value = '';
    document.getElementById('modalEmail').value = '';
    document.getElementById('modalUsername').value = '';
    document.getElementById('modalPassword').value = '';
    
    document.getElementById('modalTitle').innerText = 'Cadastrar Usuário';
    document.getElementById('passwordHint').style.display = 'none';
    document.getElementById('userModal').style.display = 'grid';
  });

  document.getElementById('closeModalBtn').addEventListener('click', () => {
    document.getElementById('userModal').style.display = 'none';
  });

  document.getElementById('saveUserBtn').addEventListener('click', async () => {
    const id = document.getElementById('modalUserId').value;
    const name = document.getElementById('modalName').value.trim();
    const email = document.getElementById('modalEmail').value.trim();
    const username = document.getElementById('modalUsername').value.trim();
    const password = document.getElementById('modalPassword').value;

    if (!name || !email || !username || (!id && !password)) {
      alert('Por favor, preencha todos os campos obrigatórios.');
      return;
    }

    const body = { name, email, username };
    if (password) body.password = password;

    try {
      const method = id ? 'PUT' : 'POST';
      const url = id ? `/api/users/${id}` : '/api/users';
      
      await api(url, {
        method,
        headers: { 
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${state.token}`
        },
        body: JSON.stringify(body)
      });

      alert(id ? 'Usuário atualizado com sucesso!' : 'Usuário cadastrado com sucesso!');
      document.getElementById('userModal').style.display = 'none';
      loadUsersList();
    } catch (e) {
      alert(e.message || 'Erro ao salvar usuário.');
    }
  });
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
    const name = document.getElementById('registerName').value.trim();
    const email = document.getElementById('registerEmail').value.trim();
    const username = document.getElementById('registerUsername').value.trim();
    const password = document.getElementById('registerPassword').value;

    if (!name || !email || !username || !password) {
      alert('Por favor, preencha todos os campos obrigatórios (nome, email, usuário e senha).');
      return;
    }

    try {
      const data = await api('/api/auth/register', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, email, username, password })
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
      
      if (tab.dataset.tab === 'users') {
        loadUsersList();
      }
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
  setupGoogleAuth();
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
setupUserCRUD();
loadAll().catch(err => {
  document.body.innerHTML = `<div style="padding:24px;color:white;font-family:Inter,sans-serif"><h1>Erro ao carregar dashboard</h1><pre>${err.message}</pre><p>Coloque o arquivo <strong>MegaSena_Dashboard.xlsx</strong> dentro da pasta <strong>data/</strong>.</p></div>`;
});
