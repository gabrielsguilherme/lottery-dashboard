const state = { stats: null, draws: [], frequency: [], charts: {} };

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
    data: { labels: freq.map(f => f.numero), datasets: [{ label: 'Frequência', data: freq.map(f => f.frequencia), backgroundColor: freq.map(f => f.categoria === 'Quente' ? '#22c55e' : f.categoria === 'Frio' ? '#f43f5e' : '#38bdf8') }] },
    options: { responsive:true, plugins:{legend:{labels:{color:'#e2e8f0'}}}, scales:{x:{ticks:{color:'#94a3b8'}},y:{ticks:{color:'#94a3b8'}}} }
  });

  const topDelay = [...freq].sort((a,b) => b.atraso - a.atraso).slice(0, 10);
  destroyChart('delayChart');
  state.charts.delayChart = new Chart(document.getElementById('delayChart'), {
    type: 'bar',
    data: { labels: topDelay.map(f => f.numero), datasets: [{ label: 'Atraso', data: topDelay.map(f => f.atraso), backgroundColor: '#f59e0b' }] },
    options: { indexAxis:'y', responsive:true, plugins:{legend:{labels:{color:'#e2e8f0'}}}, scales:{x:{ticks:{color:'#94a3b8'}},y:{ticks:{color:'#94a3b8'}}} }
  });
}

function renderStatsCharts(stats) {
  destroyChart('sumChart');
  state.charts.sumChart = new Chart(document.getElementById('sumChart'), {
    type: 'bar',
    data: { labels: Object.keys(stats.somaDist), datasets: [{ label: 'Sorteios', data: Object.values(stats.somaDist), backgroundColor: '#38bdf8' }] },
    options: { plugins:{legend:{labels:{color:'#e2e8f0'}}}, scales:{x:{ticks:{color:'#94a3b8'}},y:{ticks:{color:'#94a3b8'}}} }
  });
  destroyChart('parityChart');
  state.charts.parityChart = new Chart(document.getElementById('parityChart'), {
    type: 'doughnut',
    data: { labels: Object.keys(stats.pariDist).map(k => `${k} pares`), datasets: [{ data: Object.values(stats.pariDist), backgroundColor: ['#ef4444','#f97316','#f59e0b','#22c55e','#14b8a6','#3b82f6','#8b5cf6'] }] },
    options: { plugins:{legend:{labels:{color:'#e2e8f0'}}} }
  });
}

function destroyChart(name) { if (state.charts[name]) state.charts[name].destroy(); }

async function generateGames() {
  const body = {
    somaMin: Number(document.getElementById('somaMin').value),
    somaMax: Number(document.getElementById('somaMax').value),
    minPares: Number(document.getElementById('minPares').value),
    maxPares: Number(document.getElementById('maxPares').value),
    count: Number(document.getElementById('countGames').value),
  };
  const games = await api('/api/generate', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body) });
  document.getElementById('gamesTable').innerHTML = games.map((g, i) => `
    <tr>
      <td>Jogo ${i + 1}</td>
      <td><div class="balls">${g.numeros.map(n => `<span class="ball">${n}</span>`).join('')}</div></td>
      <td>${g.soma}</td><td>${g.pares}</td><td>${g.estrategia}</td>
    </tr>
  `).join('');
}

async function loadAll() {
  const [stats, draws, frequency] = await Promise.all([
    api('/api/stats'), api('/api/draws?limit=20'), api('/api/frequency'),
  ]);
  state.stats = stats; state.draws = draws.draws; state.frequency = frequency;
  renderStats(stats); renderDraws(draws.draws); renderFrequency(frequency); renderStatsCharts(stats);
  await generateGames();
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
loadAll().catch(err => {
  document.body.innerHTML = `<div style="padding:24px;color:white;font-family:Inter,sans-serif"><h1>Erro ao carregar dashboard</h1><pre>${err.message}</pre><p>Coloque o arquivo <strong>MegaSena_Dashboard.xlsx</strong> dentro da pasta <strong>data/</strong>.</p></div>`;
});
