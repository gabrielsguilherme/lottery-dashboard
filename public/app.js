const state = { stats: null, draws: [], frequency: [], sets: null, genMode: 'vetor_paginado', charts: {} };

async function api(url, options) {
  const res = await fetch(url, options);
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

function setTabs() {
  document.querySelectorAll('.tab[data-tab]').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab[data-tab]').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`panel-${tab.dataset.tab}`).classList.add('active');
    });
  });

  // Abas internas do Gerador (Modo A vs Modo B)
  document.querySelectorAll('.tab[data-gen-mode]').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab[data-gen-mode]').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      state.genMode = tab.dataset.genMode;

      const panelA = document.getElementById('gen-mode-panel-vetor_paginado');
      const panelB = document.getElementById('gen-mode-panel-vetor_feixe');
      const nextBtn = document.getElementById('nextPageBtnModoA');
      const resetBtn = document.getElementById('resetPageBtnModoA');
      const genBtn = document.getElementById('generateBtn');

      if (state.genMode === 'vetor_paginado') {
        panelA.style.display = 'block';
        panelB.style.display = 'none';
        nextBtn.style.display = 'inline-block';
        resetBtn.style.display = 'inline-block';
        genBtn.textContent = 'Gerar jogos agora (Modo A)';
      } else {
        panelA.style.display = 'none';
        panelB.style.display = 'block';
        nextBtn.style.display = 'none';
        resetBtn.style.display = 'none';
        genBtn.textContent = 'Gerar lote de elite agora (Modo B)';
      }
      generateGames();
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

function renderSetTheory(sets) {
  const bestPair = sets.topPairs[0];
  const bestTrio = sets.topTrios[0];

  document.getElementById('setsStatsGrid').innerHTML = `
    <div class="card">
      <div class="stat-label">Pares Possíveis C(60,2)</div>
      <div class="stat-value">${sets.totalPossiblePairs}</div>
      <div class="stat-extra">15 pares por jogo de 6 dezenas</div>
    </div>
    <div class="card">
      <div class="stat-label">Esperança Hipergeométrica (Par)</div>
      <div class="stat-value">${sets.expectedPairCount.toFixed(1)}x</div>
      <div class="stat-extra">P(A∩B) teórica: ${(sets.expectedPairProb * 100).toFixed(2)}%</div>
    </div>
    <div class="card">
      <div class="stat-label">Dupla Mais Sorteada Junta</div>
      <div class="stat-value">${String(bestPair.a).padStart(2, '0')} e ${String(bestPair.b).padStart(2, '0')}</div>
      <div class="stat-extra">${bestPair.count} vezes (${bestPair.vsExpected.toFixed(2)}x o esperado)</div>
    </div>
    <div class="card">
      <div class="stat-label">Trinca Mais Sorteada Junta</div>
      <div class="stat-value">${bestTrio.numeros.map(n => String(n).padStart(2, '0')).join('-')}</div>
      <div class="stat-extra">${bestTrio.count} vezes (esperado: ${sets.expectedTrioCount.toFixed(1)}x)</div>
    </div>
  `;

  document.getElementById('topPairsTable').innerHTML = sets.topPairs.map(p => `
    <tr>
      <td><div class="balls"><span class="ball">${String(p.a).padStart(2, '0')}</span><span class="ball">${String(p.b).padStart(2, '0')}</span></div></td>
      <td><strong>${p.count}x</strong> <small style="color:var(--muted);">(${(p.support * 100).toFixed(2)}%)</small></td>
      <td><span class="tag hot">${p.vsExpected.toFixed(2)}x</span></td>
      <td><strong>${p.lift.toFixed(2)}</strong></td>
    </tr>
  `).join('');

  document.getElementById('topTriosTable').innerHTML = sets.topTrios.map(t => `
    <tr>
      <td><div class="balls">${t.numeros.map(n => `<span class="ball">${String(n).padStart(2, '0')}</span>`).join('')}</div></td>
      <td><strong>${t.count}x</strong></td>
      <td><span class="tag hot">${t.vsExpected.toFixed(2)}x</span></td>
    </tr>
  `).join('');

  const selectEl = document.getElementById('selectBaseNumber');
  if (!selectEl.options.length) {
    selectEl.innerHTML = Array.from({ length: 60 }, (_, i) => i + 1)
      .map(n => `<option value="${n}">Dezena ${String(n).padStart(2, '0')}</option>`)
      .join('');
    selectEl.value = String(bestPair.a);
    selectEl.addEventListener('change', () => renderConditionalPartners(sets, Number(selectEl.value)));
  }
  renderConditionalPartners(sets, Number(selectEl.value) || bestPair.a);
}

function renderConditionalPartners(sets, baseNum) {
  const partners = sets.partnersByNumber[baseNum] || [];
  document.getElementById('conditionalTable').innerHTML = partners.map(p => `
    <tr>
      <td>
        <div class="balls">
          <span class="ball" style="background: var(--primary);">${String(baseNum).padStart(2, '0')}</span>
          <span style="font-weight:700; color:var(--muted);">∩</span>
          <span class="ball">${String(p.parceiro).padStart(2, '0')}</span>
        </div>
      </td>
      <td><strong>${p.count} vezes</strong></td>
      <td><strong>${(p.conf * 100).toFixed(1)}%</strong> das vezes que o ${String(baseNum).padStart(2, '0')} saiu</td>
      <td>${sets.expectedPairCount.toFixed(1)} vezes</td>
      <td><span class="tag ${p.vsExpected >= 1.2 ? 'hot' : 'mid'}">${p.vsExpected.toFixed(2)}x</span></td>
      <td><strong>${p.lift.toFixed(2)}</strong></td>
    </tr>
  `).join('');
}

function updateModoABadge() {
  const countA = Math.min(Math.max(Number(document.getElementById('countGamesModoA').value) || 10, 1), 30);
  const pageOffset = Math.max(Number(document.getElementById('pageOffsetModoA').value) || 0, 0);
  const startRank = pageOffset * countA + 1;
  const endRank = (pageOffset + 1) * countA;
  const badge = document.getElementById('modoAPageBadge');
  if (badge) {
    badge.textContent = `Página ${pageOffset + 1} (Top ${startRank} a ${endRank})`;
  }
}

async function generateGames() {
  const mode = state.genMode || 'vetor_paginado';
  const count = mode === 'vetor_paginado'
    ? Number(document.getElementById('countGamesModoA').value)
    : Number(document.getElementById('countGamesModoB').value);
  const pageOffset = mode === 'vetor_paginado'
    ? Number(document.getElementById('pageOffsetModoA').value)
    : 0;

  updateModoABadge();

  const body = {
    somaMin: Number(document.getElementById('somaMin').value),
    somaMax: Number(document.getElementById('somaMax').value),
    count,
    mode,
    pageOffset,
  };

  try {
    const games = await api('/api/generate', { 
      method: 'POST', 
      headers: {
        'Content-Type': 'application/json'
      }, 
      body: JSON.stringify(body) 
    });

    // Calcula métricas de cobertura de conjuntos do lote inteiro
    const uniqueNums = new Set();
    const uniquePairs = new Set();
    games.forEach(g => {
      g.numeros.forEach(n => uniqueNums.add(n));
      for (let i = 0; i < g.numeros.length; i++) {
        for (let j = i + 1; j < g.numeros.length; j++) {
          uniquePairs.add(`${g.numeros[i]}-${g.numeros[j]}`);
        }
      }
    });
    const totalPossiblePairsInBatch = games.length * 15;
    document.getElementById('batchSetSummary').innerHTML =
      `Cobertura do Lote (${games.length} jogos): ${uniqueNums.size}/60 dezenas distintas | ${uniquePairs.size}/${totalPossiblePairsInBatch} pares distintos (|Jᵢ ∩ Jₖ| ≤ 1)`;

    document.getElementById('gamesTable').innerHTML = games.map((g, i) => `
      <tr>
        <td><strong>Jogo ${i + 1}</strong></td>
        <td><div class="balls">${g.numeros.map(n => `<span class="ball">${String(n).padStart(2, '0')}</span>`).join('')}</div></td>
        <td>${g.soma}</td>
        <td>${g.pares}P/${6 - g.pares}I</td>
        <td><span class="tag ${g.afinidadeMedia >= 1.15 ? 'hot' : 'mid'}">${g.afinidadeMedia}x</span></td>
        <td><strong>${g.parMaisForte || '-'}</strong></td>
        <td><span class="tag mid">${g.paresIneditos || '15/15'}</span></td>
        <td><span class="tag ${mode === 'vetor_paginado' ? 'hot' : 'mid'}">${g.estrategia}</span></td>
        <td>R$ ${g.custo.toFixed(2)}</td>
        <td>${g.rateioPotencial === 'Alto' ? '<span class="tag hot">Alto</span>' : '<span class="tag mid">Normal</span>'}</td>
        <td>${g.premios.sena > 0 ? '<strong>Sena!</strong>' : `${g.premios.quina}Q / ${g.premios.quadra}q`}</td>
      </tr>
    `).join('');
  } catch (e) {
    alert('Erro ao gerar jogos: ' + e.message);
  }
}

async function loadAll() {
  const [stats, draws, frequency, sets] = await Promise.all([
    api('/api/stats'), api('/api/draws?limit=20'), api('/api/frequency'), api('/api/sets'),
  ]);
  state.stats = stats; state.draws = draws.draws; state.frequency = frequency; state.sets = sets;
  renderStats(stats); renderDraws(draws.draws); renderFrequency(frequency); renderStatsCharts(stats); renderSetTheory(sets);
  await generateGames();
}

document.getElementById('generateBtn').addEventListener('click', generateGames);

document.getElementById('nextPageBtnModoA').addEventListener('click', () => {
  const pageInput = document.getElementById('pageOffsetModoA');
  pageInput.value = Number(pageInput.value || 0) + 1;
  generateGames();
});

document.getElementById('resetPageBtnModoA').addEventListener('click', () => {
  document.getElementById('pageOffsetModoA').value = 0;
  generateGames();
});

document.getElementById('countGamesModoA').addEventListener('change', updateModoABadge);
document.getElementById('pageOffsetModoA').addEventListener('change', updateModoABadge);

document.getElementById('reloadBtn').addEventListener('click', async () => {
  await api('/api/reload');
  await loadAll();
  alert('Planilha recarregada com sucesso.');
});

setTabs();
loadAll().catch(err => {
  document.body.innerHTML = `<div style="padding:24px;color:white;font-family:Inter,sans-serif"><h1>Erro ao carregar dashboard</h1><pre>${err.message}</pre><p>Coloque o arquivo <strong>Mega-Sena.xlsx</strong> dentro da pasta <strong>data/</strong>.</p></div>`;
});
