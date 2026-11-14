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
      const panelC = document.getElementById('gen-mode-panel-vetor_nucleo');
      const panelR = document.getElementById('gen-mode-panel-regressivo_elite');
      const nextBtn = document.getElementById('nextPageBtnModoA');
      const resetBtn = document.getElementById('resetPageBtnModoA');
      const genBtn = document.getElementById('generateBtn');

      if (panelA) panelA.style.display = state.genMode === 'vetor_paginado' ? 'block' : 'none';
      if (panelB) panelB.style.display = state.genMode === 'vetor_feixe' ? 'block' : 'none';
      if (panelC) panelC.style.display = state.genMode === 'vetor_nucleo' ? 'block' : 'none';
      if (panelR) panelR.style.display = state.genMode === 'regressivo_elite' ? 'block' : 'none';

      if (state.genMode === 'vetor_paginado') {
        nextBtn.style.display = 'inline-block';
        resetBtn.style.display = 'inline-block';
        genBtn.textContent = 'Gerar jogos agora (Modo A)';
      } else if (state.genMode === 'vetor_feixe') {
        nextBtn.style.display = 'none';
        resetBtn.style.display = 'none';
        genBtn.textContent = 'Gerar lote de elite agora (Modo B)';
      } else if (state.genMode === 'vetor_nucleo') {
        nextBtn.style.display = 'none';
        resetBtn.style.display = 'none';
        genBtn.textContent = 'Gerar fechamento condensado (Modo C)';
      } else {
        nextBtn.style.display = 'none';
        resetBtn.style.display = 'none';
        genBtn.textContent = 'Gerar Top Candidatos Regressivos (A + B + C)';
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

async function loadRegressionSummary() {
  const banner = document.getElementById('regressiveSummaryBanner');
  if (!banner) return;
  try {
    const reg = await api('/api/regression');
    const sA = reg.summary.vetor_paginado;
    const sB = reg.summary.vetor_feixe;
    const sC = reg.summary.vetor_nucleo;
    banner.innerHTML =
      `Janela Walk-Forward (<strong>#${reg.fromConcurso} a #${reg.toConcurso}</strong>): ` +
      `<strong>Modo A</strong>: ${sA.approvedWFCount}/${sA.totalEvaluated} aprovados (≥3) e ${sA.discardedLowHitsCount} descartados (1-2 acertos) | ` +
      `<strong>Modo B</strong>: ${sB.approvedWFCount}/${sB.totalEvaluated} aprovados (≥3) e ${sB.discardedLowHitsCount} descartados | ` +
      `<strong>Modo C</strong>: ${sC.approvedWFCount}/${sC.totalEvaluated} aprovados (≥3, incl. Quinas e Senas) e ${sC.discardedLowHitsCount} descartados.`;
  } catch (e) {
    banner.textContent = '';
  }
}

function renderGamesRows(games, mode) {
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

  document.getElementById('gamesTable').innerHTML = games.map((g, i) => {
    const matchedSet = new Set(g.targetMatchedBalls || []);
    const weightBadge = g.pesoAposta
      ? `<span class="tag ${g.pesoAposta >= 2.2 ? 'hot' : 'mid'}" title="${g.classificacaoPeso || ''}">${g.pesoAposta}x (${g.confiancaPct || 75}% | ${g.participacaoLotePct || 10}% lote)</span>`
      : `<span class="tag mid">1.00x (Base)</span>`;
    const actualHitBadge = g.targetActualHits !== null && g.targetActualHits !== undefined
      ? `<div style="margin-top:4px;font-size:0.78rem;font-weight:700;color:${g.targetActualHits >= 3 ? '#166534' : '#64748b'};">🎯 Real no Alvo: ${g.targetActualHits} acerto(s)</div>`
      : '';
    return `
      <tr>
        <td><strong>Jogo ${i + 1}</strong></td>
        <td>
          <div class="balls">
            ${g.numeros.map(n => `<span class="ball" style="${matchedSet.has(n) ? 'background:#f29100;border-color:#b45309;transform:scale(1.08);' : ''}">${String(n).padStart(2, '0')}</span>`).join('')}
          </div>
          ${actualHitBadge}
        </td>
        <td>${weightBadge}</td>
        <td>${g.soma}</td>
        <td>${g.pares}P/${6 - g.pares}I</td>
        <td><span class="tag ${g.afinidadeMedia >= 1.15 ? 'hot' : 'mid'}">${g.afinidadeMedia}x</span></td>
        <td><strong>${g.parMaisForte || '-'}</strong></td>
        <td><span class="tag mid">${g.paresIneditos || '15/15'}</span></td>
        <td><span class="tag ${mode === 'vetor_paginado' ? 'hot' : 'mid'}">${g.estrategia}</span></td>
        <td>R$ ${(g.custo || 5).toFixed(2)}</td>
        <td>${g.rateioPotencial === 'Alto' ? '<span class="tag hot">Alto</span>' : '<span class="tag mid">Normal</span>'}</td>
        <td>${g.premios?.sena > 0 ? '<strong>Sena!</strong>' : `${g.premios?.quina || 0}Q / ${g.premios?.quadra || 0}q`}</td>
      </tr>
    `;
  }).join('');
}

async function generateGames() {
  const mode = state.genMode || 'vetor_paginado';
  const count = mode === 'vetor_paginado'
    ? Number(document.getElementById('countGamesModoA').value)
    : mode === 'vetor_feixe'
      ? Number(document.getElementById('countGamesModoB').value)
      : mode === 'vetor_nucleo'
        ? Number(document.getElementById('countGamesModoC').value)
        : Number(document.getElementById('countGamesModoR')?.value || 10);
  const pageOffset = mode === 'vetor_paginado'
    ? Number(document.getElementById('pageOffsetModoA').value)
    : mode === 'vetor_nucleo'
      ? Number(document.getElementById('pageOffsetModoC').value)
      : mode === 'regressivo_elite'
        ? Number(document.getElementById('pageOffsetModoR')?.value || 0)
        : 0;
  const regressiveFilter = document.getElementById('regressiveFilterCheck')
    ? document.getElementById('regressiveFilterCheck').checked
    : true;
  const targetConcurso = Number(document.getElementById('wfTargetConcurso')?.value) || ((state.stats?.latest?.concurso || 3085) + 1);
  const lookbackCount = Number(document.getElementById('wfLookbackCount')?.value) || 10;

  updateModoABadge();

  const body = {
    somaMin: Number(document.getElementById('somaMin').value),
    somaMax: Number(document.getElementById('somaMax').value),
    count,
    mode,
    pageOffset,
    regressiveFilter,
    targetConcurso,
    lookbackCount,
  };

  try {
    const games = await api('/api/generate', { 
      method: 'POST', 
      headers: {
        'Content-Type': 'application/json'
      }, 
      body: JSON.stringify(body) 
    });
    renderGamesRows(games, mode);
  } catch (e) {
    alert('Erro ao gerar jogos: ' + e.message);
  }
}

async function runWalkForwardInteractive() {
  const mode = state.genMode || 'vetor_nucleo';
  const count = mode === 'vetor_paginado'
    ? Number(document.getElementById('countGamesModoA').value)
    : mode === 'vetor_feixe'
      ? Number(document.getElementById('countGamesModoB').value)
      : mode === 'vetor_nucleo'
        ? Number(document.getElementById('countGamesModoC').value)
        : Number(document.getElementById('countGamesModoR')?.value || 10);
  const pageOffset = mode === 'vetor_paginado'
    ? Number(document.getElementById('pageOffsetModoA').value)
    : mode === 'vetor_nucleo'
      ? Number(document.getElementById('pageOffsetModoC').value)
      : mode === 'regressivo_elite'
        ? Number(document.getElementById('pageOffsetModoR')?.value || 0)
        : 0;
  const targetConcurso = Number(document.getElementById('wfTargetConcurso')?.value) || ((state.stats?.latest?.concurso || 3085) + 1);
  const lookbackCount = Number(document.getElementById('wfLookbackCount')?.value) || 10;
  const box = document.getElementById('wfAuditResultBox');

  try {
    const sim = await api('/api/walk-forward', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        targetConcurso,
        lookbackCount,
        model: mode,
        count,
        pageOffset,
        somaMin: Number(document.getElementById('somaMin').value),
        somaMax: Number(document.getElementById('somaMax').value),
      }),
    });

    renderGamesRows(sim.recommendedGames, mode);

    if (box) {
      box.style.display = 'block';
      const auditHeader = sim.targetAuditSummary
        ? `<div style="background:#fef3c7;border:1px solid #f59e0b;padding:10px;border-radius:6px;margin-bottom:10px;color:#92400e;">
             <strong>🎯 Auditoria Real no Concurso #${sim.targetAuditSummary.concurso} (${sim.targetAuditSummary.data}):</strong>
             Dezenas Sorteadas: <strong>${sim.targetAuditSummary.bolasSorteadas.map(n => String(n).padStart(2, '0')).join('-')}</strong> |
             Maior acerto no lote ponderado: <strong>${sim.targetAuditSummary.maxHitsAchieved} acertos</strong> (Jogo #${sim.targetAuditSummary.bestGameRank}, Peso ${sim.targetAuditSummary.bestGamePeso}x) |
             Jogos com ≥3 acertos: <strong>${sim.targetAuditSummary.gamesWith3Plus}</strong>
           </div>`
        : `<div style="background:#eff6ff;border:1px solid #93c5fd;padding:10px;border-radius:6px;margin-bottom:10px;color:#1e40af;">
             <strong>🔮 Simulação Prospectiva para o Concurso Futuro #${sim.targetConcurso}:</strong>
             Treinado estritamente com <strong>${sim.trainRangeForTarget.totalTrainDraws}</strong> sorteios (#${sim.trainRangeForTarget.fromConcurso} até #${sim.trainRangeForTarget.toConcurso}).
             Coleção Walk-Forward validada: <strong>${sim.validatedWindowLength}/${sim.lookbackCount} concursos</strong> (#${sim.windowRange.fromConcurso} a #${sim.windowRange.toConcurso}).
           </div>`;

      const stepsHtml = sim.stepByStepAudit.map(st => {
        const mKey = mode === 'regressivo_elite' ? 'vetor_nucleo' : mode;
        const b = st.bestByMode[mKey] || Object.values(st.bestByMode)[0] || {};
        return `<span class="tag ${b.maxHitsTop10 >= 3 ? 'hot' : 'mid'}" style="margin:2px;">#${st.concurso} (treino ≤#${st.trainUpTo}): Máx ${b.maxHitsTop10 ?? 0} acertos (Slot #${b.bestSlotTop10 || '-'})</span>`;
      }).join(' ');

      box.innerHTML = `
        ${auditHeader}
        <div style="font-size:0.82rem;color:#334155;margin-bottom:6px;">
          <strong>Modelo:</strong> ${sim.modelLabel} |
          <strong>Slots Avaliados:</strong> ${sim.totals.totalEvaluated} |
          <strong>Aprovados (≥3):</strong> ${sim.totals.totalApproved} |
          <strong>Descartados (1-2 acertos):</strong> ${sim.totals.totalDiscarded}
        </div>
        <div style="font-size:0.8rem;"><strong>Passo a Passo da Janela (${sim.validatedWindowLength} concursos):</strong><br/>${stepsHtml}</div>
      `;
    }
  } catch (e) {
    alert('Erro na simulação Walk-Forward: ' + e.message);
  }
}

async function loadAll() {
  const [stats, draws, frequency, sets] = await Promise.all([
    api('/api/stats'), api('/api/draws?limit=20'), api('/api/frequency'), api('/api/sets'),
  ]);
  state.stats = stats; state.draws = draws.draws; state.frequency = frequency; state.sets = sets;
  renderStats(stats); renderDraws(draws.draws); renderFrequency(frequency); renderStatsCharts(stats); renderSetTheory(sets);

  const nextConcurso = (stats.latest?.concurso || 3085) + 1;
  const wfTargetInput = document.getElementById('wfTargetConcurso');
  if (wfTargetInput && !wfTargetInput.value) wfTargetInput.value = nextConcurso;
  const quickConcursoInput = document.getElementById('quickConcursoInput');
  if (quickConcursoInput) quickConcursoInput.value = nextConcurso;
  const quickDataInput = document.getElementById('quickDataInput');
  if (quickDataInput && !quickDataInput.value) quickDataInput.value = new Date().toLocaleDateString('pt-BR');

  await generateGames();
  loadRegressionSummary();
}

document.getElementById('generateBtn').addEventListener('click', generateGames);
const regCheck = document.getElementById('regressiveFilterCheck');
if (regCheck) regCheck.addEventListener('change', generateGames);

const runWfBtn = document.getElementById('runWfAuditBtn');
if (runWfBtn) runWfBtn.addEventListener('click', runWalkForwardInteractive);

const toggleZtBtn = document.getElementById('toggleZeroTokenBtn');
if (toggleZtBtn) {
  toggleZtBtn.addEventListener('click', () => {
    const p = document.getElementById('zeroTokenPanel');
    if (p) p.style.display = p.style.display === 'none' ? 'block' : 'none';
  });
}

const quickAddBtn = document.getElementById('quickAddDrawBtn');
if (quickAddBtn) {
  quickAddBtn.addEventListener('click', async () => {
    const statusEl = document.getElementById('zeroTokenStatus');
    try {
      const concurso = Number(document.getElementById('quickConcursoInput').value);
      const data = document.getElementById('quickDataInput').value;
      const bolas = document.getElementById('quickBolasInput').value;
      const res = await api('/api/draws/add', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ concurso, data, bolas }),
      });
      if (statusEl) statusEl.textContent = '✅ ' + res.message;
      document.getElementById('quickBolasInput').value = '';
      document.getElementById('wfTargetConcurso').value = res.latest.concurso + 1;
      await loadAll();
    } catch (e) {
      if (statusEl) statusEl.textContent = '❌ Erro: ' + e.message;
    }
  });
}

const uploadXlsxBtn = document.getElementById('uploadXlsxBtn');
if (uploadXlsxBtn) {
  uploadXlsxBtn.addEventListener('click', async () => {
    const fileInput = document.getElementById('uploadXlsxInput');
    const statusEl = document.getElementById('zeroTokenStatus');
    const file = fileInput?.files?.[0];
    if (!file) {
      alert('Selecione um arquivo .xlsx primeiro.');
      return;
    }
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        if (statusEl) statusEl.textContent = '⏳ Enviando e processando planilha...';
        const res = await api('/api/upload-xlsx', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename: file.name, base64Data: reader.result }),
        });
        if (statusEl) statusEl.textContent = '✅ ' + res.message;
        document.getElementById('wfTargetConcurso').value = (res.latest?.concurso || 3085) + 1;
        await loadAll();
      } catch (e) {
        if (statusEl) statusEl.textContent = '❌ Erro: ' + e.message;
      }
    };
    reader.readAsDataURL(file);
  });
}

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
