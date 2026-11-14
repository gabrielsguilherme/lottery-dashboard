/**
 * Motor Genérico e Agnóstico de Simulação Walk-Forward e Ponderação de Apostas
 * Compatível com Mega-Sena (lottery-dashboard) e +Milionária (MMlottery)
 */

const VALID_MODELS = [
  'vetor_paginado',   // Modo A: Vetor Determinístico Paginado
  'vetor_feixe',      // Modo B: Vetor Feixe Top-3
  'vetor_nucleo',     // Modo C: Fechamento Condensado (5 Canais Estruturais)
  'regressivo_elite', // Modo R: Seletor Híbrido Ponderado (A + B + C)
];

const MODEL_LABELS = {
  vetor_paginado: 'Modo A (Vetor Determinístico)',
  vetor_feixe: 'Modo B (Vetor Feixe Top-3)',
  vetor_nucleo: 'Modo C (5 Canais de Fusão Estrutural)',
  regressivo_elite: 'Modo R (Combinado Ponderado A+B+C)',
};

const wfModeCache = new Map();

function clearWalkForwardCache() {
  wfModeCache.clear();
}

function getDrawBalls(draw) {
  return draw?.bolas || draw?.numeros || draw?.dezenas || [];
}

function getDrawTrevos(draw) {
  return draw?.trevos || [];
}

/**
 * Executa a simulação Walk-Forward estrita (sem vazamento de futuro) e pondera os jogos gerados
 * para o concurso alvo (`targetConcurso`).
 *
 * @param {Object} params
 * @param {number} params.targetConcurso - Número do concurso que vamos simular (ex: 3086 ou 3082, 408 ou 405)
 * @param {number} params.lookbackCount - Quantos concursos para trás considerar na janela Walk-Forward
 * @param {Array}  params.windowDraws - Coleção de concursos da janela Walk-Forward (deve ter tamanho === lookbackCount)
 * @param {Array}  params.allDraws - Histórico completo de concursos disponível
 * @param {string} params.model - Modelo de cálculo ('vetor_paginado' | 'vetor_feixe' | 'vetor_nucleo' | 'regressivo_elite')
 * @param {Function} params.generateGamesFn - Função geradora bruta do projeto: (trainDraws, options) => games[]
 * @param {number} [params.count=10] - Quantidade de jogos recomendados a retornar
 * @param {number} [params.pageOffset=0] - Página de resultados
 * @param {number} [params.somaMin] - Soma mínima
 * @param {number} [params.somaMax] - Soma máxima
 * @param {boolean} [params.hasTrevos=false] - Se a loteria possui segunda matriz de trevos (+Milionária)
 */
function runWalkForwardSimulation(params = {}) {
  const {
    targetConcurso,
    lookbackCount,
    windowDraws,
    allDraws,
    model = 'vetor_nucleo',
    generateGamesFn,
    count = 10,
    pageOffset = 0,
    somaMin,
    somaMax,
    hasTrevos = false,
  } = params;

  // 1. Validação do modelo de cálculo
  if (!VALID_MODELS.includes(model)) {
    throw new Error(
      `Modelo de cálculo inválido: "${model}". Modelos implementados permitidos: ${VALID_MODELS.join(', ')}.`
    );
  }

  // 2. Validação do concurso alvo
  const targetNum = Number(targetConcurso);
  if (!Number.isInteger(targetNum) || targetNum < 20) {
    throw new Error(`Número do concurso alvo inválido (${targetConcurso}). Informe um concurso >= 20.`);
  }

  // 3. Validação da quantidade de concursos para trás (lookbackCount) e da coleção (windowDraws)
  const expectedLen = Number(lookbackCount);
  if (!Number.isInteger(expectedLen) || expectedLen < 1 || expectedLen > 100) {
    throw new Error(`Parâmetro lookbackCount inválido (${lookbackCount}). Informe um inteiro entre 1 e 100.`);
  }

  if (!Array.isArray(windowDraws)) {
    throw new Error('O parâmetro windowDraws (coleção de concursos da janela) deve ser um Array.');
  }

  if (windowDraws.length !== expectedLen) {
    throw new Error(
      `Validação falhou: a coleção de concursos recebida possui ${windowDraws.length} concurso(s), ` +
      `mas o parâmetro lookbackCount exige exatamente ${expectedLen} concurso(s).`
    );
  }

  if (!Array.isArray(allDraws) || allDraws.length < expectedLen + 5) {
    throw new Error('Histórico completo (allDraws) insuficiente para realizar o corte de treino.');
  }

  if (typeof generateGamesFn !== 'function') {
    throw new Error('Função geradora (generateGamesFn) não fornecida ao motor Walk-Forward.');
  }

  // Garante que nenhum concurso da coleção windowDraws seja >= targetConcurso (zero vazamento de futuro)
  for (const wd of windowDraws) {
    if (wd.concurso >= targetNum) {
      throw new Error(
        `Vazamento de dados detectado: o concurso #${wd.concurso} na coleção windowDraws não é anterior ao concurso alvo #${targetNum}.`
      );
    }
  }

  // Base de treino estrita para o concurso alvo (apenas concursos < targetConcurso)
  const trainForTarget = allDraws.filter(d => d.concurso < targetNum);
  if (trainForTarget.length < 15) {
    throw new Error(`Concurso alvo #${targetNum} possui apenas ${trainForTarget.length} sorteios anteriores (mínimo: 15).`);
  }

  // Verifica se o concurso alvo já ocorreu no histórico (para auditoria automática de acertos reais)
  const actualTargetDraw = allDraws.find(d => d.concurso === targetNum) || null;

  // Quais modos base serão simulados no Walk-Forward
  const modesToSimulate = model === 'regressivo_elite'
    ? [
        { mode: 'vetor_nucleo', label: 'Modo C', pagesToScan: 4, beamSeedBase: 3000 },
        { mode: 'vetor_paginado', label: 'Modo A', pagesToScan: 5, beamSeedBase: 1000 },
        { mode: 'vetor_feixe', label: 'Modo B', pagesToScan: 5, beamSeedBase: 2000 },
      ]
    : [
        {
          mode,
          label: MODEL_LABELS[model] || model,
          pagesToScan: model === 'vetor_nucleo' ? 4 : 6,
          beamSeedBase: model === 'vetor_feixe' ? 2000 : 1000,
        },
      ];

  const modeResults = {};
  const stepByStepAudit = [];

  // Inicializa auditoria por concurso da janela
  for (const wd of windowDraws) {
    stepByStepAudit.push({
      concurso: wd.concurso,
      data: wd.data || '-',
      bolas: getDrawBalls(wd),
      trevos: getDrawTrevos(wd),
      trainCount: 0,
      trainUpTo: 0,
      bestByMode: {},
    });
  }

  const windowKey = windowDraws.map(w => w.concurso).join(',');
  const baseCacheKey = `${allDraws.length}_${allDraws[allDraws.length - 1]?.concurso}_${targetNum}_${expectedLen}_${windowKey}_${somaMin || 0}_${somaMax || 0}_${hasTrevos ? 1 : 0}`;

  for (const cfg of modesToSimulate) {
    const modeCacheKey = `${baseCacheKey}_${cfg.mode}`;
    if (wfModeCache.has(modeCacheKey)) {
      const cached = wfModeCache.get(modeCacheKey);
      modeResults[cfg.mode] = cached.modeResult;
      for (let wIdx = 0; wIdx < windowDraws.length; wIdx++) {
        stepByStepAudit[wIdx].trainCount = cached.stepMeta[wIdx].trainCount;
        stepByStepAudit[wIdx].trainUpTo = cached.stepMeta[wIdx].trainUpTo;
        stepByStepAudit[wIdx].bestByMode[cfg.mode] = cached.stepMeta[wIdx].best;
      }
      continue;
    }

    const totalSlots = cfg.pagesToScan * 10;
    const slotStats = Array.from({ length: totalSlots }, (_, sIdx) => ({
      slot: sIdx + 1,
      page: Math.floor(sIdx / 10) + 1,
      posInPage: (sIdx % 10) + 1,
      maxHits: 0,
      maxTrevosOnBest: 0,
      hits3Plus: 0,
      hits4Plus: 0,
      hits5Plus: 0,
      totalHitsSum: 0,
      totalTrevosSum: 0,
      bestDrawLabel: '-',
      bestConcurso: null,
      history: [],
    }));

    for (let wIdx = 0; wIdx < windowDraws.length; wIdx++) {
      const stepTarget = windowDraws[wIdx];
      const stepTargetBalls = getDrawBalls(stepTarget);
      const stepTargetTrevos = getDrawTrevos(stepTarget);

      // Corte estrito de treino: apenas concursos < stepTarget.concurso
      const stepTrain = allDraws.filter(d => d.concurso < stepTarget.concurso);
      stepByStepAudit[wIdx].trainCount = stepTrain.length;
      stepByStepAudit[wIdx].trainUpTo = stepTrain[stepTrain.length - 1]?.concurso || 0;

      const stepCands = [];
      for (let p = 0; p < cfg.pagesToScan; p++) {
        const pageGames = generateGamesFn(stepTrain, {
          somaMin,
          somaMax,
          count: 10,
          mode: cfg.mode,
          pageOffset: p,
          _beamSeed: cfg.beamSeedBase + p * 37,
          _raw: true,
        });
        stepCands.push(...pageGames);
      }

      let bestStepHits = -1;
      let bestStepTrevos = -1;
      let bestStepGame = null;
      let bestStepSlot = -1;

      // Avalia o Top 10 (Página 1) e também todos os slots varridos
      for (let sIdx = 0; sIdx < Math.min(totalSlots, stepCands.length); sIdx++) {
        const g = stepCands[sIdx];
        const h = g.numeros.filter(n => stepTargetBalls.includes(n)).length;
        const tr = hasTrevos ? (g.trevos || []).filter(t => stepTargetTrevos.includes(t)).length : 0;

        const st = slotStats[sIdx];
        st.totalHitsSum += h;
        st.totalTrevosSum += tr;
        st.history.push({ concurso: stepTarget.concurso, hits: h, trevos: tr });

        if (h >= 3) st.hits3Plus++;
        if (h >= 4) st.hits4Plus++;
        if (h >= 5) st.hits5Plus++;

        if (h > st.maxHits || (h === st.maxHits && tr > st.maxTrevosOnBest)) {
          st.maxHits = h;
          st.maxTrevosOnBest = tr;
          st.bestConcurso = stepTarget.concurso;
          st.bestDrawLabel = hasTrevos
            ? `#${stepTarget.concurso} (${h}+${tr}🍀)`
            : `#${stepTarget.concurso} (${h} acertos)`;
        }

        // Para o resumo passo a passo da janela, destaca o melhor resultado nos 10 primeiros slots (ou no pool)
        if (sIdx < 10) {
          if (h > bestStepHits || (h === bestStepHits && tr > bestStepTrevos)) {
            bestStepHits = h;
            bestStepTrevos = tr;
            bestStepGame = g.numeros;
            bestStepSlot = sIdx + 1;
          }
        }
      }

      stepByStepAudit[wIdx].bestByMode[cfg.mode] = {
        maxHitsTop10: bestStepHits,
        maxTrevosTop10: bestStepTrevos,
        bestSlotTop10: bestStepSlot,
        bestNumerosTop10: bestStepGame,
      };
    }

    // Gera os candidatos para o concurso alvo (`targetConcurso`) usando apenas `trainForTarget` (< targetConcurso)
    const targetCands = [];
    for (let p = 0; p < cfg.pagesToScan; p++) {
      const pageGames = generateGamesFn(trainForTarget, {
        somaMin,
        somaMax,
        count: 10,
        mode: cfg.mode,
        pageOffset: p,
        _beamSeed: cfg.beamSeedBase + p * 37,
        _raw: true,
      });
      targetCands.push(...pageGames);
    }

    const recent15 = trainForTarget.slice(-15);
    const resonanceSlice = trainForTarget.slice(-180, -15);

    const enriched = targetCands.map((g, sIdx) => {
      const st = slotStats[sIdx];
      const avgHits = Number((st.totalHitsSum / windowDraws.length).toFixed(2));
      const exhaustedInLast15 = recent15.some(d => g.numeros.filter(n => getDrawBalls(d).includes(n)).length >= 4);

      let maxResonance = 0;
      let bestResTrevos = 0;
      let bestResDraw = '-';
      for (const d of resonanceSlice) {
        const rh = g.numeros.filter(n => getDrawBalls(d).includes(n)).length;
        const rt = hasTrevos ? (g.trevos || []).filter(t => getDrawTrevos(d).includes(t)).length : 0;
        if (rh > maxResonance || (rh === maxResonance && rt > bestResTrevos)) {
          maxResonance = rh;
          bestResTrevos = rt;
          bestResDraw = hasTrevos ? `#${d.concurso} (${rh}+${rt}🍀)` : `#${d.concurso} (${rh} acertos)`;
        }
      }

      const wfApproved = st.maxHits >= 3 && !exhaustedInLast15;
      const resonanceApproved = maxResonance >= 3 && !exhaustedInLast15;

      // Cálculo do Peso da Aposta (wfWeight) baseado no desempenho Walk-Forward do slot
      const rawWeight =
        1.0 +
        Math.max(0, st.maxHits - 2) * 0.45 +
        st.hits3Plus * 0.25 +
        st.hits4Plus * 0.55 +
        st.hits5Plus * 0.85 +
        (hasTrevos ? st.maxTrevosOnBest * 0.18 : 0) +
        avgHits * 0.20;

      const pesoAposta = wfApproved
        ? Number(Math.min(3.5, rawWeight).toFixed(2))
        : resonanceApproved
          ? Number(Math.min(1.6, 0.85 + maxResonance * 0.15).toFixed(2))
          : 0.40;

      const confiancaPct = wfApproved
        ? Math.min(98, Math.round(52 + (st.maxHits - 3) * 12 + st.hits3Plus * 6 + st.hits4Plus * 8 + (hasTrevos ? st.maxTrevosOnBest * 3 : 0)))
        : resonanceApproved
          ? Math.min(68, Math.round(40 + maxResonance * 6))
          : 20;

      const classificacaoPeso =
        pesoAposta >= 2.4
          ? '🔥 Peso Máximo (Elite)'
          : pesoAposta >= 1.7
            ? '⭐ Peso Alto'
            : wfApproved || resonanceApproved
              ? '✅ Peso Moderado'
              : '⚠️ Descartado (≤2 acertos)';

      const regressiveScore =
        (wfApproved ? 1000 : resonanceApproved ? 300 : 0) +
        pesoAposta * 220 +
        st.maxHits * 110 +
        st.hits4Plus * 80 +
        st.hits3Plus * 45 +
        avgHits * 50 +
        (hasTrevos ? st.maxTrevosOnBest * 25 : 0);

      // Se o concurso alvo já ocorreu, calcula quantos acertos este jogo teve no targetConcurso real
      let targetActualHits = null;
      let targetActualTrevos = null;
      let targetMatchedBalls = [];
      let targetMatchedTrevos = [];
      if (actualTargetDraw) {
        const actualBalls = getDrawBalls(actualTargetDraw);
        const actualTrevos = getDrawTrevos(actualTargetDraw);
        targetMatchedBalls = g.numeros.filter(n => actualBalls.includes(n));
        targetActualHits = targetMatchedBalls.length;
        if (hasTrevos) {
          targetMatchedTrevos = (g.trevos || []).filter(t => actualTrevos.includes(t));
          targetActualTrevos = targetMatchedTrevos.length;
        }
      }

      return {
        ...g,
        modeCode: cfg.mode,
        modeLabel: cfg.label,
        slotOriginal: st.slot,
        pageOriginal: st.page,
        posInPage: st.posInPage,
        wfMaxHits: st.maxHits,
        wfMaxTrevos: st.maxTrevosOnBest,
        wfHits3Plus: st.hits3Plus,
        wfHits4Plus: st.hits4Plus,
        wfHits5Plus: st.hits5Plus,
        wfAvgHits: avgHits,
        wfBestDraw: st.bestDrawLabel,
        maxResonance,
        bestResDraw,
        cycleExhausted: exhaustedInLast15,
        wfApproved,
        resonanceApproved,
        pesoAposta,
        confiancaPct,
        classificacaoPeso,
        regressiveScore,
        targetActualHits,
        targetActualTrevos,
        targetMatchedBalls,
        targetMatchedTrevos,
      };
    });

    modeResults[cfg.mode] = {
      mode: cfg.mode,
      label: cfg.label,
      totalEvaluated: enriched.length,
      approvedWFCount: enriched.filter(c => c.wfApproved).length,
      discardedLowHitsCount: enriched.filter(c => !c.wfApproved).length,
      avgModeWeight: Number(
        (
          enriched.filter(c => c.wfApproved).reduce((s, c) => s + c.pesoAposta, 0) /
          Math.max(1, enriched.filter(c => c.wfApproved).length)
        ).toFixed(2)
      ),
      bestMaxHitsInWindow: Math.max(0, ...enriched.map(c => c.wfMaxHits)),
      candidates: enriched,
    };

    wfModeCache.set(modeCacheKey, {
      modeResult: modeResults[cfg.mode],
      stepMeta: stepByStepAudit.map(s => ({
        trainCount: s.trainCount,
        trainUpTo: s.trainUpTo,
        best: s.bestByMode[cfg.mode],
      })),
    });
  }

  // Seleciona e pondera os jogos finais para o targetConcurso descartando <= 2 acertos
  let candidatePool = [];
  if (model === 'regressivo_elite') {
    const allApproved = [
      ...(modeResults.vetor_nucleo?.candidates.filter(c => c.wfApproved) || []),
      ...(modeResults.vetor_paginado?.candidates.filter(c => c.wfApproved) || []),
      ...(modeResults.vetor_feixe?.candidates.filter(c => c.wfApproved) || []),
    ].sort((a, b) => b.regressiveScore - a.regressiveScore);

    const allResonance = [
      ...(modeResults.vetor_nucleo?.candidates.filter(c => !c.wfApproved && c.resonanceApproved) || []),
      ...(modeResults.vetor_paginado?.candidates.filter(c => !c.wfApproved && c.resonanceApproved) || []),
      ...(modeResults.vetor_feixe?.candidates.filter(c => !c.wfApproved && c.resonanceApproved) || []),
    ].sort((a, b) => b.regressiveScore - a.regressiveScore);

    candidatePool = [...allApproved, ...allResonance];
  } else {
    const mRes = modeResults[model];
    // Ordena os aprovados pelo peso Walk-Forward (ponderando os slots mais fortes no topo!)
    const wfPass = mRes.candidates.filter(c => c.wfApproved).sort((a, b) => b.regressiveScore - a.regressiveScore);
    const resPass = mRes.candidates.filter(c => !c.wfApproved && c.resonanceApproved).sort((a, b) => b.regressiveScore - a.regressiveScore);
    const rest = mRes.candidates.filter(c => !c.wfApproved && !c.resonanceApproved);
    candidatePool = [...wfPass, ...resPass, ...rest];
  }

  const totalDiscarded = Object.values(modeResults).reduce((s, m) => s + m.discardedLowHitsCount, 0);
  const totalApproved = Object.values(modeResults).reduce((s, m) => s + m.approvedWFCount, 0);
  const totalEvaluated = Object.values(modeResults).reduce((s, m) => s + m.totalEvaluated, 0);

  const selected = [];
  const seenSigs = new Set();
  const targetTotal = (pageOffset + 1) * count;

  for (const cand of candidatePool) {
    const sig = cand.numeros.join('-');
    if (seenSigs.has(sig)) continue;
    if (selected.some(s => s.numeros.filter(n => cand.numeros.includes(n)).length >= 5)) continue;
    seenSigs.add(sig);
    selected.push(cand);
    if (selected.length >= targetTotal) break;
  }

  const finalSlice = selected.slice(pageOffset * count, (pageOffset + 1) * count);
  const totalWeightSum = finalSlice.reduce((s, c) => s + (c.pesoAposta || 1), 0) || 1;

  const recommendedGames = finalSlice.map((c, idx) => {
    const badge = c.wfApproved
      ? `WF ≥3: ${c.wfBestDraw}`
      : `Ressonância ≥3: ${c.bestResDraw}`;
    const sharePct = Number(((c.pesoAposta / totalWeightSum) * 100).toFixed(1));
    return {
      ...c,
      rankPonderado: pageOffset * count + idx + 1,
      participacaoLotePct: sharePct,
      descartados1ou2: totalDiscarded,
      paresIneditos: `${badge} | Peso ${c.pesoAposta}x (${c.confiancaPct}%)`,
      estrategia: `${c.estrategia} • [Peso ${c.pesoAposta}x | ${c.wfApproved ? c.wfBestDraw : c.bestResDraw}]`,
    };
  });

  // Se o concurso alvo já existe no histórico, sumariza o resultado da auditoria no targetConcurso
  let targetAuditSummary = null;
  if (actualTargetDraw) {
    let maxH = -1;
    let maxT = -1;
    let bestGameInTarget = null;
    for (const g of recommendedGames) {
      if (
        g.targetActualHits > maxH ||
        (g.targetActualHits === maxH && (g.targetActualTrevos || 0) > maxT)
      ) {
        maxH = g.targetActualHits;
        maxT = g.targetActualTrevos || 0;
        bestGameInTarget = g;
      }
    }
    targetAuditSummary = {
      concurso: actualTargetDraw.concurso,
      data: actualTargetDraw.data || '-',
      bolasSorteadas: getDrawBalls(actualTargetDraw),
      trevosSorteados: getDrawTrevos(actualTargetDraw),
      maxHitsAchieved: maxH,
      maxTrevosAchieved: hasTrevos ? maxT : null,
      bestGameRank: bestGameInTarget?.rankPonderado || null,
      bestGameNumeros: bestGameInTarget?.numeros || [],
      bestGameTrevos: bestGameInTarget?.trevos || [],
      bestGamePeso: bestGameInTarget?.pesoAposta || null,
      gamesWith2Plus: recommendedGames.filter(g => g.targetActualHits >= 2).length,
      gamesWith3Plus: recommendedGames.filter(g => g.targetActualHits >= 3).length,
      gamesWith4Plus: recommendedGames.filter(g => g.targetActualHits >= 4).length,
    };
  }

  return {
    targetConcurso: targetNum,
    isHistoricalTarget: Boolean(actualTargetDraw),
    lookbackCount: expectedLen,
    validatedWindowLength: windowDraws.length,
    windowRange: {
      fromConcurso: windowDraws[0]?.concurso,
      toConcurso: windowDraws[windowDraws.length - 1]?.concurso,
    },
    trainRangeForTarget: {
      fromConcurso: trainForTarget[0]?.concurso,
      toConcurso: trainForTarget[trainForTarget.length - 1]?.concurso,
      totalTrainDraws: trainForTarget.length,
    },
    model,
    modelLabel: MODEL_LABELS[model] || model,
    totals: {
      totalEvaluated,
      totalApproved,
      totalDiscarded,
    },
    modesSummary: Object.fromEntries(
      Object.entries(modeResults).map(([k, v]) => [
        k,
        {
          mode: v.mode,
          label: v.label,
          totalEvaluated: v.totalEvaluated,
          approvedWFCount: v.approvedWFCount,
          discardedLowHitsCount: v.discardedLowHitsCount,
          avgModeWeight: v.avgModeWeight,
          bestMaxHitsInWindow: v.bestMaxHitsInWindow,
        },
      ])
    ),
    targetAuditSummary,
    stepByStepAudit,
    recommendedGames,
  };
}

module.exports = {
  VALID_MODELS,
  MODEL_LABELS,
  runWalkForwardSimulation,
  clearWalkForwardCache,
};
