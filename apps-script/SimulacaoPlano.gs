/**
 * SIMULAÇÃO DE CENÁRIOS — set a dez/2026
 * ────────────────────────────────────────────────────────────────────────────
 * Monta a aba SIMULAÇÃO e liga a PLANO MESTRE nela. Trocar o cenário no
 * dropdown recalcula os lotes de set a dez, o TOTAL GERAL segue, e o painel
 * lê o resultado sem saber que existe simulação — para ele nada mudou.
 *
 * Por que em aba separada: o lerPlanoMestre varre a PLANO MESTRE coluna a
 * coluna e faz plano[ano][mes] = valor. Quatro colunas rotuladas "jan./26"
 * caem todas no mesmo mês e a última sobrescreve as outras — o painel passaria
 * a ler um cenário por acidente da ordem das colunas. Uma coluna por mês é o
 * que mantém a leitura determinística.
 *
 * Só set a dez variam. Janeiro a agosto estão fechados ou comprometidos:
 * aplicar ritmo de simulação a mês que já aconteceu não é cenário, é reescrever
 * história — e corromperia todo comparativo de atingimento e desvio do painel.
 *
 * Ordem de uso:
 *   1. criarSimulacao()            — cria/reconstrói a aba SIMULAÇÃO
 *   2. ligarPlanoMestre()          — troca set..dez da PLANO MESTRE por fórmula
 * Depois disso, é só mexer no dropdown ou nos dias úteis.
 */

var SIM_ABA        = 'SIMULAÇÃO';
var SIM_ABA_PLANO  = 'PLANO MESTRE';
var SIM_VAZIO      = '-';

/* Ritmos demonstrados em 2026 (jan a jul: 141 dias trabalhados).
   A = só jornada normal · B = realizado com a hora extra atual ·
   C = ritmo da venda, o único que não deixa a carteira crescer. */
var SIM_CENARIOS = [
  { nome: 'CENARIO A', ritmo: 184556 / 141, desc: 'Jornada normal — sem hora extra estrutural' },
  { nome: 'CENARIO B', ritmo: 209251 / 141, desc: 'Ritmo realizado em 2026 — mantém a HE atual' },
  { nome: 'CENARIO C', ritmo: 233726 / 141, desc: 'Ritmo da venda — não deixa a carteira crescer' }
];
var SIM_PADRAO = 'CENARIO C';

/* Dias úteis já líquidos de feriado nacional. Dezembro tem 22 no calendário,
   mas entra com 15 por causa do recesso — mesma prática de dez/2025. Fica
   editável na aba de propósito: é a premissa mais frágil da conta. */
var SIM_MESES = [
  { mes: 'set./26', col: 9,  dias: 21 },
  { mes: 'out./26', col: 10, dias: 21 },
  { mes: 'nov./26', col: 11, dias: 19 },
  { mes: 'dez./26', col: 12, dias: 15 }
];

/* Distribuição dos lotes dentro de cada mês. Os números são a base (cenário C);
   o que a simulação usa é a participação de cada lote no mês, então trocar de
   cenário reescala tudo mantendo a proporção. */
var SIM_LOTES = [
  { lote: '032-26', q: { 9: 9000 } },
  { lote: '033-26', q: { 9: 8500 } },
  { lote: '034-26', q: { 9: 9000 } },
  { lote: '035-26', q: { 9: 8310, 10: 1500 } },
  { lote: '036-26', q: { 10: 8500 } },
  { lote: '037-26', q: { 10: 8500 } },
  { lote: '038-26', q: { 10: 8500 } },
  { lote: '039-26', q: { 10: 7810, 11: 1200 } },
  { lote: '040-26', q: { 11: 7500 } },
  { lote: '041-26', q: { 11: 7500 } },
  { lote: '042-26', q: { 11: 8000 } },
  { lote: '043-26', q: { 11: 7290, 12: 1000 } },
  { lote: '044-26', q: { 12: 8000 } },
  { lote: '045-26', q: { 12: 8000 } },
  { lote: '046-26', q: { 12: 7860 } }
];

/* Layout da aba — fixo, para as fórmulas da PLANO MESTRE poderem apontar. */
var SIM_L_CENARIO = 3;    /* B3  = dropdown do cenário ativo */
var SIM_L_RITMO   = 6;    /* B6:B8  ritmos · B9 ritmo ativo */
var SIM_L_ATIVO   = 9;
var SIM_L_MES     = 12;   /* A12:C15  mês · dias úteis · total do mês */
var SIM_L_LOTE    = 19;   /* A19:E33  lotes × meses */


/* ══ 1 · MONTAR A ABA ══ */
function criarSimulacao() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(SIM_ABA);
  if (aba) ss.deleteSheet(aba);
  aba = ss.insertSheet(SIM_ABA);

  aba.getRange('A1').setValue('SIMULAÇÃO DO PLANO — set a dez/2026').setFontWeight('bold');

  /* Cenário ativo */
  aba.getRange(SIM_L_CENARIO, 1).setValue('Cenário ativo').setFontWeight('bold');
  var alvo = aba.getRange(SIM_L_CENARIO, 2);
  alvo.setValue(SIM_PADRAO)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(SIM_CENARIOS.map(function (c) { return c.nome; }), true)
        .setAllowInvalid(false).build())
      .setFontWeight('bold').setBackground('#fff2cc');

  /* Ritmos */
  aba.getRange(SIM_L_RITMO - 1, 1).setValue('RITMO (pç/dia)').setFontWeight('bold');
  SIM_CENARIOS.forEach(function (c, i) {
    aba.getRange(SIM_L_RITMO + i, 1).setValue(c.nome);
    aba.getRange(SIM_L_RITMO + i, 2).setValue(c.ritmo).setNumberFormat('#,##0.0');
    aba.getRange(SIM_L_RITMO + i, 3).setValue(c.desc).setFontColor('#666666');
  });
  aba.getRange(SIM_L_ATIVO, 1).setValue('Ritmo ativo').setFontWeight('bold');
  aba.getRange(SIM_L_ATIVO, 2).setFormula(
    '=INDEX(B' + SIM_L_RITMO + ':B' + (SIM_L_RITMO + SIM_CENARIOS.length - 1) +
    ',MATCH(B' + SIM_L_CENARIO + ',A' + SIM_L_RITMO + ':A' +
    (SIM_L_RITMO + SIM_CENARIOS.length - 1) + ',0))'
  ).setNumberFormat('#,##0.0').setFontWeight('bold');

  /* Meses: dias úteis editáveis e total do mês por fórmula */
  aba.getRange(SIM_L_MES - 1, 1, 1, 3)
     .setValues([['MÊS', 'DIAS ÚTEIS', 'TOTAL DO MÊS']]).setFontWeight('bold');
  SIM_MESES.forEach(function (m, i) {
    var l = SIM_L_MES + i;
    aba.getRange(l, 1).setValue(m.mes);
    aba.getRange(l, 2).setValue(m.dias).setBackground('#fff2cc');
    aba.getRange(l, 3).setFormula('=ROUND($B$' + SIM_L_ATIVO + '*B' + l + ',0)')
       .setNumberFormat('#,##0');
  });
  aba.getRange(SIM_L_MES + 3, 4)
     .setValue('← dezembro com recesso; confirmar antes de fechar o plano')
     .setFontColor('#cc0000');

  /* Lotes: participação fixa dentro do mês, reescalada pelo cenário ativo.
     O último lote de cada mês absorve o arredondamento, senão a soma dos
     lotes não bate com o total do mês e o TOTAL GERAL sai torto. */
  var cab = ['LOTE'].concat(SIM_MESES.map(function (m) { return m.mes; }));
  aba.getRange(SIM_L_LOTE - 1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');

  var baseMes = {};
  SIM_MESES.forEach(function (m) {
    baseMes[m.col] = SIM_LOTES.reduce(function (s, x) { return s + (x.q[m.col] || 0); }, 0);
  });
  var ultimo = {};
  SIM_LOTES.forEach(function (x, i) {
    SIM_MESES.forEach(function (m) { if (x.q[m.col]) ultimo[m.col] = i; });
  });

  SIM_LOTES.forEach(function (x, i) {
    var l = SIM_L_LOTE + i;
    aba.getRange(l, 1).setValue(x.lote);
    SIM_MESES.forEach(function (m, j) {
      var cel = aba.getRange(l, 2 + j);
      if (!x.q[m.col]) { cel.setValue(SIM_VAZIO); return; }
      var lm = SIM_L_MES + j, colL = String.fromCharCode(66 + j);
      if (ultimo[m.col] === i) {
        /* resto do mês = total − o que já foi distribuído acima */
        cel.setFormula('=$C$' + lm + '-SUM(' + colL + SIM_L_LOTE + ':' + colL + (l - 1) + ')');
      } else {
        cel.setFormula('=ROUND($C$' + lm + '*' + (x.q[m.col] / baseMes[m.col]) + ',0)');
      }
      cel.setNumberFormat('#,##0');
    });
  });

  var fim = SIM_L_LOTE + SIM_LOTES.length;
  aba.getRange(fim, 1).setValue('TOTAL').setFontWeight('bold');
  SIM_MESES.forEach(function (m, j) {
    var c = String.fromCharCode(66 + j);
    aba.getRange(fim, 2 + j)
       .setFormula('=SUM(' + c + SIM_L_LOTE + ':' + c + (fim - 1) + ')')
       .setNumberFormat('#,##0').setFontWeight('bold');
  });

  aba.setColumnWidth(1, 110).setColumnWidth(3, 130);
  SpreadsheetApp.flush();
  simAvisar('Aba SIMULAÇÃO criada. Agora rode ligarPlanoMestre() para conectá-la ao plano.');
}


/* ══ 2 · LIGAR A PLANO MESTRE NA SIMULAÇÃO ══
   Troca só as células de set a dez dos 15 lotes. Jan a ago não são tocados. */
function ligarPlanoMestre() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(SIM_ABA_PLANO);
  if (!aba) return simErro('Aba "' + SIM_ABA_PLANO + '" não encontrada.');
  if (!ss.getSheetByName(SIM_ABA)) return simErro('Rode criarSimulacao() antes.');

  var mapa = simLocalizar(aba);
  if (mapa.erro) return simErro(mapa.erro);

  var qtd    = mapa.linhaTotal - mapa.primeiroLote;
  var codigos = aba.getRange(mapa.primeiroLote, 1, qtd, 1).getValues();
  var indice  = {};
  codigos.forEach(function (c, i) { indice[String(c[0]).trim()] = mapa.primeiroLote + i; });

  var faltando = SIM_LOTES.filter(function (x) { return !indice[x.lote]; })
                          .map(function (x) { return x.lote; });
  if (faltando.length) {
    return simErro('Estes lotes não estão na ' + SIM_ABA_PLANO + ': ' + faltando.join(', ') +
                   '. Rode carregarPlanoSetDez() primeiro.');
  }

  var n = 0;
  SIM_LOTES.forEach(function (x, i) {
    var linha = indice[x.lote];
    SIM_MESES.forEach(function (m, j) {
      if (!x.q[m.col]) return;
      var origem = String.fromCharCode(66 + j) + (SIM_L_LOTE + i);
      aba.getRange(linha, m.col + 1).setFormula("='" + SIM_ABA + "'!" + origem);
      n++;
    });
  });

  SpreadsheetApp.flush();
  simAvisar(n + ' células ligadas à SIMULAÇÃO. Troque o cenário no dropdown da aba ' +
            SIM_ABA + ' e recarregue o painel.');
}


/* ══ HELPERS ══ */
function simLocalizar(aba) {
  var v = aba.getDataRange().getValues();
  var cab = -1, total = -1;
  for (var i = 0; i < v.length; i++) {
    var a = simNormaliza(v[i][0]);
    if (cab < 0 && a === 'aba') cab = i + 1;
    if (cab > 0 && total < 0 && a === 'total geral') total = i + 1;
  }
  if (cab < 0)   return { erro: 'Linha de cabeçalho (coluna A = "ABA") não encontrada.' };
  if (total < 0) return { erro: 'Linha "TOTAL GERAL" não encontrada — o painel depende dela.' };
  return { linhaCab: cab, linhaTotal: total, primeiroLote: cab + 1 };
}

function simNormaliza(v) {
  return String(v === null || v === undefined ? '' : v)
    .trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function simErro(msg) {
  Logger.log('ERRO: ' + msg);
  simAvisar(msg);
  return { ok: false, erro: msg };
}

function simAvisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
