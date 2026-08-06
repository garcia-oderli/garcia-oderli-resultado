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
 * Fórmulas geradas aqui não usam vírgula nem decimal com ponto. setFormula
 * traduz o nome da função (SUM vira SOMA na exibição), mas NÃO traduz o
 * separador de argumentos: numa planilha pt-BR a vírgula é separador decimal,
 * e qualquer fórmula com dois argumentos vira #ERROR!. Por isso ROUND entra
 * com um argumento só, proporção entra como fração de inteiros, e a escolha do
 * cenário é soma de produtos em vez de INDEX/MATCH.
 *
 * Ordem de uso:
 *   1. criarSimulacao()            — cria/reconstrói a aba SIMULAÇÃO
 *   2. ligarPlanoMestre()          — troca set..dez da PLANO MESTRE por fórmula
 *   3. congelarPlano()             — quando decidir: fixa os números e solta o vínculo
 * Entre o 2 e o 3 é só mexer no dropdown ou nos dias úteis. Depois do 3, a
 * PLANO MESTRE volta a ser número digitável e o dropdown deixa de afetá-la —
 * rode ligarPlanoMestre() de novo se quiser voltar a simular.
 */

var SIM_ABA        = 'SIMULAÇÃO';
var SIM_ABA_PLANO  = 'PLANO MESTRE';
var SIM_VAZIO      = '-';

/* Os três cenários. O ritmo NÃO fica aqui: é lido da HISTORICO a cada execução.
   Fixar o número no código foi erro — quando fevereiro foi corrigido de 32.547
   para 28.151, a aba continuou exibindo 1.484 pç/dia enquanto o realizado já
   era 1.453, e o plano de set a dez foi carregado com o valor velho.

   A = só jornada normal · B = realizado com a hora extra atual ·
   C = ritmo da venda, o único que não deixa a carteira crescer. */
var SIM_ANO = 2026;
var SIM_ABA_HIST = 'HISTORICO';
var SIM_CENARIOS = [
  { nome: 'CENARIO A', campo: 'A', desc: 'Jornada normal — sem hora extra estrutural' },
  { nome: 'CENARIO B', campo: 'B', desc: 'Ritmo realizado — mantém a hora extra atual' },
  { nome: 'CENARIO C', campo: 'C', desc: 'Ritmo da venda — não deixa a carteira crescer' }
];

/* Mesmo critério do painel: só mês fechado e com horas lançadas entra na conta.
   Mês recém-lançado tem produção mas ainda não tem prodSemExtras nem dias, e
   entraria como zero peças em N dias, puxando a média para baixo. */
function simRitmos() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIM_ABA_HIST);
  if (!aba) return null;
  var v = aba.getDataRange().getValues();
  if (v.length < 2) return null;
  var cab = v[0].map(function (x) { return simNormaliza(x); });
  var iAno = cab.indexOf('ano'),  iReal = cab.indexOf('producaoreal');
  var iNorm = cab.indexOf('prodsemextras'), iDias = cab.indexOf('diastrabalhados');
  var iVen = cab.indexOf('qtdevendida');
  if ([iAno, iReal, iNorm, iDias, iVen].some(function (i) { return i < 0; })) return null;

  var pr = 0, pn = 0, vn = 0, d = 0;
  for (var r = 1; r < v.length; r++) {
    if (Number(v[r][iAno]) !== SIM_ANO) continue;
    var real = simNum(v[r][iReal]), norm = simNum(v[r][iNorm]), dias = simNum(v[r][iDias]);
    if (!(real > 0 && norm > 0 && dias > 0)) continue;
    pr += real; pn += norm; vn += simNum(v[r][iVen]); d += dias;
  }
  return d > 0 ? { A: pn / d, B: pr / d, C: vn / d, dias: d } : null;
}

/* A planilha grava número em formato brasileiro como texto em algumas células —
   mesmo tratamento que o Code.gs faz ao ler. */
function simNum(v) {
  if (typeof v === 'number') return v;
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return 0;
  s = s.replace(/\./g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
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
  var rit = simRitmos();
  if (!rit) return simErro('Não consegui calcular os ritmos: confira se a aba ' +
    SIM_ABA_HIST + ' tem meses de ' + SIM_ANO + ' fechados, com horas e dias lançados.');

  /* Limpa em vez de apagar: deletar a aba transforma em #REF! toda fórmula que a
     PLANO MESTRE aponta para cá, e o #REF! não se recupera quando uma aba de
     mesmo nome é recriada. */
  var aba = ss.getSheetByName(SIM_ABA);
  if (aba) { aba.clear(); aba.clearDataValidations(); }
  else     { aba = ss.insertSheet(SIM_ABA); }

  aba.getRange('A1').setValue('SIMULAÇÃO DO PLANO — set a dez/' + SIM_ANO).setFontWeight('bold');
  aba.getRange('C1').setValue('ritmos apurados sobre ' + rit.dias + ' dias fechados de ' + SIM_ANO)
     .setFontColor('#666666');

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
    aba.getRange(SIM_L_RITMO + i, 2).setValue(rit[c.campo]).setNumberFormat('#,##0.0');
    aba.getRange(SIM_L_RITMO + i, 3).setValue(c.desc).setFontColor('#666666');
  });
  aba.getRange(SIM_L_ATIVO, 1).setValue('Ritmo ativo').setFontWeight('bold');
  /* Soma de produtos em vez de INDEX/MATCH: só a linha cujo rótulo bate com o
     dropdown multiplica por 1, as outras por 0. Fica sem vírgula — ver a nota
     sobre separador no topo do arquivo. */
  var termos = SIM_CENARIOS.map(function (c, i) {
    var l = SIM_L_RITMO + i;
    return 'B' + l + '*($B$' + SIM_L_CENARIO + '=A' + l + ')';
  });
  aba.getRange(SIM_L_ATIVO, 2).setFormula('=' + termos.join('+'))
     .setNumberFormat('#,##0.0').setFontWeight('bold');

  /* Meses: dias úteis editáveis e total do mês por fórmula */
  aba.getRange(SIM_L_MES - 1, 1, 1, 3)
     .setValues([['MÊS', 'DIAS ÚTEIS', 'TOTAL DO MÊS']]).setFontWeight('bold');
  SIM_MESES.forEach(function (m, i) {
    var l = SIM_L_MES + i;
    aba.getRange(l, 1).setValue(m.mes);
    aba.getRange(l, 2).setValue(m.dias).setBackground('#fff2cc');
    aba.getRange(l, 3).setFormula('=ROUND($B$' + SIM_L_ATIVO + '*B' + l + ')')
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
        cel.setFormula('=ROUND($C$' + lm + '*' + x.q[m.col] + '/' + baseMes[m.col] + ')');
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


/* ══ 3 · CONGELAR ══
   Substitui as fórmulas de set a dez pelo número que elas estão mostrando.
   Serve para dois casos: fechar o plano depois de decidido o cenário, e abrir
   mão da simulação num mês específico para ajustar lote na mão.

   Congelar é ida sem volta automática: o vínculo com a SIMULAÇÃO some e o
   dropdown para de afetar a PLANO MESTRE. Para voltar a simular, é rodar
   ligarPlanoMestre() de novo — que sobrescreve o que tiver sido digitado. */
function congelarPlano(mesesAlvo) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(SIM_ABA_PLANO);
  if (!aba) return simErro('Aba "' + SIM_ABA_PLANO + '" não encontrada.');

  var mapa = simLocalizar(aba);
  if (mapa.erro) return simErro(mapa.erro);

  /* Sem argumento congela set a dez inteiro; com argumento (ex.: ['set./26'])
     congela só os meses pedidos e o resto continua simulando. */
  var filtro = null;
  if (mesesAlvo && mesesAlvo.length) {
    filtro = {};
    mesesAlvo.forEach(function (m) { filtro[simNormaliza(m)] = true; });
  }

  var qtd     = mapa.linhaTotal - mapa.primeiroLote;
  var codigos = aba.getRange(mapa.primeiroLote, 1, qtd, 1).getValues();
  var indice  = {};
  codigos.forEach(function (c, i) { indice[String(c[0]).trim()] = mapa.primeiroLote + i; });

  var simAba = ss.getSheetByName(SIM_ABA);
  var cenario = simAba ? simAba.getRange(SIM_L_CENARIO, 2).getValue() : '(desconhecido)';

  var congeladas = 0, jaFixas = 0;
  SIM_LOTES.forEach(function (x) {
    var linha = indice[x.lote];
    if (!linha) return;
    SIM_MESES.forEach(function (m) {
      if (!x.q[m.col]) return;
      if (filtro && !filtro[simNormaliza(m.mes)]) return;
      var cel = aba.getRange(linha, m.col + 1);
      if (!cel.getFormula()) { jaFixas++; return; }
      /* getValue() já devolve o resultado calculado — é ele que vira o número. */
      cel.setValue(cel.getValue());
      congeladas++;
    });
  });

  SpreadsheetApp.flush();
  var rel = validarPlanoMestre();   /* mesma checagem da carga: lotes x TOTAL GERAL */
  simAvisar('Congeladas ' + congeladas + ' células no cenário "' + cenario + '"' +
            (jaFixas ? ' (' + jaFixas + ' já eram número).' : '.') +
            (rel && rel.ok === false
              ? '\n\nATENÇÃO — a validação acusou:\n' + rel.problemas.join('\n')
              : '\nValidação OK.'));
  return { congeladas: congeladas, jaFixas: jaFixas, cenario: cenario };
}

/* Atalhos para chamar do editor sem digitar argumento. */
function congelarSomenteSetembro() { return congelarPlano(['set./26']); }
function congelarSetOut()          { return congelarPlano(['set./26', 'out./26']); }


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
