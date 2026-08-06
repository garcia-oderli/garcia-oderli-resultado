/**
 * CARGA DO PLANO MESTRE — set a dez/2026
 * ────────────────────────────────────────────────────────────────────────────
 * Roda dentro da própria planilha (Extensões › Apps Script › executar
 * carregarPlanoSetDez). Não passa pelo Web App: o doPost se recusa a escrever
 * na aba PLANO MESTRE de propósito, e essa proteção continua valendo — quem
 * escreve aqui é o planejamento, rodando o script à mão, não o dashboard.
 *
 * O que faz:
 *   1. Insere os lotes 032-26 a 046-26 acima da linha TOTAL GERAL;
 *   2. Completa a linha TOTAL GERAL nos meses set./26 a dez./26;
 *   3. Troca a coluna TOTAL ANO por =SOMA() da linha — é ela que já saiu de
 *      sincronia duas vezes por ser valor digitado;
 *   4. Valida coluna a coluna e aborta se algo não fechar.
 *
 * É idempotente: se o lote 032-26 já existir, não faz nada e avisa.
 */

var CARGA_ABA   = 'PLANO MESTRE';
var CARGA_TOTAL = 'TOTAL GERAL';

/* Colunas da matriz: A = código do lote, B..M = jan..dez, N = TOTAL ANO. */
var CARGA_COL_LOTE  = 1;
var CARGA_COL_JAN   = 2;
var CARGA_COL_DEZ   = 13;
var CARGA_COL_TOTAL = 14;

/* Célula vazia na matriz é traço, não zero — mesmo padrão das linhas atuais. */
var CARGA_VAZIO = '-';

/* Lotes novos: código + quantidade por mês (1 = jan ... 12 = dez). Os lotes
   que aparecem em dois meses seguidos são intencionais — seguem o padrão da
   planilha, em que o último lote do mês vira para o mês seguinte. */
var CARGA_LOTES = [
  { lote: '032-26', meses: { 9: 9000 } },
  { lote: '033-26', meses: { 9: 8500 } },
  { lote: '034-26', meses: { 9: 9000 } },
  { lote: '035-26', meses: { 9: 8310, 10: 1500 } },
  { lote: '036-26', meses: { 10: 8500 } },
  { lote: '037-26', meses: { 10: 8500 } },
  { lote: '038-26', meses: { 10: 8500 } },
  { lote: '039-26', meses: { 10: 7810, 11: 1200 } },
  { lote: '040-26', meses: { 11: 7500 } },
  { lote: '041-26', meses: { 11: 7500 } },
  { lote: '042-26', meses: { 11: 8000 } },
  { lote: '043-26', meses: { 11: 7290, 12: 1000 } },
  { lote: '044-26', meses: { 12: 8000 } },
  { lote: '045-26', meses: { 12: 8000 } },
  { lote: '046-26', meses: { 12: 7860 } }
];

/* Totais esperados do TOTAL GERAL nos meses que estavam vazios. Ficam aqui
   explícitos para a validação ter contra o que conferir — se a soma dos lotes
   não bater com isso, alguém mexeu em um dos dois lados. */
var CARGA_TOTAL_GERAL = { 9: 34810, 10: 34810, 11: 31490, 12: 24860 };


/* ══ ENTRADA ══ */
function carregarPlanoSetDez() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(CARGA_ABA);
  if (!aba) return cargaErro('Aba "' + CARGA_ABA + '" não encontrada.');

  var mapa = cargaLocalizar(aba);
  if (mapa.erro) return cargaErro(mapa.erro);

  /* Guarda de reexecução: rodar duas vezes duplicaria os 15 lotes e o total
     do ano dobraria sem ninguém perceber. */
  var jaTem = cargaLotesExistentes(aba, mapa);
  if (jaTem.length) {
    return cargaErro('Estes lotes já estão na planilha: ' + jaTem.join(', ') +
                     '. Nada foi alterado — apague-os antes de rodar de novo.');
  }

  cargaInserirLotes(aba, mapa);
  /* A linha TOTAL GERAL desceu 15 posições com a inserção. */
  mapa.linhaTotal += CARGA_LOTES.length;

  cargaEscreverTotalGeral(aba, mapa);
  cargaFormularTotalAno(aba, mapa);
  SpreadsheetApp.flush();

  var rel = validarPlanoMestre();
  cargaAvisar(rel.ok
    ? 'Plano carregado: ' + CARGA_LOTES.length + ' lotes, ano fecha em ' +
      cargaMil(rel.totalAno) + ' peças.'
    : 'Carregado, MAS a validação acusou problema:\n' + rel.problemas.join('\n'));
  return rel;
}


/* ══ LOCALIZAÇÃO ══
   Nada de número de linha fixo: a matriz já mudou de tamanho antes e vai mudar
   de novo a cada lote lançado. */
function cargaLocalizar(aba) {
  var valores = aba.getDataRange().getValues();
  var cab = -1, total = -1;

  for (var i = 0; i < valores.length; i++) {
    var a = cargaNormaliza(valores[i][CARGA_COL_LOTE - 1]);
    if (cab < 0 && a === 'aba') cab = i + 1;
    if (cab > 0 && total < 0 && a === cargaNormaliza(CARGA_TOTAL)) total = i + 1;
  }

  if (cab < 0)   return { erro: 'Linha de cabeçalho (coluna A = "ABA") não encontrada.' };
  if (total < 0) return { erro: 'Linha "' + CARGA_TOTAL + '" não encontrada.' };
  if (total <= cab + 1) return { erro: 'Matriz sem linhas de lote entre o cabeçalho e o total.' };

  return { linhaCab: cab, linhaTotal: total, primeiroLote: cab + 1 };
}

function cargaLotesExistentes(aba, mapa) {
  var qtd = mapa.linhaTotal - mapa.primeiroLote;
  var col = aba.getRange(mapa.primeiroLote, CARGA_COL_LOTE, qtd, 1).getValues();
  var atuais = {};
  col.forEach(function (l) { atuais[cargaNormaliza(l[0])] = true; });
  return CARGA_LOTES
    .filter(function (x) { return atuais[cargaNormaliza(x.lote)]; })
    .map(function (x) { return x.lote; });
}


/* ══ ESCRITA ══ */
function cargaInserirLotes(aba, mapa) {
  var n = CARGA_LOTES.length;
  /* insertRowsBefore herda o formato da linha de cima — que é um lote, então
     borda, fonte e alinhamento vêm certos sem precisar copiar à mão. */
  aba.insertRowsBefore(mapa.linhaTotal, n);

  var bloco = CARGA_LOTES.map(function (x) {
    var linha = [x.lote];
    for (var m = 1; m <= 12; m++) {
      linha.push(x.meses[m] !== undefined ? x.meses[m] : CARGA_VAZIO);
    }
    linha.push('');   /* TOTAL ANO entra como fórmula logo abaixo */
    return linha;
  });

  aba.getRange(mapa.linhaTotal, CARGA_COL_LOTE, n, CARGA_COL_TOTAL).setValues(bloco);
}

function cargaEscreverTotalGeral(aba, mapa) {
  /* A linha inteira de meses passa a ser soma da coluna: enquanto for valor
     digitado, ela volta a divergir dos lotes na primeira edição. */
  var ini = mapa.primeiroLote, fim = mapa.linhaTotal - 1;
  var f = [];
  for (var c = CARGA_COL_JAN; c <= CARGA_COL_DEZ; c++) {
    var letra = cargaLetra(c);
    f.push('=SOMA(' + letra + ini + ':' + letra + fim + ')');
  }
  aba.getRange(mapa.linhaTotal, CARGA_COL_JAN, 1, f.length).setFormulas([f]);
}

function cargaFormularTotalAno(aba, mapa) {
  var jan = cargaLetra(CARGA_COL_JAN), dez = cargaLetra(CARGA_COL_DEZ);
  var qtd = mapa.linhaTotal - mapa.primeiroLote + 1;   /* lotes + TOTAL GERAL */
  var f = [];
  for (var i = 0; i < qtd; i++) {
    var l = mapa.primeiroLote + i;
    f.push(['=SOMA(' + jan + l + ':' + dez + l + ')']);
  }
  aba.getRange(mapa.primeiroLote, CARGA_COL_TOTAL, qtd, 1).setFormulas(f);
}


/* ══ VALIDAÇÃO ══
   Roda sozinha depois da carga, mas serve para chamar a qualquer momento —
   é o teste que pega o TOTAL ANO fora de sincronia. */
function validarPlanoMestre() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CARGA_ABA);
  if (!aba) return { ok: false, problemas: ['Aba "' + CARGA_ABA + '" não encontrada.'] };

  var mapa = cargaLocalizar(aba);
  if (mapa.erro) return { ok: false, problemas: [mapa.erro] };

  var qtd  = mapa.linhaTotal - mapa.primeiroLote;
  var cab  = aba.getRange(mapa.linhaCab,      CARGA_COL_JAN, 1, 12).getDisplayValues()[0];
  var mat  = aba.getRange(mapa.primeiroLote,  CARGA_COL_JAN, qtd, 12).getValues();
  var tot  = aba.getRange(mapa.linhaTotal,    CARGA_COL_JAN, 1, 12).getValues()[0];
  var totAno = aba.getRange(mapa.linhaTotal, CARGA_COL_TOTAL).getValue();

  var problemas = [], somaMeses = 0;

  for (var c = 0; c < 12; c++) {
    var s = 0;
    for (var r = 0; r < qtd; r++) s += cargaNum(mat[r][c]);
    var t = cargaNum(tot[c]);
    somaMeses += t;
    if (s !== t) {
      problemas.push('· ' + (cab[c] || 'mês ' + (c + 1)) + ': lotes somam ' +
                     cargaMil(s) + ', TOTAL GERAL diz ' + cargaMil(t) +
                     ' (diferença ' + cargaMil(s - t) + ')');
    }
    /* Só cobra os meses recém-carregados contra o valor esperado. */
    var esp = CARGA_TOTAL_GERAL[c + 1];
    if (esp !== undefined && t !== esp) {
      problemas.push('· ' + (cab[c] || 'mês ' + (c + 1)) + ': esperado ' +
                     cargaMil(esp) + ', encontrado ' + cargaMil(t));
    }
  }

  if (cargaNum(totAno) !== somaMeses) {
    problemas.push('· TOTAL ANO (' + cargaMil(cargaNum(totAno)) +
                   ') difere da soma dos 12 meses (' + cargaMil(somaMeses) + ')');
  }

  var rel = { ok: !problemas.length, problemas: problemas, totalAno: somaMeses };
  Logger.log(rel.ok ? 'Validação OK — ano fecha em ' + cargaMil(somaMeses) + ' peças.'
                    : 'Validação falhou:\n' + problemas.join('\n'));
  return rel;
}


/* ══ HELPERS ══ */
function cargaNormaliza(v) {
  return String(v === null || v === undefined ? '' : v)
    .trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

/* Traço, vazio e texto viram 0; o resto vira número. Mesmo tratamento que o
   Code.gs dá ao ler a planilha. */
function cargaNum(v) {
  if (typeof v === 'number') return v;
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s || s === CARGA_VAZIO) return 0;
  s = s.replace(/\./g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function cargaLetra(col) {
  var s = '';
  while (col > 0) {
    var r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = (col - 1 - r) / 26;
  }
  return s;
}

function cargaMil(n) {
  return Math.round(n).toLocaleString('pt-BR');
}

function cargaErro(msg) {
  Logger.log('ERRO: ' + msg);
  cargaAvisar(msg);
  return { ok: false, problemas: [msg] };
}

/* Alert só existe com UI aberta; rodando pelo editor cai no log. */
function cargaAvisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
