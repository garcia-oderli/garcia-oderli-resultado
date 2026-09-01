/**
 * CURVA ABC — casamento dos meses por produto
 * ────────────────────────────────────────────────────────────────────────────
 * Script da planilha CURVA ABC (Extensões → Apps Script DELA, não do painel).
 *
 * A curva ABC do Lógica sai um mês por aba (JAN-26, FEV-26… colada do PDF,
 * colunas Cod.Prod. | Descrição Produto | Vlr.Vendas | Quantidade), cada uma
 * ordenada pela venda do próprio mês — o mesmo produto muda de linha a cada
 * aba e comparar meses vira caça ao código. Este script faz o casamento:
 *
 *   casarMeses() varre toda aba cujo nome é mês (JAN-26, AGO 26, AGOS 26…)
 *   e o cabeçalho começa com "Cod.Prod.", junta tudo por código e regrava a
 *   aba CONSOLIDADO: uma linha por produto, meses lado a lado (R$ e QTD),
 *   total do período, % acumulado e classe ABC no corte do próprio relatório
 *   (A = 70% do valor, B = +20%, C = +10%). A coluna FLAG vem da aba EXC,
 *   casada pelo mesmo código. Rodar de novo substitui a CONSOLIDADO inteira
 *   — colou um mês novo, é só rodar de novo.
 *
 *   De quebra confere cada aba: se a soma das linhas não bate com a linha
 *   "Valores Totais" do relatório, avisa — é o sinal de que o colar comeu
 *   linha (MAI-26 está assim: R$ 114,34 e 1 unidade a menos).
 *
 * O código chega como número (114011001) nas abas de mês e na EXC; o
 * casamento normaliza tudo para 000.000.000 antes de comparar.
 */

var CA_ABA_SAIDA = 'CONSOLIDADO';
var CA_ABA_FLAGS = 'EXC';
var CA_MESES = { JAN:1, FEV:2, MAR:3, ABR:4, MAI:5, JUN:6, JUL:7, AGO:8, SET:9, OUT:10, NOV:11, DEZ:12 };
/* Cortes do próprio relatório do Lógica (Classe ABC: 70,00 / 20,00 / 10,00). */
var CA_CORTE_A = 0.70, CA_CORTE_B = 0.90;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📈 Curva ABC')
    .addItem('Casar meses (gerar CONSOLIDADO)', 'casarMeses')
    .addToUi();
}

function casarMeses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var meses = [], produtos = {}, avisos = [];

  ss.getSheets().forEach(function (aba) {
    var nome = aba.getName().trim().toUpperCase();
    var m = nome.match(/^([A-ZÇ]{3})[^0-9]*(\d{2,4})$/);
    if (!m || !CA_MESES[m[1]]) return;
    if (String(aba.getRange(1, 1).getValue() || '').trim().toLowerCase().indexOf('cod') !== 0) return;
    var ano = parseInt(m[2], 10);
    if (ano < 100) ano += 2000;

    var dados = aba.getRange(2, 1, Math.max(aba.getLastRow() - 1, 1), 4).getValues();
    var somaRS = 0, totalRelatorio = null;
    dados.forEach(function (l) {
      var desc = String(l[1] || '').trim();
      if (/valores\s+totais/i.test(desc)) { totalRelatorio = caNum(l[2]); return; }
      var cod = caCodigo(l[0]);
      if (!cod || !desc) return;
      var rs = caNum(l[2]), qtd = caNum(l[3]);
      var p = produtos[cod] || (produtos[cod] = { desc: '', meses: {} });
      /* o PDF trunca a descrição em larguras diferentes — fica a mais longa */
      if (desc.length > p.desc.length) p.desc = desc;
      var mm = p.meses[nome] || (p.meses[nome] = { rs: 0, qtd: 0 });
      mm.rs += rs; mm.qtd += qtd;
      somaRS += rs;
    });
    meses.push({ chave: nome, ano: ano, mes: CA_MESES[m[1]] });
    if (totalRelatorio !== null && Math.abs(totalRelatorio - somaRS) > 0.01) {
      avisos.push(nome + ': as linhas somam R$ ' + somaRS.toFixed(2) + ' mas o "Valores Totais" do relatório diz R$ '
        + totalRelatorio.toFixed(2) + ' — o colar perdeu linha; confira contra o PDF.');
    }
  });
  if (!meses.length) {
    return caErro('Nenhuma aba de mês encontrada — o nome precisa ser mês (JAN-26, AGO 26…) e o cabeçalho começar com "Cod.Prod.".');
  }
  meses.sort(function (a, b) { return a.ano - b.ano || a.mes - b.mes; });

  /* FLAG (EXC) casada pelo mesmo código normalizado. */
  var flags = {}, abaFlag = ss.getSheetByName(CA_ABA_FLAGS);
  if (abaFlag) {
    abaFlag.getDataRange().getValues().forEach(function (l, i) {
      var cod = i && caCodigo(l[0]);
      if (cod) flags[cod] = String(l[1] || '').trim();
    });
  }

  var codigos = Object.keys(produtos), totalGeral = 0, totalQtd = 0;
  codigos.forEach(function (c) {
    var p = produtos[c];
    p.totRS = 0; p.totQtd = 0;
    meses.forEach(function (m) {
      var v = p.meses[m.chave];
      if (v) { p.totRS += v.rs; p.totQtd += v.qtd; }
    });
    totalGeral += p.totRS; totalQtd += p.totQtd;
  });
  codigos.sort(function (a, b) { return produtos[b].totRS - produtos[a].totRS; });

  var cab = ['CODIGO', 'DESCRICAO', 'FLAG', 'CLASSE', '% VENDA', '% ACUM', 'TOTAL R$', 'TOTAL QTD'];
  var fmt = ['@', '@', '@', '@', '0.0%', '0.0%', '#,##0.00', '#,##0'];
  meses.forEach(function (m) {
    cab.push('R$ ' + m.chave, 'QTD ' + m.chave);
    fmt.push('#,##0.00', '#,##0');
  });

  var linhas = [], acum = 0, porClasse = { A: 0, B: 0, C: 0 }, comFlag = 0;
  codigos.forEach(function (c) {
    var p = produtos[c];
    acum += p.totRS;
    var pAcum = totalGeral ? acum / totalGeral : 0;
    var classe = pAcum <= CA_CORTE_A ? 'A' : (pAcum <= CA_CORTE_B ? 'B' : 'C');
    porClasse[classe]++;
    if (flags[c]) comFlag++;
    var linha = [c, p.desc, flags[c] || '', classe,
                 totalGeral ? p.totRS / totalGeral : 0, pAcum, p.totRS, p.totQtd];
    meses.forEach(function (m) {
      var v = p.meses[m.chave];
      linha.push(v ? v.rs : '', v ? v.qtd : '');
    });
    linhas.push(linha);
  });

  var totalLinha = ['', 'TOTAL', '', '', '', '', totalGeral, totalQtd];
  meses.forEach(function (m) {
    var rs = 0, qtd = 0;
    codigos.forEach(function (c) {
      var v = produtos[c].meses[m.chave];
      if (v) { rs += v.rs; qtd += v.qtd; }
    });
    totalLinha.push(rs, qtd);
  });
  linhas.push(totalLinha);

  var aba = ss.getSheetByName(CA_ABA_SAIDA) || ss.insertSheet(CA_ABA_SAIDA);
  aba.clear();
  aba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  aba.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
  aba.getRange(2, 1, linhas.length, cab.length).setNumberFormats(linhas.map(function () { return fmt; }));
  aba.getRange(linhas.length + 1, 1, 1, cab.length).setFontWeight('bold');
  aba.setFrozenRows(1);
  aba.setFrozenColumns(4);

  /* Aviso de casamento reverso: código na EXC que não vendeu em mês nenhum. */
  var excSemVenda = Object.keys(flags).filter(function (c) { return !produtos[c]; }).length;

  caAvisar('Casado: ' + codigos.length + ' produtos × ' + meses.length + ' meses ('
    + meses[0].chave + ' a ' + meses[meses.length - 1].chave + ') na aba ' + CA_ABA_SAIDA + '. '
    + 'Classes: A=' + porClasse.A + ', B=' + porClasse.B + ', C=' + porClasse.C + ' (corte 70/20/10 do relatório). '
    + comFlag + ' produtos com flag da ' + CA_ABA_FLAGS
    + (excSemVenda ? '; ' + excSemVenda + ' códigos da ' + CA_ABA_FLAGS + ' sem venda no período.' : '.')
    + (avisos.length ? '\n\nATENÇÃO:\n' + avisos.join('\n') : ''));
}

/* ══ HELPERS ══ */

/* Código chega como número (114011001), texto com pontos ou colado com lixo;
   normaliza tudo para 000.000.000 — é essa chave que casa mês, EXC e tudo. */
function caCodigo(v) {
  if (v === null || v === undefined || v === '') return null;
  var s = String(typeof v === 'number' ? Math.round(v) : v).replace(/\D/g, '');
  return s.length === 9 ? s.slice(0, 3) + '.' + s.slice(3, 6) + '.' + s.slice(6) : null;
}

function caNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s || s === '-') return 0;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function caErro(msg) { Logger.log('ERRO: ' + msg); caAvisar(msg); }

function caAvisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
