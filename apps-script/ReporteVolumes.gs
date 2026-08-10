/**
 * REPORTE DE VOLUMES — do relatório mensal do ERP para a HISTORICO
 * ────────────────────────────────────────────────────────────────────────────
 * O ERP (Lógica) reporta a produção por SKU de VOLUME ("VOL 1/2 PENTEADEIRA
 * CAMARIM…", transação 3 - REPORTE), e a aba PRODUTO_CODIGO cadastra um SKU
 * por volume. Este script cruza os dois: cola-se o relatório do mês na aba
 * REPORTE_VOLUMES e ele calcula volumes, produtos e o fator de embalagem —
 * é esse número que alimenta a coluna volumesProduzidos da HISTORICO, que o
 * dashboard usa para converter o plano (em produtos) para volumes.
 *
 * Roda dentro da própria planilha (Extensões › Apps Script), como a carga do
 * plano: quem escreve é o planejamento, à mão, não o dashboard.
 *
 * COMO USAR:
 *   1. Rode criarAbaReporteVolumes() uma vez — monta a aba REPORTE_VOLUMES;
 *   2. Cole as linhas do relatório do mês: MES | ANO | CODIGO | QUANTIDADE
 *      (o código pode vir com pontos, "501.061.001" — tanto faz);
 *   3. Rode calcularVolumesMes() — escreve o resumo por mês na própria aba
 *      (volumes, produtos, fator, códigos não encontrados no cadastro);
 *   4. Confira o resumo e rode lancarVolumesNaHistorico() para gravar na
 *      coluna volumesProduzidos, casando mês+ano. Só grava onde o resumo
 *      tem número; não apaga nada.
 *
 * ATENÇÃO — unidade do producaoReal: enquanto não se confirmar (com o
 * relatório SEM filtro de tipo) se a produção real da HISTORICO já conta os
 * VOLs um a um ou conta produto fechado, o fator do dashboard é aproximação.
 * O resumo daqui mostra os dois números (volumes e produtos) justamente para
 * essa conferência: compare o total com o producaoReal do mês.
 */

var RV_ABA         = 'REPORTE_VOLUMES';
var RV_ABA_CADASTRO = 'PRODUTO_CODIGO';
var RV_ABA_HISTORICO = 'HISTORICO';
/* Onde o resumo por mês é escrito dentro da REPORTE_VOLUMES (colunas F..K). */
var RV_COL_RESUMO  = 6;

var RV_MESES = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

/* ══ 1 · ABA DE ENTRADA ══ */
function criarAbaReporteVolumes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(RV_ABA)) return rvAvisar('Aba ' + RV_ABA + ' já existe — nada foi alterado.');
  var aba = ss.insertSheet(RV_ABA);
  aba.getRange(1, 1, 1, 4).setValues([['MES', 'ANO', 'CODIGO', 'QUANTIDADE']]).setFontWeight('bold');
  aba.getRange(1, RV_COL_RESUMO, 1, 6).setValues([['RESUMO: MES', 'ANO', 'VOLUMES', 'PRODUTOS', 'FATOR', 'PENDENCIAS']]).setFontWeight('bold');
  aba.setFrozenRows(1);
  rvAvisar('Aba ' + RV_ABA + ' criada. Cole o relatório do mês (MES, ANO, CODIGO, QUANTIDADE) e rode calcularVolumesMes().');
}

/* ══ 2 · CADASTRO ══
   PRODUTO_CODIGO: coluna A = código do SKU de volume, B = descrição
   "VOL 1/2 NOME DO PRODUTO". O nome sem o prefixo identifica o produto;
   descrição sem prefixo VOL é tratada como produto de volume único. */
function rvLerCadastro(ss) {
  var aba = ss.getSheetByName(RV_ABA_CADASTRO);
  if (!aba) throw new Error('Aba "' + RV_ABA_CADASTRO + '" não encontrada.');
  var linhas = aba.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < linhas.length; i++) {
    var cod = rvCodigo(linhas[i][0]);
    var desc = String(linhas[i][1] || '').trim();
    if (!cod || !desc) continue;
    var m = desc.match(/^VOL\.?\s*0?(\d+)\s*\/\s*0?(\d+)\s+(.+)$/i);
    mapa[cod] = m
      ? { produto: m[3].trim(), vol: parseInt(m[1], 10), de: parseInt(m[2], 10) }
      : { produto: desc, vol: 1, de: 1 };
  }
  if (!Object.keys(mapa).length) throw new Error('Nenhum código lido em ' + RV_ABA_CADASTRO + '.');
  return mapa;
}

/* ══ 3 · CÁLCULO DO MÊS ══ */
function calcularVolumesMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(RV_ABA);
  if (!aba) return rvErro('Aba "' + RV_ABA + '" não encontrada — rode criarAbaReporteVolumes().');
  var cadastro = rvLerCadastro(ss);

  var linhas = aba.getRange(2, 1, Math.max(aba.getLastRow() - 1, 1), 4).getValues();
  /* chave mes/ano → { vol, porProduto: {nome → maior qtde entre os VOLs},
     pendentes: códigos fora do cadastro (contam como 1 volume = 1 produto,
     mas ficam listados para o cadastro ser completado) */
  var meses = {}, ordem = [];
  linhas.forEach(function (l) {
    var mes = String(l[0] || '').trim().toUpperCase().slice(0, 3);
    var ano = parseInt(l[1], 10);
    var cod = rvCodigo(l[2]);
    var qtd = rvNum(l[3]);
    if (RV_MESES.indexOf(mes) < 0 || !ano || !cod || !qtd) return;
    var chave = mes + '/' + ano;
    if (!meses[chave]) { meses[chave] = { mes: mes, ano: ano, vol: 0, porProduto: {}, pendentes: {} }; ordem.push(chave); }
    var g = meses[chave];
    g.vol += qtd;
    var cad = cadastro[cod];
    if (cad) {
      g.porProduto[cad.produto] = Math.max(g.porProduto[cad.produto] || 0, qtd);
    } else {
      g.pendentes[cod] = true;
      g.porProduto['?' + cod] = qtd;   /* sem cadastro: 1 volume = 1 produto */
    }
  });
  if (!ordem.length) return rvErro('Nenhuma linha válida em ' + RV_ABA + ' (MES, ANO, CODIGO, QUANTIDADE).');

  var saida = ordem.map(function (chave) {
    var g = meses[chave];
    var prod = 0;
    Object.keys(g.porProduto).forEach(function (p) { prod += g.porProduto[p]; });
    var pend = Object.keys(g.pendentes);
    return [g.mes, g.ano, g.vol, prod, prod > 0 ? g.vol / prod : '',
            pend.length ? pend.length + ' código(s) fora do cadastro: ' + pend.join(', ') : ''];
  });

  /* limpa o resumo anterior antes de escrever o novo */
  var alt = Math.max(aba.getLastRow() - 1, 1);
  aba.getRange(2, RV_COL_RESUMO, alt, 6).clearContent();
  aba.getRange(2, RV_COL_RESUMO, saida.length, 6).setValues(saida);
  rvAvisar('Resumo calculado para ' + saida.length + ' mês(es). Confira o total de PRODUTOS contra o producaoReal '
    + 'da HISTORICO — é essa comparação que diz em que unidade a produção real está. '
    + 'Depois rode lancarVolumesNaHistorico().');
}

/* ══ 4 · LANÇAMENTO NA HISTORICO ══
   Grava a coluna volumesProduzidos casando mês+ano com o resumo. A coluna é
   a que o planejamento criou; se não existir, o script avisa em vez de criar
   sozinho — coluna nova na HISTORICO é decisão de quem mantém a planilha. */
function lancarVolumesNaHistorico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaRV = ss.getSheetByName(RV_ABA);
  var aba = ss.getSheetByName(RV_ABA_HISTORICO);
  if (!abaRV) return rvErro('Aba "' + RV_ABA + '" não encontrada.');
  if (!aba)   return rvErro('Aba "' + RV_ABA_HISTORICO + '" não encontrada.');

  var resumo = abaRV.getRange(2, RV_COL_RESUMO, Math.max(abaRV.getLastRow() - 1, 1), 3).getValues()
    .filter(function (l) { return l[0] && l[1] && rvNum(l[2]) > 0; });
  if (!resumo.length) return rvErro('Resumo vazio — rode calcularVolumesMes() primeiro.');

  var linhas = aba.getDataRange().getValues();
  var iCab = -1, col = {};
  for (var i = 0; i < Math.min(linhas.length, 20) && iCab < 0; i++) {
    var norm = linhas[i].map(rvNormaliza);
    if (norm.indexOf('mes') >= 0 && norm.indexOf('ano') >= 0) {
      iCab = i;
      norm.forEach(function (n, c) {
        if (n === 'mes') col.mes = c;
        else if (n === 'ano') col.ano = c;
        else if (n.indexOf('volumes') === 0 && col.vol === undefined) col.vol = c;
      });
    }
  }
  if (iCab < 0) return rvErro('Cabeçalho com "mes" e "ano" não encontrado na ' + RV_ABA_HISTORICO + '.');
  if (col.vol === undefined) {
    return rvErro('Coluna de volumes não encontrada na ' + RV_ABA_HISTORICO + ' — crie o cabeçalho '
      + '"volumesProduzidos" na primeira coluna livre e rode de novo.');
  }

  var mapa = {};
  for (var r = iCab + 1; r < linhas.length; r++) {
    var m = String(linhas[r][col.mes] || '').trim().toUpperCase().slice(0, 3);
    var a = parseInt(linhas[r][col.ano], 10);
    if (RV_MESES.indexOf(m) >= 0 && a) mapa[m + '/' + a] = r + 1;   /* linha na planilha */
  }

  var gravados = [], semLinha = [];
  resumo.forEach(function (l) {
    var chave = String(l[0]).trim().toUpperCase().slice(0, 3) + '/' + parseInt(l[1], 10);
    var linha = mapa[chave];
    if (!linha) { semLinha.push(chave); return; }
    aba.getRange(linha, col.vol + 1).setValue(rvNum(l[2]));
    gravados.push(chave + '=' + rvNum(l[2]));
  });
  rvAvisar('Lançado na ' + RV_ABA_HISTORICO + ': ' + (gravados.join(', ') || 'nada')
    + (semLinha.length ? '. SEM linha na HISTORICO (mês ainda não existe lá): ' + semLinha.join(', ') : '')
    + '. O dashboard pega no próximo sync.');
}

/* ══ HELPERS ══ */
function rvCodigo(v) { return String(v === null || v === undefined ? '' : v).replace(/\D/g, ''); }

function rvNum(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s || s === '-') return 0;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function rvNormaliza(v) {
  return String(v === null || v === undefined ? '' : v)
    .trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function rvErro(msg) { Logger.log('ERRO: ' + msg); rvAvisar(msg); }

function rvAvisar(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
