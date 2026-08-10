/**
 * REPORTE DE VOLUMES — do relatório mensal do ERP para a HISTORICO
 * ────────────────────────────────────────────────────────────────────────────
 * O relatório "Mensal por Transação (3 - REPORTE), Tipo: Todos" do Lógica
 * lista a produção DUAS vezes: uma linha por produto acabado e uma linha por
 * volume do produto ("VOL 1/2 …", inclusive VOL 1/1 dos de caixa única) —
 * conferido em JAN/26: 49.209 no total = 24.491 produtos + 24.718 volumes.
 * Por isso o total geral do relatório não serve direto: é preciso separar as
 * duas famílias. Este script faz essa separação e devolve, por mês:
 *
 *   VOLUMES  = linhas VOL + produtos sem SKU de volume (1 caixa = 1 volume,
 *              ou × volumes do cadastro quando a PRODUTO_CODIGO conhecer);
 *   PRODUTOS = linhas de produto acabado;
 *   FATOR    = volumes ÷ produtos — mesma fonte, mesma competência.
 *
 * É esse VOLUMES que alimenta a coluna volumesProduzidos da HISTORICO; o
 * PRODUTOS pode ir para uma coluna produtosReportados (opcional) — com ela o
 * dashboard calcula o fator na base coerente em vez de dividir pelo
 * producaoReal, que vem de outro corte (~9% menor que o relatório em JAN/26).
 *
 * COMO USAR:
 *   1. criarAbaReporteVolumes() — uma vez; monta a aba REPORTE_VOLUMES;
 *   2. Cole o relatório do mês: MES | ANO | CODIGO | DESCRICAO | QUANTIDADE.
 *      A DESCRICAO é obrigatória: é ela que diz o que é volume ("VOL x/y…")
 *      e o que é produto. Código com pontos ("501.061.001") funciona;
 *   3. calcularVolumesMes() — escreve o resumo por mês na própria aba e
 *      lista os produtos sem SKU de volume (pendências de cadastro);
 *   4. Confira e rode lancarVolumesNaHistorico() — grava volumesProduzidos
 *      (e produtosReportados, se a coluna existir) casando mês+ano.
 *
 * Roda dentro da própria planilha, à mão, como a carga do plano — o Web App
 * não escreve nessas colunas.
 */

var RV_ABA           = 'REPORTE_VOLUMES';
var RV_ABA_CADASTRO  = 'PRODUTO_CODIGO';
var RV_ABA_HISTORICO = 'HISTORICO';
/* Onde o resumo por mês é escrito dentro da REPORTE_VOLUMES (colunas G..L). */
var RV_COL_RESUMO = 7;

var RV_MESES = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

/* "VOL 1/2 NOME", "VOL. 01/03 NOME"… → captura número, total e nome. */
var RV_RE_VOL = /^VOL\.?\s*0?(\d+)\s*\/\s*0?(\d+)\s+(.+)$/i;

/* ══ 1 · ABA DE ENTRADA ══ */
function criarAbaReporteVolumes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(RV_ABA)) return rvAvisar('Aba ' + RV_ABA + ' já existe — nada foi alterado.');
  var aba = ss.insertSheet(RV_ABA);
  aba.getRange(1, 1, 1, 5).setValues([['MES', 'ANO', 'CODIGO', 'DESCRICAO', 'QUANTIDADE']]).setFontWeight('bold');
  aba.getRange(1, RV_COL_RESUMO, 1, 6).setValues([['RESUMO: MES', 'ANO', 'VOLUMES', 'PRODUTOS', 'FATOR', 'PENDENCIAS']]).setFontWeight('bold');
  aba.setFrozenRows(1);
  rvAvisar('Aba ' + RV_ABA + ' criada. Cole o relatório do mês (MES, ANO, CODIGO, DESCRICAO, QUANTIDADE) e rode calcularVolumesMes().');
}

/* ══ 2 · CADASTRO ══
   PRODUTO_CODIGO: A = código do SKU de volume, B = descrição "VOL 1/2 NOME".
   Devolve nome do produto → quantos volumes ele tem. Usado só para o produto
   que aparecer no mês SEM linhas de volume: aí a conta usa o cadastro em vez
   de assumir caixa única. A aba não existir não é erro — vira aproximação. */
function rvLerCadastro(ss) {
  var aba = ss.getSheetByName(RV_ABA_CADASTRO);
  var porProduto = {};
  if (!aba) return porProduto;
  var linhas = aba.getDataRange().getValues();
  for (var i = 1; i < linhas.length; i++) {
    var m = String(linhas[i][1] || '').trim().match(RV_RE_VOL);
    if (!m) continue;
    var nome = m[3].trim().toUpperCase();
    porProduto[nome] = Math.max(porProduto[nome] || 0, parseInt(m[2], 10));
  }
  return porProduto;
}

/* ══ 3 · CÁLCULO DO MÊS ══ */
function calcularVolumesMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(RV_ABA);
  if (!aba) return rvErro('Aba "' + RV_ABA + '" não encontrada — rode criarAbaReporteVolumes().');
  var cadastro = rvLerCadastro(ss);

  var linhas = aba.getRange(2, 1, Math.max(aba.getLastRow() - 1, 1), 5).getValues();
  var meses = {}, ordem = [];
  linhas.forEach(function (l) {
    var mes = String(l[0] || '').trim().toUpperCase().slice(0, 3);
    var ano = parseInt(l[1], 10);
    var desc = String(l[3] || '').trim();
    var qtd = rvNum(l[4]);
    if (RV_MESES.indexOf(mes) < 0 || !ano || !desc || !qtd) return;
    var chave = mes + '/' + ano;
    if (!meses[chave]) {
      meses[chave] = { mes: mes, ano: ano, vol: 0, prod: 0, nomesVol: [], produtos: [] };
      ordem.push(chave);
    }
    var g = meses[chave], m = desc.match(RV_RE_VOL);
    if (m) { g.vol += qtd; g.nomesVol.push(m[3].trim().toUpperCase()); }
    else   { g.prod += qtd; g.produtos.push({ nome: desc.toUpperCase(), qtd: qtd }); }
  });
  if (!ordem.length) return rvErro('Nenhuma linha válida em ' + RV_ABA + ' (MES, ANO, CODIGO, DESCRICAO, QUANTIDADE).');

  var saida = ordem.map(function (chave) {
    var g = meses[chave];
    /* Produto espelhado = algum nome de VOL do mês é prefixo do nome dele
       (as descrições do relatório truncam, e a linha VOL trunca mais cedo
       por causa do prefixo "VOL x/y "). Espelhado: volumes já contados nas
       linhas VOL. Sem espelho: usa o cadastro; sem cadastro, 1 caixa. */
    var pend = {};
    g.produtos.forEach(function (p) {
      var espelhado = g.nomesVol.some(function (nv) { return p.nome.indexOf(nv) === 0; });
      if (espelhado) return;
      var nVols = 0;
      Object.keys(cadastro).some(function (nc) {
        if (p.nome.indexOf(nc) === 0) { nVols = cadastro[nc]; return true; }
        return false;
      });
      if (nVols > 0) { g.vol += p.qtd * nVols; }
      else           { g.vol += p.qtd; pend[p.nome] = true; }
    });
    var pendentes = Object.keys(pend);
    return [g.mes, g.ano, g.vol, g.prod, g.prod > 0 ? g.vol / g.prod : '',
            pendentes.length
              ? pendentes.length + ' produto(s) sem SKU de volume (contados como 1 caixa): ' + pendentes.slice(0, 8).join('; ')
                + (pendentes.length > 8 ? '…' : '')
              : ''];
  });

  var alt = Math.max(aba.getLastRow() - 1, 1);
  aba.getRange(2, RV_COL_RESUMO, alt, 6).clearContent();
  aba.getRange(2, RV_COL_RESUMO, saida.length, 6).setValues(saida);
  rvAvisar('Resumo calculado para ' + saida.length + ' mês(es). Confira VOLUMES e PRODUTOS e rode lancarVolumesNaHistorico(). '
    + 'Produto listado como pendência: cadastre os SKUs de volume dele na ' + RV_ABA_CADASTRO + ' para a conta ficar exata.');
}

/* ══ 4 · LANÇAMENTO NA HISTORICO ══
   Grava volumesProduzidos casando mês+ano; se existir uma coluna cujo
   cabeçalho comece com "produtosReport", grava também os produtos do
   relatório — é com ela que o dashboard calcula o fator na base coerente.
   Não cria coluna sozinho e não apaga nada. */
function lancarVolumesNaHistorico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaRV = ss.getSheetByName(RV_ABA);
  var aba = ss.getSheetByName(RV_ABA_HISTORICO);
  if (!abaRV) return rvErro('Aba "' + RV_ABA + '" não encontrada.');
  if (!aba)   return rvErro('Aba "' + RV_ABA_HISTORICO + '" não encontrada.');

  var resumo = abaRV.getRange(2, RV_COL_RESUMO, Math.max(abaRV.getLastRow() - 1, 1), 4).getValues()
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
        else if (n.indexOf('produtosreport') === 0 && col.prod === undefined) col.prod = c;
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
    if (col.prod !== undefined && rvNum(l[3]) > 0) aba.getRange(linha, col.prod + 1).setValue(rvNum(l[3]));
    gravados.push(chave + '=' + rvNum(l[2]));
  });
  rvAvisar('Lançado na ' + RV_ABA_HISTORICO + ': ' + (gravados.join(', ') || 'nada')
    + (semLinha.length ? '. SEM linha na HISTORICO (mês ainda não existe lá): ' + semLinha.join(', ') : '')
    + (col.prod === undefined ? '. Dica: crie também a coluna "produtosReportados" para o dashboard usar o fator exato.' : '')
    + '. O dashboard pega no próximo sync.');
}

/* ══ HELPERS ══ */
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
