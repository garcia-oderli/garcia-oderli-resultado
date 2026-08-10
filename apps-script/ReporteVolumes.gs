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
 * MODO AUTOMÁTICO (o do dia a dia):
 *   1. Crie no Drive uma pasta chamada REPORTES DE VOLUMES;
 *   2. Todo mês, salve nela o PDF do relatório (3 - REPORTE, Tipo: Todos);
 *   3. O resto é sozinho: processarReportesDrive() converte o PDF, extrai as
 *      linhas, atualiza a REPORTE_VOLUMES, recalcula e lança na HISTORICO.
 *      Rode instalarProcessamentoDiario() uma vez e nem o clique precisa —
 *      a pasta é varrida todo dia de manhã; PDF processado vai para a
 *      subpasta PROCESSADOS. Também dá para disparar na hora pelo menu
 *      "📦 Volumes" que aparece na planilha.
 *
 * MODO MANUAL (continua valendo, e é o plano B se o PDF mudar de cara):
 *   1. criarAbaReporteVolumes() — uma vez; monta a aba REPORTE_VOLUMES;
 *   2. Cole o relatório do mês: MES | ANO | CODIGO | DESCRICAO | QUANTIDADE.
 *      A DESCRICAO é obrigatória: é ela que diz o que é volume ("VOL x/y…")
 *      e o que é produto. Código com pontos ("501.061.001") funciona;
 *   3. calcularVolumesMes() — escreve o resumo por mês na própria aba e
 *      lista os produtos sem SKU de volume (pendências de cadastro);
 *   4. Confira e rode lancarVolumesNaHistorico() — grava volumesProduzidos
 *      (e produtosReportados, se a coluna existir) casando mês+ano.
 *
 * Roda dentro da própria planilha, como a carga do plano — o Web App não
 * escreve nessas colunas.
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
      meses[chave] = { mes: mes, ano: ano, vol: 0, prod: 0, grupos: {}, produtos: [] };
      ordem.push(chave);
    }
    var g = meses[chave], m = desc.match(RV_RE_VOL);
    if (m) {
      /* grupo por produto: quantas caixas ele tem (o "de" do x/de) e quanto
         foi reportado nessas linhas */
      var nome = m[3].trim().toUpperCase();
      var grp = g.grupos[nome] || (g.grupos[nome] = { n: 1, q: 0 });
      grp.n = Math.max(grp.n, parseInt(m[2], 10) || 1);
      grp.q += qtd;
    } else {
      g.prod += qtd;
      g.produtos.push({ nome: desc.toUpperCase(), qtd: qtd });
    }
  });
  if (!ordem.length) return rvErro('Nenhuma linha válida em ' + RV_ABA + ' (MES, ANO, CODIGO, DESCRICAO, QUANTIDADE).');

  var saida = ordem.map(function (chave) {
    var g = meses[chave];
    /* VOLUMES = quantidade do produto × nº de caixas dele.
       Somar as linhas VOL parecia mais direto, mas elas divergem do produto
       quando a caixa é apontada em mês diferente (39 de 138 produtos em
       JAN/26), e aí o fator chega a ficar abaixo de 1 — impossível, já que
       nenhum produto sai em menos de uma caixa. Multiplicar é imune a essa
       defasagem: cada unidade produzida gera as caixas que a embalagem dela
       exige. O nº de caixas vem da estrutura VOL x/de do próprio mês; sem
       linha VOL, vem do cadastro; sem cadastro, 1 caixa. */
    var pend = {}, usados = {};
    g.produtos.forEach(function (p) {
      var achou = null;
      Object.keys(g.grupos).some(function (nv) {
        if (p.nome.indexOf(nv) === 0 || rvAbrevia(nv, p.nome)) { achou = nv; return true; }
        return false;
      });
      if (achou) { g.vol += p.qtd * g.grupos[achou].n; usados[achou] = true; return; }
      var nVols = 0;
      Object.keys(cadastro).some(function (nc) {
        if (p.nome.indexOf(nc) === 0 || rvAbrevia(nc, p.nome)) { nVols = cadastro[nc]; return true; }
        return false;
      });
      if (nVols > 0) { g.vol += p.qtd * nVols; }
      else           { g.vol += p.qtd; pend[p.nome] = true; }
    });
    /* Linhas VOL sem produto correspondente no mês: entram como reportadas,
       senão a caixa apontada some da conta. */
    Object.keys(g.grupos).forEach(function (nv) {
      if (!usados[nv]) g.vol += g.grupos[nv].q;
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

/* ══ 5 · AUTOMÁTICO — PDF do Drive direto para a HISTORICO ══
   O ERP é local e não conversa com a planilha; o combinado é: o PDF do mês
   cai na pasta REPORTES DE VOLUMES do Drive e daqui para frente ninguém
   digita nada. A conversão PDF→texto usa o próprio Drive (copiar o arquivo
   como Documento Google extrai o texto, o mesmo truque do OCR), as linhas
   são lidas com a mesma régua do modo manual e o mês inteiro é substituído
   na REPORTE_VOLUMES — rodar duas vezes não duplica. */

var RV_PASTA      = 'REPORTES DE VOLUMES';
var RV_PROCESSADOS = 'PROCESSADOS';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 Volumes')
    .addItem('Processar PDFs da pasta do Drive', 'processarReportesDrive')
    .addItem('Recalcular resumo (dados colados)', 'calcularVolumesMes')
    .addItem('Lançar na HISTORICO', 'lancarVolumesNaHistorico')
    .addSeparator()
    .addItem('Instalar processamento diário', 'instalarProcessamentoDiario')
    .addItem('Criar linha TOTAL VOLUMES no plano mestre', 'criarLinhaTotalVolumes')
    .addToUi();
}

function processarReportesDrive() {
  var pastas = DriveApp.getFoldersByName(RV_PASTA);
  if (!pastas.hasNext()) {
    return rvErro('Pasta "' + RV_PASTA + '" não encontrada no Drive — crie a pasta e solte os PDFs do relatório nela.');
  }
  var pasta = pastas.next();
  var arquivos = pasta.getFilesByType(MimeType.PDF);
  var feitos = [], falhas = [];

  while (arquivos.hasNext()) {
    var pdf = arquivos.next();
    try {
      var texto = rvPdfParaTexto(pdf.getId());
      var mesAno = rvPeriodoDoTexto(texto);
      if (!mesAno) throw new Error('não achei "Período: dd/mm/aa" no PDF');
      var linhas = rvLinhasDoTexto(texto, mesAno);
      if (!linhas.length) throw new Error('nenhuma linha de produto/volume reconhecida');
      rvSubstituirMes(mesAno, linhas);
      rvMoverParaProcessados(pasta, pdf);
      feitos.push(pdf.getName() + ' → ' + mesAno.mes + '/' + mesAno.ano + ' (' + linhas.length + ' linhas)');
    } catch (e) {
      falhas.push(pdf.getName() + ': ' + (e && e.message || e));
    }
  }

  if (feitos.length) {
    calcularVolumesMes();
    lancarVolumesNaHistorico();
  }
  rvAvisar((feitos.length ? 'Processados:\n' + feitos.join('\n') : 'Nenhum PDF novo na pasta.')
    + (falhas.length ? '\n\nFALHARAM (ficaram na pasta):\n' + falhas.join('\n') : ''));
}

/* Um gatilho por dia, de manhã. Reinstalar não duplica. */
function instalarProcessamentoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processarReportesDrive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processarReportesDrive').timeBased().everyDays(1).atHour(6).create();
  rvAvisar('Instalado: a pasta "' + RV_PASTA + '" é varrida todo dia por volta das 6h. '
    + 'É só salvar o PDF do mês lá dentro; o resultado aparece na HISTORICO e o dashboard pega no sync.');
}

/* Copia o PDF como Documento Google (o Drive extrai o texto), lê e apaga a
   cópia. Não mexe no PDF original. */
function rvPdfParaTexto(fileId) {
  var resp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ name: 'tmp_reporte_volumes', mimeType: 'application/vnd.google-apps.document' }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error('conversão do PDF falhou (HTTP ' + resp.getResponseCode() + ')');
  }
  var docId = JSON.parse(resp.getContentText()).id;
  try {
    return DocumentApp.openById(docId).getBody().getText();
  } finally {
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (ignore) {}
  }
}

function rvPeriodoDoTexto(texto) {
  var m = texto.match(/Per[íi]odo:\s*\d{2}\/(\d{2})\/(\d{2,4})/);
  if (!m) return null;
  var ano = parseInt(m[2], 10);
  if (ano < 100) ano += 2000;
  return { mes: RV_MESES[parseInt(m[1], 10) - 1], ano: ano };
}

/* Mesma régua do PDF: código 000.000.000, descrição, três números no fim
   (quantidade, peso, custo). Linha de cabeçalho/rodapé não casa e é
   ignorada de graça. */
function rvLinhasDoTexto(texto, mesAno) {
  var out = [];
  texto.split('\n').forEach(function (ln) {
    var m = ln.trim().match(/^(\d{3}\.\d{3}\.\d{3})\s+(.+?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$/);
    if (!m) return;
    var qtd = rvNum(m[3]);
    if (!qtd) return;
    out.push([mesAno.mes, mesAno.ano, m[1], m[2].trim(), qtd]);
  });
  return out;
}

/* Troca as linhas do mês na REPORTE_VOLUMES pelas recém-lidas — reprocessar
   o mesmo mês (PDF corrigido, por exemplo) substitui em vez de somar. */
function rvSubstituirMes(mesAno, novas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(RV_ABA);
  if (!aba) { criarAbaReporteVolumes(); aba = ss.getSheetByName(RV_ABA); }
  var alt = Math.max(aba.getLastRow() - 1, 0);
  var atuais = alt ? aba.getRange(2, 1, alt, 5).getValues() : [];
  var mantidas = atuais.filter(function (l) {
    var m = String(l[0] || '').trim().toUpperCase().slice(0, 3);
    return l[0] && !(m === mesAno.mes && parseInt(l[1], 10) === mesAno.ano);
  });
  var tudo = mantidas.concat(novas);
  if (alt) aba.getRange(2, 1, alt, 5).clearContent();
  if (tudo.length) aba.getRange(2, 1, tudo.length, 5).setValues(tudo);
}

function rvMoverParaProcessados(pasta, pdf) {
  var sub = pasta.getFoldersByName(RV_PROCESSADOS);
  var destino = sub.hasNext() ? sub.next() : pasta.createFolder(RV_PROCESSADOS);
  pdf.moveTo(destino);
}

/* ══ 6 · PLANO MESTRE EM VOLUMES ══
   A capacidade da embalagem é por volume, então o plano também precisa
   existir nessa unidade — mas a matriz de lotes continua em produtos, que é
   o que o resto do painel (ritmos, simulação, demanda) consome. A saída é
   uma linha oficial TOTAL VOLUMES logo abaixo da TOTAL GERAL: nasce como
   fórmula (produtos do mês × fator, editável na coluna O) e o planejamento
   pode sobrescrever qualquer mês com o número decidido. O Code.gs lê a
   linha e o dashboard passa a usar o plano oficial em vez do fator médio. */

var RV_PM_ABA    = 'PLANO MESTRE';
var RV_PM_TOTAL  = 'TOTAL GERAL';
var RV_PM_LINHA  = 'TOTAL VOLUMES';
/* Colunas da matriz: A = rótulo, B..M = jan..dez, N = TOTAL ANO, O = fator. */
var RV_PM_COL_JAN = 2, RV_PM_COL_DEZ = 13, RV_PM_COL_TOTAL = 14, RV_PM_COL_FATOR = 15;
/* Fator inicial: realizado JAN–JUL/26 (254.821 volumes ÷ 210.961 produtos). */
var RV_PM_FATOR_INICIAL = 1.208;

function criarLinhaTotalVolumes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(RV_PM_ABA);
  if (!aba) return rvErro('Aba "' + RV_PM_ABA + '" não encontrada.');

  var valores = aba.getDataRange().getValues();
  var iTotal = -1, iVol = -1;
  for (var i = 0; i < valores.length; i++) {
    var a = rvNormaliza(valores[i][0]);
    if (iTotal < 0 && a === rvNormaliza(RV_PM_TOTAL)) iTotal = i + 1;
    if (iVol < 0 && a === rvNormaliza(RV_PM_LINHA)) iVol = i + 1;
  }
  if (iVol > 0)   return rvAvisar('Linha ' + RV_PM_LINHA + ' já existe (linha ' + iVol + ') — nada foi alterado.');
  if (iTotal < 0) return rvErro('Linha "' + RV_PM_TOTAL + '" não encontrada na ' + RV_PM_ABA + '.');

  aba.insertRowsAfter(iTotal, 1);
  var nova = iTotal + 1;
  aba.getRange(nova, 1).setValue(RV_PM_LINHA).setFontWeight('bold');
  aba.getRange(nova, RV_PM_COL_FATOR).setValue(RV_PM_FATOR_INICIAL);
  aba.getRange(nova, RV_PM_COL_FATOR + 1)
     .setValue('← fator vol/produto — edite aqui, ou digite o volume direto no mês')
     .setFontStyle('italic');

  /* Mês sem plano fica traço, como o resto da matriz; com plano, produtos ×
     fator. É fórmula de propósito: enquanto o planejamento não decidir o
     número, a linha acompanha o plano de produtos sozinha. */
  var f = [];
  for (var c = RV_PM_COL_JAN; c <= RV_PM_COL_DEZ; c++) {
    var letra = String.fromCharCode(64 + c);
    f.push('=IF(N(' + letra + iTotal + ')=0,"-",ROUND(' + letra + iTotal + '*$O$' + nova + ',0))');
  }
  aba.getRange(nova, RV_PM_COL_JAN, 1, f.length).setFormulas([f]);
  aba.getRange(nova, RV_PM_COL_TOTAL).setFormula('=SUM(B' + nova + ':M' + nova + ')');

  rvAvisar('Linha ' + RV_PM_LINHA + ' criada abaixo da ' + RV_PM_TOTAL + ' com fator '
    + RV_PM_FATOR_INICIAL + ' (realizado JAN–JUL/26). Ajuste o fator na coluna O ou digite '
    + 'os meses direto. Reimplante o Web App para o dashboard passar a ler o plano em volumes.');
}

/* ══ HELPERS ══ */

/* A linha VOL abrevia e trunca a descrição do produto: "PENT CAMARIM 1PT
   2GAV" contra "PENTEADEIRA CAMARIM 1PT 2GAV BRANCO", "CONJ 2 MESAS" contra
   "CONJUNTO 2 MESAS", "CANT CAFE" contra "CANTINHO DO CAFE". Só comparar
   começo de texto não casa esses, e o produto era contado como caixa única
   ALÉM das linhas VOL dele — 5% de volume a mais no ano. Aqui cada palavra
   do VOL precisa ser começo da palavra correspondente do produto (ou vice-
   versa), ignorando conectores; qualquer palavra diferente reprova o
   casamento, então "MESA CABECEIRA SLEEP" não casa com "MESA CABECEIRA
   MAVIE". */
var RV_STOP = { DO:1, DA:1, DE:1, DOS:1, DAS:1, E:1 };
function rvPalavras(s) {
  return String(s || '').toUpperCase().split(/[\s\/]+/).filter(function (t) {
    return t && !RV_STOP[t];
  });
}
function rvAbrevia(nomeVol, nomeProduto) {
  var a = rvPalavras(nomeVol), b = rvPalavras(nomeProduto);
  var n = Math.min(a.length, b.length);
  if (n < 2) return false;
  for (var i = 0; i < n; i++) {
    if (b[i].indexOf(a[i]) !== 0 && a[i].indexOf(b[i]) !== 0) return false;
  }
  return true;
}

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
