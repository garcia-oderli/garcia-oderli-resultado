/**
 * DASHBOARD PPCP — Web App de leitura da planilha
 * ────────────────────────────────────────────────────────────────────────────
 * Publique como Web App (Implantar › Nova implantação › Tipo: App da Web,
 * "Executar como: eu", "Quem tem acesso: qualquer pessoa") e cole a URL /exec
 * em SHEETS_WEBAPP_URL no index.html.
 *
 * doGet — o dashboard lê a planilha ao abrir:
 *   1. Lê a aba HISTORICO (uma linha por mês) e devolve os campos que o
 *      dashboard conhece;
 *   2. Puxa a previsão de produção da aba PLANO MESTRE (linha TOTAL GERAL),
 *      que passa a ser a fonte única do plano — não é preciso repetir o
 *      número no HISTORICO nem manter as duas pontas em dia na mão;
 *   3. Normaliza número em formato brasileiro. Sem isso, uma célula gravada
 *      como texto "33.291" chega ao JavaScript como 33,291 (trinta e três) e
 *      "11,52" chega como NaN → 0. Os dois casos existem hoje na planilha.
 *
 * doPost — o dashboard grava de volta (importar Excel, editar ou excluir mês).
 *   Casa cada registro por mês+ano: atualiza a linha existente e acrescenta a
 *   que não existir, preservando a ordem das colunas da planilha. NÃO escreve
 *   na coluna de previsão nem na aba PLANO MESTRE: o plano é do planejamento,
 *   e deixar o app sobrescrevê-lo permitiria apagar o plano sem querer.
 */

/* Aba com uma linha por mês (a que alimenta o dashboard). */
var ABA_HISTORICO = 'HISTORICO';
/* Aba matriz do Planejamento Mestre (lotes nas linhas × meses nas colunas). */
var ABA_PLANO = 'PLANO MESTRE';
/* Rótulo da linha de totais dentro da aba do plano. */
var LINHA_TOTAL = 'TOTAL GERAL';

var MESES = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

/* Precisa ser igual a SHEETS_SECRET no index.html. Não é segurança de
   verdade — o valor viaja no JavaScript da página, que é público. Serve para
   evitar escrita acidental por quem esbarrar na URL. */
var SEGREDO = 'TROQUE_ESTA_SENHA_2026';

/* Colunas que o app NÃO pode sobrescrever ao gravar. A previsão vem do
   PLANO MESTRE; se o app pudesse escrevê-la, um dado local desatualizado
   apagaria o plano. */
var COLUNAS_PROTEGIDAS = ['previsaoproducao', 'previsao'];

function doGet() {
  var saida;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dados = lerHistorico(ss);
    var plano = lerPlanoMestre(ss);
    aplicarPlano(dados, plano);
    saida = { ok: true, dados: dados, plano: plano, geradoEm: new Date().toISOString() };
  } catch (e) {
    saida = { ok: false, erro: String(e && e.message || e) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(saida))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ══ ESCRITA ══
   O app envia { secret, dados:[...] } como text/plain, em modo no-cors — ele
   não lê a resposta, mas devolvemos JSON assim mesmo para dar para testar a
   URL na mão. */
function doPost(e) {
  var saida;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);   /* duas abas salvando ao mesmo tempo não se atropelam */
    var corpo = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (SEGREDO && corpo.secret !== SEGREDO) throw new Error('Segredo inválido.');
    var recebidos = corpo.dados;
    if (!recebidos || !recebidos.length) throw new Error('Nada para gravar.');

    var res = gravarHistorico(SpreadsheetApp.getActiveSpreadsheet(), recebidos);
    saida = { ok: true, atualizados: res.atualizados, incluidos: res.incluidos };
  } catch (err) {
    saida = { ok: false, erro: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
  return ContentService
    .createTextOutput(JSON.stringify(saida))
    .setMimeType(ContentService.MimeType.JSON);
}

function gravarHistorico(ss, recebidos) {
  var aba = ss.getSheetByName(ABA_HISTORICO);
  if (!aba) throw new Error('Aba "' + ABA_HISTORICO + '" não encontrada.');

  var linhas = aba.getDataRange().getValues();
  var iCab = acharCabecalho(linhas);
  var cab = linhas[iCab];

  /* nome normalizado da coluna → índice, para casar com os campos do app */
  var col = {};
  for (var c = 0; c < cab.length; c++) {
    var n = normaliza(cab[c]);
    if (!n) continue;
    if (n.indexOf('previsao') === 0) n = 'previsaoproducao';
    if (n.indexOf('ferias') >= 0)    n = 'horasferias';
    if (col[n] === undefined) col[n] = c;
  }
  if (col.mes === undefined || col.ano === undefined) {
    throw new Error('Colunas "mes" e "ano" não encontradas em ' + ABA_HISTORICO + '.');
  }

  /* mês+ano → número da linha na planilha */
  var mapa = {};
  for (var r = iCab + 1; r < linhas.length; r++) {
    var m = String(linhas[r][col.mes] || '').trim().toUpperCase().slice(0, 3);
    var a = num(linhas[r][col.ano]);
    if (MESES.indexOf(m) >= 0 && a) mapa[m + '/' + a] = r;
  }

  var atualizados = 0, incluidos = 0;
  recebidos.forEach(function (rec) {
    var mes = String(rec.mes || '').trim().toUpperCase().slice(0, 3);
    var ano = num(rec.ano);
    if (MESES.indexOf(mes) < 0 || !ano) return;

    var chave = mes + '/' + ano;
    var idx = mapa[chave];
    var linha;
    if (idx === undefined) {
      linha = new Array(cab.length).fill('');
      linhas.push(linha);
      idx = linhas.length - 1;
      mapa[chave] = idx;
      incluidos++;
    } else {
      linha = linhas[idx];
      atualizados++;
    }

    Object.keys(rec).forEach(function (campo) {
      var n = normaliza(campo);
      if (n === 'id') return;
      if (COLUNAS_PROTEGIDAS.indexOf(n) >= 0) return;
      var c = col[n];
      if (c === undefined) return;      /* campo que a planilha não tem: ignora */
      linha[c] = rec[campo];
    });
    linha[col.mes] = mes;
    linha[col.ano] = ano;
  });

  /* uma escrita só — muito mais rápido que célula a célula */
  var largura = cab.length;
  var bloco = linhas.slice(iCab + 1).map(function (l) {
    var out = l.slice(0, largura);
    while (out.length < largura) out.push('');
    return out;
  });
  if (bloco.length) {
    aba.getRange(iCab + 2, 1, bloco.length, largura).setValues(bloco);
  }
  return { atualizados: atualizados, incluidos: incluidos };
}

function acharCabecalho(linhas) {
  for (var i = 0; i < Math.min(linhas.length, 20); i++) {
    var norm = linhas[i].map(function (c) { return normaliza(c); });
    if (norm.indexOf('mes') >= 0 && norm.indexOf('ano') >= 0) return i;
  }
  throw new Error('Cabeçalho com "mes" e "ano" não encontrado em ' + ABA_HISTORICO + '.');
}

/* ══ NÚMEROS ══
   Aceita number puro, "1.234", "1.234,56", "1234,56", "R$ 1.234,56", "12%",
   "-", vazio. Devolve sempre Number. */
function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (Object.prototype.toString.call(v) === '[object Date]') return 0;

  var s = String(v).trim();
  if (!s || s === '-' || s === '—') return 0;
  s = s.replace(/\s| /g, '').replace(/R\$/gi, '').replace(/%/g, '');

  if (s.indexOf(',') >= 0) {
    /* vírgula é o decimal: pontos que restarem são separador de milhar */
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    /* só pontos e todos em grupos de 3 → separador de milhar, não decimal.
       É este caso que salva "33.291" de virar 33,291. */
    s = s.replace(/\./g, '');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function txt(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

/* ══ HISTORICO ══ */
function lerHistorico(ss) {
  var aba = ss.getSheetByName(ABA_HISTORICO);
  if (!aba) throw new Error('Aba "' + ABA_HISTORICO + '" não encontrada.');
  var linhas = aba.getDataRange().getValues();
  if (!linhas.length) return [];

  /* A linha de cabeçalho é a que traz "mes" e "ano" — não assumimos que é a
     primeira, porque a planilha pode ganhar título ou linha em branco. */
  var iCab = acharCabecalho(linhas);
  var cab = linhas[iCab];
  var out = [];
  for (var r = iCab + 1; r < linhas.length; r++) {
    var linha = linhas[r];
    var rec = {};
    for (var c = 0; c < cab.length; c++) {
      var nome = txt(cab[c]);
      if (!nome) continue;
      rec[nome] = linha[c];
    }
    var mes = txt(rec.mes || rec.Mes || rec.MES).toUpperCase().slice(0, 3);
    var ano = num(rec.ano || rec.Ano || rec.ANO);
    if (MESES.indexOf(mes) < 0 || !ano) continue;   /* linha vazia ou de rodapé */

    /* converte tudo que não é texto para número normalizado */
    Object.keys(rec).forEach(function (k) {
      if (normaliza(k) === 'mes') { rec[k] = mes; return; }
      rec[k] = num(rec[k]);
    });
    rec.mes = mes;
    rec.ano = ano;
    out.push(rec);
  }
  return out;
}

/* ══ PLANO MESTRE ══
   Devolve { '2026': { JAN: 22828, FEV: 27893, ... } } lendo a linha TOTAL
   GERAL e o cabeçalho de meses da própria aba. */
function lerPlanoMestre(ss) {
  var plano = {};
  var aba = ss.getSheetByName(ABA_PLANO);
  if (!aba) return plano;

  var linhas = aba.getDataRange().getValues();
  var iTotal = -1, iMes = -1;
  for (var i = 0; i < linhas.length; i++) {
    var a = normaliza(linhas[i][0]);
    if (iTotal < 0 && a === normaliza(LINHA_TOTAL)) iTotal = i;
    if (iMes < 0 && contaMeses(linhas[i]) >= 6) iMes = i;
  }
  if (iTotal < 0 || iMes < 0) return plano;

  var cab = linhas[iMes], tot = linhas[iTotal];
  for (var c = 1; c < cab.length; c++) {
    var mv = mesEAno(cab[c]);
    if (!mv) continue;
    var v = num(tot[c]);
    if (!v) continue;
    var chave = String(mv.ano);
    if (!plano[chave]) plano[chave] = {};
    plano[chave][mv.mes] = v;
  }
  return plano;
}

/* Copia a previsão do plano para os registros do histórico. O plano manda:
   onde ele tem número, ele vence o que estiver no HISTORICO. */
function aplicarPlano(dados, plano) {
  dados.forEach(function (r) {
    var doAno = plano[String(r.ano)];
    if (!doAno) return;
    var v = doAno[r.mes];
    if (v > 0) r.previsaoProducao = v;
  });
}

/* ══ HELPERS DE CABEÇALHO ══ */
function normaliza(v) {
  return String(v === null || v === undefined ? '' : v)
    .trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

var _MES_TXT = { jan:'JAN', fev:'FEV', mar:'MAR', abr:'ABR', mai:'MAI', jun:'JUN',
                 jul:'JUL', ago:'AGO', set:'SET', out:'OUT', nov:'NOV', dez:'DEZ' };

/* Aceita "jan./26", "JAN/2026", "jan-26" ou uma Data — o Sheets converte
   alguns desses cabeçalhos em data sozinho. */
function mesEAno(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return { mes: MESES[v.getMonth()], ano: v.getFullYear() };
  }
  var s = normaliza(v);
  if (!s) return null;
  var m = s.match(/^([a-z]{3})[a-z]*\.?[\/\-\s]?(\d{2,4})?/);
  if (!m || !_MES_TXT[m[1]]) return null;
  var ano = m[2] ? parseInt(m[2], 10) : new Date().getFullYear();
  if (ano < 100) ano += 2000;
  return { mes: _MES_TXT[m[1]], ano: ano };
}

function contaMeses(linha) {
  var n = 0;
  for (var i = 0; i < linha.length; i++) if (mesEAno(linha[i])) n++;
  return n;
}
