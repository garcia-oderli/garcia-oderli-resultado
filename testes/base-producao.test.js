/* ══════════════════════════════════════════════════════════════════════════
   Teste da base de cálculo de produção — DashUtils.periodo()

   Rode com:  node testes/base-producao.test.js
   Sem dependência, sem build. Sai com código 1 se algo quebrar.

   POR QUE ISTO EXISTE
   O painel é um HTML único de 1 MB sem build e sem teste. A conta que decide
   o que a diretoria vê estava errada de três formas ao mesmo tempo:
     • produção COM hora extra comparada contra `meta`, que é calculada sobre
       horas normais — semáforo verde em 18 de 18 meses;
     • absenteísmo dividido por horasCarga quando a planilha usa horasNormais
       — erro de 0,23 a 0,74 p.p. por mês;
     • percentuais tirados por média aritmética entre meses de 15 e de 23 dias.
   Nenhuma dessas três dá erro de sintaxe, e nenhuma aparece na tela como
   defeito: aparece como número plausível. Só teste pega.

   O teste lê DashUtils e o dataset do próprio index.html — não há cópia dos
   dados aqui para sair de sincronia.
══════════════════════════════════════════════════════════════════════════ */
'use strict';
const fs = require('fs');
const path = require('path');

const HTML = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

/* ── Extrai DashUtils do index.html e avalia num escopo isolado ── */
function carregarDashUtils() {
  const ini = HTML.indexOf('window.DashUtils = (function()');
  const fim = HTML.indexOf('function sortD(d){', ini);
  if (ini < 0 || fim < 0) throw new Error('DashUtils não encontrado no index.html');
  const escopo = { window: {} };
  new Function('window', HTML.slice(ini, fim)).call(escopo, escopo.window);
  return escopo.window.DashUtils;
}

/* ── Extrai o dataset embutido (var PLANILHA = [...]) ── */
function carregarPlanilha() {
  const m = HTML.match(/var\s+PLANILHA\s*=\s*(\[[\s\S]*?\n\];)/);
  if (!m) throw new Error('PLANILHA não encontrada no index.html');
  return JSON.parse(m[1].replace(/;\s*$/, ''));
}

const U = carregarDashUtils();
const TODOS = carregarPlanilha();
const REGS = TODOS.filter(r => r.producaoReal > 0);

let falhas = 0, total = 0;
function ok(nome, obtido, esperado, tol = 0.01) {
  total++;
  const bom = Math.abs(obtido - esperado) <= tol;
  if (!bom) falhas++;
  console.log(`  ${bom ? 'ok  ' : 'FALHA'} ${nome.padEnd(46)} ${(+obtido).toFixed(2).padStart(12)} (esperado ${(+esperado).toFixed(2)})`);
}
function afirma(nome, cond) {
  total++;
  if (!cond) falhas++;
  console.log(`  ${cond ? 'ok  ' : 'FALHA'} ${nome}`);
}
const sec = t => console.log('\n' + t);

/* ══ Identidades da planilha — se quebrarem, o dado de origem mudou ══ */
sec('Identidades do dataset (18 meses fechados)');
const quebra = (f) => REGS.filter(r => Math.abs(f(r)) > 0.5).map(r => `${r.mes}/${r.ano}`);
afirma('horasTotais = horasNormais + totalExtras',
  quebra(r => r.horasTotais - (r.horasNormais + r.totalExtras)).length === 0);
afirma('totalExtras = extra50 + extra100',
  quebra(r => r.totalExtras - (r.extra50 + r.extra100)).length === 0);
afirma('totalFaltaAtraso = faltas + atraso',
  quebra(r => r.totalFaltaAtraso - (r.faltas + r.atraso)).length === 0);
afirma('prodSemExtras <= producaoReal',
  REGS.every(r => r.prodSemExtras <= r.producaoReal));
afirma('absenteismo = totalFaltaAtraso / horasNormais  (NAO / horasCarga)',
  REGS.every(r => Math.abs(r.absenteismo - r.totalFaltaAtraso / r.horasNormais * 100) <= 0.02));
afirma('eficiencia = prodSemExtras / meta',
  REGS.every(r => !r.meta || Math.abs(r.eficiencia - r.prodSemExtras / r.meta * 100) <= 0.02));

/* A `meta` da planilha é derivada do próprio realizado. Enquanto isso for
   verdade, nenhum indicador pode usá-la como se fosse capacidade. Se este
   teste passar a falhar, a planilha mudou de método — e aí `meta` volta a
   valer como alvo. */
sec('meta é circular (derivada do realizado) — premissa do desenho da tela');
afirma('meta = producaoReal x (horasNormais / horasCarga), erro < 0,01%',
  REGS.every(r => Math.abs(r.meta - r.producaoReal * r.horasNormais / r.horasCarga) / r.meta * 100 < 0.01));
afirma('logo producaoReal >= meta em 18 de 18 meses (verde por construção)',
  REGS.filter(r => r.producaoReal >= r.meta).length === REGS.length);

/* ══ Agregado do período ══ */
sec('periodo() — 18 meses');
const o = U.periodo(TODOS);
ok('meses com produção',        o.mesesProd,   18, 0);
ok('produção COM extras',       o.comExtras, 556827, 0);
ok('produção SEM extras',       o.semExtras, 472441, 0);
ok('peças vindas de hora extra',o.deExtras,   84386, 0);
ok('dependência de hora extra %',o.depExtras, 15.15);
ok('absenteísmo ponderado %',   o.absenteismo, 12.10);
ok('% horas extras',            o.pctExtras,   12.68);
ok('ticket ponderado R$',       o.ticket,     247.69);
ok('ritmo hora normal pç/h',    o.ritmoNormal,  1.79);
ok('aderência ao plano %',      o.aderPlano,  101.79);
ok('gap vs demanda faturada',   o.gapDemanda, 91883, 0);

sec('periodo() — recortes');
ok('2026: meses',        U.periodo(TODOS.filter(r => r.ano === 2026)).mesesProd, 6, 0);
ok('2026: aderência plano %', U.periodo(TODOS.filter(r => r.ano === 2026)).aderPlano, 100.47);
const abr = U.periodo(TODOS.filter(r => r.ano === 2026 && r.mes === 'ABR'));
ok('ABR/26: prod c/ extras', abr.comExtras, 29195, 0);
ok('ABR/26: prod s/ extras', abr.semExtras, 26147, 0);
ok('ABR/26: aderência = eficiência da planilha', abr.aderencia, 92.26, 0.02);

sec('Bordas');
afirma('lista vazia não quebra',       U.periodo([]).mesesProd === 0);
afirma('null não quebra',              U.periodo(null).mesesProd === 0);
afirma('vazio devolve null, não NaN',  U.periodo([]).aderencia === null);
afirma('meses todos zerados não contam', U.periodo(TODOS.filter(r => r.producaoReal === 0)).mesesProd === 0);

/* Mês de fábrica parada: horas apontadas, produção zero. Tem de entrar no
   denominador das horas — se saísse, o absenteísmo melhoraria justamente no
   mês em que a fábrica parou. */
sec('Mês de fábrica parada entra nas horas, não na produção');
const parado = { mes:'XXX', ano:2099, producaoReal:0, prodSemExtras:0, meta:0,
  horasCarga:10000, horasNormais:9000, totalExtras:0, horasTotais:9000,
  totalFaltaAtraso:1000, qtdeFaturado:0, qtdeVendida:0, ticketMedio:0,
  colaboradores:0, diasTrabalhados:20, previsaoProducao:0 };
const cp = U.periodo(TODOS.concat([parado]));
ok('conta nos meses de horas',      cp.meses,     o.meses + 1, 0);
ok('não conta nos meses de produção',cp.mesesProd, o.mesesProd, 0);
ok('horas perdidas somam',          cp.hPerdidas, o.hPerdidas + 1000, 0.1);
ok('horas normais somam',           cp.hNormais,  o.hNormais + 9000, 0.1);
ok('produção não muda',             cp.comExtras, o.comExtras, 0);
afirma('absenteísmo reage ao mês parado, não o ignora', cp.absenteismo !== o.absenteismo);

/* getProd cai em producaoReal quando prodSemExtras falta — reintroduz em
   silêncio o bug que esta função existe para evitar. Tem de ser sinalizado. */
sec('Sinalização de base "sem extras" ausente');
const semCol = REGS.map(r => Object.assign({}, r, { prodSemExtras: 0 }));
const sc = U.periodo(semCol);
ok('detecta os 18 meses sem base', sc.semBase, 18, 0);
afirma('e nesse caso semExtras vira comExtras (o risco sinalizado)',
  sc.semExtras === sc.comExtras);

/* ══ PORTÃO DE SANIDADE DA IMPORTAÇÃO ══
   Reproduz as identidades que o importarExcel usa como gate. O que importa
   testar é que um DESLOCAMENTO DE COLUNA seja pego — o modo de falha que
   grava dado errado em silêncio e só aparece na reunião. */
sec('Portão de sanidade do import');
function reprova(r){
  var f=[];
  if (Math.abs((r.horasNormais + r.totalExtras) - r.horasTotais) > 1) f.push('horas');
  if (Math.abs((r.extra50 + r.extra100) - r.totalExtras) > 1)        f.push('extras');
  if (Math.abs((r.faltas + r.atraso) - r.totalFaltaAtraso) > 1)      f.push('faltaAtraso');
  if (r.prodSemExtras > 0 && r.prodSemExtras > r.producaoReal + 1)   f.push('prodSE>prodCE');
  if (r.diasTrabalhados < 0 || r.diasTrabalhados > 31)               f.push('dias');
  if (r.absenteismo < 0 || r.absenteismo > 60)                       f.push('absenteismo');
  if (r.ticketMedio < 0 || r.ticketMedio > 10000)                    f.push('ticket');
  if (r.custoCap < 0)                                                f.push('custoCap');
  return f;
}
afirma('planilha atual passa no portão (18/18)',
  REGS.every(r => reprova(r).length === 0));

/* Simula o cenário real: uma coluna inserida no meio da planilha. Tudo a
   partir dali roda uma casa. O bloco de horas, que é o único que alguém
   confere de olho, sai perfeito — por isso a validação tem de ser automática. */
const CAMPOS = ['colaboradores','horasCarga','faltas','atraso','totalFaltaAtraso',
  'absenteismo','horasNormais','extra50','extra100','totalExtras','horasTotais',
  '_x1','_x2','_x3','producaoReal','prodSemExtras','_x4','meta','_x5','eficiencia',
  'eficienciaAdj','margem','ticketMedio','custoCap','qtdeFaturado','diasTrabalhados'];
function deslocar(r, aPartirDe){
  const linha = CAMPOS.map(c => (c.startsWith('_') ? 0 : r[c]));
  linha.splice(aPartirDe, 0, 999);          /* insere uma coluna */
  const o = Object.assign({}, r);
  CAMPOS.forEach((c, i) => { if (!c.startsWith('_')) o[c] = linha[i] === undefined ? 0 : linha[i]; });
  return o;
}
const base = REGS[REGS.length - 1];
[3, 8, 12, 14, 18, 20].forEach(pos => {
  const d = deslocar(base, pos);
  const f = reprova(d);
  afirma(`coluna inserida na posição ${String(pos).padStart(2)} é detectada  [${f.join(',') || 'NENHUMA'}]`,
    f.length > 0);
});

console.log(`\n${total - falhas}/${total} passaram` + (falhas ? ` — ${falhas} FALHA(S)\n` : '\n'));
process.exit(falhas ? 1 : 0);
