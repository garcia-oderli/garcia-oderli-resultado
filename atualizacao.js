/* ══════════════════════════════════════════════════════════════════════════
   AVISO DE ATUALIZAÇÃO — módulo avulso para os painéis da Patrimar
   ─────────────────────────────────────────────────────────────────────────
   Mostra um cartão "Nova versão disponível" quando o service worker detecta
   uma publicação nova, e só recarrega quando o usuário manda. Sem isso a
   página ou recarrega sozinha no meio de uma reunião, ou serve versão antiga
   sem ninguém perceber — os dois já aconteceram aqui.

   COMO USAR NOS OUTROS APPS (curva-abc, rastreio-corte, ritmoprod-patrimar,
   painel-esteira):

   1. Copie este arquivo para a raiz do projeto como `atualizacao.js`;
   2. Antes de </body>, inclua:
          <script src="/atualizacao.js"></script>
   3. Ajuste APP_VERSAO e APP_HISTORICO abaixo para o app em questão;
   4. Garanta que existe um sw.js registrado e que ele responde à mensagem
      { tipo: 'ATIVAR_AGORA' } com self.skipWaiting() — veja o rodapé deste
      arquivo para o trecho pronto;
   5. A cada publicação, suba APP_VERSAO aqui e o CACHE_NAME no sw.js. É a
      troca do CACHE_NAME que faz o navegador enxergar a versão nova.

   Se o app não tiver service worker, o módulo apenas desenha o selo e não
   avisa nada — não quebra.
══════════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  /* ── AJUSTE POR APP ───────────────────────────────────────────────────── */
  var APP_NOME    = 'Painel';
  var APP_VERSAO  = 'b1';
  var APP_DATA    = '05/08/2026';
  var APP_HISTORICO = [
    { v: 'b1', data: '05/08/2026', itens: ['Primeira versão com aviso de atualização.'] }
  ];
  /* ─────────────────────────────────────────────────────────────────────── */

  var COR = '#F5A623';   /* laranja Patrimar */

  /* ── Selo de versão no rodapé ──
     Reaproveita um #app-version que já exista; se não houver, cria. */
  function selo() {
    var el = document.getElementById('app-version');
    if (!el) {
      el = document.createElement('div');
      el.id = 'app-version';
      el.style.cssText = 'position:fixed;right:8px;bottom:6px;z-index:99999;'
        + 'font:600 10px/1.2 Consolas,monospace;color:#7aa0c0;background:rgba(0,0,0,.5);'
        + 'border:1px solid rgba(120,144,156,.3);border-radius:5px;padding:3px 7px;'
        + 'letter-spacing:.04em';
      document.body.appendChild(el);
    }
    el.textContent = 'v' + APP_DATA.split('/').reverse().join('.') + ' · ' + APP_VERSAO;
    el.title = 'Versão publicada — clique para ver o que mudou';
    el.style.cursor = 'pointer';
    el.style.pointerEvents = 'auto';
    el.onclick = historico;
  }

  /* ── Painel com o histórico de versões ── */
  function historico() {
    if (document.getElementById('ver-overlay')) return;
    var ov = document.createElement('div');
    ov.id = 'ver-overlay';
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(6,14,26,.9);z-index:100000;'
      + 'display:flex;align-items:center;justify-content:center;padding:16px;'
      + 'font-family:system-ui,Segoe UI,Arial,sans-serif';
    ov.onclick = function (e) { if (e.target === ov) ov.remove(); };
    ov.innerHTML =
      '<div style="background:#1F1F1F;border:1px solid #2A2A2A;border-top:3px solid ' + COR + ';'
      + 'border-radius:8px;width:100%;max-width:540px;max-height:80vh;overflow:auto;padding:24px 26px;color:#F5F5F5">'
      + '<div style="display:flex;align-items:baseline;justify-content:space-between;gap:12px">'
        + '<div style="font-size:.95rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase">'
          + APP_NOME + ' — versões</div>'
        + '<button onclick="document.getElementById(\'ver-overlay\').remove()" '
          + 'style="background:none;border:none;color:#888;font-size:1.2rem;cursor:pointer;line-height:1">✕</button>'
      + '</div>'
      + '<div style="font-size:.72rem;color:#888;margin:4px 0 18px">Em uso: <b style="color:' + COR + '">'
        + APP_VERSAO + '</b> · ' + APP_DATA + '</div>'
      + APP_HISTORICO.map(function (h, i) {
          var atual = i === 0;
          return '<div style="margin-bottom:16px;padding-left:14px;border-left:2px solid '
            + (atual ? COR : '#2A2A2A') + '">'
            + '<div style="font-family:Consolas,monospace;font-size:.78rem;font-weight:700;color:'
              + (atual ? COR : '#aaa') + ';margin-bottom:6px">' + h.v
              + ' <span style="font-weight:400;color:#888">· ' + h.data + '</span>'
              + (atual ? ' <span style="font-size:.62rem;background:rgba(245,166,35,.15);padding:2px 7px;border-radius:10px">EM USO</span>' : '')
            + '</div>'
            + '<ul style="margin:0;padding-left:16px;font-size:.76rem;line-height:1.65;color:#bbb">'
            + h.itens.map(function (t) { return '<li>' + t + '</li>'; }).join('')
            + '</ul></div>';
        }).join('')
      + '</div>';
    document.body.appendChild(ov);
  }

  /* ── Cartão de aviso ── */
  function avisar(aplicar) {
    if (document.getElementById('ver-update')) return;
    var box = document.createElement('div');
    box.id = 'ver-update';
    box.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:100001;max-width:330px;'
      + 'background:#1F1F1F;border:1px solid rgba(245,166,35,.5);border-left:4px solid ' + COR + ';'
      + 'border-radius:8px;padding:14px 16px;box-shadow:0 8px 30px rgba(0,0,0,.55);'
      + 'font-family:system-ui,Segoe UI,Arial,sans-serif;font-size:.8rem;color:#F5F5F5';
    box.innerHTML =
      '<div style="font-weight:700;margin-bottom:5px;color:' + COR + '">↻ Nova versão disponível</div>'
      + '<div style="font-size:.74rem;color:#aaa;line-height:1.5;margin-bottom:11px">'
        + 'Atualizar recarrega a página. Os dados ficam salvos — só o que estiver '
        + 'em edição na tela se perde.</div>'
      + '<div style="display:flex;gap:8px">'
        + '<button id="ver-up-ok" style="flex:1;padding:7px 12px;border:1px solid ' + COR + ';border-radius:5px;'
          + 'background:rgba(245,166,35,.12);color:' + COR + ';font-weight:700;font-size:.74rem;cursor:pointer">Atualizar</button>'
        + '<button id="ver-up-no" style="padding:7px 12px;border:1px solid #2A2A2A;border-radius:5px;'
          + 'background:transparent;color:#888;font-size:.74rem;cursor:pointer">Depois</button>'
      + '</div>';
    document.body.appendChild(box);
    document.getElementById('ver-up-ok').onclick = function () {
      box.innerHTML = '<div style="color:#aaa;font-size:.76rem">Atualizando…</div>';
      aplicar();
    };
    document.getElementById('ver-up-no').onclick = function () { box.remove(); };
  }

  /* ── Ligação com o service worker ── */
  function ligar() {
    selo();
    if (!('serviceWorker' in navigator)) return;

    navigator.serviceWorker.register('/sw.js').then(function (reg) {
      reg.update().catch(function () {});

      function vigiar(sw) {
        if (!sw) return;
        sw.addEventListener('statechange', function () {
          /* "installed" com controller ativo = versão nova esperando.
             Sem controller é primeira instalação, não há o que avisar. */
          if (sw.state === 'installed' && navigator.serviceWorker.controller) {
            avisar(function () { sw.postMessage({ tipo: 'ATIVAR_AGORA' }); });
          }
        });
      }

      if (reg.waiting && navigator.serviceWorker.controller) {
        avisar(function () { reg.waiting.postMessage({ tipo: 'ATIVAR_AGORA' }); });
      }
      vigiar(reg.installing);
      reg.addEventListener('updatefound', function () { vigiar(reg.installing); });

      var recarregando = false;
      navigator.serviceWorker.addEventListener('controllerchange', function () {
        if (recarregando) return;
        recarregando = true;
        window.location.reload();
      });

      /* Painel costuma ficar aberto o dia todo em TV/monitor. */
      setInterval(function () { reg.update().catch(function () {}); }, 5 * 60 * 1000);
    }).catch(function (err) {
      console.warn('SW não registrou (app segue funcionando):', err);
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ligar);
  else ligar();
})();

/* ══════════════════════════════════════════════════════════════════════════
   TRECHO PARA O sw.js DE CADA APP
   ─────────────────────────────────────────────────────────────────────────
   Sem isto o botão "Atualizar" não tem efeito: o worker novo fica preso em
   espera até o navegador decidir trocar sozinho.

     var CACHE_NAME = 'nome-do-app-v1';   // suba a cada publicação

     self.addEventListener('message', function (event) {
       if (event.data && event.data.tipo === 'ATIVAR_AGORA') self.skipWaiting();
     });

   E, no install, NÃO chame self.skipWaiting() — é justamente o que provoca a
   troca silenciosa por baixo do usuário.
══════════════════════════════════════════════════════════════════════════ */
