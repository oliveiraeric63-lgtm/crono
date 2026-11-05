/* Crono · ponte com a planilha (Google Apps Script)
   Substitui os recursos do claude.ai (db, user, mcp, downloads) por chamadas
   ao Apps Script da planilha. Funciona offline: as gravações ficam numa fila
   no aparelho e são enviadas quando a internet volta. Nada é apagado no servidor. */
(function () {
  'use strict';
  var CFG_K = 'crono-servidor', CACHE_K = 'crono-db-cache', FILA_K = 'crono-fila';

  function lerJSON(k, d) { try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function gravarJSON(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function cfg() { return lerJSON(CFG_K, null); }
  function conectado() { var c = cfg(); return !!(c && c.url && c.token); }

  // conexão por link: crono.../#conectar=<base64 de "url|chave">
  (function () {
    var m = (location.hash || '').match(/conectar=([^&]+)/);
    if (!m) return;
    try {
      var p = atob(decodeURIComponent(m[1])).split('|');
      if (/^https:\/\/script\.google\.com\//.test(p[0]) && p[1]) gravarJSON(CFG_K, { url: p[0], token: p[1] });
    } catch (e) {}
    history.replaceState(null, '', location.pathname);
  })();

  function chamar(corpo) {
    var c = cfg();
    if (!c) return Promise.reject({ code: 'server_not_connected' });
    corpo.token = c.token;
    return fetch(c.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(corpo), redirect: 'follow' })
      .then(function (r) { if (!r.ok) throw { code: 'server_unavailable' }; return r.json(); })
      .then(function (j) { if (!j || !j.ok) throw { code: (j && j.code) || 'tool_error', message: j && j.msg }; return j; },
        function (e) { throw e && e.code ? e : { code: 'offline' }; });
  }

  /* ---------- cache local e fila de envio ---------- */
  var cache = lerJSON(CACHE_K, {});
  var fila = lerJSON(FILA_K, []);
  function salvarCache() { gravarJSON(CACHE_K, cache); }
  function salvarFila() { gravarJSON(FILA_K, fila); }
  function mesclar(a, b) {
    var o = Object.assign({}, a);
    Object.keys(b).forEach(function (k) {
      var x = o[k], y = b[k], obj = function (v) { return v && typeof v === 'object' && !Array.isArray(v); };
      o[k] = obj(x) && obj(y) ? mesclar(x, y) : y;
    });
    return o;
  }
  function aplicarLocal(op) {
    if (op.op === 'delete') delete cache[op.id];
    else if (op.op === 'update') cache[op.id] = mesclar(cache[op.id] || {}, op.data || {});
    else cache[op.id] = op.data;
    salvarCache();
  }
  function avisar() { window.dispatchEvent(new Event('crono-sync')); }
  window.__cronoFila = function () { return fila.length; };

  var enviando = false;
  function enviarFila() {
    if (enviando || !fila.length || !conectado()) return Promise.resolve();
    enviando = true;
    var lote = fila.slice(0, 20);
    return chamar({ action: 'write', ops: lote }).then(function (j) {
      var res = j.res || [];
      // update de documento que não existe no servidor: reenvia como set com o conteúdo local
      res.forEach(function (r, i) {
        if (!r.ok && r.code === 'invalid_argument' && cache[lote[i].id]) fila.push({ op: 'set', id: lote[i].id, data: cache[lote[i].id] });
      });
      fila.splice(0, lote.length); salvarFila();
      enviando = false; avisar();
      if (fila.length) return enviarFila();
    }, function () { enviando = false; avisar(); });
  }
  setInterval(enviarFila, 20000);
  window.addEventListener('online', enviarFila);

  /* ---------- db no formato que o Crono já usa ---------- */
  var ouvintes = [], ultimoHash = '';
  function snapshot(docs, fromCache) {
    var lista = Object.keys(docs).sort().map(function (id) { var v = docs[id]; return { id: id, data: function () { return JSON.parse(JSON.stringify(v)); } }; });
    return { docs: lista, empty: !lista.length, metadata: { fromCache: fromCache } };
  }
  function buscar(forcar) {
    if (fila.length) return enviarFila();
    return chamar({ action: 'list' }).then(function (j) {
      if (fila.length) return;
      var docs = {};
      (j.docs || []).forEach(function (d) { docs[d.id] = d.data; });
      // planilha vazia no primeiro uso: não grava nada antes de perguntar
      if (!Object.keys(docs).length && !Object.keys(cache).length && !lerJSON('crono-inicio-ok', false)) {
        window.cronoPrimeiroUso(function () { buscar(true); });
        return;
      }
      var h = JSON.stringify(docs);
      cache = docs; salvarCache();
      if (!forcar && h === ultimoHash) return;
      ultimoHash = h;
      ouvintes.forEach(function (o) { o.cb(snapshot(docs, false)); });
    }, function (e) {
      // sem internet mas com cópia no aparelho: segue usando a cópia e envia depois
      if (forcar && !Object.keys(cache).length) ouvintes.forEach(function (o) { o.err && o.err(e); });
      avisar();
    });
  }
  setInterval(function () { if (document.visibilityState === 'visible') buscar(false); }, 45000);
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') { enviarFila().then(function () { buscar(false); }); } });

  var db = {
    collection: function () {
      return {
        doc: function (id) {
          function gravar(op) {
            aplicarLocal(op);
            fila.push(op); salvarFila();
            enviarFila();
            return Promise.resolve();
          }
          return {
            set: function (body) { return gravar({ op: 'set', id: id, data: body }); },
            update: function (patch) {
              if (!cache[id]) return Promise.reject({ code: 'invalid_argument' });
              return gravar({ op: 'update', id: id, data: patch });
            },
            delete: function () { return gravar({ op: 'delete', id: id }); },
            get: function () {
              return chamar({ action: 'get', id: id }).then(function (j) {
                return { exists: !!j.exists, data: function () { return j.data; } };
              }, function () {
                var v = cache[id];
                return { exists: !!v, data: function () { return v; } };
              });
            }
          };
        },
        onSnapshot: function (cb, err) {
          ouvintes.push({ cb: cb, err: err });
          if (Object.keys(cache).length) setTimeout(function () { cb(snapshot(cache, true)); }, 0);
          enviarFila().then(function () { buscar(true); });
          return function () {};
        }
      };
    }
  };

  var mcp = {
    callTool: function (servidor, ferramenta, args) {
      if (ferramenta === 'batch_clear_values') return chamar({ action: 'clear', spreadsheetId: args.spreadsheetId, ranges: args.ranges });
      if (ferramenta === 'update_values') return chamar({ action: 'values', spreadsheetId: args.spreadsheetId, range: args.range, values: args.values });
      return Promise.reject({ code: 'tool_error' });
    }
  };

  /* ---------- ClaudIA: conversa via servidor, ferramentas rodam aqui ---------- */
  var IA_MSG = {
    sem_ia: 'A IA da ClaudIA ainda não foi ativada no servidor da planilha.',
    ia_chave: 'A chave da Anthropic foi recusada. Confira em Propriedades do script.',
    ia_credito: 'A conta da Anthropic está sem créditos.',
    offline: 'Estou sem internet agora. Assim que a conexão voltar, posso ajudar.'
  };
  function sample(msgs, opts) {
    opts = opts || {};
    var ferramentas = opts.tools || [];
    var porNome = {};
    ferramentas.forEach(function (t) { porNome[t.name] = t; });
    var tools = ferramentas.map(function (t) { return { name: t.name, description: t.description, input_schema: t.inputSchema || { type: 'object', properties: {} } }; });
    var conv = msgs.map(function (m) { return { role: m.role, content: m.content }; });
    var voltas = 0;
    function cancelado() { return opts.signal && opts.signal.aborted; }
    function rodada() {
      if (cancelado()) return Promise.reject({ code: 'cancelled' });
      voltas++;
      return chamar({ action: 'ia', body: { max_tokens: 1024, messages: conv, tools: tools } }).then(function (j) {
        var r = j.resp || {}, blocos = r.content || [];
        var texto = blocos.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('').trim();
        var usos = blocos.filter(function (b) { return b.type === 'tool_use'; });
        if (r.stop_reason === 'tool_use' && usos.length && voltas < 7) {
          conv.push({ role: 'assistant', content: blocos });
          return Promise.all(usos.map(function (u) {
            var t = porNome[u.name];
            return Promise.resolve().then(function () {
              if (!t) throw new Error('Ferramenta desconhecida: ' + u.name);
              return t.execute(u.input || {}, {});
            }).then(function (out) {
              return { type: 'tool_result', tool_use_id: u.id, content: JSON.stringify(out === undefined ? { ok: true } : out) };
            }, function (e) {
              return { type: 'tool_result', tool_use_id: u.id, is_error: true, content: String(e && e.message || e) };
            });
          })).then(function (res) { conv.push({ role: 'user', content: res }); return rodada(); });
        }
        texto = texto || 'Pronto, Sr. Eric.';
        if (opts.onText) opts.onText({ text: texto });
        return { text: texto };
      }, function (e) {
        var c = e && e.code;
        throw { code: c === 'rate_limited' ? 'rate_limited' : (c || 'ia_erro'), text: IA_MSG[c] || (e && e.message) || '' };
      });
    }
    return rodada();
  }
  sample.limits = function () { return Promise.resolve({ tools: true }); };

  var downloads = {
    save: function (o) {
      var url = URL.createObjectURL(o.data), a = document.createElement('a');
      a.href = url; a.download = o.filename || 'crono.csv';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
      return Promise.resolve();
    }
  };

  if (conectado()) {
    window.claude = {
      use: function (nome) {
        if (nome === 'db') return Promise.resolve(db);
        if (nome === 'user') return Promise.resolve({ id: function () { return Promise.resolve('eric'); } });
        if (nome === 'mcp') return Promise.resolve(mcp);
        if (nome === 'downloads') return Promise.resolve(downloads);
        if (nome === 'sample') return Promise.resolve(sample);
        return Promise.reject({ code: 'unavailable' });
      }
    };
  } else {
    window.claude = { use: function (nome) { return nome === 'downloads' ? Promise.resolve(downloads) : Promise.reject({ code: 'unavailable' }); } };
  }

  /* ---------- tela de conexão ---------- */
  window.cronoConectar = function () {
    var c = cfg() || {};
    var ov = document.createElement('div');
    ov.className = 'cx-ov';
    ov.innerHTML =
      '<form class="cx-card">' +
      '<h2>Conectar à planilha</h2>' +
      '<p>Cole a URL do App da Web (termina em <b>/exec</b>) e a chave que apareceu no Apps Script.</p>' +
      '<label>URL do servidor<input name="url" type="url" required placeholder="https://script.google.com/macros/s/…/exec"></label>' +
      '<label>Chave<input name="token" type="text" required autocomplete="off"></label>' +
      '<p class="cx-msg" aria-live="polite"></p>' +
      '<div class="cx-acts"><button type="button" class="cx-sec">Agora não</button><button type="submit" class="cx-pri">Testar e conectar</button></div>' +
      '</form>';
    document.body.appendChild(ov);
    var f = ov.querySelector('form'), msg = ov.querySelector('.cx-msg');
    f.url.value = c.url || ''; f.token.value = c.token || '';
    ov.querySelector('.cx-sec').onclick = function () { ov.remove(); };
    f.onsubmit = function (e) {
      e.preventDefault();
      var nova = { url: f.url.value.trim(), token: f.token.value.trim() };
      if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(nova.url)) { msg.textContent = 'A URL precisa começar com https://script.google.com e terminar em /exec.'; return; }
      var antiga = cfg(); gravarJSON(CFG_K, nova); msg.textContent = 'Testando…';
      chamar({ action: 'ping' }).then(function (j) {
        msg.textContent = 'Conectado à planilha “' + j.planilha + '”. Recarregando…';
        setTimeout(function () { location.reload(); }, 900);
      }, function (er) {
        if (antiga) gravarJSON(CFG_K, antiga); else { try { localStorage.removeItem(CFG_K); } catch (x) {} }
        msg.textContent = er.code === 'unauthorized' ? 'Chave incorreta.' : 'Não consegui falar com o servidor. Confira a URL e se a implantação está como “Qualquer pessoa”.';
      });
    };
  };

  window.cronoPrimeiroUso = function (seguir) {
    if (document.querySelector('.cx-ov')) return;
    var ov = document.createElement('div');
    ov.className = 'cx-ov';
    ov.innerHTML = '<div class="cx-card"><h2>Planilha conectada</h2>' +
      '<p>Ela ainda não tem registros do Crono. Se você já usava o Crono, importe o arquivo de backup (.json) para trazer seu histórico.</p>' +
      '<div class="cx-acts"><button type="button" class="cx-sec">Começar do zero</button><button type="button" class="cx-pri">Importar backup</button></div></div>';
    document.body.appendChild(ov);
    ov.querySelector('.cx-pri').onclick = function () { ov.remove(); window.cronoImportar(function () { window.cronoPrimeiroUso(seguir); }); };
    ov.querySelector('.cx-sec').onclick = function () { gravarJSON('crono-inicio-ok', true); ov.remove(); seguir(); };
  };

  /* ---------- importar backup (.json) ----------
     Só acrescenta: documentos que já existem na planilha não são tocados. */
  window.cronoImportar = function (aoCancelar) {
    if (!conectado()) { window.cronoConectar(); return; }
    var inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    if (typeof aoCancelar === 'function') inp.addEventListener('cancel', aoCancelar);
    inp.onchange = function () {
      var f = inp.files && inp.files[0]; if (!f) return;
      f.text().then(function (t) {
        var dados = JSON.parse(t);
        if (!dados || typeof dados !== 'object' || Array.isArray(dados)) throw new Error('formato');
        return chamar({ action: 'list' }).then(function (j) {
          var existe = {}; (j.docs || []).forEach(function (d) { existe[d.id] = 1; });
          var ops = Object.keys(dados).filter(function (id) { return !existe[id] && dados[id] && typeof dados[id] === 'object'; })
            .map(function (id) { return { op: 'set', id: id, data: dados[id] }; });
          var pulados = Object.keys(dados).length - ops.length;
          function lote() {
            if (!ops.length) return Promise.resolve();
            return chamar({ action: 'write', ops: ops.splice(0, 15) }).then(lote);
          }
          var total = ops.length;
          return lote().then(function () {
            alert('Importação concluída: ' + total + ' registro(s) novos' + (pulados ? ', ' + pulados + ' já existiam e foram mantidos.' : '.'));
            gravarJSON('crono-inicio-ok', true);
            location.reload();
          });
        });
      }).catch(function (e) { alert('Não consegui importar: ' + (e && (e.code || e.message) || 'arquivo inválido')); });
    };
    inp.click();
  };

  var css = document.createElement('style');
  css.textContent =
    '.cx-ov{position:fixed;inset:0;z-index:9999;background:rgb(10 17 28/.6);display:grid;place-items:center;padding:16px}' +
    '.cx-card{background:var(--surface,#fff);color:var(--ink,#0F1B2D);border-radius:18px;padding:22px;max-width:440px;width:100%;display:grid;gap:12px;font-family:var(--f-body,system-ui);box-shadow:0 20px 50px -20px #000}' +
    '.cx-card h2{font-family:var(--f-display,system-ui);font-size:22px;margin:0}.cx-card p{margin:0;color:var(--muted,#566172);font-size:14px}' +
    '.cx-card label{display:grid;gap:6px;font-weight:600;font-size:13px}.cx-card input{background:var(--sunk,#F4F3EE);border:1px solid var(--line,#DEDCD3);border-radius:10px;padding:10px 12px;font:inherit;color:inherit}' +
    '.cx-acts{display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap}.cx-acts button{border:0;border-radius:10px;padding:10px 16px;font-weight:700;cursor:pointer;font:inherit}' +
    '.cx-pri{background:var(--amber,#F2A541);color:#0F1B2D}.cx-sec{background:transparent;color:var(--muted,#566172)}.cx-msg{min-height:1.2em}' +
    '.cx-btn{border:1px solid var(--line,#ddd);border-bottom-width:2px;background:var(--surface,#fff);color:inherit;border-radius:10px;padding:9px 14px;font-weight:600;cursor:pointer;font:inherit}';
  document.head.appendChild(css);

  document.addEventListener('DOMContentLoaded', function () {
    // botão em Ajustes, dentro do cartão "Planilha Google"
    var h = Array.prototype.find.call(document.querySelectorAll('h2'), function (x) { return /Planilha Google/.test(x.textContent); });
    if (h) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'cx-btn';
      b.textContent = conectado() ? 'Servidor conectado · alterar' : 'Conectar servidor da planilha';
      b.onclick = window.cronoConectar;
      var bi = document.createElement('button');
      bi.type = 'button'; bi.className = 'cx-btn'; bi.textContent = 'Importar backup (.json)';
      bi.onclick = window.cronoImportar;
      var w = document.createElement('div'); w.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap';
      w.appendChild(b); w.appendChild(bi);
      h.insertAdjacentElement('afterend', w);
    }
    if (!conectado() && !lerJSON('crono-cx-adiado', false)) {
      setTimeout(function () { window.cronoConectar(); gravarJSON('crono-cx-adiado', true); }, 1200);
    }
  });
})();
