/**
 * CRONO · servidor da planilha (Google Apps Script)
 * ------------------------------------------------------------
 * Este script fica DENTRO da planilha "Cartão de Ponto".
 * Ele guarda os dados do app e atualiza as abas de relatório.
 *
 * Regras de segurança:
 *  - NUNCA apaga dados: "apagar" só marca a linha como apagada na aba _dados.
 *  - Toda gravação é registrada na aba _historico (só acrescenta linhas).
 *  - As abas _dados e _historico não podem ser limpas pelo app.
 *  - Só aceita pedidos com a chave secreta gerada em configurar().
 *
 * Como instalar (uma vez só):
 *  1. Na planilha: Extensões › Apps Script. Apague o conteúdo do arquivo
 *     Código.gs e cole este arquivo inteiro. Salve.
 *  2. Escolha a função "configurar" no topo e clique em Executar.
 *     Autorize o acesso. A chave aparece no "Registro de execução".
 *  3. Implantar › Nova implantação › tipo "App da Web".
 *     Executar como: Eu. Quem pode acessar: Qualquer pessoa.
 *     Copie a URL que termina em /exec.
 *  4. No Crono: Ajustes › Conectar planilha. Cole a URL e a chave.
 */

const ABA_DADOS = '_dados';
const ABA_HIST = '_historico';
const ABAS_RELATORIO = ['Ponto', 'Clientes', 'Blocos', 'Pausas'];
const PROP_CHAVE = 'CRONO_CHAVE';

/* ---------- instalação ---------- */

function configurar() {
  const props = PropertiesService.getScriptProperties();
  let chave = props.getProperty(PROP_CHAVE);
  if (!chave) {
    chave = Utilities.getUuid().replace(/-/g, '');
    props.setProperty(PROP_CHAVE, chave);
  }
  garantirAbas_();
  Logger.log('Chave do Crono (cole no app): ' + chave);
  return chave;
}

/** Opcional: cópia completa da planilha no Drive toda segunda às 6h. */
function ativarCopiaSemanal() {
  const ja = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'copiaSemanal');
  if (!ja) {
    ScriptApp.newTrigger('copiaSemanal').timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  }
  Logger.log('Cópia semanal ativada.');
}

function copiaSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nome = ss.getName() + ' · cópia ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  DriveApp.getFileById(ss.getId()).makeCopy(nome);
}

function garantirAbas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let d = ss.getSheetByName(ABA_DADOS);
  if (!d) {
    d = ss.insertSheet(ABA_DADOS);
    d.getRange(1, 1, 1, 4).setValues([['id', 'json', 'atualizado', 'apagado']]);
    d.setFrozenRows(1);
  }
  let h = ss.getSheetByName(ABA_HIST);
  if (!h) {
    h = ss.insertSheet(ABA_HIST);
    h.getRange(1, 1, 1, 4).setValues([['quando', 'id', 'acao', 'json']]);
    h.setFrozenRows(1);
  }
  return { d: d, h: h };
}

/* ---------- entrada HTTP ---------- */

function doGet(e) {
  return responder_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let corpo = {};
  try { corpo = JSON.parse(e.postData.contents); } catch (x) { corpo = {}; }
  return responder_(corpo);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function responder_(p) {
  const chave = PropertiesService.getScriptProperties().getProperty(PROP_CHAVE);
  if (!chave) return json_({ ok: false, code: 'sem_chave', msg: 'Rode configurar() no Apps Script.' });
  if (String(p.token || '') !== chave) return json_({ ok: false, code: 'unauthorized', msg: 'Chave inválida.' });

  // a IA não usa a trava da planilha, para não atrasar as gravações
  if (p.action === 'ia') {
    try { return json_(ia_(p.body || {})); }
    catch (err) { return json_({ ok: false, code: 'ia_erro', msg: String(err && err.message || err) }); }
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const abas = garantirAbas_();
    switch (p.action) {
      case 'ping': return json_({ ok: true, planilha: SpreadsheetApp.getActiveSpreadsheet().getName() });
      case 'list': return json_({ ok: true, docs: listar_(abas.d) });
      case 'get': return json_(Object.assign({ ok: true }, obter_(abas.d, p.id)));
      case 'write': return json_({ ok: true, res: (p.ops || []).map(op => gravar_(abas, op)) });
      case 'clear': return json_(limpar_(p));
      case 'values': return json_(valores_(p));
      default: return json_({ ok: false, code: 'acao_desconhecida' });
    }
  } catch (err) {
    return json_({ ok: false, code: err && err.code ? err.code : 'tool_error', msg: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- banco de dados (aba _dados) ---------- */

const FUSO = 'America/Sao_Paulo';
function agora_() { return Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm:ss'); }

// O id é sempre texto. Se a planilha tiver convertido "2026-10-01" em data, desfaz aqui.
function idTexto_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function linhas_(d) {
  const n = d.getLastRow() - 1;
  if (n < 1) return [];
  return d.getRange(2, 1, n, 4).getValues().map(r => { r[0] = idTexto_(r[0]); return r; });
}

function apagada_(r) { return r[3] === true || String(r[3]).toUpperCase() === 'TRUE'; }

// procura de baixo para cima: se houver linhas repetidas, vale a ativa mais recente
// (se todas estiverem marcadas como apagadas, devolve a mais recente delas)
function achar_(rows, id) {
  let apagadaIdx = -1;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] !== String(id)) continue;
    if (!apagada_(rows[i])) return i;
    if (apagadaIdx < 0) apagadaIdx = i;
  }
  return apagadaIdx;
}

function listar_(d) {
  const ult = {};
  linhas_(d).forEach(r => { if (r[0] && !apagada_(r)) ult[r[0]] = r; });
  return Object.keys(ult).map(id => ult[id])
    .map(r => { let v = null; try { v = JSON.parse(r[1]); } catch (x) {} return { id: r[0], data: v }; })
    .filter(x => x.data);
}

// grava uma linha nova com a coluna do id em formato texto
function novaLinha_(aba, valores) {
  const linha = aba.getLastRow() + 1;
  aba.getRange(linha, 1, 1, 2).setNumberFormat('@');
  aba.getRange(linha, 1, 1, valores.length).setValues([valores]);
}

function obter_(d, id) {
  const rows = linhas_(d), i = achar_(rows, id);
  if (i < 0 || apagada_(rows[i])) return { exists: false };
  try { return { exists: true, data: JSON.parse(rows[i][1]) }; } catch (x) { return { exists: false }; }
}

function mesclar_(base, patch) {
  const out = Object.assign({}, base);
  Object.keys(patch).forEach(k => {
    const a = out[k], b = patch[k];
    const ehObj = v => v && typeof v === 'object' && !Array.isArray(v);
    out[k] = (ehObj(a) && ehObj(b)) ? mesclar_(a, b) : b;
  });
  return out;
}

function gravar_(abas, op) {
  const d = abas.d, h = abas.h;
  const rows = linhas_(d), i = achar_(rows, op.id);
  const agora = agora_();
  let antes = null;
  if (i >= 0) { try { antes = JSON.parse(rows[i][1]); } catch (x) {} }
  const apagada = i >= 0 && apagada_(rows[i]);

  if (op.op === 'delete') {
    // nunca apaga: só marca como apagado e guarda o conteúdo no histórico
    if (i >= 0) d.getRange(i + 2, 3, 1, 2).setValues([[agora, true]]);
    novaLinha_(h, [agora, op.id, 'apagado (marcado)', JSON.stringify(antes)]);
    return { id: op.id, ok: true };
  }

  let novo;
  if (op.op === 'update') {
    if (i < 0 || apagada || !antes) return { id: op.id, ok: false, code: 'invalid_argument' };
    novo = mesclar_(antes, op.data || {});
  } else {
    novo = op.data || {};
  }
  const txt = JSON.stringify(novo);
  if (txt.length > 49000) return { id: op.id, ok: false, code: 'muito_grande' };
  if (i >= 0) d.getRange(i + 2, 2, 1, 3).setValues([[txt, agora, '']]);
  else novaLinha_(d, [String(op.id), txt, agora, '']);
  novaLinha_(h, [agora, String(op.id), op.op, txt]);
  return { id: op.id, ok: true };
}

/* ---------- ClaudIA (IA via API da Anthropic) ----------
   A chave da Anthropic fica só aqui, em Configurações do projeto ›
   Propriedades do script › ANTHROPIC_KEY. Nunca vai para o app nem para o GitHub. */

const IA_MODELO = 'claude-haiku-5-5';
const IA_LIMITE_DIA = 150; // trava de custo: no máximo 150 chamadas por dia

function ia_(body) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, code: 'sem_ia', msg: 'A IA ainda não foi ativada no servidor (falta a chave da Anthropic).' };

  const dia = 'IA_USO_' + Utilities.formatDate(new Date(), FUSO, 'yyyyMMdd');
  const usados = Number(props.getProperty(dia) || 0);
  if (usados >= IA_LIMITE_DIA) return { ok: false, code: 'rate_limited', msg: 'Limite diário da ClaudIA atingido.' };
  props.setProperty(dia, String(usados + 1));

  const pedido = {
    model: IA_MODELO,
    max_tokens: Math.min(Number(body.max_tokens) || 1024, 2048),
    messages: body.messages || []
  };
  if (body.system) pedido.system = body.system;
  if (Array.isArray(body.tools) && body.tools.length) pedido.tools = body.tools;

  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(pedido),
    muteHttpExceptions: true
  });
  const st = r.getResponseCode();
  let j = null; try { j = JSON.parse(r.getContentText()); } catch (x) {}
  if (st === 200 && j) return { ok: true, resp: j };
  const msg = j && j.error && j.error.message ? j.error.message : ('HTTP ' + st);
  if (st === 401 || st === 403) return { ok: false, code: 'ia_chave', msg: 'A chave da Anthropic foi recusada.' };
  if (st === 429 || st === 529) return { ok: false, code: 'rate_limited', msg: msg };
  if (/credit|balance/i.test(msg)) return { ok: false, code: 'ia_credito', msg: 'Sem créditos na conta da Anthropic.' };
  return { ok: false, code: 'ia_erro', msg: msg };
}

/** Rode no editor uma vez depois de colar a chave: autoriza e testa a IA. */
function testarIA() {
  const r = ia_({ max_tokens: 50, messages: [{ role: 'user', content: 'Responda só: ClaudIA pronta.' }] });
  Logger.log(r.ok ? 'OK: ' + r.resp.content.map(c => c.text || '').join('') : 'ERRO: ' + r.code + ' · ' + r.msg);
}

/* ---------- abas de relatório (Ponto, Clientes, Blocos, Pausas) ---------- */

function planilha_(id) {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  return (!id || id === ativa.getId()) ? ativa : SpreadsheetApp.openById(id);
}

function abaPermitida_(range) {
  const nome = String(range).split('!')[0].replace(/'/g, '');
  return ABAS_RELATORIO.indexOf(nome) >= 0;
}

function limpar_(p) {
  const ss = planilha_(p.spreadsheetId);
  const rs = p.ranges || [];
  for (const r of rs) if (!abaPermitida_(r)) return { ok: false, code: 'tool_error', msg: 'Só as abas de relatório podem ser regravadas.' };
  rs.forEach(r => { const g = ss.getRange(r); g.clearContent(); });
  return { ok: true };
}

function valores_(p) {
  if (!abaPermitida_(p.range)) return { ok: false, code: 'tool_error', msg: 'Só as abas de relatório podem ser regravadas.' };
  const ss = planilha_(p.spreadsheetId);
  const v = p.values || [];
  if (!v.length) return { ok: true };
  const nome = String(p.range).split('!')[0].replace(/'/g, '');
  const aba = ss.getSheetByName(nome);
  if (!aba) return { ok: false, code: 'tool_error', msg: 'Aba não encontrada: ' + nome };
  const inicio = String(p.range).split('!')[1].split(':')[0];
  const cel = aba.getRange(inicio);
  const largura = Math.max.apply(null, v.map(r => r.length));
  const linhas = v.map(r => { const c = r.slice(); while (c.length < largura) c.push(''); return c; });
  aba.getRange(cel.getRow(), cel.getColumn(), linhas.length, largura).setValues(linhas);
  return { ok: true };
}
