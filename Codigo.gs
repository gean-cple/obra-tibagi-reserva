/**
 * App de campo · LD 34,5 kV Tibagi – Reserva
 * Backend na planilha Google (Apps Script). Cole este código em Extensões > Apps Script.
 *
 * Aba "eventos": um registro por marcação feita no app (nada é apagado; desfazer é um novo registro).
 * Aba "projetos": qual empresa pode marcar cada projeto (pode ser editada aqui na planilha).
 * Aba "chaves":   códigos de acesso. Cada linha é um código, com a empresa dele. Ativa = SIM/NÃO.
 *                 Empresa "Copel" = fiscal: pode marcar e desfazer em qualquer projeto.
 *                 Para criar mais códigos: escreva a empresa (e o responsável) numa linha nova, deixe a
 *                 chave vazia e use o menu "App de campo > Gerar códigos que faltam".
 */

const ABA_EVENTOS = 'eventos';
const ABA_PROJETOS = 'projetos';
const ABA_CHAVES = 'chaves';
const CAB = ['id', 'recebido_em', 'marcado_em', 'data_execucao', 'nome', 'empresa', 'tipo', 'projeto', 'alvo', 'ref', 'lat', 'lon', 'precisao_m', 'obs', 'chave'];
const N_PUBLICO = 10; // colunas que a leitura devolve (sem GPS e sem chave)
const TIPOS = ['cava', 'poste', 'cabo', 'desfazer'];
const FISCAL = 'Copel';
const PREFIXO = { 'Serpa': 'SER', 'Campos Gerais': 'CGE', 'Eleng': 'ELE', 'Klem': 'KLE', 'Copel': 'COP' };

const PROJETOS_INICIAIS = [
  ['1749308', 'Serpa', 'Serpa (Paulo)', 5],
  ['1753234', 'Campos Gerais', 'Campos Gerais', 77],
  ['1751947', 'Eleng', 'Eleng', 42],
  ['1751918', 'Serpa', 'Serpa (Paulo)', 39],
  ['1751789', 'Serpa', 'Serpa (Paulo)', 67],
  ['1751790', 'Klem', 'Klem', 51],
  ['1751793', 'Klem', 'Klem', 53],
  ['1751794', 'Eleng', 'Eleng', 55],
  ['1751796', 'Serpa', 'Serpa (Ari)', 59],
  ['1751798', 'Serpa', 'Serpa (Ari)', 51],
];

const CHAVES_INICIAIS = [
  ['Serpa', '', 'frente Paulo'], ['Serpa', '', 'frente Paulo'], ['Serpa', '', 'frente Paulo'],
  ['Serpa', '', 'frente Ari'], ['Serpa', '', 'frente Ari'], ['Serpa', '', 'frente Ari'],
  ['Campos Gerais', '', ''], ['Campos Gerais', '', ''], ['Campos Gerais', '', ''],
  ['Eleng', '', ''], ['Eleng', '', ''], ['Eleng', '', ''],
  ['Klem', '', ''], ['Klem', '', ''], ['Klem', '', ''],
  ['Copel', '', 'fiscal 1'], ['Copel', '', 'fiscal 2'], ['Copel', 'Gean', 'gestor'],
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('App de campo')
    .addItem('Gerar códigos que faltam', 'gerarChaves')
    .addToUi();
}

/** Rode pelo editor (botão Executar). Pode rodar de novo: só cria o que estiver faltando. */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  let ev = ss.getSheetByName(ABA_EVENTOS);
  if (!ev) ev = ss.insertSheet(ABA_EVENTOS);
  ev.getRange(1, 1, 1, CAB.length).setValues([CAB]).setFontWeight('bold').setBackground('#f47b20').setFontColor('#ffffff');
  ev.setFrozenRows(1);

  let pj = ss.getSheetByName(ABA_PROJETOS);
  if (!pj) pj = ss.insertSheet(ABA_PROJETOS);
  if (pj.getLastRow() === 0) {
    pj.appendRow(['projeto', 'empresa', 'frente', 'meta_postes']);
    pj.getRange(2, 1, PROJETOS_INICIAIS.length, 4).setValues(PROJETOS_INICIAIS);
    pj.setFrozenRows(1);
    pj.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#101826').setFontColor('#ffffff');
    pj.getRange('A:A').setNumberFormat('@');
  }

  let ch = ss.getSheetByName(ABA_CHAVES);
  if (!ch) ch = ss.insertSheet(ABA_CHAVES);
  if (ch.getLastRow() === 0) {
    ch.appendRow(['chave', 'empresa', 'responsavel', 'ativa', 'obs']);
    ch.getRange(2, 1, CHAVES_INICIAIS.length, 5).setValues(CHAVES_INICIAIS.map(r => ['', r[0], r[1], 'SIM', r[2]]));
    ch.setFrozenRows(1);
    ch.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#101826').setFontColor('#ffffff');
    ch.getRange('A:A').setNumberFormat('@');
    ch.setColumnWidth(1, 110);
  }
  gerarChaves();

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 3) ss.deleteSheet(padrao);
}

/** Preenche um código aleatório em toda linha da aba "chaves" que tem empresa e não tem código. */
function gerarChaves() {
  const ch = SpreadsheetApp.getActive().getSheetByName(ABA_CHAVES);
  const n = ch.getLastRow() - 1;
  if (n < 1) return;
  const v = ch.getRange(2, 1, n, 4).getValues();
  const usados = new Set(v.map(r => String(r[0]).trim().toUpperCase()).filter(Boolean));
  const ALF = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  v.forEach((r, i) => {
    const emp = String(r[1]).trim();
    if (String(r[0]).trim() || !emp) return;
    let c;
    do {
      c = (PREFIXO[emp] || 'OBR') + '-';
      for (let k = 0; k < 5; k++) c += ALF[Math.floor(Math.random() * ALF.length)];
    } while (usados.has(c));
    usados.add(c);
    ch.getRange(i + 2, 1).setValue(c);
    if (!String(r[3]).trim()) ch.getRange(i + 2, 4).setValue('SIM');
  });
}

function empresaDoProjeto_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA_PROJETOS);
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  const m = {};
  v.forEach(r => { if (r[0]) m[String(r[0]).trim()] = String(r[1]).trim(); });
  return m;
}

/** código -> empresa, só dos códigos ativos */
function chaves_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA_CHAVES);
  const m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    const c = String(r[0]).trim().toUpperCase(), emp = String(r[1]).trim(), at = String(r[3]).trim().toUpperCase();
    if (c && emp && at !== 'NÃO' && at !== 'NAO' && at !== 'N') m[c] = emp;
  });
  return m;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Leitura.
 *  ?chave=XXX  -> confere um código: {ok, empresa}
 *  ?desde=N    -> linhas depois da linha N (sem GPS e sem chave)
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.chave !== undefined) {
    const emp = chaves_()[String(p.chave).trim().toUpperCase()];
    return json_(emp ? { ok: true, empresa: emp, fiscal: emp === FISCAL } : { ok: false, erro: 'código inválido ou desativado' });
  }
  const desde = Math.max(0, parseInt(p.desde || '0', 10) || 0);
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA_EVENTOS);
  const total = Math.max(sh.getLastRow() - 1, 0);
  let linhas = [];
  if (total > desde) {
    linhas = sh.getRange(2 + desde, 1, total - desde, N_PUBLICO).getDisplayValues();
  }
  return json_({ ok: true, total: total, cab: CAB.slice(0, N_PUBLICO), linhas: linhas, projetos: empresaDoProjeto_() });
}

/** Gravação: recebe {eventos:[...]} em texto JSON. Cada marcação precisa de um código ativo. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const lista = Array.isArray(body.eventos) ? body.eventos : [];
    const sh = SpreadsheetApp.getActive().getSheetByName(ABA_EVENTOS);
    const n = Math.max(sh.getLastRow() - 1, 0);
    const exist = n ? sh.getRange(2, 1, n, 6).getValues() : [];
    const empDoEvento = {};
    exist.forEach(r => { empDoEvento[String(r[0])] = String(r[5]); });
    const emp = empresaDoProjeto_();
    const CH = chaves_();
    const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
    const aceitos = [], recusados = [], novas = [];
    lista.forEach(ev => {
      const id = String(ev.id || '');
      if (!id) return;
      if (empDoEvento[id] !== undefined) { aceitos.push(id); return; }
      const chave = String(ev.chave || '').trim().toUpperCase();
      const quem = CH[chave];
      const recusa = m => recusados.push({ id: id, motivo: m });
      if (!quem) return recusa('código de acesso inválido ou desativado');
      const fiscal = quem === FISCAL;
      const proj = String(ev.projeto || '');
      if (TIPOS.indexOf(ev.tipo) < 0) return recusa('tipo inválido');
      if (!emp[proj]) return recusa('projeto desconhecido');
      if (!fiscal && emp[proj] !== quem) return recusa('projeto de outra empresa');
      if (!String(ev.nome || '').trim()) return recusa('sem nome');
      if (ev.tipo === 'desfazer') {
        const dono = empDoEvento[String(ev.ref || '')];
        if (dono === undefined) return recusa('marcação a desfazer não encontrada');
        if (!fiscal && dono !== quem) return recusa('só a própria empresa ou a fiscalização desfaz');
      }
      novas.push([id, agora, ev.marcado_em || '', ev.data_execucao || '', String(ev.nome).trim(), quem, ev.tipo, proj,
        ev.alvo || '', ev.ref || '', ev.lat || '', ev.lon || '', ev.precisao_m || '', ev.obs || '', chave]);
      empDoEvento[id] = quem;
      aceitos.push(id);
    });
    if (novas.length) {
      const r0 = sh.getLastRow() + 1;
      sh.getRange(r0, 1, novas.length, CAB.length).setNumberFormat('@').setValues(novas);
    }
    return json_({ ok: true, aceitos: aceitos, recusados: recusados, total: Math.max(sh.getLastRow() - 1, 0) });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}
