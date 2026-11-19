// Recebe os ensaios do app "Ensaios de Campo" e grava o arquivo na pasta ENTRADA do lançador (Google Drive).
// Como instalar: script.google.com > Novo projeto > cole este código > Implantar > Nova implantação >
// tipo "App da Web", executar como "Eu", acesso "Qualquer pessoa". Copie o endereço que termina em /exec.
// Ao trocar o código: Implantar > Gerenciar implantações > lápis > Versão: Nova versão > Implantar.
//
// SENHAS: cada pessoa do campo tem a sua, no arquivo "SENHAS ENVIO.txt" da pasta do lançador
// (a pasta acima da ENTRADA), uma linha por pessoa no formato   Nome;senha
// Para cortar o acesso de alguém, apague a linha dela. Não precisa reimplantar.
// A CHAVE abaixo é a do administrador e vale sempre.
var PASTA_ENTRADA = 'COLE_AQUI_O_ID_DA_PASTA_ENTRADA';
var CHAVE = 'COLE_AQUI_A_CHAVE';
var ARQ_SENHAS = 'SENHAS ENVIO.txt';
// PARÂMETROS DE PROJETO: planilha Google com os parâmetros (pasta do lançador). O app lê por ?param=1 e o
// Resumo do Excel lê as cópias .csv que este código grava na pasta PARAMETROS (ao lado da ENTRADA).
// Depois de colar o código, rode UMA VEZ a função ligarCopiaAutomatica (menu Executar) e autorize.
var PLANILHA_PARAMETROS = 'COLE_AQUI_O_ID_DA_PLANILHA_PARAMETROS';
var PASTA_COPIA_PARAMETROS = 'PARAMETROS';

function quemEnviou(senha) {
  senha = String(senha || '').trim();
  if (!senha) return null;
  if (senha === CHAVE) return 'adm';
  var pais = DriveApp.getFolderById(PASTA_ENTRADA).getParents();
  if (!pais.hasNext()) return null;
  var arqs = pais.next().getFilesByName(ARQ_SENHAS);
  if (!arqs.hasNext()) return null;
  var linhas = arqs.next().getBlob().getDataAsString('UTF-8').split(/\r?\n/);
  for (var i = 0; i < linhas.length; i++) {
    var p = linhas[i].split(';');
    if (p.length >= 2 && p[0].trim() && p[1].trim() === senha) return p[0].trim();
  }
  return null;
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var quem = quemEnviou(d.chave);
    if (!quem) return resposta({ ok: false, erro: 'senha errada' });
    if (d.acao === 'sync') return resposta(sincronizar(d, quem));
    var conteudo = String(d.conteudo || '');
    var pac = JSON.parse(conteudo);                       // precisa ser um JSON do app
    if (pac.app !== 'ensaios-campo-trincheira') return resposta({ ok: false, erro: 'arquivo não é do app' });
    var nome = String(d.nome || 'ensaios campo.json').replace(/[\\\/:*?"<>|]/g, ' ').replace(/\.json$/i, '');
    if (quem !== 'adm') nome += ' - ' + quem.replace(/[\\\/:*?"<>|;]/g, ' ');
    nome = nome.slice(0, 115) + '.json';
    DriveApp.getFolderById(PASTA_ENTRADA).createFile(nome, conteudo, 'application/json');
    return resposta({ ok: true, nome: nome, quem: quem });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}
// GET ?lista=1&chave=SENHA devolve os ensaios já lançados no PC (arquivo lancados.csv da ENTRADA),
// para o app de todos os celulares saber o que já foi feito e qual é o próximo número de registro.
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.param) {
    if (!quemEnviou(p.chave)) return resposta({ ok: false, erro: 'senha errada' });
    try { return resposta(lerParametros()); } catch (err) { return resposta({ ok: false, erro: String(err) }); }
  }
  if (!p.lista) return resposta({ ok: true, app: 'envio-ensaios-campo' });
  try {
    if (!quemEnviou(p.chave)) return resposta({ ok: false, erro: 'senha errada' });
    var arqs = DriveApp.getFolderById(PASTA_ENTRADA).getFilesByName('lancados.csv');
    if (!arqs.hasNext()) return resposta({ ok: true, itens: [] });
    var txt = arqs.next().getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
    var linhas = Utilities.parseCsv(txt);
    var cab = linhas.shift() || [];
    var col = {}; for (var i = 0; i < cab.length; i++) col[String(cab[i]).trim()] = i;
    var itens = [];
    for (var k = 0; k < linhas.length; k++) {
      var l = linhas[k]; if (!l || !l[col.id]) continue;
      var arq = String(l[col.arquivo] || '');
      var m = arq.match(/ - ([^-]+?)\.json(\.txt)?$/i);
      itens.push({ id: l[col.id], tipo: l[col.tipo], registro: l[col.registro], rua: l[col.rua], data: l[col.data],
                   lancado: l[col.lancado], quem: m && !/ensaios campo/i.test(m[1]) ? m[1].trim() : '' });
    }
    return resposta({ ok: true, itens: itens });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}
// ---------- Parâmetros de projeto ----------
// cada aba vira uma lista de linhas {título da coluna: valor}; o cabeçalho é a 1ª linha com 2 ou mais células
function tabelasParametros() {
  var ss = SpreadsheetApp.openById(PLANILHA_PARAMETROS), tz = ss.getSpreadsheetTimeZone(), abas = {};
  ss.getSheets().forEach(function (sh) {
    var nome = sh.getName(); if (/^LEIA/i.test(nome)) return;
    var v = sh.getDataRange().getValues(), h = -1;
    for (var i = 0; i < v.length && h < 0; i++) { if (v[i].filter(function (x) { return x !== ''; }).length >= 2) h = i; }
    if (h < 0) { abas[nome] = { cab: [], linhas: [] }; return; }
    var cab = v[h].map(function (x) { return String(x).replace(/\s+/g, ' ').trim(); });
    var linhas = [];
    for (var r = h + 1; r < v.length; r++) {
      if (v[r].every(function (x) { return x === ''; })) continue;
      linhas.push(v[r].map(function (x) { return x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x; }));
    }
    abas[nome] = { cab: cab, linhas: linhas };
  });
  return abas;
}
function lerParametros() {
  var t = tabelasParametros(), abas = {};
  Object.keys(t).forEach(function (n) {
    abas[n] = t[n].linhas.map(function (l) { var o = {}; t[n].cab.forEach(function (c, j) { if (c) o[c] = l[j]; }); return o; });
  });
  var f = DriveApp.getFileById(PLANILHA_PARAMETROS);
  return { ok: true, atualizado: f.getLastUpdated().toISOString(), abas: abas };
}
// grava uma cópia .csv de cada aba na pasta PARAMETROS (o Google Drive do PC sincroniza; o Excel lê dali)
function copiarParametros() {
  var t = tabelasParametros();
  var raiz = DriveApp.getFolderById(PASTA_ENTRADA).getParents().next();
  var it = raiz.getFoldersByName(PASTA_COPIA_PARAMETROS);
  var pasta = it.hasNext() ? it.next() : raiz.createFolder(PASTA_COPIA_PARAMETROS);
  var q = function (x) {
    if (typeof x === 'number') return String(x);
    x = String(x == null ? '' : x);
    return /[",\r\n]/.test(x) ? '"' + x.replace(/"/g, '""') + '"' : x;
  };
  Object.keys(t).forEach(function (n) {
    var txt = '\ufeff' + [t[n].cab].concat(t[n].linhas).map(function (l) { return l.map(q).join(','); }).join('\r\n') + '\r\n';
    var nome = n + '.csv', fs = pasta.getFilesByName(nome);
    if (fs.hasNext()) fs.next().setContent(txt); else pasta.createFile(nome, txt, 'text/csv');
  });
}
function aoMudarParametros(e) { copiarParametros(); }
// rodar UMA VEZ pelo editor: liga a cópia automática a cada alteração da planilha e já faz a primeira cópia
function ligarCopiaAutomatica() {
  ScriptApp.getProjectTriggers().forEach(function (g) { if (g.getHandlerFunction() === 'aoMudarParametros') ScriptApp.deleteTrigger(g); });
  ScriptApp.newTrigger('aoMudarParametros').forSpreadsheet(PLANILHA_PARAMETROS).onChange().create();
  copiarParametros();
}
// ---------- Ensaios compartilhados entre os celulares ----------
// Todos os ensaios ficam na planilha Google "ENSAIOS APP" (criada sozinha na pasta do lançador, uma linha por ensaio).
// Cada celular manda o que mudou e recebe o que os outros mudaram (POST {acao:'sync'}). Vale a última alteração.
// Quando um ensaio é CONCLUÍDO no app, este código grava o arquivo dele na ENTRADA e o PC lança sozinho, sem exportar.
var NOME_PLANILHA_ENSAIOS = 'ENSAIOS APP';
var COLS = ['id', 'tipo', 'registro', 'data', 'rua / local', 'autor', 'última alteração por', 'alterado em', 'seq', 'situação', 'enviado ao PC', 'dados (não mexer)'];
function planilhaEnsaios() {
  var pr = PropertiesService.getScriptProperties(), id = pr.getProperty('PLANILHA_ENSAIOS'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(NOME_PLANILHA_ENSAIOS);
    var f = DriveApp.getFileById(ss.getId()), raiz = DriveApp.getFolderById(PASTA_ENTRADA).getParents().next();
    f.moveTo(raiz);
    var sh0 = ss.getSheets()[0]; sh0.setName('ENSAIOS');
    sh0.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold').setBackground('#DDEBF7');
    sh0.setFrozenRows(1); sh0.setColumnWidth(COLS.length, 120);
    pr.setProperty('PLANILHA_ENSAIOS', ss.getId());
  }
  return ss.getSheetByName('ENSAIOS') || ss.getSheets()[0];
}
// rodar UMA VEZ pelo editor depois de colar o código: cria a planilha ENSAIOS APP e pede a autorização
function criarPlanilhaEnsaios() { var sh = planilhaEnsaios(); Logger.log('Planilha ENSAIOS APP: ' + sh.getParent().getUrl()); }
function chaveNum(o) { return o.tipo === 'taxa' ? 'taxa|' + (o.material === 'RR-1C' ? 'RR-1C' : 'EAI') : o.tipo; }
function sincronizar(d, quem) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sh = planilhaEnsaios(), pr = PropertiesService.getScriptProperties();
    var seq = Number(pr.getProperty('SEQ') || 0);
    var n = sh.getLastRow(), linhas = n > 1 ? sh.getRange(2, 1, n - 1, COLS.length).getValues() : [];
    var porId = {}; linhas.forEach(function (l, i) { if (l[0]) porId[l[0]] = i; });
    var recebidos = {}, renumerados = [];
    (d.itens || []).forEach(function (it) {
      if (!it || !it.id) return;
      var i = porId[it.id], atual = i == null ? null : linhas[i];
      var antigo = atual ? JSON.parse(atual[11] || '{}') : null;
      if (antigo && Number(antigo.atualizado || 0) > Number(it.atualizado || 0)) return;   // a planilha já tem uma alteração mais nova
      var o = it.apagado ? (antigo || { id: it.id }) : it.dados;
      if (!o) return;
      o.atualizado = Number(it.atualizado || Date.now());
      if (!antigo) o.autor = o.autor || quem;
      else o.autor = antigo.autor || o.autor || quem;
      o.alteradoPor = quem;
      // ensaio novo com número automático já usado por outro celular: pega o próximo livre
      if (!antigo && !it.apagado && o.regAuto) {
        var k = chaveNum(o), usados = {}, maior = 0;
        linhas.forEach(function (l) { if (!l[0] || l[9] === 'apagado') return; var x = JSON.parse(l[11] || '{}');
          if (chaveNum(x) !== k || /teste/i.test(String(x.registro))) return; usados[String(x.registro).trim()] = 1; maior = Math.max(maior, parseInt(x.registro, 10) || 0); });
        if (usados[String(o.registro).trim()]) { o.registro = String(maior + 1); renumerados.push({ id: o.id, registro: o.registro }); }
      }
      seq++;
      var sit = it.apagado ? 'apagado' : (o.concluido ? 'concluído' : 'rascunho');
      var enviado = atual ? atual[10] : '';
      if (sit === 'concluído' && it.pacote && o.tipo !== 'cimento' && String(enviado) !== String(o.atualizado)) {
        var nome = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyMMdd HHmm') + ' - ' + o.tipo + ' R' + (o.registro || '') + ' - ' + quem;
        nome = nome.replace(/[\\\/:*?"<>|;]/g, ' ').slice(0, 110) + '.json';
        var pac = { app: 'ensaios-campo-trincheira', versao: 1, exportado: new Date().toISOString(), ensaios: [it.pacote] };
        DriveApp.getFolderById(PASTA_ENTRADA).createFile(nome, JSON.stringify(pac, null, 1), 'application/json');
        enviado = String(o.atualizado);
      }
      if (sit === 'apagado') o.apagado = true;
      var linha = [o.id, o.tipo || '', String(o.registro || ''), o.data || '', o.rua || o.local || '', o.autor || '', quem,
                   new Date(o.atualizado), seq, sit, enviado, JSON.stringify(o)];
      if (i == null) { linhas.push(linha); porId[o.id] = linhas.length - 1; sh.appendRow(linha); }
      else { linhas[i] = linha; sh.getRange(i + 2, 1, 1, COLS.length).setValues([linha]); }
      sh.getRange(porId[o.id] + 2, 3).setNumberFormat('@').setValue(String(o.registro || ''));
      recebidos[o.id] = 1;
    });
    pr.setProperty('SEQ', String(seq));
    var desde = Number(d.desde || 0), mudados = [];
    linhas.forEach(function (l) {
      if (!l[0] || Number(l[8]) <= desde) return;
      var o = JSON.parse(l[11] || '{}'); o.enviadoPC = !!l[10] && String(l[10]) === String(o.atualizado);
      mudados.push(o);
    });
    return { ok: true, quem: quem, seq: seq, itens: mudados, renumerados: renumerados };
  } finally { lock.releaseLock(); }
}
function resposta(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
