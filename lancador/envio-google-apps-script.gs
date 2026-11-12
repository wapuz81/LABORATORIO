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
function resposta(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
