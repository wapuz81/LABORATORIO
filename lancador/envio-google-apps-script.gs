// Recebe os ensaios do app "Ensaios de Campo" e grava o arquivo na pasta ENTRADA do lançador (Google Drive).
// Como instalar: script.google.com > Novo projeto > cole este código > Implantar > Nova implantação >
// tipo "App da Web", executar como "Eu", acesso "Qualquer pessoa". Copie o endereço que termina em /exec.
var PASTA_ENTRADA = 'COLE_AQUI_O_ID_DA_PASTA_ENTRADA';
var CHAVE = 'COLE_AQUI_A_CHAVE';

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.chave !== CHAVE) return resposta({ ok: false, erro: 'chave errada' });
    var conteudo = String(d.conteudo || '');
    var pac = JSON.parse(conteudo);                       // precisa ser um JSON do app
    if (pac.app !== 'ensaios-campo-trincheira') return resposta({ ok: false, erro: 'arquivo não é do app' });
    var nome = String(d.nome || 'ensaios campo.json').replace(/[\\\/:*?"<>|]/g, ' ').slice(0, 120);
    if (!/\.json$/i.test(nome)) nome += '.json';
    DriveApp.getFolderById(PASTA_ENTRADA).createFile(nome, conteudo, 'application/json');
    return resposta({ ok: true, nome: nome });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}
function doGet() { return resposta({ ok: true, app: 'envio-ensaios-campo' }); }
function resposta(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
