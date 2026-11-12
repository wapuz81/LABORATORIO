# LABORATORIO

Ferramentas do laboratório da obra das Trincheiras Vila São Pedro.

## Ensaios de Campo (app de celular) — `app/`

App web que funciona sem internet depois de aberto uma vez. O laboratorista lança os dados no campo,
vê o resultado na hora e exporta um arquivo `.json` para o PC.

- **Densidade in situ (frasco de areia):** mesmas fórmulas da ficha `SL - DENSIDADE IN-SITU`
  (K = 1−2, L = 9−6, M = 4/5, N = 10/7 − 11, P = (4−3)/12, R = 13/(100+h), S = 15/8).
  Critérios: GC ≥ 100% e umidade entre −2% e +1% da ótima.
- **Viga Benkelman:** D = (Li − Lf) × K por lado (LE, eixo, LD). Deflexão admissível pela tabela do
  projeto (camada × Número N da rua), a mesma da aba `Viga. Benk` do Resumo. Mostra também Dc = média + desvio.
- **Moldagem de concreto** e **Imprimação / Pintura de ligação (EAI, RR-1C)**: ver telas do app.
- **CBUQ teor de betume:** ligante = mistura − agregado, teor = ligante / mistura (±0,3% do projeto);
  granulometria por retido acumulado, passante comparado com a Faixa B ou C (mesmos limites do Resumo).
  O lançador grava a linha na aba `CBUQ Gl Betume` do Resumo e cria a aba nova na ficha TEOR DE BETUME da faixa (`fichasBetume` no config.json).
- **Cadastros:** ruas com o N, materiais com umidade ótima e densidades, padrões (areia, K da viga).
- Os dados ficam no aparelho (armazenamento do navegador) até serem exportados.

Para instalar no celular: abrir o endereço publicado no Chrome e usar "Adicionar à tela inicial".

## Lançador (PC com Excel) — `lancador/`

1. Copie a pasta `lancador` para o PC e ajuste os caminhos em `config.json`.
2. Salve os `.json` exportados pelo app na `pastaEntrada`.
3. Dê dois cliques em `LANCAR ENSAIOS.bat`.

Para cada ensaio ele copia a última aba da ficha do material como modelo, cria a aba nova
(`IN SITU-R52-MOLEDO-RUA...` ou `VIGA-R47-BGS-RUA...`) e acrescenta as linhas no Resumo
(`Mat. Granular` ou `Viga. Benk`), copiando formatos e fórmulas da linha de cima.

- `"modo": "teste"` (padrão): trabalha em cópias na `pastaSaidaTeste\aammdd-hhmm\`; os originais não são tocados.
- `"modo": "gravar"`: altera os originais; antes guarda a versão anterior na subpasta `ANTIGO`,
  registra os ensaios em `lancados.csv` (não lança duas vezes) e move os `.json` para `PROCESSADOS`.

### Lançamento automático (opcional)

- No app, o botão **Enviar para o Drive** compartilha o arquivo; salve na pasta `ENTRADA` do lançador (Google Drive).
- No PC, `LIGAR LANCAMENTO AUTOMATICO.bat` cria uma tarefa agendada que confere a `ENTRADA` a cada 5 minutos (`Vigiar-Entrada.ps1`) e roda o lançador sozinho, só no modo `gravar`. Ligue em um computador só. `DESLIGAR LANCAMENTO AUTOMATICO.bat` remove a tarefa.
- Arquivos lançados vão para `ENTRADA\PROCESSADOS`; os que tiveram erro, para `ENTRADA\COM ERRO`. Os logs ficam em `LOGS`.
