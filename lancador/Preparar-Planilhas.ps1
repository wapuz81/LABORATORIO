# Preparar-Planilhas.ps1
# Prepara as planilhas do lançador para começar do zero:
#   - Resumo: corrige fórmulas/critérios, deixa uma linha-modelo com fórmulas em cada aba e apaga as linhas de ensaio
#   - Fichas: deixa só uma aba "MODELO" limpa (a que o lançador copia), corrige fórmulas
#   - Formatação: fórmulas em azul claro, cabeçalho travado, filtro no cabeçalho
# NÃO altera os originais: trabalha em cópias gravadas em SAIDA\<aammdd-HHmm> PREPARO.
param([string]$Pasta = $PSScriptRoot, [switch]$SemPausa)
$ErrorActionPreference = "Stop"
function Pausa { if (-not $SemPausa) { Read-Host "Enter para fechar" | Out-Null } }
$log = New-Object System.Collections.Generic.List[string]
function Log($m) { Write-Host $m; $log.Add($m) }

$cfg = Get-Content (Join-Path $Pasta "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$saida = Join-Path $Pasta ("SAIDA\" + (Get-Date).ToString("yyMMdd-HHmm") + " PREPARO")
New-Item -ItemType Directory -Force -Path $saida | Out-Null

# lista: caminho original -> tipo
$arquivos = [ordered]@{}
$arquivos[$cfg.resumo] = "resumo"
foreach ($par in @(@("fichasDensidade","densidade"), @("fichasBenkelman","benkelman"), @("fichasTaxa","taxa"), @("fichasBetume","betume"))) {
  $m = $cfg.($par[0])
  if ($m) { foreach ($p in $m.PSObject.Properties) { if (-not $arquivos.Contains($p.Value)) { $arquivos[$p.Value] = $par[1] } } }
}
if ($cfg.fichaConcreto) { $arquivos[$cfg.fichaConcreto] = "concreto" }

$AZUL = 16247773      # DDEBF7 (R=DD G=EB B=F7) em BGR
$xlFormulas = -4123; $xlUp = -4162

function PintaFormulas($rng) {
  try { $f = $rng.SpecialCells($xlFormulas) } catch { return }
  $f.Interior.Color = $AZUL
}
function Limpa($ws, $enderecos) {
  foreach ($e in $enderecos) {
    foreach ($cel in $ws.Range($e).Cells) {
      if ($cel.HasFormula) { continue }
      if ($cel.MergeCells) { if ($cel.Address() -eq $cel.MergeArea.Cells(1, 1).Address()) { $cel.MergeArea.ClearContents() | Out-Null } }
      else { $cel.ClearContents() | Out-Null }
    }
  }
}
function SoModelo($wb) {
  # mantém só a última aba (a que o lançador usa como modelo) e renomeia para MODELO
  $n = $wb.Worksheets.Count
  $modelo = $wb.Worksheets.Item($n)
  $modelo.Visible = -1
  for ($i = $n - 1; $i -ge 1; $i--) { $wb.Worksheets.Item($i).Visible = -1; $wb.Worksheets.Item($i).Delete() | Out-Null }
  # folhas de gráfico soltas, se houver
  for ($i = $wb.Charts.Count; $i -ge 1; $i--) { $wb.Charts.Item($i).Delete() | Out-Null }
  $modelo.Name = "MODELO"
  return $modelo
}

function PreparaResumo($wb) {
  # aba, linha do cabeçalho, 1ª linha de dados, última coluna, fórmulas da linha-modelo, textos de cabeçalho
  $q = '"'
  $abas = @(
    @{ nome = "Mat. Granular"; cab = 9; prim = 10; ult = "V"; f = @{
        "K" = "=IFERROR(LOOKUP(2,1/(('Proj. Exec.'!`$A`$3:`$A`$12=`$E10)*('Proj. Exec.'!`$C`$3:`$C`$12<=`$B10)),'Proj. Exec.'!`$H`$3:`$H`$12),$q$q)"
        "L" = "=IF(J10=$q$q,$q$q,IF(N(K10)=0,${q}proj.?$q,IF(AND(ROUND(J10-K10,4)<=0.01,ROUND(J10-K10,4)>=-0.02),${q}ok$q,${q}verif.$q)))"
        "O" = "=IF(N10=$q$q,$q$q,IF(N10<=IF(ISNUMBER(SEARCH(${q}subleito$q,`$F10)),2%,0.5%)+0.000000001,${q}ok$q,${q}verif.$q))"
        "R" = "=IFERROR(LOOKUP(2,1/(('Proj. Exec.'!`$A`$3:`$A`$12=`$E10)*('Proj. Exec.'!`$C`$3:`$C`$12<=`$B10)),'Proj. Exec.'!`$J`$3:`$J`$12),$q$q)"
      }; txt = @{ "N8" = "Exp(%): ≤0,5% (subleito ≤2%)" } },
    @{ nome = "Viga. Benk"; cab = 5; prim = 6; ult = "L"; f = @{}; txt = @{} },
    @{ nome = "Concreto"; cab = 5; prim = 6; ult = "S"; f = @{}; txt = @{} },
    @{ nome = "Imprimação"; cab = 5; prim = 6; ult = "I"; textoA = $true; f = @{
        "H" = "=IF(E6=$q$q,$q$q,1.2)"
        "I" = "=IF(OR(G6=$q$q,H6=$q$q),$q$q,IF(ABS(G6-H6)<=0.2+0.000000001,${q}ok$q,${q}verif.$q))"
      }; txt = @{ "G4" = "Tolerância: ±0,2 l/m² da taxa de projeto (EAI 1,20 l/m²)" } },
    @{ nome = "P. Ligação"; cab = 5; prim = 6; ult = "I"; textoA = $true; f = @{
        "H" = "=IF(E6=$q$q,$q$q,1)"
        "I" = "=IF(OR(G6=$q$q,H6=$q$q),$q$q,IF(ABS(G6-H6)<=0.2+0.000000001,${q}ok$q,${q}verif.$q))"
      }; txt = @{ "G4" = "Tolerância: ±0,2 l/m² da taxa de projeto" } },
    @{ nome = "CBUQ %V. Compactação"; cab = 14; prim = 15; ult = "AG"; f = @{}; txt = @{} },
    @{ nome = "CBUQ Gl Betume"; cab = 14; prim = 15; ult = "AG"; f = @{}; txt = @{} }
  )
  foreach ($a in $abas) {
    $ws = $wb.Worksheets.Item($a.nome)
    $ws.Activate()
    if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
    $ws.Rows.Hidden = $false
    $prim = $a.prim
    $fim = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    $antes = 0
    for ($c = 1; $c -le 7; $c++) { $r = $ws.Cells($ws.Rows.Count, $c).End($xlUp).Row; if ($r -gt $antes) { $antes = $r } }
    if ($fim -gt $prim) {
      $bloco = $ws.Range("A$($prim):$($a.ult)$fim")
      $bloco.UnMerge() | Out-Null
      $ws.Rows("$($prim + 1):$fim").Delete() | Out-Null
    }
    # linha-modelo: fórmulas corrigidas + dados limpos
    foreach ($col in $a.f.Keys) { $ws.Range("$col$prim").Formula = $a.f[$col] }
    $ultCol = $ws.Range("$($a.ult)1").Column
    for ($c = 1; $c -le $ultCol; $c++) {
      $cel = $ws.Cells($prim, $c)
      if (-not $cel.HasFormula) { $cel.ClearContents() | Out-Null }
    }
    foreach ($k in $a.txt.Keys) { $ws.Range($k).MergeArea.Cells(1, 1).Value2 = [string]$a.txt[$k] }
    # formatação
    $linha = $ws.Range("A$($prim):$($a.ult)$prim")
    $linha.Interior.Pattern = -4142     # sem preenchimento
    PintaFormulas $linha
    $ws.Range("B$prim").NumberFormatLocal = "dd/mm/aa"
    if ($a.textoA) { $ws.Range("A$prim").NumberFormat = "@" }
    $ws.Rows($prim).RowHeight = 15
    $ws.Parent.Windows.Item(1).FreezePanes = $false
    $ws.Application.Goto($ws.Range("A1"), $true) | Out-Null
    $ws.Range("A$prim").Select() | Out-Null
    $ws.Parent.Windows.Item(1).FreezePanes = $true
    $ws.Range("A$($a.cab):$($a.ult)$($a.cab)").AutoFilter() | Out-Null
    Log ("   {0}: {1} linha(s) de ensaio apagadas; linha-modelo = {2}" -f $a.nome, [math]::Max(0, $antes - $prim + 1), $prim)
  }
  $wb.Worksheets.Item("Mat. Granular").Activate()
}

function PreparaFicha($wb, $tipo) {
  $q = '"'
  $ws = SoModelo $wb
  switch ($tipo) {
    "densidade" {
      Limpa $ws @("A10:J17", "Q10:Q17", "I2", "B3", "B4")
      # GC dava #VALOR! com o retido 3/4" em branco
      for ($l = 10; $l -le 17; $l++) { if ($ws.Range("N$l").HasFormula) { $ws.Range("N$l").Formula = "=IF(L$l=$q$q,$q$q,(L$l/`$R`$5)-N(M$l))" } }
      for ($l = 18; $l -le 30; $l++) { if ("$($ws.Range("A$l").Value2)" -like "OBSERVA*") { $ws.Range("A$l").Value2 = "OBSERVAÇÃO: "; break } }
    }
    "benkelman" {
      Limpa $ws @("A12:D27", "F12:I27", "K12:N27", "H4", "M4", "C5", "C6")
      for ($l = 12; $l -le 27; $l++) {
        $ws.Range("E$l").Formula = "=IF(D$l=$q$q,$q$q,(B$l-D$l)*H`$6)"
        $ws.Range("J$l").Formula = "=IF(I$l=$q$q,$q$q,(G$l-I$l)*H`$6)"
        $ws.Range("O$l").Formula = "=IF(N$l=$q$q,$q$q,(L$l-N$l)*H`$6)"
      }
    }
    "taxa" {
      Limpa $ws @("A11:G15", "I11:I15", "G4", "B5", "B6")
      for ($l = 11; $l -le 15; $l++) {
        $ws.Range("H$l").Formula = "=IF(G$l=$q$q,$q$q,G$l-F$l)"
        $ws.Range("J$l").Formula = "=IF(OR(H$l=$q$q,N(I$l)=0),$q$q,H$l/I$l/1000)"
      }
      for ($l = 20; $l -le 35; $l++) { if ("$($ws.Range("A$l").Value2)" -like "OBSERVA*") { $ws.Range("A$l").Value2 = "OBSERVAÇÃO: "; break } }
    }
    "concreto" {
      Limpa $ws @("B12:I26", "K12:Q26", "P4", "C5")
      $ws.Range("J12").Formula = "=IF(I12=$q$q,$q$q,I12-F12)"
      # datas de rompimento davam #VALOR! com a data de moldagem (P4) em branco
      $ws.Range("AD9").Formula = "=IF(OR(AD8=$q$q,W4=$q$q),$q$q,W4+AD8)"
      $ws.Range("AF9").Formula = "=IF(OR(AF8=$q$q,W4=$q$q),$q$q,W4+AF8)"
      for ($l = 27; $l -le 32; $l++) { if ("$($ws.Range("B$l").Value2)" -like "OBSERVA*") { $ws.Range("B$l").Value2 = "OBSERVAÇÃO: "; break } }
    }
    "betume" {
      Limpa $ws @("A8", "C8", "D8", "E8", "A9", "C9", "D9", "E9", "C15:C27", "F15:F27", "H2", "H3", "C3", "C4")
      if (-not $ws.Range("G8").HasFormula) { $ws.Range("G8").Formula = "=IF(A8=$q$q,$q$q,A8-E8)" }
      if (-not $ws.Range("I8").HasFormula) { $ws.Range("I8").Formula = "=IF(G8=$q$q,$q$q,(A8-E8)/A8)" }
      if (-not $ws.Range("J8").HasFormula) { $ws.Range("J8").Formula = "=IF(I8=$q$q,$q$q,AVERAGE(I8:I9))" }
      # nomes de peneira que o Excel transformou em data (3/4" e 3/8") e fórmulas de cada linha
      for ($l = 15; $l -le 27; $l++) {
        $mm = $ws.Range("A$l").Value2
        if ($mm -is [double]) {
          if ([math]::Abs($mm - 19.1) -lt 0.2) { $ws.Range("B$l").NumberFormat = "@"; $ws.Range("B$l").Value2 = "3/4''" }
          if ([math]::Abs($mm - 9.5) -lt 0.1) { $ws.Range("B$l").NumberFormat = "@"; $ws.Range("B$l").Value2 = "3/8''" }
          $ws.Range("D$l").Formula = "=IF(`$C$l=$q$q,$q$q,C$l/`$E`$12)"
          $ws.Range("E$l").Formula = "=IF(`$C$l=$q$q,$q$q,1-D$l)"
          $ws.Range("I$l").Formula = "=IF(E$l=$q$q,$q$q,AVERAGE(H$l,E$l))"
        }
      }
      for ($l = 40; $l -le 60; $l++) { if ("$($ws.Range("A$l").Value2)" -like "OBSERVA*") { $ws.Range("A$l").Value2 = "OBSERVAÇÃO: "; break } }
    }
  }
  PintaFormulas $ws.UsedRange
  $ws.Calculate()
  # erros que sobraram
  $erros = @()
  try { foreach ($cel in $ws.UsedRange.SpecialCells($xlFormulas, 16).Cells) { $erros += "$($cel.Address($false,$false))=$($cel.Text)" } } catch {}
  $ws.Activate(); $ws.Range("A1").Select() | Out-Null
  if ($erros) { Log ("   aviso: fórmulas com erro na aba MODELO: " + (($erros | Select-Object -First 12) -join ", ")) }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
try { $xl.AskToUpdateLinks = $false } catch {}
$okTudo = $true
try {
  foreach ($orig in $arquivos.Keys) {
    $tipo = $arquivos[$orig]
    $nomeArq = Split-Path $orig -Leaf
    if (-not (Test-Path -LiteralPath $orig)) { Log "ERRO $nomeArq -> arquivo não encontrado"; $okTudo = $false; continue }
    $dest = Join-Path $saida $nomeArq
    Copy-Item -LiteralPath $orig -Destination $dest -Force
    $wb = $null
    try {
      $wb = $xl.Workbooks.Open($dest, 0, $false)
      $nAbas = $wb.Worksheets.Count
      if ($tipo -eq "resumo") { Log "Resumo: $nomeArq"; PreparaResumo $wb }
      else { PreparaFicha $wb $tipo; Log "OK  ficha ($tipo) $nomeArq -> $nAbas aba(s) viraram 1 aba MODELO" }
      $xl.CalculateFull()
      if ($tipo -eq "resumo") {
        $erros = @()
        foreach ($ws in $wb.Worksheets) { try { foreach ($cel in $ws.UsedRange.SpecialCells($xlFormulas, 16).Cells) { $erros += "$($ws.Name)!$($cel.Address($false,$false))=$($cel.Text)" } } catch {} }
        if ($erros) { Log ("   aviso: fórmulas com erro: " + (($erros | Select-Object -First 15) -join ", ")) } else { Log "OK  Resumo sem fórmulas com erro" }
      }
      $wb.Save()
    }
    catch { Log "ERRO $nomeArq -> $($_.Exception.Message) (linha $($_.InvocationInfo.ScriptLineNumber))"; $okTudo = $false }
    finally { if ($wb) { $wb.Close($false) | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } }
  }
}
finally {
  $xl.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Log ""
Log "Planilhas preparadas em: $saida"
Log $(if ($okTudo) { "Tudo certo. Os originais NÃO foram alterados." } else { "Houve erro em algum arquivo. Os originais NÃO foram alterados." })
$log | Set-Content -Path (Join-Path $saida "preparo.log") -Encoding UTF8
Pausa
