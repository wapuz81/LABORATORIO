# Lançador dos ensaios de campo
# Lê os arquivos .json exportados pelo app "Ensaios de Campo", cria a aba de cada ensaio
# na ficha certa (copiando a última aba da ficha como modelo) e acrescenta as linhas no Resumo.
# Usa o próprio Excel (logos, gráficos e fórmulas das fichas são preservados).
#
# Modo "teste": trabalha em CÓPIAS dentro de pastaSaidaTeste\aammdd-hhmm\ e não toca nos originais.
# Modo "gravar": antes de alterar, copia a versão anterior de cada arquivo para a subpasta ANTIGO.

param([string]$Config = (Join-Path $PSScriptRoot "config.json"), [switch]$SemPausa)
function Pausa { if (-not $SemPausa) { Read-Host "Enter para fechar" | Out-Null } }

$ErrorActionPreference = "Stop"
$cfg = Get-Content -Raw -Encoding UTF8 $Config | ConvertFrom-Json
$agora = Get-Date
$carimbo = $agora.ToString("yyMMdd-HHmm")
$log = New-Object System.Collections.Generic.List[string]
function Log($m) { Write-Host $m; $log.Add($m) }

function Valor($obj, $nome) { if ($null -eq $obj) { return $null }; $p = $obj.PSObject.Properties[$nome]; if ($p) { return $p.Value } else { return $null } }
function MapaCfg($nome, $chave) { $m = Valor $cfg $nome; return (Valor $m $chave) }

# ---------- arquivos de entrada ----------
$entrada = $cfg.pastaEntrada
if (-not (Test-Path $entrada)) { throw "Pasta de entrada não encontrada: $entrada" }
$jsons = Get-ChildItem -Path $entrada -Filter *.json -File | Sort-Object Name
if (-not $jsons) { Log "Nenhum arquivo .json em $entrada. Nada a fazer."; Pausa; exit }

$registroLanc = Join-Path $entrada "lancados.csv"
$jaLancados = @{}
if (Test-Path $registroLanc) { Import-Csv $registroLanc -Encoding UTF8 | ForEach-Object { $jaLancados[$_.id] = $_ } }

$ensaios = @()
foreach ($j in $jsons) {
  $pac = Get-Content -Raw -Encoding UTF8 $j.FullName | ConvertFrom-Json
  if ($pac.app -ne "ensaios-campo-trincheira") { Log "Ignorado (não é do app): $($j.Name)"; continue }
  foreach ($e in $pac.ensaios) {
    if ($jaLancados.ContainsKey($e.id)) { Log "Já lançado antes, pulando: $($e.tipo) R$($e.registro) ($($e.rua))"; continue }
    $e | Add-Member -NotePropertyName arquivo -NotePropertyValue $j.Name -Force
    $ensaios += $e
  }
}
if (-not $ensaios) { Log "Todos os ensaios desses arquivos já foram lançados."; Pausa; exit }

# ---------- preparar arquivos de destino (cópia de teste ou backup em ANTIGO) ----------
$modo = if ($cfg.modo) { $cfg.modo.ToLower() } else { "teste" }
$pastaTeste = Join-Path $cfg.pastaSaidaTeste $carimbo
if ($modo -ne "gravar") { New-Item -ItemType Directory $pastaTeste -Force | Out-Null }
$destinos = @{}
function Destino($original) {
  if ($destinos.ContainsKey($original)) { return $destinos[$original] }
  if (-not (Test-Path $original)) { throw "Arquivo não encontrado: $original" }
  if ($modo -eq "gravar") {
    $antigo = Join-Path (Split-Path $original) "ANTIGO"
    if (-not (Test-Path $antigo)) { New-Item -ItemType Directory $antigo | Out-Null }
    $nome = [IO.Path]::GetFileNameWithoutExtension($original) + " (antes do app $carimbo)" + [IO.Path]::GetExtension($original)
    Copy-Item $original (Join-Path $antigo $nome)
    $destinos[$original] = $original
  } else {
    if (-not (Test-Path $pastaTeste)) { New-Item -ItemType Directory $pastaTeste -Force | Out-Null }
    $copia = Join-Path $pastaTeste (Split-Path $original -Leaf)
    Copy-Item $original $copia -Force
    $destinos[$original] = $copia
  }
  return $destinos[$original]
}

# ---------- Excel ----------
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$abertos = @{}
function Abrir($caminho) {
  if (-not $abertos.ContainsKey($caminho)) { $abertos[$caminho] = $xl.Workbooks.Open($caminho) }
  return $abertos[$caminho]
}
$M = [System.Reflection.Missing]::Value

function DataOA($iso) { return ([datetime]::ParseExact($iso, "yyyy-MM-dd", $null)).ToOADate() }
function DataBR($iso) { return ([datetime]::ParseExact($iso, "yyyy-MM-dd", $null)).ToString("dd/MM/yyyy") }
function Rua($e) { return "Rua: " + $e.rua }
function NomeAba($wb, $nome) {
  $n = ($nome -replace '[\\/\?\*\[\]:]', ' ').Trim()
  if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq $n) { return $null } }
  return $n
}
function Por($v) { if ($null -eq $v) { return $null } else { return [double]$v } }
# No PowerShell 5.1 o COM do Excel recusa valores vindos de variável/JSON sem cast explícito
function Escreve($ws, $end, $v) {
  if ($null -eq $v -or "$v" -eq "") { return }
  if ($v -is [string]) { $ws.Range($end).Value2 = [string]$v }
  elseif ($v -is [ValueType]) { $ws.Range($end).Value2 = [double]$v }
  else { $ws.Range($end).Value2 = "$v" }
}
function EstacaM($s) { if ("$s" -match '^\s*(\d+)\s*\+\s*(\d+(?:[.,]\d+)?)\s*$') { return [double]$matches[1] * 20 + [double]($matches[2] -replace ',', '.') } else { return 1e9 } }
function UltimaLinha($ws, $col, $min) {
  $r = $ws.Cells($ws.Rows.Count, $col).End(-4162).Row   # xlUp
  if ($r -lt $min) { return $min - 1 } else { return $r }
}
# copia a linha anterior (formatos e fórmulas) para a nova e limpa as células de dado
function NovaLinha($ws, $col) {
  $ult = UltimaLinha $ws $col 6
  $nova = $ult + 1
  $null = $ws.Rows($ult).Copy($ws.Rows($nova))
  return $nova
}
# repõe na linha nova as fórmulas da linha de referência (primeira linha de dados da aba)
function FormulasDe($ws, $linhaRef, $nova) {
  for ($c = 1; $c -le 26; $c++) { $ref = $ws.Cells($linhaRef, $c); if ($ref.HasFormula) { $ws.Cells($nova, $c).FormulaR1C1 = $ref.FormulaR1C1 } }
}
function NumOuTexto($t) { $d = 0.0; if ([double]::TryParse(($t -replace ',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d } else { return $t } }
function LimpaDados($ws, $linha, $colunas) { foreach ($c in $colunas) { $cel = $ws.Range("$c$linha"); if (-not $cel.HasFormula) { $cel.ClearContents() | Out-Null } } }
function MaterialResumo($m) { $x = MapaCfg "materialNoResumo" $m; if ($x) { return $x } else { return $m } }

$lancados = @()
$resumoOrig = $cfg.resumo
$wbRes = $null

try {
  foreach ($e in $ensaios) {
    $rot = "$($e.tipo) R$($e.registro) - $($e.rua) - $($e.data)"
    try {
      if ($e.tipo -eq "densidade") {
        # ===== ficha =====
        $orig = MapaCfg "fichasDensidade" $e.material
        if (-not $orig) { throw "sem ficha de densidade configurada para o material '$($e.material)' (config.json, fichasDensidade)" }
        $furos = @($e.furos | Where-Object { $null -ne $_.antes -or $null -ne $_.umidade })
        if ($furos.Count -gt 8) { throw "a ficha tem 8 linhas e o ensaio tem $($furos.Count) furos" }
        $wb = Abrir (Destino $orig)
        $nome = NomeAba $wb ("IN SITU-R$($e.registro)-$(("$($e.material)".ToUpper()))-$(("$($e.rua)".ToUpper()))")
        if (-not $nome) { throw "a ficha já tem uma aba com esse nome (ensaio lançado à mão?)" }
        $modelo = $wb.Worksheets.Item($wb.Worksheets.Count)
        $modelo.Copy($M, $modelo)
        $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
        $ws.Name = $nome
        $ws.Range("A10:J17").ClearContents() | Out-Null
        $ws.Range("Q10:Q17").ClearContents() | Out-Null
        Escreve $ws "I2" (NumOuTexto "$($e.registro)")
        Escreve $ws "B3" ("RUA: " + ("$($e.rua)".ToUpper()))
        Escreve $ws "B4" $e.segmento
        Escreve $ws "B5" ("$($e.camada)".ToUpper())
        Escreve $ws "I3" ("$($e.procedencia)".ToUpper())
        Escreve $ws "I4" ("$($e.empreiteira)".ToUpper())
        Escreve $ws "I5" ("$($e.material)".ToUpper())
        Escreve $ws "R2" $e.otima
        Escreve $ws "R3" $e.densReal
        Escreve $ws "R4" $e.areiaCone
        Escreve $ws "R5" $e.densAreia
        $l = 10
        foreach ($f in $furos) {
          Escreve $ws "A$l" (DataOA $e.data)
          Escreve $ws "B$l" $e.segmento
          Escreve $ws "C$l" $f.estaca
          Escreve $ws "D$l" ("$($f.posicao)".ToUpper())
          Escreve $ws "E$l" $f.altura
          Escreve $ws "F$l" $f.antes
          Escreve $ws "G$l" $f.depois
          Escreve $ws "H$l" $f.material
          # retido vazio vira 0: com I em branco a fórmula N da ficha dá #VALOR!
          if ($null -ne $f.retido) { Escreve $ws "I$l" $f.retido } else { $ws.Range("I$l").Value2 = [double]0 }
          Escreve $ws "J$l" $e.densLab
          Escreve $ws "Q$l" $f.umidade
          $l++
        }
        if ($e.obs) { Escreve $ws "A18" ("OBSERVAÇÃO: " + $e.obs) }
        $xl.Calculate()
        # ===== resumo: aba Mat. Granular =====
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $wr = $wbRes.Worksheets.Item("Mat. Granular")
        $i = 0
        foreach ($f in $furos) {
          $i++
          $gc = $ws.Range("S$(9 + $i)").Value2
          $r = NovaLinha $wr 4
          LimpaDados $wr $r @("A","B","C","D","E","F","G","H","J","K","N","Q","R","T","U")
          FormulasDe $wr 10 $r
          $reg = if ("$($e.registro)" -match '\.') { "$($e.registro)" } else { "$($e.registro).$i" }
          Escreve $wr "A$r" (NumOuTexto $reg)
          Escreve $wr "B$r" (DataOA $e.data)
          Escreve $wr "C$r" (Rua $e)
          Escreve $wr "D$r" ($f.estaca)
          Escreve $wr "E$r" (MaterialResumo $e.material)
          Escreve $wr "F$r" ($e.camada)
          Escreve $wr "G$r" (@{ "LE" = "Esquerdo"; "Eixo" = "Eixo"; "LD" = "Direito" }[$f.posicao])
          if ($gc -is [double]) { $wr.Range("H$r").Value2 = [math]::Round($gc / 100, 4) }
          if ($null -ne $f.umidade) { $wr.Range("J$r").Value2 = [math]::Round([double]$f.umidade / 100, 4) }
          # ótima: mantém a fórmula da Proj. Exec.; se o valor usado no campo for outro, grava o do campo
          $xl.Calculate()
          $k = $wr.Range("K$r")
          if ($null -ne $e.otima -and (-not ($k.Value2 -is [double]) -or [math]::Abs($k.Value2 * 100 - [double]$e.otima) -gt 0.05)) {
            $k.Value2 = [math]::Round([double]$e.otima / 100, 4)
            Log "   aviso: R$($e.registro).$i - umidade ótima do Resumo diferente da usada no campo; gravei $($e.otima)% na coluna proj."
          }
        }
        Log "OK  $rot -> aba '$nome' + $($furos.Count) linha(s) no Resumo (Mat. Granular)"
      }
      elseif ($e.tipo -eq "benkelman") {
        $orig = MapaCfg "fichasBenkelman" $e.material
        if (-not $orig) { throw "sem ficha de Benkelman configurada para o material '$($e.material)' (config.json, fichasBenkelman)" }
        foreach ($lado in "LE","EIXO","LD") { if (@($e.leituras.$lado).Count -gt 16) { throw "a ficha tem 16 linhas e o lado $lado tem mais leituras" } }
        $wb = Abrir (Destino $orig)
        $nome = NomeAba $wb ("VIGA-R$($e.registro)-$(("$($e.material)".ToUpper()))-$(("$($e.rua)".ToUpper()))")
        if (-not $nome) { throw "a ficha já tem uma aba com esse nome (ensaio lançado à mão?)" }
        $modelo = $wb.Worksheets.Item($wb.Worksheets.Count)
        $modelo.Copy($M, $modelo)
        $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
        $ws.Name = $nome
        foreach ($faixa in "A12:D27","F12:I27","K12:N27") { $ws.Range($faixa).ClearContents() | Out-Null }
        # repõe fórmulas de deflexão que tenham sido digitadas como número na ficha modelo
        for ($l = 12; $l -le 27; $l++) {
          $ws.Range("E$l").Formula = "=IF(D$l=`"`",`"`",(B$l-D$l)*H`$6)"
          $ws.Range("J$l").Formula = "=IF(I$l=`"`",`"`",(G$l-I$l)*H`$6)"
          $ws.Range("O$l").Formula = "=IF(N$l=`"`",`"`",(L$l-N$l)*H`$6)"
        }
        Escreve $ws "H4" (DataBR $e.data)
        Escreve $ws "M4" (NumOuTexto "$($e.registro)")
        Escreve $ws "C5" ("RUA: " + ("$($e.rua)".ToUpper()))
        Escreve $ws "H5" $e.admissivel
        Escreve $ws "M5" ("$($e.camada)".ToUpper())
        Escreve $ws "C6" $e.segmento
        Escreve $ws "H6" $e.k
        Escreve $ws "M6" ("$($e.procedencia)".ToUpper())
        Escreve $ws "C7" ("$($e.material)".ToUpper())
        Escreve $ws "H7" ("$($e.empreiteira)".ToUpper())
        $cols = @{ "LE" = @("A","B","C","D"); "EIXO" = @("F","G","H","I"); "LD" = @("K","L","M","N") }
        foreach ($lado in "LE","EIXO","LD") {
          $l = 12
          foreach ($p in @($e.leituras.$lado)) {
            $c = $cols[$lado]
            Escreve $ws "$($c[0])$l" $p.estaca
            Escreve $ws "$($c[1])$l" $p.li
            if ($null -ne $p.l25) { Escreve $ws "$($c[2])$l" $p.l25 } else { $ws.Range("$($c[2])$l").Value2 = "-" }
            Escreve $ws "$($c[3])$l" $p.lf
            $l++
          }
        }
        # ===== resumo: aba Viga. Benk — uma linha por estaca, LE / EIXO / LD lado a lado =====
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $wr = $wbRes.Worksheets.Item("Viga. Benk")
        $porEstaca = [ordered]@{}
        foreach ($lado in "LE","EIXO","LD") {
          foreach ($p in @($e.leituras.$lado)) {
            if ($null -eq $p.deflexao -or -not $p.estaca) { continue }
            if (-not $porEstaca.Contains($p.estaca)) { $porEstaca[$p.estaca] = @{} }
            $porEstaca[$p.estaca][$lado] = $p.deflexao
          }
        }
        $estacas = $porEstaca.Keys | Sort-Object { EstacaM $_ }
        $prim = $true
        foreach ($est in $estacas) {
          $r = NovaLinha $wr 4
          LimpaDados $wr $r @("A","B","C","D","E","F","G","H","I")
          FormulasDe $wr 6 $r
          if ($prim) {
            Escreve $wr "A$r" (NumOuTexto "$($e.registro)")
            Escreve $wr "B$r" (DataOA $e.data)
            $wr.Range("B$r").NumberFormatLocal = "dd/mm/aa"
            Escreve $wr "C$r" (Rua $e)
            $prim = $false
          }
          Escreve $wr "D$r" ($est)
          Escreve $wr "E$r" (MaterialResumo $e.material)
          Escreve $wr "F$r" ($e.camada)
          $d = $porEstaca[$est]
          if ($d.ContainsKey("LE")) { $wr.Range("G$r").Value2 = [double]$d["LE"] }
          if ($d.ContainsKey("EIXO")) { $wr.Range("H$r").Value2 = [double]$d["EIXO"] }
          if ($d.ContainsKey("LD")) { $wr.Range("I$r").Value2 = [double]$d["LD"] }
        }
        Log "OK  $rot -> aba '$nome' + $(@($estacas).Count) linha(s) no Resumo (Viga. Benk)"
      }
      else { throw "tipo de ensaio desconhecido: $($e.tipo)" }
      $lancados += [pscustomobject]@{ id = $e.id; tipo = $e.tipo; registro = $e.registro; rua = $e.rua; data = $e.data; lancado = $agora.ToString("yyyy-MM-dd HH:mm"); modo = $modo; arquivo = $e.arquivo }
    }
    catch { Log "ERRO $rot -> $($_.Exception.Message) (linha $($_.InvocationInfo.ScriptLineNumber))" }
  }
  foreach ($wb in $abertos.Values) { $wb.Save() }
}
finally {
  foreach ($wb in $abertos.Values) { try { $wb.Close($false) } catch {} }
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

# só marca como lançado no modo gravar (no modo teste dá para repetir quantas vezes quiser)
if ($modo -eq "gravar" -and $lancados) {
  if (Test-Path $registroLanc) { $lancados | Export-Csv $registroLanc -Append -NoTypeInformation -Encoding UTF8 }
  else { $lancados | Export-Csv $registroLanc -NoTypeInformation -Encoding UTF8 }
  $proc = Join-Path $entrada "PROCESSADOS"
  if (-not (Test-Path $proc)) { New-Item -ItemType Directory $proc | Out-Null }
  foreach ($j in $jsons) { Move-Item $j.FullName (Join-Path $proc $j.Name) -Force }
}

Log ""
if ($modo -eq "gravar") { Log "Pronto. Arquivos originais atualizados; versões anteriores em ANTIGO." }
else { Log "Pronto (modo TESTE). Cópias em: $pastaTeste" }
$log | Set-Content -Encoding UTF8 (Join-Path $(if ($modo -eq "gravar") { $entrada } else { $pastaTeste }) "$carimbo - log lancador.txt")
Pausa
