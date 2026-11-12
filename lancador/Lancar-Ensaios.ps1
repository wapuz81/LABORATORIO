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
# o app envia pelo Android como "... .json.txt" (o Android não compartilha .json)
$jsons = @(Get-ChildItem -Path $entrada -File | Where-Object { $_.Name -like "*.json" -or $_.Name -like "*.json.txt" } | Sort-Object Name)
if (-not $jsons) { Log "Nenhum arquivo .json em $entrada. Nada a fazer."; Pausa; exit }

$registroLanc = Join-Path $entrada "lancados.csv"
$jaLancados = @{}
if (Test-Path $registroLanc) { Import-Csv $registroLanc -Encoding UTF8 | ForEach-Object { $jaLancados[$_.id] = $_ } }

$ensaios = @()
$comErro = @{}      # arquivos .json com problema (ilegível, de outro app ou com algum ensaio que deu erro)
foreach ($j in $jsons) {
  try { $pac = Get-Content -Raw -Encoding UTF8 $j.FullName | ConvertFrom-Json }
  catch { Log "ERRO arquivo ilegível (não é um .json válido): $($j.Name)"; $comErro[$j.Name] = $true; continue }
  if ($pac.app -ne "ensaios-campo-trincheira") { Log "ERRO arquivo não é do app Ensaios de Campo: $($j.Name)"; $comErro[$j.Name] = $true; continue }
  foreach ($e in $pac.ensaios) {
    if ($jaLancados.ContainsKey($e.id)) { Log "Já lançado antes, pulando: $($e.tipo) R$($e.registro) ($($e.rua))"; continue }
    $e | Add-Member -NotePropertyName arquivo -NotePropertyValue $j.Name -Force
    $ensaios += $e
  }
}
$modo = if ($cfg.modo) { $cfg.modo.ToLower() } else { "teste" }
$raiz = Split-Path $Config -Parent
function MoveUnico($arq, $pasta) {
  if (-not (Test-Path $pasta)) { New-Item -ItemType Directory $pasta | Out-Null }
  $dest = Join-Path $pasta $arq.Name
  if (Test-Path -LiteralPath $dest) { $dest = Join-Path $pasta ($arq.BaseName + " ($carimbo)" + $arq.Extension) }
  Move-Item -LiteralPath $arq.FullName -Destination $dest
}
# no modo gravar a ENTRADA sempre termina vazia: sem erro vai para PROCESSADOS; com erro vai para COM ERRO
# (corrija e devolva para a ENTRADA: só o que faltou é lançado)
function GuardaEntrada() {
  if ($modo -ne "gravar") { return }
  foreach ($j in $jsons) {
    if ($comErro.ContainsKey($j.Name)) { MoveUnico $j (Join-Path $entrada "COM ERRO"); Log "Arquivo com problema guardado em ENTRADA\COM ERRO: $($j.Name)" }
    else { MoveUnico $j (Join-Path $entrada "PROCESSADOS") }
  }
}
function GravaLog($pasta) {
  if (-not (Test-Path $pasta)) { New-Item -ItemType Directory $pasta -Force | Out-Null }
  $log | Set-Content -Encoding UTF8 (Join-Path $pasta "$carimbo - log lancador.txt")
}
if (-not $ensaios) {
  Log "Nenhum ensaio novo para lançar nesses arquivos."
  GuardaEntrada
  if ($modo -eq "gravar") { GravaLog (Join-Path $raiz "LOGS") }
  Pausa
  if ($comErro.Count) { exit 2 } else { exit 0 }
}

# ---------- preparar arquivos de destino (cópia de teste ou backup em ANTIGO) ----------
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
# Filtro de mensagens COM: quando o Excel está ocupado (recalculando o Resumo), o Windows
# repete a chamada por até 2 minutos em vez de dar RPC_E_CALL_REJECTED.
if (-not ("FiltroCom" -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
  [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
  [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
  [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}
public class FiltroCom : IOleMessageFilter {
  [DllImport("Ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
  public static void Ligar() { IOleMessageFilter old; CoRegisterMessageFilter(new FiltroCom(), out old); }
  public static void Desligar() { IOleMessageFilter old; CoRegisterMessageFilter(null, out old); }
  int IOleMessageFilter.HandleInComingCall(int a, IntPtr b, int c, IntPtr d) { return 0; }
  int IOleMessageFilter.RetryRejectedCall(IntPtr a, int tick, int tipo) { if (tipo == 2 && tick < 120000) return 200; return -1; }
  int IOleMessageFilter.MessagePending(IntPtr a, int b, int c) { return 2; }
}
"@
}
[FiltroCom]::Ligar()
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$abertos = @{}
# As fichas abrem num segundo Excel, que é fechado depois de cada ensaio: depois de umas 5 cópias
# de aba na mesma instância o Excel passa a recusar Worksheet.Copy. O Resumo fica no Excel principal.
$xf = $null
function ExcelFichas() {
  if (-not $script:xf) {
    $script:xf = New-Object -ComObject Excel.Application
    $script:xf.Visible = $false
    $script:xf.DisplayAlerts = $false
  }
  return $script:xf
}
function Abrir($caminho) {
  if (-not $abertos.ContainsKey($caminho)) {
    $app = if ($caminho -eq $destinos[$resumoOrig]) { $xl } else { ExcelFichas }
    $abertos[$caminho] = $app.Workbooks.Open($caminho)
  }
  return $abertos[$caminho]
}
$M = [System.Reflection.Missing]::Value
# copia a aba modelo para o fim da pasta; se o Excel recusar, ativa a pasta e tenta de novo
# aba que serve de modelo: a "MODELO" (ficha preparada); se não existir, a última aba
function AbaModelo($wb) {
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq "MODELO") { return $s } }
  return $wb.Worksheets.Item($wb.Worksheets.Count)
}
function CopiaAba($wb, $modelo) {
  # devolve a aba nova (achada pelo nome, não pela posição) e garante que ela fica no fim
  $antes = @{}; foreach ($s in $wb.Worksheets) { $antes[$s.Name] = $true }
  for ($t = 1; $t -le 3; $t++) {
    try { $modelo.Copy([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); break }
    catch {
      if ($t -eq 3) { throw }
      Start-Sleep -Seconds 2
      try { $wb.Activate(); $modelo.Activate() } catch {}
    }
  }
  $nova = $null
  foreach ($s in $wb.Worksheets) { if (-not $antes.ContainsKey($s.Name)) { $nova = $s } }
  if (-not $nova) { throw "a cópia da aba modelo não apareceu na ficha" }
  if ($nova.Index -ne $wb.Worksheets.Count) {
    Log "   aviso: a cópia da aba modelo ficou na posição $($nova.Index) de $($wb.Worksheets.Count); movida para o fim"
    $nova.Move([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
  }
  return $nova
}
function FechaFichas() {
  $chaveRes = $destinos[$resumoOrig]
  foreach ($k in @($abertos.Keys)) {
    if ($k -ne $chaveRes) {
      $w = $abertos[$k]
      try { $w.Save() } catch { Log "   aviso: não consegui salvar $k ($($_.Exception.Message))" }
      try { $w.Close($false) } catch {}
      $abertos.Remove($k)
    }
  }
  if ($script:xf) {
    try { $script:xf.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:xf) | Out-Null
    $script:xf = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

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
function NovaLinha($ws, $col, $prim = 6) {
  $ult = UltimaLinha $ws $col $prim
  if ($ult -lt $prim) { return $prim }   # aba vazia: usa a linha-modelo (já tem as fórmulas)
  $nova = $ult + 1
  $null = $ws.Rows($ult).Copy($ws.Rows($nova))
  return $nova
}
# repõe na linha nova as fórmulas da linha de referência (primeira linha de dados da aba)
# para cada coluna, procura de baixo para cima a última linha que tem fórmula (até a linha de referência)
function FormulasDe($ws, $linhaRef, $nova) {
  for ($c = 1; $c -le 26; $c++) {
    for ($l = $nova - 1; $l -ge $linhaRef; $l--) {
      $ref = $ws.Cells($l, $c)
      if ($ref.HasFormula) { $ws.Cells($nova, $c).FormulaR1C1 = $ref.FormulaR1C1; break }
      if ($null -ne $ref.Value2 -and "$($ref.Value2)" -ne "") { break }   # célula com valor digitado: mantém como está
    }
  }
}
function NumOuTexto($t) { $d = 0.0; if ([double]::TryParse(($t -replace ',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d } else { return $t } }
function LimpaDados($ws, $linha, $colunas) { foreach ($c in $colunas) { $cel = $ws.Range("$c$linha"); if (-not $cel.HasFormula) { $cel.ClearContents() | Out-Null } } }
function MaterialResumo($m) { $x = MapaCfg "materialNoResumo" $m; if ($x) { return $x } else { return $m } }

$lancados = @()
$resumoOrig = $cfg.resumo
$wbRes = $null

try {
  foreach ($e in $ensaios) {
    $rot = "$($e.tipo) R$($e.registro) - $(if ($e.rua) { $e.rua } else { $e.local }) - $($e.data)"
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
        $modelo = AbaModelo $wb
        $ws = CopiaAba $wb $modelo
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
        $ws.Calculate()
        # ===== resumo: aba Mat. Granular =====
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $wr = $wbRes.Worksheets.Item("Mat. Granular")
        $i = 0; $otProj = $null
        foreach ($f in $furos) {
          $i++
          $gc = $ws.Range("S$(9 + $i)").Value2
          $r = NovaLinha $wr 4 10
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
          # ótima: vale sempre a da Proj. Exec. (fórmula da coluna K); o valor do campo só entra se o material não tiver projeto
          $xl.Calculate()
          $k = $wr.Range("K$r")
          if ($k.Value2 -is [double] -and $k.Value2 -gt 0) {
            $otProj = [math]::Round($k.Value2 * 100, 2)
            if ($i -eq 1 -and $null -ne $e.otima -and [math]::Abs($otProj - [double]$e.otima) -gt 0.05) {
              Log "   aviso: R$($e.registro) - o campo usou umidade ótima $($e.otima)%; vale a da Proj. Exec.: $otProj%"
            }
          }
          elseif ($null -ne $e.otima) {
            $k.Value2 = [math]::Round([double]$e.otima / 100, 4)
            Log "   aviso: R$($e.registro).$i - material sem umidade ótima na Proj. Exec.; gravei a do campo ($($e.otima)%) na coluna proj."
          }
        }
        # a ficha mostra a mesma ótima do Resumo
        if ($otProj) { $ws.Range("R2").Value2 = [double]$otProj; $ws.Calculate() }
        Log "OK  $rot -> aba '$nome' + $($furos.Count) linha(s) no Resumo (Mat. Granular)"
      }
      elseif ($e.tipo -eq "benkelman") {
        $orig = MapaCfg "fichasBenkelman" $e.material
        if (-not $orig) { throw "sem ficha de Benkelman configurada para o material '$($e.material)' (config.json, fichasBenkelman)" }
        foreach ($lado in "LE","EIXO","LD") { if (@($e.leituras.$lado).Count -gt 16) { throw "a ficha tem 16 linhas e o lado $lado tem mais leituras" } }
        $wb = Abrir (Destino $orig)
        $nome = NomeAba $wb ("VIGA-R$($e.registro)-$(("$($e.material)".ToUpper()))-$(("$($e.rua)".ToUpper()))")
        if (-not $nome) { throw "a ficha já tem uma aba com esse nome (ensaio lançado à mão?)" }
        $modelo = AbaModelo $wb
        $ws = CopiaAba $wb $modelo
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
      elseif ($e.tipo -eq "concreto") {
        # ===== Moldagem de concreto: uma aba por série =====
        $orig = $cfg.fichaConcreto
        if (-not $orig) { throw "ficha de concreto não configurada (config.json, fichaConcreto)" }
        $wb = Abrir (Destino $orig)
        $nome = NomeAba $wb ("SÉRIE $($e.registro)-$("$($e.elemento)".ToUpper())")
        if (-not $nome) { throw "a ficha já tem uma aba com esse nome (série lançada à mão?)" }
        $modelo = AbaModelo $wb
        $ws = CopiaAba $wb $modelo
        $ws.Name = $nome
        $ws.Range("B12:I26").ClearContents() | Out-Null
        $ws.Range("K12:Q26").ClearContents() | Out-Null
        $ws.Range("J12").Formula = '=IF(I12="","",I12-F12)'
        Escreve $ws "J4" ("$($e.concreteira)".ToUpper())
        Escreve $ws "P4" (DataOA $e.data)
        if ($e.rua) { Escreve $ws "C5" ("RUA: " + "$($e.rua)".ToUpper()) }
        Escreve $ws "AE4" $e.fck
        Escreve $ws "G8" $e.fck
        Escreve $ws "L7" $e.cimento
        Escreve $ws "B12" (NumOuTexto "$($e.registro)")
        Escreve $ws "C12" $e.cps
        Escreve $ws "D12" $e.slump
        Escreve $ws "E12" $e.nf
        foreach ($par in @(@("F12","hDosagem"), @("G12","hChegada"), @("H12","hInicio"), @("I12","hTermino"))) {
          $h = "$($e.($par[1]))"
          if ($h -match '^(\d{1,2}):(\d{2})') { $ws.Range($par[0]).Value2 = [double](([int]$matches[1] * 60 + [int]$matches[2]) / 1440.0) }
        }
        Escreve $ws "K12" $e.volume
        Escreve $ws "L12" $e.acumulado
        Escreve $ws "M12" ("$($e.placa)".ToUpper())
        $elem = "$($e.elemento)".ToUpper(); if ($e.lado -and $e.lado -ne "–") { $elem = "$elem - $("$($e.lado)".ToUpper())" }
        Escreve $ws "N12" $elem
        if ($e.idades) { $ws.Range("Q12").NumberFormat = "@"; $ws.Range("Q12").Value2 = [string]$e.idades }   # texto: "7/28" não pode virar data
        Escreve $ws "N30" ("$($e.moldador)".ToUpper())
        if ($e.obs) { Escreve $ws "B29" ("OBSERVAÇÃO: " + $e.obs) }
        # ===== resumo: aba Concreto =====
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $wr = $wbRes.Worksheets.Item("Concreto")
        $r = NovaLinha $wr 1
        LimpaDados $wr $r @("A","B","C","D","E","F","G","H","I","J","K","M","N","Q","R")
        FormulasDe $wr 6 $r
        Escreve $wr "A$r" (NumOuTexto "$($e.registro)")
        Escreve $wr "B$r" (DataOA $e.data)
        $wr.Range("B$r").NumberFormatLocal = "dd/mm/aa"
        Escreve $wr "C$r" $(if ($e.local) { $e.local } else { Rua $e })
        Escreve $wr "E$r" $e.elemento
        if ($e.lado -and $e.lado -ne "–") { Escreve $wr "F$r" $e.lado }
        Escreve $wr "G$r" $e.slump
        Escreve $wr "H$r" "10 X 20"
        Escreve $wr "I$r" "Compressão"
        Escreve $wr "J$r" $e.estrutura
        Escreve $wr "K$r" $e.fck
        Log "OK  $rot -> aba '$nome' + linha $r no Resumo (Concreto)"
      }
      elseif ($e.tipo -eq "taxa") {
        # ===== Imprimação (EAI) / Pintura de ligação (RR-1C) =====
        $orig = MapaCfg "fichasTaxa" $e.emulsao
        if (-not $orig) { throw "sem ficha configurada para a emulsão '$($e.emulsao)' (config.json, fichasTaxa)" }
        $bands = @($e.bandejas)
        if ($bands.Count -gt 5) { throw "a ficha tem 5 linhas e o ensaio tem $($bands.Count) bandejas" }
        $wb = Abrir (Destino $orig)
        $pref = if ($e.emulsao -eq "RR-1C") { "RR1C" } else { "EAI" }
        $nome = NomeAba $wb ("$pref-R$($e.registro)-$("$($e.camada)".ToUpper().Replace(' - ', ' '))-$("$($e.rua)".ToUpper())")
        if (-not $nome) { throw "a ficha já tem uma aba com esse nome (ensaio lançado à mão?)" }
        $modelo = AbaModelo $wb
        $ws = CopiaAba $wb $modelo
        $ws.Name = $nome
        $ws.Range("A11:G15").ClearContents() | Out-Null
        $ws.Range("I11:I15").ClearContents() | Out-Null
        # a aba modelo pode não ter as fórmulas em todas as linhas: repõe H e J nas 5
        for ($l = 11; $l -le 15; $l++) {
          $ws.Range("H$l").Formula = "=IF(G$l=`"`",`"`",G$l-F$l)"
          $ws.Range("J$l").Formula = "=IF(B$l=0,`"`",H$l/I$l/1000)"
        }
        $ws.Range("G4").NumberFormat = "@"
        Escreve $ws "G4" ([string]$e.registro)
        Escreve $ws "B5" ("RUA: " + "$($e.rua)".ToUpper())
        Escreve $ws "G5" ("$($e.camada)".ToUpper())
        Escreve $ws "B6" $e.segmento
        Escreve $ws "G6" $pref
        Escreve $ws "B7" ("$($e.empreiteira)".ToUpper())
        if ($null -ne $e.proj -and $null -ne $e.tolerancia) {
          Escreve $ws "G7" (("{0:0.0#} A {1:0.0#}" -f ([double]$e.proj - [double]$e.tolerancia), ([double]$e.proj + [double]$e.tolerancia)) -replace '\.', ',')
        }
        $seg = if ($e.segmento) { "$($e.segmento)" } else { "-" }
        $l = 11
        foreach ($b in $bands) {
          Escreve $ws "A$l" (DataOA $e.data)
          Escreve $ws "B$l" $seg
          Escreve $ws "C$l" $b.estaca
          Escreve $ws "D$l" ("$($b.posicao)".ToUpper())
          Escreve $ws "E$l" $b.temp
          Escreve $ws "F$l" $b.bandeja
          Escreve $ws "G$l" $b.total
          Escreve $ws "I$l" $b.area
          $l++
        }
        if ($e.obs) { Escreve $ws "A27" ("OBSERVAÇÃO: " + $e.obs) }
        $ws.Calculate()
        # ===== resumo: aba Imprimação ou P. Ligação =====
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $aba = if ($e.emulsao -eq "RR-1C") { "P. Ligação" } else { "Imprimação" }
        $wr = $wbRes.Worksheets.Item($aba)
        $i = 0
        foreach ($b in $bands) {
          $i++
          $tx = $ws.Range("J$(10 + $i)").Value2
          $r = NovaLinha $wr 4
          LimpaDados $wr $r @("A","B","C","D","E","F","G","H","I")
          FormulasDe $wr 6 $r
          $reg = if ("$($e.registro)" -match '\.') { "$($e.registro)" } else { "$($e.registro).$i" }
          $wr.Range("A$r").NumberFormat = "@"
          Escreve $wr "A$r" ([string]$reg)
          Escreve $wr "B$r" (DataOA $e.data)
          $wr.Range("B$r").NumberFormatLocal = "dd/mm/aa"
          Escreve $wr "C$r" (Rua $e)
          Escreve $wr "D$r" ($b.estaca)
          Escreve $wr "E$r" ($e.emulsao)
          Escreve $wr "F$r" ($e.camada)
          if ($tx -is [double]) { $wr.Range("G$r").Value2 = [math]::Round($tx, 2) }
          if (-not $wr.Range("H$r").HasFormula -and $null -ne $e.proj) { Escreve $wr "H$r" ([double]$e.proj) }
        }
        Log "OK  $rot -> aba '$nome' + $($bands.Count) linha(s) no Resumo ($aba)"
      }
      elseif ($e.tipo -eq "betume") {
        # ===== CBUQ teor de betume + granulometria =====
        # ficha "TEOR DE BETUME" (uma por faixa): cabeçalho H2/H3/C3/C4/H4/C5/H5, pesagens na linha 8, retidos na coluna C das linhas 15–27
        $avisoFicha = ""
        $nome = $null
        $orig = MapaCfg "fichasBetume" $e.faixa
        if (-not $orig) { $avisoFicha = "; sem ficha configurada para a faixa '$($e.faixa)' (config.json, fichasBetume): lancei só no Resumo" }
        else {
          $wb = Abrir (Destino $orig)
          $nome = NomeAba $wb ("BETUME-R$($e.registro)-FAIXA $($e.faixa)-$("$($e.rua)".ToUpper())")
          if (-not $nome) { throw "a ficha já tem uma aba com esse nome (ensaio lançado à mão?)" }
          $modelo = AbaModelo $wb
          $ws = CopiaAba $wb $modelo
          $ws.Name = $nome
          try {
          # células mescladas (A8:B8, E8:F8...): limpar a mescla inteira
          foreach ($c in @("A8","C8","E8","A9","C9","E9")) { $ws.Range($c).MergeArea.ClearContents() | Out-Null }
          foreach ($c in @("D8","D9")) { if (-not $ws.Range($c).HasFormula) { $ws.Range($c).MergeArea.ClearContents() | Out-Null } }
          for ($l = 15; $l -le 27; $l++) { foreach ($c in @("C","F")) { $ws.Range("$c$l").MergeArea.ClearContents() | Out-Null } }
          Escreve $ws "H2" (NumOuTexto "$($e.registro)")
          Escreve $ws "H3" (DataOA $e.data)
          Escreve $ws "C3" ("Rua " + $e.rua)
          Escreve $ws "C4" $e.segmento
          Escreve $ws "H4" ("$($e.empreiteira)".ToUpper())
          Escreve $ws "C5" ("CBUQ - FAIXA `"$($e.faixa)`"")
          Escreve $ws "H5" ("$($e.procedencia)".ToUpper())
          Escreve $ws "K12" $e.faixa
          Escreve $ws "A8" $e.mistura
          Escreve $ws "C8" $e.papel
          if (-not $ws.Range("D8").HasFormula) { Escreve $ws "D8" $e.papelAgreg }
          Escreve $ws "E8" $e.agregado
          # a aba modelo pode estar sem as fórmulas do teor (ex.: "TEOR DE BETUM copia"): repõe G8, I8 e J8
          if (-not $ws.Range("G8").HasFormula) { $ws.Range("G8").Formula = '=IF(A8="","",A8-E8)' }
          if (-not $ws.Range("I8").HasFormula) { $ws.Range("I8").Formula = '=IF(G8="","",(A8-E8)/A8)' }
          if (-not $ws.Range("J8").HasFormula) { $ws.Range("J8").Formula = '=IF(I8="","",AVERAGE(I8:I9))' }
          # linha de cada peneira pela abertura (mm) da coluna A
          $mmPen = @{ "2" = @(50.8, 50); "112" = @(38.1, 37.5); "1" = @(25.4, 25); "34" = @(19.1, 19); "12" = @(12.7, 12.5); "38" = @(9.5); "4" = @(4.8, 4.75); "10" = @(2.0); "40" = @(0.42); "80" = @(0.18); "200" = @(0.075) }
          $faltou = @()
          foreach ($pn in @($e.peneiras)) {
            if ($null -eq $pn.retidoAcum) { continue }
            $linha = 0
            for ($l = 15; $l -le 27; $l++) {
              $v = $ws.Range("A$l").Value2
              if ($v -is [double]) { foreach ($m in $mmPen["$($pn.peneira)"]) { if ([math]::Abs($v - $m) -lt 0.011) { $linha = $l } } }
              if ($linha) { break }
            }
            if (-not $linha) { $faltou += $pn.nome; continue }
            Escreve $ws "C$linha" ([double]$pn.retidoAcum)
            # a aba modelo pode estar sem as fórmulas: repõe retido/passante/média nessa linha
            $ws.Range("D$linha").Formula = "=IF(`$C$linha=`"`",`"`",C$linha/`$E`$12)"
            $ws.Range("E$linha").Formula = "=IF(`$C$linha=`"`",`"`",1-D$linha)"
            $ws.Range("I$linha").Formula = "=IF(E$linha=`"`",`"`",AVERAGE(H$linha,E$linha))"
          }
          if ($faltou) { $avisoFicha = "; peneira(s) sem linha na ficha: $($faltou -join ', ')" }
          for ($l = 40; $l -le 60; $l++) {
            if ("$($ws.Range("A$l").Value2)" -like "OBSERVA*") { $ws.Range("A$l").Value2 = [string]("OBSERVAÇÃO: " + $(if ($e.obs) { $e.obs } else { "" })); break }
          }
          $ws.Calculate()
          }
          catch {
            # não deixa aba pela metade (senão ela vira o modelo do próximo lançamento)
            $ws.Application.DisplayAlerts = $false; try { $ws.Delete() | Out-Null } catch {}
            throw
          }
        }
        if (-not $wbRes) { $wbRes = Abrir (Destino $resumoOrig) }
        $wr = $wbRes.Worksheets.Item("CBUQ Gl Betume")
        $r = NovaLinha $wr 2 15          # coluna B (data): a coluna A já vem numerada até o fim da aba
        if ($r -lt 15) { $r = 15 }
        LimpaDados $wr $r @("A","B","C","D","E","F","G","I","J","K","L","M","N","O","P","Q","R","S","U","X","Y","AB","AC","AF")
        FormulasDe $wr 15 $r
        Escreve $wr "A$r" (NumOuTexto "$($e.registro)")
        Escreve $wr "B$r" (DataOA $e.data)
        $wr.Range("B$r").NumberFormatLocal = "dd/mm/aa"
        Escreve $wr "C$r" (Rua $e)
        Escreve $wr "D$r" $e.segmento
        Escreve $wr "E$r" $e.faixa
        Escreve $wr "F$r" $e.camada
        $lado = switch ("$($e.lado)") { "LE" { "Esquerdo" } "LD" { "Direito" } default { "Eixo" } }
        Escreve $wr "G$r" $lado
        $ordem = @("2","112","1","34","12","38","4","10","40","80","200")   # colunas I..S
        foreach ($pn in @($e.peneiras)) {
          $ix = [array]::IndexOf($ordem, "$($pn.peneira)")
          if ($ix -ge 0 -and $null -ne $pn.passante) { $wr.Cells($r, 9 + $ix).Value2 = [double][math]::Round([double]$pn.passante, 3) }
        }
        # teor e projeto como fração (as colunas X e Y são formatadas em %)
        if ($null -ne $e.calc.teor) { $wr.Range("X$r").Value2 = [double][math]::Round([double]$e.calc.teor / 100, 4); $wr.Range("X$r").NumberFormatLocal = "0,00%" }
        if ($null -ne $e.proj) { $wr.Range("Y$r").Value2 = [double]([double]$e.proj / 100); $wr.Range("Y$r").NumberFormatLocal = "0,0%" }
        $txtAba = if ($nome) { "aba '$nome' + " } else { "" }
        Log "OK  $rot -> $($txtAba)linha $r no Resumo (CBUQ Gl Betume)$avisoFicha"
      }
      else { throw "tipo de ensaio desconhecido: $($e.tipo)" }
      $lancados += [pscustomobject]@{ id = $e.id; tipo = $e.tipo; registro = $e.registro; rua = $e.rua; data = $e.data; lancado = $agora.ToString("yyyy-MM-dd HH:mm"); modo = $modo; arquivo = $e.arquivo }
    }
    catch { Log "ERRO $rot -> $($_.Exception.Message) (linha $($_.InvocationInfo.ScriptLineNumber))"; $comErro[$e.arquivo] = $true }
    finally { FechaFichas }
  }
  foreach ($wb in $abertos.Values) { $wb.Save() }
}
finally {
  foreach ($wb in $abertos.Values) { try { $wb.Close($false) } catch {} }
  if ($xf) { try { $xf.Quit() } catch {} }
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  [FiltroCom]::Desligar()
}

# só marca como lançado no modo gravar (no modo teste dá para repetir quantas vezes quiser)
if ($modo -eq "gravar" -and $lancados) {
  if (Test-Path $registroLanc) { $lancados | Export-Csv $registroLanc -Append -NoTypeInformation -Encoding UTF8 }
  else { $lancados | Export-Csv $registroLanc -NoTypeInformation -Encoding UTF8 }
}
GuardaEntrada

Log ""
if ($modo -eq "gravar") { Log "Pronto. Arquivos originais atualizados; versões anteriores em ANTIGO." }
else { Log "Pronto (modo TESTE). Cópias em: $pastaTeste" }
GravaLog $(if ($modo -eq "gravar") { Join-Path $raiz "LOGS" } else { $pastaTeste })
Pausa
if ($comErro.Count) { exit 2 } else { exit 0 }
