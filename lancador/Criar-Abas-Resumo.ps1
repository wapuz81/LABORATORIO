# Cria (ou refaz) a aba RESUMO nas fichas: cópia da aba correspondente do Resumo, só valores,
# com link do Registro para a aba do ensaio e "← RESUMO" em cada aba de ensaio.
# Uso: -Pasta <pasta com as fichas e o Resumo> (altera os arquivos dessa pasta; faça cópia antes)
param([Parameter(Mandatory = $true)][string]$Pasta, [string]$Resumo = "2611 - Resumo Ensaios Trincheira.xlsx")
$ErrorActionPreference = "Stop"
function AbaDoEnsaio($wb, $serie) {
  if (-not $serie) { return $null }
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq $serie) { return $s.Name } }
  foreach ($s in $wb.Worksheets) { if ($s.Name -eq "SÉRIE $serie" -or $s.Name -like "SÉRIE $serie-*") { return $s.Name } }
  foreach ($s in $wb.Worksheets) { if ($s.Name -like "*-R$serie-*") { return $s.Name } }
  return $null
}
function ResumoNaFicha($wbF, $wsRes) {
  $app = $wbF.Application
  $app.DisplayAlerts = $false
  foreach ($s in @($wbF.Worksheets)) { if ($s.Name -eq "RESUMO") { $s.Delete() | Out-Null } }
  $wsRes.Copy($wbF.Worksheets.Item(1))
  $r = $wbF.Worksheets.Item(1)
  $r.Name = "RESUMO"
  # só valores: nada de fórmula ligada ao arquivo do Resumo
  $u = $r.UsedRange
  $null = $u.Copy(); $null = $u.PasteSpecial(-4163); $app.CutCopyMode = $false
  try { $r.Cells.Validation.Delete() } catch {}
  foreach ($n in @($wbF.Names)) { try { if ("$($n.RefersTo)" -match '\[|\.xls') { $n.Delete() } } catch {} }
  $links = $wbF.LinkSources(1)
  if ($links) { foreach ($k in @($links)) { try { $wbF.BreakLink([string]$k, 1) } catch {} } }
  # linha do cabeçalho: "Registro" na coluna A
  $h = 0
  for ($l = 1; $l -le 30; $l++) { if ("$($r.Cells($l, 1).Value2)".Trim() -eq "Registro") { $h = $l; break } }
  if (-not $h) { throw "não achei o cabeçalho 'Registro' na aba $($wsRes.Name)" }
  $ult = $r.Cells($r.Rows.Count, 1).End(-4162).Row
  $n = 0
  for ($l = $h + 1; $l -le $ult; $l++) {
    $t = "$($r.Cells($l, 1).Text)".Trim()
    if (-not $t) { continue }
    $aba = AbaDoEnsaio $wbF (($t -split '[.,]')[0].Trim())
    if ($aba) { $null = $r.Hyperlinks.Add($r.Cells($l, 1), "", "'$aba'!A1", "Abrir a aba $aba"); $n++ }
  }
  # em cada aba de ensaio: "← RESUMO" logo depois da área de impressão, na linha 1
  foreach ($s in $wbF.Worksheets) {
    if ($s.Name -eq "RESUMO" -or $s.Name -like "MODELO*") { continue }
    $col = 0
    $pa = "$($s.PageSetup.PrintArea)"
    if ($pa) { try { $a = $s.Range(($pa -split '[,;]')[0]); $col = $a.Column + $a.Columns.Count } catch {} }
    if (-not $col) { $col = $s.UsedRange.Column + $s.UsedRange.Columns.Count }
    $c = $s.Cells(1, $col)
    if ($null -eq $c.Value2 -or "$($c.Value2)" -eq "← RESUMO") {
      $c.Value2 = "← RESUMO"
      $null = $s.Hyperlinks.Add($c, "", "RESUMO!A1", "Voltar para o resumo")
    }
  }
  $r.Activate()
  $null = $r.Range("A1").Select()
  return $n
}

$mapa = [ordered]@{
  "DENSIDADE IN SITU.xlsx" = "Mat. Granular"
  "VIGA BENKELMAN.xlsx"    = "Viga. Benk"
  "TEOR DE BETUME.xlsx"    = "CBUQ Gl Betume"
  "CONCRETO.xlsx"          = "Concreto"
  "EAI.xlsx"               = "Imprimação"
  "RR1C.xlsx"              = "P. Ligação"
}
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wbR = $xl.Workbooks.Open((Join-Path $Pasta $Resumo), 0, $true)   # Resumo só leitura
  foreach ($f in $mapa.Keys) {
    $arq = Join-Path $Pasta $f
    if (-not (Test-Path -LiteralPath $arq)) { Write-Host "pulei (não existe): $f"; continue }
    $wbF = $xl.Workbooks.Open($arq)
    try {
      $qtd = ResumoNaFicha $wbF ($wbR.Worksheets.Item($mapa[$f]))
      $wbF.Save()
      $lk = $wbF.LinkSources(1)
      Write-Host ("OK  {0}: aba RESUMO ({1}), {2} registro(s) com link; abas: {3}; vínculos externos: {4}" -f $f, $mapa[$f], $qtd, (@($wbF.Worksheets | ForEach-Object { $_.Name }) -join ' | '), $(if ($lk) { @($lk).Count } else { 0 }))
    }
    catch { Write-Host "ERRO $f -> $($_.Exception.Message)" }
    finally { $wbF.Close($false) }
  }
  $wbR.Close($false)
}
finally { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
