# Vigiar-Entrada.ps1
# Chamado pela tarefa agendada "Lancador Ensaios Campo" a cada 5 minutos.
# Se houver arquivo .json novo na ENTRADA (enviado pelo app via Drive), roda o lançador e mostra o resultado.
# Só trabalha no modo "gravar". Não faz nada se alguma planilha estiver aberta no Excel.
param([string]$Pasta = $PSScriptRoot, [switch]$SemAviso)   # -SemAviso: escreve o aviso no console em vez de abrir janela (para teste)
$ErrorActionPreference = "Stop"
function Aviso($txt, $icone = 64) { if ($SemAviso) { Write-Output "AVISO[$icone]: $txt"; return }; try { (New-Object -ComObject WScript.Shell).Popup($txt, 0, "Lançador de ensaios de campo", $icone + 4096) | Out-Null } catch {} }

$mutex = New-Object System.Threading.Mutex($false, "LancadorEnsaiosCampo")
if (-not $mutex.WaitOne(0)) { exit }
try {
  $cfg = Get-Content (Join-Path $Pasta "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
  if ("$($cfg.modo)".ToLower() -ne "gravar") { exit }
  $entrada = $cfg.pastaEntrada
  if (-not (Test-Path $entrada)) { exit }          # Drive ainda não montou
  $parado = Join-Path $entrada "PARADO - leia.txt"
  if (Test-Path $parado) { exit }
  $jsons = @(Get-ChildItem -Path $entrada -File | Where-Object { $_.Name -like "*.json" -or $_.Name -like "*.json.txt" })
  if (-not $jsons) { exit }

  # espera o Drive terminar de baixar: arquivo com conteúdo, parado há 30 s e sem ninguém usando
  foreach ($j in $jsons) {
    if ($j.Length -eq 0 -or $j.LastWriteTime -gt (Get-Date).AddSeconds(-30)) { exit }
    try { $f = [IO.File]::Open($j.FullName, 'Open', 'Read', 'None'); $f.Close() } catch { exit }
  }

  # planilhas livres? (se o Resumo ou uma ficha estiver aberta no Excel, espera a próxima rodada)
  $planilhas = @($cfg.resumo)
  foreach ($n in "fichasDensidade", "fichasBenkelman", "fichasTaxa", "fichasBetume") { $m = $cfg.$n; if ($m) { foreach ($p in $m.PSObject.Properties) { $planilhas += $p.Value } } }
  if ($cfg.fichaConcreto) { $planilhas += $cfg.fichaConcreto }
  $marca = Join-Path $entrada "aguardando.tmp"
  foreach ($p in ($planilhas | Select-Object -Unique)) {
    if (-not (Test-Path -LiteralPath $p)) { continue }
    try { $f = [IO.File]::Open($p, 'Open', 'ReadWrite', 'None'); $f.Close() }
    catch {
      # avisa uma vez por hora, para não incomodar
      if (-not (Test-Path $marca) -or (Get-Item $marca).LastWriteTime -lt (Get-Date).AddHours(-1)) {
        Set-Content -Path $marca -Value (Get-Date) -Encoding UTF8
        Aviso ("Chegaram ensaios do app, mas esta planilha está aberta:`n`n" + (Split-Path $p -Leaf) + "`n`nFeche o Excel. O lançamento é feito sozinho em até 5 minutos.") 48
      }
      exit
    }
  }
  if (Test-Path $marca) { Remove-Item $marca -Force }

  $saida = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Pasta "Lancar-Ensaios.ps1") -SemPausa 2>&1 | ForEach-Object { "$_" }
  $codigo = $LASTEXITCODE
  $ok = @($saida | Where-Object { $_ -like "OK *" })
  $erros = @($saida | Where-Object { $_ -like "ERRO*" })
  $avisos = @($saida | Where-Object { $_ -like "*aviso:*" })
  if ($codigo -eq 0 -or $codigo -eq 2) {
    $txt = "$($ok.Count) ensaio(s) lançado(s) nas fichas e no Resumo."
    if ($ok) { $txt += "`n`n" + (($ok | Select-Object -First 12) -join "`n") }
    if ($avisos) { $txt += "`n`nAvisos:`n" + (($avisos | Select-Object -First 6) -join "`n") }
    if ($erros) { $txt += "`n`n$($erros.Count) com ERRO (arquivo guardado em ENTRADA\COM ERRO):`n" + (($erros | Select-Object -First 6) -join "`n") }
    if ($ok -or $erros) { Aviso $txt $(if ($erros) { 48 } else { 64 }) }
  }
  else {
    # o lançador parou no meio: não tenta de novo sozinho
    $fim = ($saida | Select-Object -Last 15) -join "`r`n"
    Set-Content -Path $parado -Encoding UTF8 -Value ("O lançador automático parou com erro em " + (Get-Date).ToString("dd/MM/yyyy HH:mm") + ".`r`nEnquanto este arquivo existir, nada é lançado sozinho.`r`nDepois de resolver, apague este arquivo.`r`n`r`n" + $fim)
    Aviso ("O lançador parou com um erro e NÃO vai tentar de novo sozinho.`n`nOs arquivos continuam na ENTRADA. Veja o arquivo 'PARADO - leia.txt' nessa pasta.`n`n" + (($saida | Select-Object -Last 4) -join "`n")) 16
  }
}
catch { Aviso ("Erro no vigia do lançador:`n`n" + $_.Exception.Message) 16 }
finally { try { $mutex.ReleaseMutex() } catch {} }
