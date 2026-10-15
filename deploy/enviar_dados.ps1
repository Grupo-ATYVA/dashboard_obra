<#
  Envia a pasta dados\ (planilhas, usuarios, certificado, caches) para o servidor.
  Faca isso ANTES de abrir o painel para o publico: a pasta contem os usuarios e senhas (hash)
  e, sem ela, o primeiro visitante poderia criar o administrador.

  CUIDADO: sobrescreve o dados\ do servidor. Por isso pede confirmacao.
  Uso:
    powershell -ExecutionPolicy Bypass -File deploy\enviar_dados.ps1
    powershell -ExecutionPolicy Bypass -File deploy\enviar_dados.ps1 -SemModelosGrandes   # pula os .nwd (~500 MB)
#>
param([switch]$SemModelosGrandes)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$confPath = Join-Path $PSScriptRoot 'server.conf'
if (-not (Test-Path $confPath)) { Write-Host "falta deploy\server.conf" -ForegroundColor Red; exit 1 }
$c = @{}
Get-Content $confPath | Where-Object { $_ -match '^\s*[A-Za-z_]+\s*=' } | ForEach-Object { $k, $v = $_ -split '=', 2; $c[$k.Trim()] = $v.Trim() }
$port = if ($c['PORT']) { $c['PORT'] } else { '22' }
$remote = if ($c['REMOTE_ROOT']) { $c['REMOTE_ROOT'] } else { '/opt/dashboard' }
$target = "$($c['USER'])@$($c['HOST'])"
$sshOpt = @('-i', $c['KEY'], '-o', 'StrictHostKeyChecking=accept-new', '-o', 'ServerAliveInterval=30')

Write-Host "Isto vai SOBRESCREVER $remote/dados no servidor $($c['HOST'])." -ForegroundColor Yellow
$r = Read-Host "Digite SIM para continuar"
if ($r -ne 'SIM') { Write-Host "Cancelado."; exit 0 }

$tgz = Join-Path $env:TEMP 'dados_envio.tgz'
$excl = @('--exclude=.melius_profile*', '--exclude=.sienge_profile*', '--exclude=*.log')
if ($SemModelosGrandes) { $excl += '--exclude=*.nwd' }
Write-Host "Compactando dados\ ..."
tar -czf $tgz @excl -C $root dados
if ($LASTEXITCODE -ne 0) { Write-Host "falha ao compactar" -ForegroundColor Red; exit 1 }
Write-Host ("Pacote: {0} MB" -f [math]::Round((Get-Item $tgz).Length / 1MB, 1))

scp -P $port @sshOpt $tgz "${target}:/tmp/dados_envio.tgz"
if ($LASTEXITCODE -ne 0) { Write-Host "falha no envio" -ForegroundColor Red; exit 1 }
ssh -p $port @sshOpt $target "set -e; mkdir -p $remote; tar -xzf /tmp/dados_envio.tgz -C $remote; chown -R 1000:1000 $remote/dados; chmod 700 $remote/dados; rm -f /tmp/dados_envio.tgz; echo 'dados enviados:'; du -sh $remote/dados"
Remove-Item $tgz -ErrorAction SilentlyContinue
Write-Host "Pronto." -ForegroundColor Green
