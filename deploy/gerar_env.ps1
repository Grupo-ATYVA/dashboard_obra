<#
  Monta o arquivo de segredos do servidor (deploy\servidor.env) a partir do iniciar_dashboard.bat.
  As senhas NAO sao mostradas na tela: so o nome de cada variavel e se foi preenchida.

  Uso:
    powershell -ExecutionPolicy Bypass -File deploy\gerar_env.ps1 -Dominio painel.suaempresa.com.br -Email ti@suaempresa.com.br

  Depois envie para o servidor (so por SSH) e proteja:
    scp deploy\servidor.env root@SERVIDOR:/opt/dashboard/.env
    ssh root@SERVIDOR "chmod 600 /opt/dashboard/.env"
  E apague o servidor.env local quando terminar.
#>
param(
    [Parameter(Mandatory = $true)][string]$Dominio,
    [Parameter(Mandatory = $true)][string]$Email
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bat = Join-Path $root 'iniciar_dashboard.bat'
if (-not (Test-Path $bat)) { Write-Host "iniciar_dashboard.bat nao encontrado" -ForegroundColor Red; exit 1 }

$nomes = 'AUTODOC_EMAIL','AUTODOC_SENHA','CERT_PFX_SENHA','MELIUS_USER','MELIUS_PASSWORD','MELIUS_SITE_SENHA','MELIUS_TOKEN','DIRECTD_API_TOKEN'
$vals = @{}
foreach ($line in (Get-Content $bat -Encoding UTF8)) {
    if ($line -match '^\s*set\s+"([A-Za-z_0-9]+)=(.*)"\s*$') {
        if ($nomes -contains $Matches[1]) { $vals[$Matches[1]] = $Matches[2] }
    }
}

function NovoSegredo { $b = New-Object byte[] 32; [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($b); -join ($b | ForEach-Object { $_.ToString('x2') }) }

$out = New-Object System.Text.StringBuilder
[void]$out.AppendLine("DASH_DOMAIN=$Dominio")
[void]$out.AppendLine("ACME_EMAIL=$Email")
[void]$out.AppendLine("DASHBOARD_HTTPS=1")
[void]$out.AppendLine("DASHBOARD_SECRET_KEY=$(NovoSegredo)")
[void]$out.AppendLine("DASHBOARD_SETUP_TOKEN=$(NovoSegredo)")
foreach ($n in $nomes) {
    $v = if ($vals.ContainsKey($n)) { $vals[$n] } else { '' }
    [void]$out.AppendLine("$n=$v")
}
[void]$out.AppendLine("GUNICORN_THREADS=16")

$dest = Join-Path $PSScriptRoot 'servidor.env'
[System.IO.File]::WriteAllText($dest, ($out.ToString() -replace "`r`n", "`n"), (New-Object System.Text.UTF8Encoding($false)))

Write-Host "Arquivo gerado: $dest" -ForegroundColor Green
foreach ($n in $nomes) {
    $st = if ($vals.ContainsKey($n) -and $vals[$n]) { 'preenchida' } else { 'VAZIA (preencha a mao se for usar)' }
    Write-Host ("  {0,-20} {1}" -f $n, $st)
}
Write-Host "  DASHBOARD_SECRET_KEY / DASHBOARD_SETUP_TOKEN: gerados agora (aleatorios)."
Write-Host "`nGuarde o DASHBOARD_SETUP_TOKEN: so e usado no primeiro acesso, se a pasta dados for criada do zero."
