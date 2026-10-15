<#
  ATUALIZA O PAINEL NO SERVIDOR LINUX SEM DERRUBAR (blue/green, ver deploy\release.sh)

  Uso:
    powershell -ExecutionPolicy Bypass -File deploy\deploy.ps1            # atualiza de verdade
    powershell -ExecutionPolicy Bypass -File deploy\deploy.ps1 -SoVerificar   # so confere e empacota, nao envia

  Configuracao: deploy\server.conf (copie de server.conf.example)
#>
param([switch]$SoVerificar)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$deployDir = $PSScriptRoot
$id = Get-Date -Format 'yyyyMMdd-HHmmss'

function Fail($msg) { Write-Host "ERRO: $msg" -ForegroundColor Red; exit 1 }
function Step($msg) { Write-Host "`n== $msg" -ForegroundColor Cyan }

# ---------- 1) conferencia local (nada vai pro ar se isso falhar) ----------
Step "1/4 Conferindo o codigo antes de enviar"
$pyFiles = Get-ChildItem -Path $root -Filter *.py -File
foreach ($f in $pyFiles) {
    python -c "import ast,sys; ast.parse(open(sys.argv[1],encoding='utf-8').read())" $f.FullName
    if ($LASTEXITCODE -ne 0) { Fail "erro de sintaxe em $($f.Name)" }
}
Write-Host "   Python OK ($($pyFiles.Count) arquivos)"

$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
    $tmpJs = Join-Path $env:TEMP "dash_inline_$id.js"
    python -c "import re,sys; h=open(sys.argv[1],encoding='utf-8').read(); s=re.findall(r'<script>(.*?)</script>',h,re.S); open(sys.argv[2],'w',encoding='utf-8').write(max(s,key=len))" (Join-Path $root 'templates\index.html') $tmpJs
    node --check $tmpJs
    if ($LASTEXITCODE -ne 0) { Fail "erro de sintaxe no JavaScript do index.html" }
    Remove-Item $tmpJs -ErrorAction SilentlyContinue
    Write-Host "   JavaScript do painel OK"
} else {
    Write-Host "   (node nao encontrado: JavaScript nao conferido)" -ForegroundColor Yellow
}

# ---------- 2) empacotar so o que roda no servidor ----------
Step "2/4 Empacotando a release $id"
$stage = Join-Path $env:TEMP "dash_stage_$id"
New-Item -ItemType Directory -Path $stage | Out-Null
Copy-Item (Join-Path $root '*.py') $stage
foreach ($n in 'requirements.txt','Dockerfile','.dockerignore') { Copy-Item (Join-Path $root $n) $stage }
foreach ($d in 'templates','static') { Copy-Item (Join-Path $root $d) (Join-Path $stage $d) -Recurse }
# arquivos que rodam no Linux precisam de LF (nao CRLF)
foreach ($n in 'requirements.txt','Dockerfile','.dockerignore','gunicorn.conf.py') {
    $p = Join-Path $stage $n
    if (Test-Path $p) {
        $t = [System.IO.File]::ReadAllText($p) -replace "`r`n", "`n"
        [System.IO.File]::WriteAllText($p, $t, (New-Object System.Text.UTF8Encoding($false)))
    }
}
$tgz = Join-Path $env:TEMP "dash_$id.tgz"
tar -czf $tgz -C $stage .
if ($LASTEXITCODE -ne 0) { Fail "nao consegui gerar o pacote" }
$mb = [math]::Round((Get-Item $tgz).Length / 1MB, 1)
Write-Host "   Pacote: $tgz ($mb MB)"

# scripts do servidor (LF)
$srvStage = Join-Path $env:TEMP "dash_srv_$id"
New-Item -ItemType Directory -Path $srvStage | Out-Null
foreach ($n in 'release.sh','rollback.sh','run-in-active.sh','backup.sh','Caddyfile','docker-compose.proxy.yml') {
    $t = [System.IO.File]::ReadAllText((Join-Path $deployDir $n)) -replace "`r`n", "`n"
    [System.IO.File]::WriteAllText((Join-Path $srvStage $n), $t, (New-Object System.Text.UTF8Encoding($false)))
}

if ($SoVerificar) {
    Write-Host "`nModo -SoVerificar: tudo certo, nada foi enviado." -ForegroundColor Green
    Write-Host "Conteudo do pacote:"; tar -tzf $tgz | Select-Object -First 25
    Remove-Item $stage,$srvStage -Recurse -Force; exit 0
}

# ---------- 3) enviar ----------
Step "3/4 Enviando para o servidor"
$confPath = Join-Path $deployDir 'server.conf'
if (-not (Test-Path $confPath)) { Fail "falta deploy\server.conf (copie de server.conf.example)" }
$c = @{}
Get-Content $confPath | Where-Object { $_ -match '^\s*[A-Za-z_]+\s*=' } | ForEach-Object {
    $k, $v = $_ -split '=', 2; $c[$k.Trim()] = $v.Trim()
}
foreach ($k in 'HOST','USER','KEY') { if (-not $c[$k]) { Fail "server.conf sem $k" } }
$port = if ($c['PORT']) { $c['PORT'] } else { '22' }
$remote = if ($c['REMOTE_ROOT']) { $c['REMOTE_ROOT'] } else { '/opt/dashboard' }
$target = "$($c['USER'])@$($c['HOST'])"
$sshOpt = @('-i', $c['KEY'], '-o', 'StrictHostKeyChecking=accept-new', '-o', 'ServerAliveInterval=30')

scp -P $port @sshOpt $tgz "${target}:${remote}/$id.tgz"
if ($LASTEXITCODE -ne 0) { Fail "falha ao enviar o pacote" }
scp -P $port @sshOpt (Join-Path $srvStage '*') "${target}:${remote}/"
if ($LASTEXITCODE -ne 0) { Fail "falha ao enviar os scripts do servidor" }

# ---------- 4) instalar e trocar sem derrubar ----------
Step "4/4 Instalando a release e trocando sem derrubar o painel"
$cmd = "set -e; cd $remote; mkdir -p releases/$id; tar -xzf $id.tgz -C releases/$id; rm -f $id.tgz; chmod +x *.sh; ./release.sh $id"
ssh -p $port @sshOpt $target $cmd
$rc = $LASTEXITCODE
Remove-Item $stage,$srvStage,$tgz -Recurse -Force -ErrorAction SilentlyContinue
if ($rc -ne 0) { Fail "a atualizacao falhou. A versao anterior continua no ar (veja o log acima)." }
Write-Host "`nPronto: release $id no ar. $($c['URL'])" -ForegroundColor Green
