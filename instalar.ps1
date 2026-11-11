# Instalador do app "Relatórios de Vendas".
# Cria C:\Relatórios e C:\Relatórios\_entrada, instala o sistema em C:\Relatórios\_sistema,
# garante o Python, monta o app (PainelRelatorios.exe), registra os agendamentos e cria os atalhos.
# Pode ser executado de novo para atualizar: o config.json existente (seus clientes) é mantido.
$ErrorActionPreference = 'Stop'
$Raiz = 'C:\Relatórios'
$Sistema = Join-Path $Raiz '_sistema'
$Origem = Join-Path $PSScriptRoot 'sistema'
$NomeApp = 'Relatórios de Vendas'

function Passo($t) { Write-Host ""; Write-Host "== $t" -ForegroundColor Cyan }

# Outros programas (ex.: VLibras) definem PYTHONPATH/PYTHONHOME no Windows e fazem qualquer Python
# carregar as bibliotecas deles. Ignoramos isso SÓ dentro deste instalador; nada muda no Windows.
foreach ($v in 'PYTHONPATH', 'PYTHONHOME', 'PYTHONSTARTUP', 'PYTHONUSERBASE') { Remove-Item "Env:$v" -ErrorAction SilentlyContinue }

Passo 'Fechando versões em execução (se houver)'
Get-Process PainelRelatorios -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like 'python*' -and $_.CommandLine -like '*\_sistema\*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1
Write-Host "[ok]"

Passo 'Pastas'
foreach ($p in $Raiz, (Join-Path $Raiz '_entrada'), $Sistema) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null; Write-Host "[criada] $p" } else { Write-Host "[ok] $p" }
}

Passo 'Arquivos do sistema'
Get-ChildItem $Origem -File | ForEach-Object {
    $dest = Join-Path $Sistema $_.Name
    if ($_.Name -eq 'config.json' -and (Test-Path $dest)) {
        Write-Host "[mantido] config.json existente (seus clientes e ajustes)"
    } elseif ($_.Name -eq 'magalu-ids.csv' -and (Test-Path $dest)) {
        Write-Host "[mantido] magalu-ids.csv existente (de-para Magalu)"
    } else {
        Copy-Item $_.FullName $dest -Force
        Write-Host "[copiado] $($_.Name)"
    }
}

Passo 'Python 3.10 ou mais novo (usado para montar o app)'
# Aceita só Python >= 3.10 de instalação própria; ignora Pythons embutidos em outros programas (ex.: VLibras).
function Testar-Python($exe, $extra) {
    try {
        $r = & $exe @extra -c "import sys; assert sys.version_info >= (3, 10); print(sys.executable)" 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $r) { return $null }
        $r = ($r | Select-Object -Last 1).Trim()
        if (-not (Test-Path $r) -or $r -match 'WindowsApps') { return $null }
        $oficial = ($r -like "$env:LOCALAPPDATA\Programs\Python\*") -or ($r -like "$env:ProgramFiles\Python*") -or
                   ($r -like "${env:ProgramFiles(x86)}\Python*") -or ($r -like 'C:\Python*')
        if (-not $oficial) { Write-Host "  (ignorado: $r - nao e uma instalacao propria do Python)"; return $null }
        return $r
    } catch { return $null }
}
$py = $null
foreach ($v in '3.13', '3.12', '3.11', '3.10') {
    if (-not $py -and (Get-Command py -ErrorAction SilentlyContinue)) { $py = Testar-Python 'py' @("-$v") }
}
if (-not $py) {
    foreach ($v in '313', '312', '311', '310') {
        foreach ($c in (Join-Path $env:LOCALAPPDATA "Programs\Python\Python$v\python.exe"), (Join-Path $env:ProgramFiles "Python$v\python.exe")) {
            if (-not $py -and (Test-Path $c)) { $py = Testar-Python $c @() }
        }
    }
}
if (-not $py -and (Get-Command python -ErrorAction SilentlyContinue)) { $py = Testar-Python 'python' @() }
if (-not $py) {
    Write-Host "Nao encontrei um Python 3.10+ instalado neste PC (Pythons embutidos em outros programas nao servem)."
    $resp = Read-Host "Instalar o Python 3.12 agora pelo winget (oficial da Microsoft)? [S/N]"
    if ($resp -notmatch '^[sS]') { throw "Instale o Python 3.12 em https://www.python.org/downloads/ e rode o instalador de novo." }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw "winget indisponivel. Instale o Python 3.12 em https://www.python.org/downloads/ e rode o instalador de novo." }
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    $c = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
    if (Test-Path $c) { $py = $c } else { throw "Python instalado, mas nao localizado. Feche esta janela e rode o instalador de novo." }
}
Write-Host "[ok] Python base: $py"

Passo 'Ambiente isolado do app (nao mexe em outros Pythons do PC)'
$venv = Join-Path $Sistema '.venv'
$vpy = Join-Path $venv 'Scripts\python.exe'
if (Test-Path $vpy) {
    $venvOk = $false
    try { & $vpy -c "import sys; assert sys.version_info >= (3, 10)" 2>$null; $venvOk = ($LASTEXITCODE -eq 0) } catch {}
    if (-not $venvOk) { Write-Host "Ambiente antigo incompativel; recriando."; Rename-Item $venv ($venv + '_antigo_' + (Get-Date -Format 'yyyyMMddHHmmss')) }
}
if (-not (Test-Path $vpy)) { & $py -m venv $venv; if ($LASTEXITCODE -ne 0) { throw "Nao consegui criar o ambiente do app." } }
$py = $vpy
Write-Host "[ok] $py"
Set-Content -Path (Join-Path $Sistema 'python.txt') -Value $py -Encoding UTF8

Passo 'Bibliotecas (janela, icone na bandeja, Excel, tratamento e empacotador)'
& $py -m pip install --quiet --disable-pip-version-check --no-warn-script-location --upgrade pip setuptools wheel
& $py -m pip install --quiet --disable-pip-version-check --no-warn-script-location --upgrade openpyxl pandas pywebview pystray pillow python-calamine xlsxwriter pyinstaller
if ($LASTEXITCODE -ne 0) { throw "Nao consegui instalar as bibliotecas (verifique a internet e rode de novo)." }
Write-Host "[ok] bibliotecas"

Passo 'Teste rapido do sistema'
& $py -c "import sys; sys.path.insert(0, r'$Sistema'); import comum, organizador, servidor, tratador, tratamento_mercado_livre, tratamento_shopee, tratamento_magalu, tratamento_tiktok, tratamento_amazon; cfg=comum.carregar_config(); comum.criar_pastas_base(cfg); i,f=comum.periodo(); print('[ok] periodo de hoje:', comum.fmt(i), 'a', comum.fmt(f)); print('[ok] perfis do Chrome encontrados:', len(comum.listar_perfis(cfg))); print('[ok] chrome.exe:', comum.chrome_exe(cfg) or 'NAO ENCONTRADO - informe no app')"
if ($LASTEXITCODE -ne 0) { throw "O teste do sistema falhou (veja a mensagem acima)." }

Passo 'Montando o app (PainelRelatorios.exe) - leva 1 a 3 minutos'
$build = Join-Path $env:TEMP 'relatorios_build'
$appDir = Join-Path $Sistema 'app'
& $py -m PyInstaller --noconfirm --clean --windowed --log-level WARN --name PainelRelatorios `
    --icon (Join-Path $Sistema 'icone.ico') --distpath $appDir --workpath $build --specpath $build `
    --paths $Sistema --hidden-import openpyxl --hidden-import pystray._win32 `
    --hidden-import servidor --hidden-import organizador --hidden-import abrir_chrome --hidden-import executor `
    --hidden-import tratador --hidden-import planilhas --hidden-import consolidado --hidden-import lancamentos --hidden-import forecast --hidden-import painel_dados --hidden-import rede --hidden-import estacao --hidden-import atividade --hidden-import xlsxwriter --hidden-import compactados --hidden-import python_calamine --hidden-import tratamento_mercado_livre --hidden-import tratamento_shopee --hidden-import tratamento_magalu --hidden-import tratamento_tiktok --hidden-import tratamento_amazon --hidden-import tratamento_casas_bahia --hidden-import tratamento_temu --hidden-import tratamento_madeira_madeira `
    (Join-Path $Sistema 'app.py')
$exe = Join-Path $appDir 'PainelRelatorios\PainelRelatorios.exe'
if ($LASTEXITCODE -eq 0 -and (Test-Path $exe)) {
    Write-Host "[ok] $exe"
    $Programa = $exe; $Args0 = ''
} else {
    Write-Host "[aviso] Nao consegui montar o .exe (antivirus?). O app vai rodar pelo Python instalado - funciona igual." -ForegroundColor Yellow
    $Programa = Join-Path (Split-Path $py) 'pythonw.exe'; if (-not (Test-Path $Programa)) { $Programa = $py }
    $Args0 = '-I "' + (Join-Path $Sistema 'app.py') + '"'
}

Passo 'Agendamentos do Windows'
& (Join-Path $Sistema 'agendar.ps1')

Passo 'Atalhos (area de trabalho e menu Iniciar)'
$antigo = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Painel Relatorios.lnk'
if (Test-Path $antigo) { Remove-Item $antigo -Force; Write-Host "[substituido] atalho antigo 'Painel Relatorios'" }
$ws = New-Object -ComObject WScript.Shell
foreach ($pasta in [Environment]::GetFolderPath('Desktop'), (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs')) {
    $lnk = $ws.CreateShortcut((Join-Path $pasta "$NomeApp.lnk"))
    $lnk.TargetPath = $Programa
    $lnk.Arguments = $Args0
    $lnk.WorkingDirectory = $Sistema
    $lnk.IconLocation = (Join-Path $Sistema 'icone.ico')
    $lnk.Description = 'Painel da rotina de relatorios de vendas'
    $lnk.Save()
    Write-Host "[ok] atalho em $pasta"
}

Passo 'Abrindo o app'
if ($Args0) { Start-Process -FilePath $Programa -ArgumentList $Args0 -WorkingDirectory $Sistema }
else { Start-Process -FilePath $Programa -WorkingDirectory $Sistema }

Write-Host ""
Write-Host "LEMBRETE: deixe o PC ligado e SEM suspensao entre 4h50 e 8h." -ForegroundColor Yellow
Write-Host "Abrindo Configuracoes > Energia e suspensao: em 'Quando conectado, suspender apos', escolha 'Nunca'." -ForegroundColor Yellow
Start-Process 'ms-settings:powersleep'
Write-Host ""
Write-Host "Pronto! O app fica no icone perto do relogio." -ForegroundColor Green
Write-Host "Proximo passo: no app, Horarios e ajustes > Claude Code > Instalar e Fazer login." -ForegroundColor Green
Write-Host "Se o Windows mostrar 'O Windows protegeu o computador' ao abrir o app: Mais informacoes > Executar assim mesmo." -ForegroundColor Green
