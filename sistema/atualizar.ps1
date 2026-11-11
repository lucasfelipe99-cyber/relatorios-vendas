# Atualiza o app "Relatorios de Vendas": espera o app fechar, remonta o PainelRelatorios.exe com os
# arquivos atuais de _sistema (mesmo comando do instalador) e abre o app de novo.
# Chamado pelo botao "Atualizar sistema" do painel (com -AppPid) ou a mao pelo ATUALIZAR.bat.
# Se a montagem falhar, a versao anterior do app continua (nada e apagado).
param([int]$AppPid = 0)
$ErrorActionPreference = 'Stop'
$Sistema = $PSScriptRoot
$AppDir  = Join-Path $Sistema 'app'
$Atual   = Join-Path $AppDir 'PainelRelatorios'
$Exe     = Join-Path $Atual 'PainelRelatorios.exe'
$Log     = Join-Path $Sistema 'logs\atualizar.log'
New-Item -ItemType Directory -Force -Path (Split-Path $Log) | Out-Null
$Host.UI.RawUI.WindowTitle = 'Relatorios de Vendas - atualizando'
# PYTHONPATH/PYTHONHOME de outros programas (ex.: VLibras) e variaveis do .exe que chamou este script
foreach ($v in 'PYTHONPATH','PYTHONHOME') { Remove-Item "Env:$v" -ErrorAction SilentlyContinue }
Get-ChildItem Env: | Where-Object { $_.Name -like '_PYI*' -or $_.Name -like '_MEI*' } | ForEach-Object { Remove-Item "Env:$($_.Name)" }

function Msg($t, $cor = 'Gray') {
    $linha = '{0:dd/MM HH:mm:ss}  {1}' -f (Get-Date), $t
    Write-Host $linha -ForegroundColor $cor
    Add-Content -Path $Log -Value $linha -Encoding UTF8
}

function Abrir-App {
    Remove-Item Env:PYINSTALLER_RESET_ENVIRONMENT -ErrorAction SilentlyContinue
    if (Test-Path $Exe) {
        Start-Process -FilePath $Exe -WorkingDirectory $Sistema
    } else {
        $py = (Get-Content -Raw -Encoding UTF8 (Join-Path $Sistema 'python.txt')).Trim()
        $pyw = Join-Path (Split-Path $py) 'pythonw.exe'
        if (-not (Test-Path $pyw)) { $pyw = $py }
        Start-Process -FilePath $pyw -ArgumentList @('-I', ('"' + (Join-Path $Sistema 'app.py') + '"')) -WorkingDirectory $Sistema
    }
    Msg 'App aberto.' 'Green'
}

function Esperar-Fechar {
    if ($AppPid -gt 0) {
        Msg "Esperando o app fechar (processo $AppPid)..."
        $p = Get-Process -Id $AppPid -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(30000)) { Msg 'O app nao fechou sozinho: encerrando.' 'Yellow'; Stop-Process -Id $AppPid -Force -ErrorAction SilentlyContinue }
    }
    # qualquer outra copia do .exe (app aberto, organizador agendado) precisa terminar antes da troca
    for ($i = 0; $i -lt 120; $i++) {
        $outros = Get-Process -Name 'PainelRelatorios' -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path -like "$AppDir*" }
        if (-not $outros) { return }
        if ($i -eq 0 -and $AppPid -eq 0) {
            $r = Read-Host 'O app esta aberto. Fechar agora para atualizar? [S/N]'
            if ($r -notmatch '^[sS]') { throw 'Atualizacao cancelada: o app continua aberto.' }
            $outros | Stop-Process -Force -ErrorAction SilentlyContinue
        } elseif ($i -eq 60) {
            Msg 'Ainda ha uma copia do app rodando: encerrando.' 'Yellow'
            $outros | Stop-Process -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 1
    }
    throw 'O app continua aberto; feche pelo icone da bandeja (Sair) e rode ATUALIZAR.bat.'
}

try {
    Msg '=== Atualizacao do app ===' 'Cyan'
    Esperar-Fechar
    $py = (Get-Content -Raw -Encoding UTF8 (Join-Path $Sistema 'python.txt')).Trim()
    if (-not (Test-Path $py)) { throw "Python do app nao encontrado ($py). Rode o INSTALAR.bat." }

    Msg 'Conferindo os arquivos do sistema...'
    & $py -I -c "import sys, py_compile, glob, os; [py_compile.compile(f, doraise=True) for f in glob.glob(os.path.join(sys.argv[1], '*.py'))]; print('ok')" $Sistema
    if ($LASTEXITCODE -ne 0) { throw 'Ha um arquivo .py com erro (veja acima). O app anterior foi mantido.' }

    if (-not (Test-Path $Exe)) {
        # o .exe sumiu (antivirus, montagem interrompida): remonta em vez de seguir pelo Python (icone errado)
        & $py -I -c "import PyInstaller" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Msg 'App roda pelo Python e o PyInstaller nao esta instalado: so reabrindo (rode o INSTALAR.bat para voltar ao .exe).' 'Yellow'
            Abrir-App
            exit 0
        }
        Msg 'O PainelRelatorios.exe nao esta na pasta do app: montando de novo.' 'Yellow'
    }

    Msg 'Montando o app com as alteracoes (1 a 3 minutos)...' 'Cyan'
    $novo  = Join-Path $AppDir '_novo'
    $build = Join-Path $env:TEMP 'relatorios_build'
    if (Test-Path $novo) { Remove-Item $novo -Recurse -Force }
    $ocultos = @('--hidden-import', 'openpyxl', '--hidden-import', 'pystray._win32')
    Get-ChildItem -Path $Sistema -Filter '*.py' | Where-Object { $_.BaseName -ne 'app' } |
        ForEach-Object { $ocultos += @('--hidden-import', $_.BaseName) }
    & $py -I -m PyInstaller --noconfirm --clean --windowed --log-level WARN --name PainelRelatorios `
        --icon (Join-Path $Sistema 'icone.ico') --distpath $novo --workpath $build --specpath $build `
        --paths $Sistema @ocultos (Join-Path $Sistema 'app.py')
    $novoExe = Join-Path $novo 'PainelRelatorios\PainelRelatorios.exe'
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $novoExe)) { throw 'Nao consegui montar o app novo (antivirus?). O app anterior foi mantido.' }

    Msg 'Trocando pela versao nova...'
    $anterior = Join-Path $AppDir 'PainelRelatorios_anterior'
    if (Test-Path $anterior) { Remove-Item $anterior -Recurse -Force }
    $tinha = Test-Path $Atual
    if ($tinha) { Rename-Item -Path $Atual -NewName 'PainelRelatorios_anterior' }
    try {
        Move-Item -Path (Join-Path $novo 'PainelRelatorios') -Destination $Atual
    } catch {
        if ($tinha) { Rename-Item -Path $anterior -NewName 'PainelRelatorios' }   # volta a anterior
        throw "Nao consegui trocar a pasta do app: $($_.Exception.Message). O app anterior foi mantido."
    }
    Remove-Item $novo -Recurse -Force -ErrorAction SilentlyContinue
    Msg 'App atualizado.' 'Green'
    Abrir-App
    Start-Sleep -Seconds 3
} catch {
    Msg ("ERRO: " + $_.Exception.Message) 'Red'
    try { Abrir-App } catch { Msg ("Nao consegui abrir o app: " + $_.Exception.Message) 'Red' }
    Write-Host ''
    Read-Host 'Pressione Enter para fechar esta janela'
    exit 1
}
