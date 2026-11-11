# Registra (ou atualiza) as tarefas do Agendador de Tarefas do Windows a partir do config.json.
# Não precisa de administrador: as tarefas rodam na sua conta, com você conectado.
$ErrorActionPreference = 'Stop'
$Sistema = $PSScriptRoot
# Só precisa de "horarios" e "execucao". O config inteiro pode ter chaves que só diferem em maiúscula/minúscula
# (nomes de produto), que o ConvertFrom-Json do PowerShell recusa: em caso de erro, o Python lê só essas duas partes.
$arqCfg = Join-Path $Sistema 'config.json'
try {
    $cfg = Get-Content -Raw -Encoding UTF8 $arqCfg | ConvertFrom-Json
} catch {
    $pyCfg = Join-Path $Sistema '.venv\Scripts\python.exe'
    if (-not (Test-Path $pyCfg) -and (Test-Path (Join-Path $Sistema 'python.txt'))) {
        $pyCfg = (Get-Content -Raw -Encoding UTF8 (Join-Path $Sistema 'python.txt')).Trim()
    }
    # sem aspas duplas no código: o PowerShell 5 as perde ao passar para outro programa
    $codigo = 'import json,sys; c=json.load(open(sys.argv[1],encoding=''utf-8'')); print(json.dumps({''horarios'':c.get(''horarios'') or {},''execucao'':c.get(''execucao'') or {}}))'
    $cfg = (& $pyCfg -I -c $codigo $arqCfg) -join "`n" | ConvertFrom-Json
}
# Preferência: o app (.exe com Python embutido). Sem ele, usa o Python instalado + app.py.
$exe = Join-Path $Sistema 'app\PainelRelatorios\PainelRelatorios.exe'
if (Test-Path $exe) {
    $Programa = $exe; $Prefixo = ''
} else {
    $py = (Get-Content -Raw -Encoding UTF8 (Join-Path $Sistema 'python.txt')).Trim()
    $Programa = Join-Path (Split-Path $py) 'pythonw.exe'
    if (-not (Test-Path $Programa)) { $Programa = $py }
    $Prefixo = '-I "' + (Join-Path $Sistema 'app.py') + '" '   # -I: ignora PYTHONPATH de outros programas
}
Write-Host "Programa usado nas tarefas: $Programa"

$usuario = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited

function Settings($limiteMin) {
    $s = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -AllowStartIfOnBatteries `
         -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    if ($limiteMin -eq 0) { $s.ExecutionTimeLimit = 'PT0S' } else { $s.ExecutionTimeLimit = "PT$($limiteMin)M" }
    return $s
}

function Acao($argumento) {
    New-ScheduledTaskAction -Execute $Programa -Argument ($Prefixo + $argumento) -WorkingDirectory $Sistema
}

$h = $cfg.horarios

# 1) Baixar relatórios: o app aciona o Claude Code + extensão, um cliente por vez
$antiga = Get-ScheduledTask -TaskName 'Relatorios - 1 Abrir Chrome' -ErrorAction SilentlyContinue
if ($antiga) { Unregister-ScheduledTask -TaskName 'Relatorios - 1 Abrir Chrome' -Confirm:$false; Write-Host "[removida] tarefa antiga 'Abrir Chrome' (o app agora abre cada perfil na hora certa)" }
$agendado = $false
if ($cfg.execucao -and $cfg.execucao.agendado) { $agendado = [bool]$cfg.execucao.agendado }
if ($agendado) {
    $t1 = New-ScheduledTaskTrigger -Daily -At $h.inicio_atalhos
    Register-ScheduledTask -TaskName 'Relatorios - 1 Baixar relatorios' -Action (Acao '--rodar-todos') -Trigger $t1 `
        -Principal $principal -Settings (Settings 0) -Description 'Baixa os relatórios de todos os clientes (Claude Code + Claude in Chrome), um cliente por vez.' -Force | Out-Null
    Write-Host ("[ok] Baixar relatorios: todo dia as {0}" -f $h.inicio_atalhos)
} else {
    if (Get-ScheduledTask -TaskName 'Relatorios - 1 Baixar relatorios' -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName 'Relatorios - 1 Baixar relatorios' -Confirm:$false
    }
    Write-Host "[ok] Download automatico DESLIGADO: os relatorios so sao baixados quando voce clicar em Baixar no app."
}

# 2) Organizador a cada N minutos entre inicio e fim
$ini = [datetime]::ParseExact($h.organizador_inicio, 'HH:mm', $null)
$fim = [datetime]::ParseExact($h.organizador_fim, 'HH:mm', $null)
if ($fim -le $ini) { $fim = $fim.AddDays(1) }
$duracao = New-TimeSpan -Minutes ([int]($fim - $ini).TotalMinutes + 1)
$t2 = New-ScheduledTaskTrigger -Daily -At $h.organizador_inicio
$t2.Repetition = (New-ScheduledTaskTrigger -Once -At $h.organizador_inicio `
    -RepetitionInterval (New-TimeSpan -Minutes ([int]$h.organizador_intervalo_min)) -RepetitionDuration $duracao).Repetition
Register-ScheduledTask -TaskName 'Relatorios - 2 Organizador' -Action (Acao '--organizar') -Trigger $t2 `
    -Principal $principal -Settings (Settings 15) -Description 'Move os relatórios da _entrada para a pasta do dia.' -Force | Out-Null
Write-Host ("[ok] Organizador: todo dia das {0} as {1}, a cada {2} min" -f $h.organizador_inicio, $h.organizador_fim, $h.organizador_intervalo_min)

# 3) Painel: inicia junto com o Windows (ao entrar na sua conta)
$t3 = New-ScheduledTaskTrigger -AtLogOn -User $usuario
Register-ScheduledTask -TaskName 'Relatorios - 0 Painel' -Action (Acao '--bandeja') -Trigger $t3 `
    -Principal $principal -Settings (Settings 0) -Description 'App Relatórios de Vendas no ícone da bandeja (servidor local usado pela extensão).' -Force | Out-Null
Write-Host "[ok] App: inicia na bandeja quando voce entra no Windows"

Write-Host ""
Write-Host "Nao agende atalhos na extensao: o app aciona a extensao quando voce clica em Baixar (ou no horario, se o download automatico estiver ligado)."
