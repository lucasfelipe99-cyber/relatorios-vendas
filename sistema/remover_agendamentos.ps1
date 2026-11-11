# Remove apenas as tarefas agendadas. Não apaga pastas, relatórios, config nem logs.
foreach ($n in 'Relatorios - 0 Painel', 'Relatorios - 1 Abrir Chrome', 'Relatorios - 2 Organizador', 'Relatorios - 1 Baixar relatorios') {
    if (Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $n -Confirm:$false
        Write-Host "[removida] $n"
    }
}
Write-Host "Pronto. Arquivos e pastas foram mantidos."
