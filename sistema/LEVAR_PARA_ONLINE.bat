@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  == Levar os dados deste computador para o sistema online ==
echo  Monta um pacote com clientes, custos, ADS/afiliados/visitas, tratados e pacotes
echo  e envia para https://relatorios-vendas-np0v.onrender.com
echo  Depois disso o sistema passa a ser usado so pelo link (sem o app e sem baixar sozinho).
echo  Feche o app Relatorios de Vendas antes de continuar.
echo.
set /p SENHA=Senha do admin do sistema online: 
".venv\Scripts\python.exe" -I exportar_nuvem.py --url https://relatorios-vendas-np0v.onrender.com --senha "%SENHA%" --estacao ""
if errorlevel 1 (
  echo.
  echo  Deu erro. Leia a mensagem acima e me mande um print.
  pause
  exit /b 1
)
echo.
echo  Dados enviados. Tirando os agendamentos deste computador (o download automatico das 05:00)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0remover_agendamentos.ps1"
taskkill /IM PainelRelatorios.exe /F >nul 2>&1
echo.
echo  Pronto! Use o sistema pelo link: https://relatorios-vendas-np0v.onrender.com
start "" https://relatorios-vendas-np0v.onrender.com
pause
