@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  == Levar os dados deste computador para o sistema online ==
echo  Monta um pacote com clientes, custos, ADS/afiliados/visitas, tratados e pacotes,
echo  envia para https://relatorios-vendas-np0v.onrender.com e liga este computador ao online.
echo  Feche o app Relatorios de Vendas antes de continuar.
echo.
set /p SENHA=Senha do admin do sistema online: 
".venv\Scripts\python.exe" -I exportar_nuvem.py --url https://relatorios-vendas-np0v.onrender.com --senha "%SENHA%" --estacao DESKTOP-86CATLG --ligar-estacao
if errorlevel 1 (
  echo.
  echo  Deu erro. Leia a mensagem acima e me mande um print.
  pause
  exit /b 1
)
echo.
echo  Dados enviados. Agora o app sera remontado e aberto ja ligado ao online.
call "%~dp0ATUALIZAR.bat"
