@echo off
chcp 65001 >nul
echo Instalando o app Relatorios de Vendas...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1"
if errorlevel 1 (echo. & echo Ocorreu um erro. Leia a mensagem acima.)
pause
