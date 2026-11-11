@echo off
rem Remonta o app com os arquivos atuais de _sistema e abre de novo (o mesmo que o botao "Atualizar sistema").
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1"
