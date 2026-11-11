# Relatórios de Vendas

Sistema de relatórios de vendas dos marketplaces: download dos relatórios (nos computadores, pelo Chrome),
tratamento das vendas, ADS, afiliados, visitas, pacote final e painel.

## Partes
- `sistema/` — o código (Python + `painel.html`).
- `INSTALAR.bat` / `instalar.ps1` — instalação no Windows (app de mesa, agendamentos, Chrome).
- `render.yaml` — sistema online no Render (Starter + disco de 10 GB em `/var/data`).

## Online
- Variáveis: `RV_NUVEM=1`, `RELATORIOS_SISTEMA` e `RELATORIOS_RAIZ` (no disco), `ADMIN_USUARIO`, `ADMIN_SENHA`, `TZ`.
- Login: usuário e senha do admin. As estações usam a mesma senha.
- Os computadores ficam como **estação** apontando para o endereço online (`https://...onrender.com`):
  baixam no Chrome deles e enviam os relatórios; o tratamento e o pacote final acontecem online.
- Levar os dados de um computador: `python sistema/exportar_nuvem.py --url https://... --senha ... --estacao NOME_DO_PC`.

Os dados dos clientes (config, custos, relatórios) ficam só no disco do Render e nos computadores — nunca no repositório.
