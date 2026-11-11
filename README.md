# Relatórios de Vendas

Sistema online de relatórios de vendas dos marketplaces: os relatórios são subidos pelo site, e o sistema
faz o tratamento das vendas, ADS, afiliados, visitas, pacote final, forecast e o painel.

## Partes
- `sistema/` — o código (Python + `painel.html`).
- `render.yaml` — sistema online no Render (Starter + disco de 10 GB em `/var/data`).

## Online
- Variáveis: `RV_NUVEM=1`, `RELATORIOS_SISTEMA` e `RELATORIOS_RAIZ` (no disco), `ADMIN_USUARIO`, `ADMIN_SENHA`, `TZ`.
- Login: cada pessoa tem o seu usuário (aba **Usuários**); o admin principal vem de `ADMIN_USUARIO` / `ADMIN_SENHA`.
- Levar dados de um computador (se precisar de novo): `python sistema/exportar_nuvem.py --url https://... --senha ...`.

Os dados dos clientes (config, custos, relatórios) ficam só no disco do Render — nunca no repositório.
