# -*- coding: utf-8 -*-
"""Visualizador de planilhas dentro do painel (somente leitura).

Lista os relatórios de cada dia (brutos e tratados, por cliente/conta) e as tabelas de custos,
e lê qualquer um deles em páginas, com busca, ordenação e totais das colunas numéricas.
Nunca altera os arquivos.
"""
import csv
import datetime as dt
import glob
import io
import os
import re
import threading
import zipfile

import comum

EXTS = (".xlsx", ".xlsm", ".csv", ".txt", ".zip")
_CACHE = {}          # chave (caminho, mtime, aba) -> dados
_ORDEM_CACHE = []
_TRAVA = threading.Lock()
MAX_CACHE = 4
MAX_LINHAS = 300000


# ---------------------------------------------------------------- segurança
def _permitido(cfg, caminho):
    a = os.path.normcase(os.path.abspath(caminho))
    bases = [comum.raiz(cfg), os.path.join(comum.PASTA_SISTEMA, "dados", "custos")]
    return any(a.startswith(os.path.normcase(os.path.abspath(b)) + os.sep) for b in bases)


# ---------------------------------------------------------------- listagem
def _dias(cfg):
    """Pastas DD-MM da raiz, da mais recente para a mais antiga."""
    r = comum.raiz(cfg)
    if not os.path.isdir(r):
        return []
    dias = [d for d in os.listdir(r) if re.fullmatch(r"\d{2}-\d{2}", d) and os.path.isdir(os.path.join(r, d))]
    return sorted(dias, key=lambda d: os.path.getmtime(os.path.join(r, d)), reverse=True)


def _arquivos(pasta):
    if not os.path.isdir(pasta):
        return []
    return sorted(f for f in os.listdir(pasta) if f.lower().endswith(EXTS) and os.path.isfile(os.path.join(pasta, f)))


def _rotulo_mk(cfg, nome):
    m = re.match(r"vendas_([a-z0-9-]+)_", nome)
    if m and m.group(1) in cfg["marketplaces"]:
        return cfg["marketplaces"][m.group(1)]["nome"]
    return nome


def listar(cfg, dia=None):
    dias = _dias(cfg)
    hoje = comum.rotulo_dia()
    if hoje not in dias:
        dias.insert(0, hoje)
    dia = dia if dia in dias else dias[0]
    grupos = []
    for cli in cfg.get("clientes", []):
        contas = []
        for u in comum.contas_do_cliente(cfg, cli):
            pasta = os.path.join(comum.raiz(cfg), dia, comum.pasta_de(u))
            brutos = sorted(({"nome": f, "rotulo": _rotulo_mk(cfg, f), "caminho": os.path.join(pasta, f)} for f in _arquivos(pasta)), key=lambda x: x["rotulo"])
            trat = os.path.join(pasta, "tratados")
            tratados = sorted(({"nome": f, "rotulo": _rotulo_mk(cfg, f), "caminho": os.path.join(trat, f)} for f in _arquivos(trat)), key=lambda x: x["rotulo"])
            if brutos or tratados:
                contas.append({"id": u["id"], "nome": u["conta_nome"], "brutos": brutos, "tratados": tratados})
        custos = None
        import tratador
        info_c, _ = tratador.escolher_custos(cli)
        arq = (info_c or {}).get("arquivo")
        if arq and os.path.isfile(os.path.join(comum.PASTA_SISTEMA, "dados", arq)):
            custos = {"nome": "Tabela de custos", "rotulo": "Tabela de custos", "caminho": "custos:" + cli["id"]}
        pasta_cli = os.path.join(comum.raiz(cfg), dia, cli["nome"])
        consolidados = [{"nome": f, "rotulo": "Consolidado " + (re.search(r"_(\d{2}-\d{4})\.xlsx$", f) or [None, ""])[1],
                         "caminho": os.path.join(pasta_cli, f)}
                        for f in _arquivos(pasta_cli) if f.lower().startswith("relatorio_consolidado")]
        if contas or custos or consolidados:
            grupos.append({"id": cli["id"], "nome": cli["nome"], "contas": contas, "custos": custos,
                           "consolidados": consolidados})
    return {"dias": dias, "dia": dia, "grupos": grupos}


# ---------------------------------------------------------------- leitura
def _valor(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return "VERDADEIRO" if v else "FALSO"
    if isinstance(v, float):
        return None if v != v else (int(v) if v.is_integer() and abs(v) < 1e15 else v)
    if isinstance(v, int):
        return v
    if isinstance(v, dt.datetime):
        return v.strftime("%d/%m/%Y %H:%M") if (v.hour or v.minute) else v.strftime("%d/%m/%Y")
    if isinstance(v, dt.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    s = str(v)
    return s if s.strip() != "" else None


def _decodificar(bruto):
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return bruto.decode(cod)
        except UnicodeDecodeError:
            continue


def _linhas_texto(txt, tab=False):
    amostra = txt[:20000]
    if tab or amostra.count("\t") > amostra.count(";") and amostra.count("\t") > amostra.count(","):
        delim = "\t"
    else:
        try:
            delim = csv.Sniffer().sniff(amostra, delimiters=";,|\t").delimiter
        except csv.Error:
            delim = ";" if amostra.count(";") > amostra.count(",") else ","
    return [[_valor(c) for c in r] for r in csv.reader(io.StringIO(txt), delimiter=delim)]


def abas(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(caminho, read_only=True)
        try:
            return [ws.title for ws in wb.worksheets]
        finally:
            wb.close()
    if ext == ".zip":
        with zipfile.ZipFile(caminho) as z:
            return [n for n in z.namelist() if n.lower().endswith((".csv", ".txt", ".xlsx"))]
    return ["Dados"]


def _ler_bruto(caminho, aba):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:   # leitor rápido (python-calamine); sem ele, usa o openpyxl
            from python_calamine import CalamineWorkbook
            wb = CalamineWorkbook.from_path(caminho)
            nome = aba if aba in wb.sheet_names else wb.sheet_names[0]
            linhas = []
            for r in wb.get_sheet_by_name(nome).to_python(skip_empty_area=False):
                linhas.append([_valor(c) for c in r])
                if len(linhas) > MAX_LINHAS:
                    break
            return linhas
        except ImportError:
            pass
        from openpyxl import load_workbook
        wb = load_workbook(caminho, read_only=True, data_only=True)
        try:
            ws = wb[aba] if aba in wb.sheetnames else wb.worksheets[0]
            linhas = []
            for r in ws.iter_rows(values_only=True):
                linhas.append([_valor(c) for c in r])
                if len(linhas) > MAX_LINHAS:
                    break
            return linhas
        finally:
            wb.close()
    if ext == ".zip":
        with zipfile.ZipFile(caminho) as z:
            nomes = abas(caminho)
            nome = aba if aba in nomes else (nomes[0] if nomes else None)
            if not nome:
                return []
            dados = z.read(nome)
            if nome.lower().endswith(".xlsx"):
                tmp = os.path.join(comum.pasta_dados(None, "dados"), ".zip_visualizar.xlsx")
                with open(tmp, "wb") as f:
                    f.write(dados)
                return _ler_bruto(tmp, None)
            return _linhas_texto(_decodificar(dados))
    with open(caminho, "rb") as f:
        return _linhas_texto(_decodificar(f.read()), tab=ext == ".txt")


def _cabecalho(linhas):
    """Índice da linha de cabeçalho: entre as 30 primeiras, a com mais textos (pelo menos 2,
    e quase só texto). Empate fica com a de cima. Ex.: Mercado Livre tem o cabeçalho na linha 6."""
    melhor, idx = 1, 0
    for i, r in enumerate(linhas[:30]):
        textos = sum(1 for c in r if isinstance(c, str))
        preenchidas = sum(1 for c in r if c is not None)
        if textos > melhor and textos >= 0.6 * preenchidas:
            melhor, idx = textos, i
    return idx


def _carregar(caminho, aba):
    chave = (caminho, os.path.getmtime(caminho), aba)
    with _TRAVA:
        if chave in _CACHE:
            return _CACHE[chave]
    linhas = _ler_bruto(caminho, aba)
    while linhas and all(c is None for c in linhas[-1]):
        linhas.pop()
    h = _cabecalho(linhas) if linhas else 0
    ncol = max((len(r) for r in linhas), default=0)
    while ncol and all((len(r) < ncol or r[ncol - 1] is None) for r in linhas):
        ncol -= 1
    cab = list(linhas[h]) if linhas else []
    cab = [(str(c) if c is not None else "Coluna %d" % (j + 1)) for j, c in enumerate(cab[:ncol] + [None] * (ncol - len(cab)))]
    corpo = [list(r[:ncol]) + [None] * (ncol - len(r)) for r in linhas[h + 1:] if any(c is not None for c in r)]
    numericas = [j for j in range(ncol)
                 if sum(1 for r in corpo if isinstance(r[j], (int, float))) >= max(1, 0.8 * sum(1 for r in corpo if r[j] is not None))
                 and any(isinstance(r[j], (int, float)) for r in corpo)]
    dados = {"colunas": cab, "linhas": corpo, "numericas": numericas, "linha_cabecalho": h + 1,
             "acima": [[c for c in r if c is not None] for r in linhas[:h] if any(c is not None for c in r)][:6],
             "truncado": len(linhas) > MAX_LINHAS}
    with _TRAVA:
        _CACHE[chave] = dados
        _ORDEM_CACHE.append(chave)
        while len(_ORDEM_CACHE) > MAX_CACHE:
            _CACHE.pop(_ORDEM_CACHE.pop(0), None)
    return dados


def _chave_ordem(v):
    if v is None:
        return (2, 0, "")
    if isinstance(v, (int, float)):
        return (0, v, "")
    return (1, 0, str(v).lower())


def _resolver(cfg, caminho):
    """'custos:<cliente>' aponta para a tabela de custos guardada do cliente."""
    if caminho and caminho.startswith("custos:"):
        partes = caminho.split(":")
        cli = next((c for c in cfg.get("clientes", []) if c["id"] == partes[1]), None)
        if not cli:
            raise ValueError("Cliente não encontrado.")
        import tratador
        if len(partes) > 2:
            a, m = (int(x) for x in partes[2].split("-"))
            info, mes = tratador.escolher_custos(cli, a, m)
        else:
            info, mes = tratador.escolher_custos(cli)
        if not info or not info.get("arquivo"):
            raise ValueError("Este cliente ainda não tem tabela de custos.")
        rot = ("%s/%s" % (mes[5:], mes[:4])) if mes else ""
        return (os.path.join(comum.PASTA_SISTEMA, "dados", info["arquivo"]),
                "Tabela de custos — %s%s (versão %s)" % (cli["nome"], (" · " + rot) if rot else "", info.get("versao", 1)))
    return caminho, None


def ler(cfg, caminho, aba=None, pagina=1, por_pagina=100, busca="", ordem=None, desc=False, filtros=None):
    pedido = caminho
    caminho, titulo = _resolver(cfg, caminho)
    if not caminho or not _permitido(cfg, caminho):
        raise ValueError("Só abro planilhas da pasta de relatórios ou da tabela de custos.")
    if not os.path.isfile(caminho):
        raise ValueError("Arquivo não encontrado (pode ter sido movido).")
    lista_abas = abas(caminho)
    aba = aba if aba in lista_abas else (lista_abas[0] if lista_abas else None)
    d = _carregar(caminho, aba)
    linhas = d["linhas"]
    busca = (busca or "").strip().lower()
    if busca:
        linhas = [r for r in linhas if any(c is not None and busca in str(c).lower() for c in r)]
    for j, termo in (filtros or {}).items():
        try:
            j = int(j)
        except ValueError:
            continue
        termo = str(termo).strip().lower()
        if termo and 0 <= j < len(d["colunas"]):
            linhas = [r for r in linhas if r[j] is not None and termo in str(r[j]).lower()]
    if ordem is not None and str(ordem).lstrip("-").isdigit() and 0 <= int(ordem) < len(d["colunas"]):
        j = int(ordem)
        vazias = [r for r in linhas if r[j] is None]
        linhas = sorted((r for r in linhas if r[j] is not None), key=lambda r: _chave_ordem(r[j]), reverse=bool(desc)) + vazias
    total = len(linhas)
    somas = {j: sum(r[j] for r in linhas if isinstance(r[j], (int, float))) for j in d["numericas"]}
    por_pagina = max(10, min(int(por_pagina or 100), 1000))
    paginas = max(1, -(-total // por_pagina))
    pagina = max(1, min(int(pagina or 1), paginas))
    ini = (pagina - 1) * por_pagina
    st = os.stat(caminho)
    return {"ok": True, "arquivo": titulo or os.path.basename(caminho), "caminho": caminho, "pedido": pedido, "abas": lista_abas, "aba": aba,
            "colunas": d["colunas"], "numericas": d["numericas"], "linhas": linhas[ini:ini + por_pagina],
            "total": total, "total_arquivo": len(d["linhas"]), "pagina": pagina, "paginas": paginas,
            "por_pagina": por_pagina, "somas": somas, "linha_cabecalho": d["linha_cabecalho"], "acima": d["acima"],
            "truncado": d["truncado"], "tamanho": st.st_size,
            "modificado": dt.datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds")}
