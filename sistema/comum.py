# -*- coding: utf-8 -*-
"""Funções compartilhadas: configuração, pastas, período, contas e histórico."""
import datetime as dt
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import unicodedata

def _achar_pasta_sistema():
    """Pasta com config.json. No .exe (PyInstaller) sobe a partir do executável até achar."""
    if os.environ.get("RELATORIOS_SISTEMA"):
        return os.environ["RELATORIOS_SISTEMA"]
    if getattr(sys, "frozen", False):
        p = os.path.dirname(os.path.abspath(sys.executable))
        for _ in range(5):
            if os.path.exists(os.path.join(p, "config.json")):
                return p
            p = os.path.dirname(p)
    return os.path.dirname(os.path.abspath(__file__))


PASTA_SISTEMA = _achar_pasta_sistema()
# Online (Render): RV_NUVEM=1. Os dados ficam no disco (RELATORIOS_SISTEMA / RELATORIOS_RAIZ) e o código na pasta do app.
NUVEM = os.environ.get("RV_NUVEM") == "1"
PASTA_CODIGO = PASTA_SISTEMA if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
RAIZ_NUVEM = os.environ.get("RELATORIOS_RAIZ") or os.path.join(os.path.dirname(PASTA_SISTEMA), "Relatorios")
TRAVA_CFG = threading.RLock()   # quem lê-altera-grava o config.json (app e tratamentos)
ARQ_CONFIG = os.environ.get("RELATORIOS_CONFIG") or os.path.join(PASTA_SISTEMA, "config.json")
CARACTERES_INVALIDOS = '\\/:*?"<>|'


# ---------------------------------------------------------------- config
def _config_inicial():
    """Online, na primeira vez (disco vazio): começa do modelo que vem com o código."""
    if NUVEM and not os.path.exists(ARQ_CONFIG):
        os.makedirs(os.path.dirname(ARQ_CONFIG), exist_ok=True)
        modelo = os.path.join(PASTA_CODIGO, "config.json")
        shutil.copy2(modelo, ARQ_CONFIG)


def carregar_config():
    _config_inicial()
    with open(ARQ_CONFIG, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    migrar_contas(cfg)  # em memória; o servidor grava a migração ao iniciar
    if NUVEM:
        cfg["pasta_raiz"] = RAIZ_NUVEM          # o config pode ter vindo do Windows (C:\\Relatórios)
        cfg["pasta_entrada"] = "_entrada"
        cfg["nuvem"] = True
    return cfg


def salvar_config(cfg):
    pasta = os.path.dirname(ARQ_CONFIG)
    fd, tmp = tempfile.mkstemp(prefix="config_", suffix=".json", dir=pasta)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    if os.path.exists(ARQ_CONFIG):
        os.makedirs(pasta_dados(cfg, "backups"), exist_ok=True)
        shutil.copy2(ARQ_CONFIG, os.path.join(pasta_dados(cfg, "backups"), "config_ultimo.json"))
    os.replace(tmp, ARQ_CONFIG)


def pasta_dados(cfg, nome):
    """Subpastas do sistema: logs, dados, backups."""
    p = os.path.join(PASTA_SISTEMA, nome)
    os.makedirs(p, exist_ok=True)
    return p


# ---------------------------------------------------------------- pastas
def raiz(cfg):
    return cfg["pasta_raiz"]


def pasta_entrada(cfg, cliente_nome=None):
    p = os.path.join(raiz(cfg), cfg.get("pasta_entrada", "_entrada"))
    return os.path.join(p, cliente_nome) if cliente_nome else p


def rotulo_dia(data=None):
    data = data or dt.date.today()
    return data.strftime("%d-%m")


def pasta_dia(cfg, cliente_nome, data=None):
    return os.path.join(raiz(cfg), rotulo_dia(data), cliente_nome)


def nome_arquivo(mk_id, cfg, data=None):
    return "vendas_%s_%s%s" % (mk_id, rotulo_dia(data), cfg["marketplaces"][mk_id]["extensao"])


def criar_pastas_base(cfg):
    os.makedirs(pasta_entrada(cfg), exist_ok=True)


def pasta_de(unidade):
    """Subcaminho da conta: Cliente\\Conta (cai para o nome se for um cadastro antigo)."""
    return unidade.get("pasta") or unidade["nome"]


def criar_pastas_cliente(cfg, cliente):
    """Aceita um cliente (grupo) ou uma conta já 'achatada'."""
    if "contas" in cliente:
        for u in contas_do_cliente(cfg, cliente):
            os.makedirs(pasta_entrada(cfg, pasta_de(u)), exist_ok=True)
    else:
        os.makedirs(pasta_entrada(cfg, pasta_de(cliente)), exist_ok=True)


# ---------------------------------------------------------------- grupos e contas
SEP_CONTA = "--"  # gerar_id nunca produz "--", então o id da conta é único e reversível


SHOPEE_ANTECIPA_PADRAO = 0.025


def _unidade(cliente, conta, total_contas):
    return {
        "id": cliente["id"] + SEP_CONTA + conta["id"],
        "cliente_id": cliente["id"], "cliente_nome": cliente["nome"],
        "conta_id": conta["id"], "conta_nome": conta["nome"],
        "nome": cliente["nome"] if conta["nome"] == cliente["nome"] else "%s / %s" % (cliente["nome"], conta["nome"]),
        "loja": conta["nome"],
        "pasta": os.path.join(cliente["nome"], conta["nome"]),
        "unica_conta": total_contas == 1,
        "marketplaces": conta.get("marketplaces", []),
        "ativo": bool(cliente.get("ativo", True)) and bool(conta.get("ativo", True)),
        "imposto": cliente.get("imposto"), "custos": cliente.get("custos"), "custos_meses": cliente.get("custos_meses") or {},
        "ids_shopee": cliente.get("ids_shopee"),
        # antecipação Shopee da conta, em fração (padrão 2,5%; 0 = a loja não usa antecipação)
        "shopee_antecipa": (cliente.get("shopee_antecipa") or {}).get(conta["id"], SHOPEE_ANTECIPA_PADRAO),
    }


# Marketplaces que entraram depois (o relatório é subido pelo app). Entram nas configurações antigas pelo _migrar do servidor.
MARKETPLACES_NOVOS = {
    "temu": {
        "nome": "Temu",
        "padroes_nome": [
            "temu",
            "order_export"
        ],
        "palavras_conteudo": [
            "temu",
            "sku de contribuicao",
            "preco-base dos produtos",
            "desconto da temu",
            "approved courier list"
        ],
        "extensao": ".xlsx"
    },
    "viavarejo": {
        "nome": "Via Varejo",
        "padroes_nome": [
            "via.?varejo",
            "casas.?bahia",
            "grupo.?casas"
        ],
        "palavras_conteudo": [
            "via varejo",
            "casas bahia"
        ],
        "extensao": ".xlsx"
    },
    "madeiramadeira": {
        "nome": "Madeira Madeira",
        "padroes_nome": [
            "madeira.?madeira",
            "report_pedido"
        ],
        "palavras_conteudo": [
            "madeiramadeira",
            "madeira madeira",
            "pedido site mm",
            "nome do seller",
            "status cliente entrega"
        ],
        "extensao": ".xlsx"
    },
    "site": {
        "nome": "Site",
        "padroes_nome": [
            "^site[ _.-]",
            "loja.?propria",
            "vendas.?site"
        ],
        "palavras_conteudo": [],
        "extensao": ".xlsx",
        "sem_tratador": True
    }
}


# o que era só do download automático (Chrome/extensão/Claude Code): sai do cadastro
CHAVES_DOWNLOAD_CFG = ("chrome_user_data", "chrome_exe", "horarios", "claude_exe", "execucao")
CHAVES_DOWNLOAD_MK = ("link", "link_status", "link_fonte", "instrucoes", "roteiros", "sem_download")
CHAVES_DOWNLOAD_CONTA = ("modo", "perfil", "computador", "links")


def limpar_download(cfg):
    """Tira do cadastro os campos que só o download automático usava. Devolve True se mudou algo."""
    mudou = False
    for k in CHAVES_DOWNLOAD_CFG:
        mudou |= cfg.pop(k, None) is not None
    for m in (cfg.get("marketplaces") or {}).values():
        for k in CHAVES_DOWNLOAD_MK:
            mudou |= m.pop(k, None) is not None
    for c in cfg.get("clientes", []):
        for k in CHAVES_DOWNLOAD_CONTA:
            mudou |= c.pop(k, None) is not None
        for ct in c.get("contas", []):
            for k in CHAVES_DOWNLOAD_CONTA:
                mudou |= ct.pop(k, None) is not None
    return mudou


def _arq_ok_vendas():
    return os.path.join(PASTA_SISTEMA, "dados", "ok_vendas.json")


def oks_vendas():
    """{conta_id: {mk: {"AAAA-MM-DD": {"motivo", "quando"}}}} — dia sem relatório de vendas aceito à mão."""
    try:
        with open(_arq_ok_vendas(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def ok_vendas(conta_id, mk, dia):
    return (oks_vendas().get(conta_id, {}).get(mk, {}) or {}).get(dia.isoformat() if hasattr(dia, "isoformat") else str(dia))


def marcar_ok_vendas(conta_id, mk, dia, motivo="", remover=False):
    todos = oks_vendas()
    d = (dia.isoformat() if hasattr(dia, "isoformat") else str(dia))[:10]
    m = todos.setdefault(conta_id, {}).setdefault(mk, {})
    if remover:
        m.pop(d, None)
    else:
        m[d] = {"motivo": (motivo or "OK manual").strip()[:200], "quando": dt.datetime.now().isoformat(timespec="seconds")}
    os.makedirs(os.path.dirname(_arq_ok_vendas()), exist_ok=True)
    with open(_arq_ok_vendas() + ".novo", "w", encoding="utf-8") as f:
        json.dump(todos, f, ensure_ascii=False, indent=1)
    os.replace(_arq_ok_vendas() + ".novo", _arq_ok_vendas())


def contas_do_cliente(cfg, cliente):
    lista = cliente.get("contas") or []
    return [_unidade(cliente, ct, len(lista)) for ct in lista]


def contas(cfg, so_ativas=False):
    """Todas as contas de todos os clientes, no formato usado pela rotina."""
    migrar_contas(cfg)  # aceita cadastro antigo mesmo se vier de fora do carregar_config
    saida = []
    for c in cfg.get("clientes", []):
        for u in contas_do_cliente(cfg, c):
            if not so_ativas or u["ativo"]:
                saida.append(u)
    return saida


def conta(cfg, uid):
    return next((u for u in contas(cfg) if u["id"] == uid), None)


def filtrar_contas(cfg, ids, so_ativas=True):
    """ids pode ter ids de conta ou de cliente (grupo inteiro)."""
    return [u for u in contas(cfg, so_ativas) if not ids or u["id"] in ids or u["cliente_id"] in ids]


def migrar_contas(cfg):
    """Cadastro antigo (perfil/marketplaces no cliente) -> cliente com uma conta de mesmo nome."""
    mudou = False
    for c in cfg.get("clientes", []):
        if "contas" not in c:
            c.pop("perfil", None)
            c.pop("links", None)
            c["contas"] = [{"id": gerar_id(c["nome"], set()), "nome": c["nome"],
                            "marketplaces": c.pop("marketplaces", []), "ativo": True}]
            mudou = True
    return mudou


def validar_nome_cliente(nome):
    nome = (nome or "").strip()
    if not nome:
        return "Informe o nome do cliente."
    if any(c in CARACTERES_INVALIDOS for c in nome):
        return "O nome não pode conter: " + " ".join(CARACTERES_INVALIDOS)
    if nome.startswith("_") or nome.endswith(".") or nome.endswith(" "):
        return "O nome não pode começar com _ nem terminar com ponto ou espaço."
    return None


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFKD", txt or "") if not unicodedata.combining(c)).lower()


def gerar_id(nome, existentes):
    base = re.sub(r"[^a-z0-9]+", "-", sem_acento(nome)).strip("-") or "cliente"
    base = base[:30]
    novo, n = base, 2
    while novo in existentes:
        novo = "%s-%d" % (base, n)
        n += 1
    return novo


# ---------------------------------------------------------------- período
def periodo(hoje=None):
    """(inicio, fim) do relatório. Dia 1 -> mês anterior inteiro; senão dia 1 do mês até ontem."""
    hoje = hoje or dt.date.today()
    if hoje.day == 1:
        fim = hoje - dt.timedelta(days=1)
        inicio = fim.replace(day=1)
    else:
        inicio = hoje.replace(day=1)
        fim = hoje - dt.timedelta(days=1)
    return inicio, fim


def fmt(d):
    return d.strftime("%d/%m/%Y")


def devolver_memoria():
    """Depois de um trabalho pesado (pandas/planilhas grandes), devolve ao sistema a memória já liberada.
    No Render o limite é de 512 MB: sem isso o processo fica com a memória do pico mesmo parado."""
    import gc
    gc.collect()
    try:
        import ctypes
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except (OSError, AttributeError):
        pass


# ---------------------------------------------------------------- histórico
def arq_historico():
    return os.path.join(PASTA_SISTEMA, "dados", "historico.jsonl")


MUDOU = [0]   # conta alterações (log/eventos) para as telas se atualizarem sozinhas


def registrar_evento(ev):
    MUDOU[0] += 1
    os.makedirs(os.path.dirname(arq_historico()), exist_ok=True)
    ev = dict(ev)
    ev.setdefault("ts", dt.datetime.now().isoformat(timespec="seconds"))
    with open(arq_historico(), "a", encoding="utf-8") as f:
        f.write(json.dumps(ev, ensure_ascii=False) + "\n")


_HIST = {"chave": None, "pos": 0, "evs": []}
_HIST_MAX = 20000
_TRAVA_HIST = threading.Lock()


def ler_historico(limite=20000):
    """Últimos eventos do histórico. Fica em memória: a cada leitura só as linhas novas do arquivo são lidas."""
    arq = arq_historico()
    try:
        st = os.stat(arq)
    except OSError:
        return []
    with _TRAVA_HIST:
        chave = (st.st_ino, st.st_dev)
        if _HIST["chave"] != chave or st.st_size < _HIST["pos"]:      # arquivo trocado ou encolheu: lê de novo
            _HIST.update(chave=chave, pos=0, evs=[])
        if st.st_size > _HIST["pos"]:
            with open(arq, "rb") as f:
                f.seek(_HIST["pos"])
                bruto = f.read()
            fim = bruto.rfind(b"\n") + 1          # só linhas completas (a última pode estar sendo gravada)
            for l in bruto[:fim].decode("utf-8", errors="replace").splitlines():
                try:
                    _HIST["evs"].append(json.loads(l))
                except ValueError:
                    pass
            _HIST["pos"] += fim
            if len(_HIST["evs"]) > _HIST_MAX:
                del _HIST["evs"][:-_HIST_MAX]
        evs = _HIST["evs"][-limite:] if limite else []
    return [dict(e) for e in evs]


def log(msg, cfg=None):
    MUDOU[0] += 1
    pasta = os.path.join(PASTA_SISTEMA, "logs")
    os.makedirs(pasta, exist_ok=True)
    agora = dt.datetime.now()
    with open(os.path.join(pasta, "organizador_%s.log" % agora.strftime("%Y-%m")), "a", encoding="utf-8") as f:
        f.write("%s  %s\n" % (agora.strftime("%Y-%m-%d %H:%M:%S"), msg))

