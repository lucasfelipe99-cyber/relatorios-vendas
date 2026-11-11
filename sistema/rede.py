# -*- coding: utf-8 -*-
"""Uso em dois (ou mais) computadores na mesma rede.

Papéis (arquivo local dados/rede.json — cada computador tem o seu, nunca é compartilhado):
  sozinho  — como sempre: só este computador (padrão).
  servidor — este computador guarda os dados e responde aos outros pela rede (com senha).
  estacao  — este computador usa os dados do servidor: a tela é a mesma (o app repassa tudo para o servidor)
             e ele baixa, no Chrome dele, os relatórios das contas marcadas para ele.

Na estação, o servidor local (127.0.0.1) vira um repasse: a extensão do Chrome continua chamando
localhost/periodo e /registrar e o registro chega no servidor.
"""
import base64
import datetime as dt
import hashlib
import json
import os
import socket
import threading
import time
import urllib.error
import urllib.request
from urllib.parse import quote

import comum

ARQ = os.path.join(comum.PASTA_SISTEMA, "dados", "rede.json")
PADRAO = {"papel": "sozinho", "nome": "", "senha": "", "servidor": "", "usuario": "", "principal": ""}


def ler():
    if comum.NUVEM:   # online: é sempre o principal, com a senha do admin (variável ADMIN_SENHA do Render)
        return dict(PADRAO, papel="servidor", nome="Online", senha=os.environ.get("ADMIN_SENHA", ""),
                    usuario=os.environ.get("ADMIN_USUARIO", "admin"))
    try:
        with open(ARQ, "r", encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    r = dict(PADRAO, **{k: v for k, v in d.items() if k in PADRAO})
    if not r["nome"]:
        r["nome"] = (socket.gethostname() or "computador").split(".")[0]
    return r


def salvar(d):
    atual = ler()
    novo = dict(atual)
    for k in PADRAO:
        if k in d and d[k] is not None:
            novo[k] = str(d[k]).strip()
    novo["usuario"] = novo["usuario"][:40]
    if novo["papel"] not in ("sozinho", "servidor", "estacao"):
        raise ValueError("Papel inválido.")
    if novo["papel"] in ("servidor", "estacao") and len(novo["senha"]) < 4:
        raise ValueError("Defina uma senha com pelo menos 4 caracteres (a mesma nos dois computadores).")
    if novo["papel"] == "estacao":
        s = novo["servidor"]
        if not s:
            raise ValueError("Informe o endereço do computador principal (ex.: 192.168.0.10).")
        if not s.startswith("http"):
            s = "http://" + s
        if s.count(":") < 2 and not s.startswith("https://"):     # endereço online (https) usa a porta padrão
            s += ":%d" % int(comum.carregar_config().get("porta_painel", 8765))
        novo["servidor"] = s.rstrip("/")
    os.makedirs(os.path.dirname(ARQ), exist_ok=True)
    with open(ARQ + ".novo", "w", encoding="utf-8") as f:
        json.dump(novo, f, ensure_ascii=False, indent=1)
    os.replace(ARQ + ".novo", ARQ)
    return novo


def papel():
    return ler()["papel"]


def token(senha=None):
    s = ler()["senha"] if senha is None else senha
    return hashlib.sha256(("relatorios-vendas:" + (s or "")).encode("utf-8")).hexdigest() if s else ""


def ips_locais():
    ips = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    try:
        for i in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(i[4][0])
    except OSError:
        pass
    return sorted(ip for ip in ips if not ip.startswith("127."))


# ---------------------------------------------------------------- servidor: estações conectadas
ESTACOES = {}          # nome -> {"ultimo": iso, "perfis": [...], "execucao": {...}, "ip": ...}
FILA = {}              # nome -> [trabalhos]
_TRAVA = threading.Lock()
ARQ_EST = os.path.join(comum.PASTA_SISTEMA, "dados", "estacoes.json")


def _guardar_estacoes():
    try:
        with open(ARQ_EST + ".novo", "w", encoding="utf-8") as f:
            json.dump({k: {kk: vv for kk, vv in v.items() if kk != "execucao"} for k, v in ESTACOES.items()}, f, ensure_ascii=False)
        os.replace(ARQ_EST + ".novo", ARQ_EST)
    except OSError:
        pass


def _carregar_estacoes():
    try:
        with open(ARQ_EST, "r", encoding="utf-8") as f:
            for k, v in json.load(f).items():
                ESTACOES.setdefault(k, v)
    except (OSError, ValueError):
        pass


_carregar_estacoes()


def registrar_ping(dados, ip=None):
    nome = str(dados.get("nome") or "").strip()[:60]
    if not nome:
        return {"ok": False, "erro": "Estação sem nome."}
    with _TRAVA:
        e = ESTACOES.setdefault(nome, {})
        primeiro = "perfis" not in e
        e.update({"ultimo": dt.datetime.now().isoformat(timespec="seconds"), "ip": ip,
                  "perfis": dados.get("perfis") or [], "execucao": dados.get("execucao") or {},
                  "claude": bool(dados.get("claude")), "chrome": bool(dados.get("chrome")), "versao": dados.get("versao")})
        trabalhos = FILA.pop(nome, [])
    if primeiro:
        comum.log("Estação conectada: %s (%s)" % (nome, ip or "?"))
    _guardar_estacoes()
    return {"ok": True, "trabalhos": trabalhos}


def online(nome, segundos=90):
    e = ESTACOES.get(nome)
    if not e or not e.get("ultimo"):
        return False
    try:
        return (dt.datetime.now() - dt.datetime.fromisoformat(e["ultimo"])).total_seconds() < segundos
    except ValueError:
        return False


def enfileirar(nome, trabalho):
    with _TRAVA:
        FILA.setdefault(nome, []).append(dict(trabalho, criado=dt.datetime.now().isoformat(timespec="seconds")))


def resumo_estacoes():
    return {k: {"online": online(k), "ultimo": v.get("ultimo"), "ip": v.get("ip"), "perfis": v.get("perfis") or [],
                "execucao": v.get("execucao") or {}, "claude": v.get("claude"), "chrome": v.get("chrome")}
            for k, v in ESTACOES.items()}


# ---------------------------------------------------------------- estação: falar com o servidor
class ErroServidor(Exception):
    pass


def chamar(caminho, dados=None, metodo=None, timeout=30, bruto=False):
    r = ler()
    url = r["servidor"] + caminho
    corpo = None if dados is None else json.dumps(dados).encode("utf-8")
    req = urllib.request.Request(url, data=corpo, method=metodo or ("POST" if corpo is not None else "GET"),
                                 headers={"X-Painel": "1", "Content-Type": "application/json",
                                          "Authorization": "Bearer " + token(r["senha"]), "X-Estacao": quote(r["nome"]),
                                          "X-Usuario": quote(r["usuario"] or r["nome"])})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            dados_r = resp.read()
            return (dados_r, resp.headers) if bruto else json.loads(dados_r or b"{}")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise ErroServidor("Senha recusada pelo computador principal.")
        raise ErroServidor("O computador principal respondeu com erro %s." % e.code)
    except (urllib.error.URLError, OSError) as e:
        raise ErroServidor("Não consegui falar com o computador principal (%s): %s" % (r["servidor"], getattr(e, "reason", e)))


LOCAIS = ("chrome_user_data", "chrome_exe", "claude_exe", "porta_painel", "pasta_raiz", "pasta_entrada")


def mesclar(cfg, local=None):
    """Config do servidor + o que é deste computador: Chrome, Claude, pastas, horários e o 'baixar sozinho'."""
    local = local or comum.carregar_config()
    for k in LOCAIS:
        if k in local:
            cfg[k] = local[k]
        elif k in ("chrome_user_data", "chrome_exe", "claude_exe"):
            cfg[k] = ""          # caminhos do outro computador não valem aqui: detectar
    if local.get("horarios"):
        cfg["horarios"] = local["horarios"]
    cfg.setdefault("execucao", {})["agendado"] = bool(local.get("execucao", {}).get("agendado", False))
    return cfg


def config_mesclada():
    """Configuração do servidor (clientes, marketplaces, execução) + o que é deste computador."""
    return mesclar(chamar("/api/agente/config", {})["config"])


def contas_desta(cfg):
    nome = ler()["nome"]
    return [u for u in comum.contas(cfg) if (u.get("computador") or "") == nome]
