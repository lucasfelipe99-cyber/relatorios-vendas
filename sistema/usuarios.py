# -*- coding: utf-8 -*-
"""Usuários do sistema online: login e senha de cada pessoa, sessão por cookie assinado.

- O admin principal vem das variáveis do Render (ADMIN_USUARIO / ADMIN_SENHA) e não pode ser removido daqui.
- Os outros usuários ficam em dados/usuarios.json, com a senha guardada só como hash (PBKDF2).
- Perfis: "admin" (também cadastra usuários) e "usuario" (usa o sistema todo, menos o cadastro de usuários).
- Trocar a senha de alguém derruba as sessões abertas daquela pessoa.
"""
import base64
import datetime as dt
import hashlib
import hmac
import json
import os
import re
import secrets
import threading
import time

import comum

ARQ = os.path.join(comum.PASTA_SISTEMA, "dados", "usuarios.json")
ARQ_SEGREDO = os.path.join(comum.PASTA_SISTEMA, "dados", "segredo_sessao.txt")
DURACAO = 30 * 24 * 3600          # sessão vale 30 dias
_TRAVA = threading.RLock()


def _hash(senha, sal=None, voltas=200000):
    sal = sal or secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), sal.encode("ascii"), voltas).hex()
    return "pbkdf2$%d$%s$%s" % (voltas, sal, h)


def _confere(senha, guardado):
    try:
        _, voltas, sal, h = guardado.split("$")
        return hmac.compare_digest(_hash(senha, sal, int(voltas)).split("$")[3], h)
    except (ValueError, AttributeError):
        return False


def _segredo():
    with _TRAVA:
        try:
            with open(ARQ_SEGREDO, "r", encoding="ascii") as f:
                s = f.read().strip()
            if len(s) >= 32:
                return s.encode("ascii")
        except OSError:
            pass
        s = secrets.token_hex(32)
        os.makedirs(os.path.dirname(ARQ_SEGREDO), exist_ok=True)
        with open(ARQ_SEGREDO, "w", encoding="ascii") as f:
            f.write(s)
        return s.encode("ascii")


def _ler():
    try:
        with open(ARQ, "r", encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    d.setdefault("usuarios", [])
    return d


def _gravar(d):
    os.makedirs(os.path.dirname(ARQ), exist_ok=True)
    with open(ARQ + ".tmp", "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(ARQ + ".tmp", ARQ)


def _admin_env():
    return (os.environ.get("ADMIN_USUARIO") or "admin").strip().lower(), os.environ.get("ADMIN_SENHA") or ""


def login_admin():
    return _admin_env()[0]


def senha_admin():
    return _admin_env()[1]


def _versao_env():
    return hashlib.sha256(("env:" + _admin_env()[1]).encode("utf-8")).hexdigest()[:12]


def _publico(u):
    return {k: u.get(k) for k in ("login", "nome", "perfil", "ativo", "criado", "alterado", "ultimo_acesso")}


def achar(login):
    """Usuário pelo login (o admin das variáveis do Render incluído)."""
    login = (login or "").strip().lower()
    adm, senha_adm = _admin_env()
    if login == adm:
        return {"login": adm, "nome": "Admin", "perfil": "admin", "ativo": bool(senha_adm), "principal": True,
                "versao": _versao_env()}
    for u in _ler()["usuarios"]:
        if u["login"] == login:
            return dict(u, versao=u.get("hash", "")[-12:])
    return None


def entrar(login, senha):
    """Confere login e senha. Devolve o usuário ou None."""
    u = achar(login)
    if not u or not u.get("ativo") or not senha:
        return None
    if u.get("principal"):
        ok = hmac.compare_digest(senha, _admin_env()[1])
    else:
        ok = _confere(senha, u.get("hash", ""))
    if not ok:
        return None
    if not u.get("principal"):
        with _TRAVA:
            d = _ler()
            for x in d["usuarios"]:
                if x["login"] == u["login"]:
                    x["ultimo_acesso"] = dt.datetime.now().isoformat(timespec="seconds")
            _gravar(d)
    return u


def sessao_nova(u):
    exp = int(time.time()) + DURACAO
    corpo = base64.urlsafe_b64encode(json.dumps([u["login"], exp, u["versao"]]).encode("utf-8")).decode("ascii").rstrip("=")
    sig = hmac.new(_segredo(), corpo.encode("ascii"), hashlib.sha256).hexdigest()
    return corpo + "." + sig


def sessao(valor):
    """Cookie -> usuário logado (ou None se vencido, adulterado, desativado ou com senha trocada)."""
    try:
        corpo, sig = (valor or "").split(".")
        if not hmac.compare_digest(sig, hmac.new(_segredo(), corpo.encode("ascii"), hashlib.sha256).hexdigest()):
            return None
        login, exp, versao = json.loads(base64.urlsafe_b64decode(corpo + "=" * (-len(corpo) % 4)))
        if exp < time.time():
            return None
        u = achar(login)
        if not u or not u.get("ativo") or u.get("versao") != versao:
            return None
        return u
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------- cadastro (só admin)
def listar():
    adm = achar(_admin_env()[0])
    return [dict(_publico(adm), principal=True)] + [_publico(u) for u in _ler()["usuarios"]]


def salvar(dados, por):
    """Cria ou altera um usuário. Senha vazia na alteração = mantém a atual."""
    login = re.sub(r"[^a-z0-9._-]", "", str(dados.get("login") or "").strip().lower())[:40]
    nome = str(dados.get("nome") or "").strip()[:60]
    perfil = "admin" if dados.get("perfil") == "admin" else "usuario"
    senha = str(dados.get("senha") or "")
    if len(login) < 3:
        raise ValueError("O login precisa de pelo menos 3 letras ou números (sem espaço nem acento).")
    if login == _admin_env()[0]:
        raise ValueError("Esse é o login do admin principal: a senha dele é trocada no Render (ADMIN_SENHA).")
    if not nome:
        raise ValueError("Informe o nome da pessoa.")
    with _TRAVA:
        d = _ler()
        u = next((x for x in d["usuarios"] if x["login"] == login), None)
        agora = dt.datetime.now().isoformat(timespec="seconds")
        if u is None:
            if dados.get("so_alterar"):
                raise ValueError("Usuário não encontrado.")
            if len(senha) < 8:
                raise ValueError("A senha precisa ter pelo menos 8 caracteres.")
            u = {"login": login, "criado": agora, "criado_por": por}
            d["usuarios"].append(u)
        elif senha and len(senha) < 8:
            raise ValueError("A senha precisa ter pelo menos 8 caracteres.")
        u.update(nome=nome, perfil=perfil, ativo=bool(dados.get("ativo", True)), alterado=agora)
        if senha:
            u["hash"] = _hash(senha)
        _gravar(d)
    return _publico(u)


def remover(login):
    login = (login or "").strip().lower()
    with _TRAVA:
        d = _ler()
        antes = len(d["usuarios"])
        d["usuarios"] = [u for u in d["usuarios"] if u["login"] != login]
        _gravar(d)
        return antes != len(d["usuarios"])


def trocar_minha_senha(login, atual, nova):
    u = achar(login)
    if not u or u.get("principal"):
        raise ValueError("A senha do admin principal é trocada no Render (Environment → ADMIN_SENHA).")
    if not entrar(login, atual):
        raise ValueError("Senha atual incorreta.")
    if len(nova or "") < 8:
        raise ValueError("A senha nova precisa ter pelo menos 8 caracteres.")
    with _TRAVA:
        d = _ler()
        for x in d["usuarios"]:
            if x["login"] == u["login"]:
                x["hash"] = _hash(nova)
                x["alterado"] = dt.datetime.now().isoformat(timespec="seconds")
        _gravar(d)
    return True
