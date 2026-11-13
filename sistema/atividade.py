# -*- coding: utf-8 -*-
"""Trabalho em equipe (vários computadores no mesmo principal).

- AGORA: o que cada pessoa está fazendo neste momento (ações em andamento + em qual tela ela está).
- Atividade recente: quem fez o quê e quando (dados/atividade.jsonl).
- Proteção: duas ações no mesmo relatório (mesma conta + marketplace + tipo) nunca rodam juntas;
  se outra pessoa já está fazendo uma ação longa no mesmo relatório, a segunda recebe um aviso.
- VERSÃO: número que muda a cada alteração, para as telas se atualizarem em poucos segundos.
"""
import datetime as dt
import json
import os
import threading
import time
import uuid

import comum

ARQ = os.path.join(comum.PASTA_SISTEMA, "dados", "atividade.jsonl")
_TRAVA = threading.Lock()
_TRAVAS_CHAVE = {}
AGORA = {}            # token -> entrada (ações em andamento)
PRESENCA = {}         # quem -> {"usuario","pc","aba","cliente","editando","ts"}
VERSAO = {"n": int(time.time()) % 100000}
TTL_PRESENCA = 45
TTL_ACAO = 30 * 60

ABAS = {"status": "Painel geral", "vendas": "Vendas", "tratamento": "Tratamento", "ads": "ADS", "afiliados": "Afiliados",
        "pacote": "Pacote final", "meses": "Mês a mês", "forecast": "Forecast", "planilhas": "Planilhas", "clientes": "Clientes", "acomp": "Acompanhamento",
        "visitas": "Visita & Conversão", "traducao": "Tradução de SKU", "custos": "Custos e imposto", "usuarios": "Usuários", "resumo": "Resumo IA", "depara": "De-para",
        "ajuda": "Ajuda"}


def tocar():
    with _TRAVA:
        VERSAO["n"] += 1


def versao():
    return VERSAO["n"] + comum.MUDOU[0]


def _agora():
    return dt.datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------- descrição das ações
def _nomes(cfg, dados):
    cliente = conta = mk = ""
    u = comum.conta(cfg, dados.get("conta_id") or dados.get("cliente_id") or "")
    if u:
        cliente, conta = u["cliente_nome"], (u["conta_nome"] if u["conta_nome"] != u["cliente_nome"] else "")
        cid, uid = u["cliente_id"], u["id"]
    else:
        cid = dados.get("cliente_id") or dados.get("id") or dados.get("para") or ""
        c = next((x for x in cfg.get("clientes", []) if x["id"] == cid), None)
        cliente = c["nome"] if c else (dados.get("nome") or "")
        uid = cid
    m = dados.get("mk")
    if m and m in cfg.get("marketplaces", {}):
        mk = cfg["marketplaces"][m]["nome"]
    return cid, uid, cliente, conta, m or "", mk


def descrever(caminho, dados, cfg):
    """None = ação que não precisa aparecer (leituras)."""
    cid, uid, cliente, conta, m, mk = _nomes(cfg, dados)
    tipo = dados.get("tipo") or "ads"
    nome_tipo = {"ads": "ADS", "visitas": "visitas"}.get(tipo, "afiliados")
    camp = (" · campanha " + str(dados.get("campanha"))) if dados.get("campanha") else ""
    T = {
        "/api/lancamento/subir": ("subindo " + nome_tipo + camp, "%s|%s|%s" % (tipo, uid, m), True),
        "/api/lancamento/ok": (("tirando OK de " if dados.get("remover") else "dando OK em ") + nome_tipo, "%s|%s|%s" % (tipo, uid, m), False),
        "/api/lancamento/ajuste": ("ajuste manual de " + nome_tipo, "%s|%s|%s" % (tipo, uid, m), True),
        "/api/lancamento/campanha": ("editando campanhas de ADS", "ads|%s|%s" % (uid, m), False),
        "/api/lancamento/depara/salvar": ("preenchendo o de-para", "ads|%s|%s" % (uid, m), True),
        "/api/lancamento/depara": ("subindo de-para (mass update)", "ads|%s|%s" % (uid, m), True),
        "/api/lancamento/reprocessar": ("reprocessando ADS e afiliados", "pacote|%s" % cid, True),
        "/api/vendas/subir": ("subindo relatório de vendas", "vendas|%s|%s" % (uid, m), True),
        "/api/vendas/limpar": ("limpando vendas" + ("" if dados.get("mk") else " (todos os marketplaces)"), "vendas|%s|%s" % (uid, m), True),
        "/api/lancamento/limpar": ("limpando " + nome_tipo + camp + (" (mês inteiro)" if dados.get("tudo") else ""), "%s|%s|%s" % (tipo, uid, m), True),
        "/api/vendas/ok": (("tirando OK de vendas" if dados.get("remover") else "dando OK em vendas"), "vendas|%s|%s" % (uid, m), False),
        "/api/classificar": ("classificando arquivo", "vendas|%s|%s" % (uid, m), False),
        "/api/tratar": ("tratando vendas", "trat|%s|%s" % (uid, m), True),
        "/api/custos/enviar": ("subindo tabela de custos", "custos|%s" % cid, True),
        "/api/custos/salvar": ("editando a tabela de custos", "custos|%s" % cid, True),
        "/api/custos/copiar": ("copiando tabela de custos", "custos|%s" % (dados.get("para") or cid), False),
        "/api/cliente/imposto": ("alterando imposto", "cliente|%s" % cid, False),
        "/api/cliente/antecipa": ("alterando antecipação Shopee", "cliente|%s" % cid, False),
        "/api/cliente": (("salvando cadastro" if dados.get("id") else "cadastrando cliente novo"), "cadastro|%s" % (dados.get("id") or dados.get("nome")), True),
        "/api/cliente/remover": ("removendo cliente", "cadastro|%s" % dados.get("id"), True),
        "/api/acompanhamento/salvar": ("registrando acompanhamento", "acomp|%s|%s" % (cid, dados.get("id") or "novo"), False),
        "/api/acompanhamento/remover": ("removendo registro de acompanhamento", "acomp|%s|%s" % (cid, dados.get("id")), False),
        "/api/acompanhamento/ficha": ("editando a ficha de acompanhamento", "acomp-ficha|%s" % cid, False),
        "/api/acompanhamento/visao": ("editando o direcionamento geral", "acomp-geral", False),
        "/api/acompanhamento/importar": ("importando acompanhamento", "acomp-importar", True),
        "/api/ia/contexto": ("escrevendo o contexto do cliente", "ia-contexto|%s" % cid, False),
        "/api/ia/gerar": ("gerando o resumo com a IA", "ia-gerar|%s" % cid, False),
        "/api/ia/editar": ("editando o resumo do cliente", "ia-editar|%s" % cid, False),
        "/api/tratamento/retratar-mes": ("pediu para retratar o mês todo", "retratar-mes", True),
        "/api/consolidado/gerar": ("gerando pacote final", "pacote|%s" % cid, True),
        "/api/forecast/definir": ("definindo forecast", "forecast|%s" % cid, False),
        "/api/organizar": ("rodando o organizador", "organizar", True),
    }
    F = {
        "/api/lancamento/subir": "subiu " + nome_tipo + camp,
        "/api/lancamento/ok": ("tirou o OK de " if dados.get("remover") else "deu OK em ") + nome_tipo,
        "/api/lancamento/ajuste": "fez ajuste manual de " + nome_tipo,
        "/api/lancamento/campanha": "editou campanhas de ADS",
        "/api/lancamento/depara/salvar": "preencheu o de-para no app",
        "/api/lancamento/depara": "subiu o de-para (mass update)",
        "/api/lancamento/reprocessar": "reprocessou ADS e afiliados",
        "/api/vendas/subir": "subiu relatório de vendas",
        "/api/vendas/limpar": "limpou o relatório de vendas" + ("" if dados.get("mk") else " (todos os marketplaces)"),
        "/api/lancamento/limpar": "limpou " + nome_tipo + camp + (" do mês inteiro" if dados.get("tudo") else (" de %d dia(s)" % len(dados.get("dias") or []))),
        "/api/vendas/ok": "tirou o OK de vendas" if dados.get("remover") else "deu OK em vendas",
        "/api/classificar": "classificou arquivo",
        "/api/tratar": "tratou vendas",
        "/api/custos/enviar": "subiu tabela de custos",
        "/api/custos/salvar": "salvou a tabela de custos" + ((" de " + str(dados.get("mes"))[5:] + "/" + str(dados.get("mes"))[:4]) if dados.get("mes") else ""),
        "/api/custos/copiar": "copiou tabela de custos",
        "/api/cliente/imposto": "alterou o imposto",
        "/api/cliente/antecipa": "alterou a antecipação Shopee",
        "/api/cliente": "salvou o cadastro" if dados.get("id") else "cadastrou cliente novo",
        "/api/cliente/remover": "removeu cliente",
        "/api/acompanhamento/salvar": "atualizou um registro de acompanhamento" if dados.get("id") else "registrou acompanhamento",
        "/api/acompanhamento/remover": "removeu um registro de acompanhamento",
        "/api/acompanhamento/ficha": "atualizou a ficha de acompanhamento",
        "/api/acompanhamento/visao": "atualizou o direcionamento geral",
        "/api/acompanhamento/importar": "importou registros de acompanhamento",
        "/api/ia/contexto": "atualizou o contexto do cliente (Resumo IA)",
        "/api/ia/gerar": "gerou o resumo do cliente com a IA",
        "/api/ia/editar": "editou o resumo do cliente",
        "/api/tratamento/retratar-mes": "pediu para retratar o mês todo",
        "/api/consolidado/gerar": "gerou o pacote final",
        "/api/forecast/definir": "definiu o forecast",
        "/api/organizar": "rodou o organizador",
    }
    if caminho not in T:
        return None
    acao, chave, longa = T[caminho]
    if caminho == "/api/cliente" and dados.get("nome"):
        cliente = dados["nome"]
    if caminho == "/api/tratar" and not dados.get("cliente_id"):
        cliente = "todos os clientes"
    return {"acao": acao, "feito": F.get(caminho, acao), "chave": chave, "longa": longa, "cliente": cliente, "conta": conta, "mk": mk, "cliente_id": cid}


# ---------------------------------------------------------------- andamento
def _trava_da(chave):
    with _TRAVA:
        return _TRAVAS_CHAVE.setdefault(chave, threading.Lock())


def conflito(desc, quem):
    """Outra pessoa já está fazendo algo longo no mesmo relatório (ou com o cadastro aberto)?"""
    if not desc or not desc.get("chave") or not desc["longa"]:
        return None
    _limpar()
    for e in list(AGORA.values()):
        if e.get("chave") == desc["chave"] and e.get("quem") != quem:
            alvo = " · ".join(x for x in (e.get("cliente"), e.get("conta"), e.get("mk")) if x)
            return "%s (%s) está %s%s agora. Espere terminar para não sobrescrever." % (
                e["usuario"], e["pc"], e["acao"], (" · " + alvo) if alvo else "")
    if desc["chave"].startswith("cadastro|"):
        cid = desc["chave"].split("|", 1)[1]
        for k, p in list(PRESENCA.items()):
            if k != quem and p.get("editando") == cid:
                return "%s (%s) está com o cadastro deste cliente aberto. Combine com ele antes de salvar." % (p["usuario"], p["pc"])
    return None


def executar(desc, quem, usuario, pc, func):
    """Roda a ação registrando em AGORA e na atividade recente; ações no mesmo relatório rodam uma de cada vez."""
    token = uuid.uuid4().hex
    e = dict(desc, quem=quem, usuario=usuario, pc=pc, inicio=_agora(), t=time.time())
    with _TRAVA:
        AGORA[token] = e
    tocar()
    trava = _trava_da(desc["chave"]) if desc.get("chave") else None
    ok, erro = False, None
    try:
        if trava:
            trava.acquire()
        r = func()
        ok = bool(isinstance(r, dict) and r.get("ok", True))
        if isinstance(r, dict) and not ok:
            erro = r.get("erro")
        return r
    except Exception as ex:
        erro = str(ex)
        raise
    finally:
        if trava:
            trava.release()
        with _TRAVA:
            AGORA.pop(token, None)
        registrar({"usuario": usuario, "pc": pc, "acao": desc.get("feito") or desc["acao"], "cliente": desc.get("cliente"), "conta": desc.get("conta"),
                   "mk": desc.get("mk"), "cliente_id": desc.get("cliente_id"), "ok": ok, "erro": (erro or "")[:200] or None,
                   "duracao_s": round(time.time() - e["t"], 1)})


def registrar(ev):
    ev = dict(ev, ts=_agora())
    try:
        os.makedirs(os.path.dirname(ARQ), exist_ok=True)
        with _TRAVA:
            with open(ARQ, "a", encoding="utf-8") as f:
                f.write(json.dumps(ev, ensure_ascii=False) + "\n")
    except OSError:
        pass
    tocar()


def recentes(n=40):
    try:
        with open(ARQ, "rb") as f:
            f.seek(0, 2)
            tam = f.tell()
            f.seek(max(0, tam - 64 * 1024))
            linhas = f.read().decode("utf-8", "ignore").splitlines()
    except OSError:
        return []
    out = []
    for l in reversed(linhas):
        try:
            out.append(json.loads(l))
        except ValueError:
            continue
        if len(out) >= n:
            break
    return out


# ---------------------------------------------------------------- presença (em qual tela cada um está)
def presenca(quem, usuario, pc, dados):
    novo = {"usuario": usuario, "pc": pc, "aba": str(dados.get("aba") or "")[:30], "cliente": str(dados.get("cliente") or "")[:80],
            "editando": str(dados.get("editando") or "")[:80], "ts": time.time()}
    with _TRAVA:
        antes = PRESENCA.get(quem)
        PRESENCA[quem] = novo
    if not antes or any(antes.get(k) != novo[k] for k in ("aba", "cliente", "editando", "usuario")):
        tocar()
    return {"ok": True}


def sair(quem):
    with _TRAVA:
        PRESENCA.pop(quem, None)
    tocar()
    return {"ok": True}


def _limpar():
    agora = time.time()
    with _TRAVA:
        for k in [k for k, p in PRESENCA.items() if agora - p["ts"] > TTL_PRESENCA]:
            PRESENCA.pop(k)
        for k in [k for k, e in AGORA.items() if agora - e["t"] > TTL_ACAO]:
            AGORA.pop(k)


def quadro(cfg=None):
    """O que a tela mostra no quadro 'Equipe agora'."""
    _limpar()
    nomes = {}
    if cfg:
        nomes = {c["id"]: c["nome"] for c in cfg.get("clientes", [])}
    agora = time.time()
    acoes = [{"quem": e["quem"], "usuario": e["usuario"], "pc": e["pc"], "acao": e["acao"], "chave": e.get("chave"), "cliente": e.get("cliente"),
              "conta": e.get("conta"), "mk": e.get("mk"), "ha_s": int(agora - e["t"])} for e in AGORA.values()]
    pres = [{"quem": k, "usuario": p["usuario"], "pc": p["pc"], "aba": ABAS.get(p["aba"], p["aba"]),
             "cliente": nomes.get(p["cliente"], p["cliente"]) if p["cliente"] else "",
             "editando": nomes.get(p["editando"], p["editando"]) if p["editando"] else ""} for k, p in PRESENCA.items()]
    return {"versao": versao(), "acoes": acoes, "presenca": pres, "downloads": []}
