# -*- coding: utf-8 -*-
"""Resumo do cliente escrito pela IA (Claude, da Anthropic).

Junta, para o cliente e o mês escolhidos:
- o contexto que a equipe escreve na aba "Resumo IA" (texto livre, por cliente);
- a ficha de acompanhamento (status, situação, prioridade, objetivo);
- os números de venda do mês (pacote final): totais, mês anterior até o mesmo dia, forecast, canais, produtos;
- as atividades em aberto e as realizadas no mês;
- o direcionamento geral;
e pede para a IA um resumo executivo em português. O resumo fica guardado (com histórico) e pode ser editado.

Tudo fica no arquivo do cliente em dados/acompanhamento/<cliente_id>.json, na chave "ia".
A chave da API vem da variável de ambiente ANTHROPIC_API_KEY (Render > Environment); o modelo, de IA_MODELO.
"""
import datetime as dt
import json
import os
import secrets
import urllib.error
import urllib.request

import acompanhamento as A
import comum

MODELO_PADRAO = "claude-sonnet-5-5"
MAX_HIST = 30


def configurada():
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def modelo():
    return (os.environ.get("IA_MODELO") or MODELO_PADRAO).strip()


def _ia(d):
    ia = d.setdefault("ia", {})
    ia.setdefault("contexto", {})
    ia.setdefault("resumos", [])
    return ia


def estado(cfg, cliente_id, mes):
    if cliente_id not in A._clientes(cfg):
        raise ValueError("Escolha o cliente.")
    d = A._ler_tudo(cliente_id)
    ia = _ia(d)
    resumos = sorted(ia["resumos"], key=lambda r: r.get("gerado_em", ""), reverse=True)
    return {"contexto": ia["contexto"], "resumos": [r for r in resumos if r.get("mes") == mes],
            "outros_meses": sorted({r.get("mes") for r in resumos if r.get("mes") and r.get("mes") != mes}, reverse=True),
            "configurada": configurada(), "modelo": modelo()}


def salvar_contexto(cfg, cliente_id, texto, por):
    if cliente_id not in A._clientes(cfg):
        raise ValueError("Escolha o cliente.")
    with A._TRAVA:
        d = A._ler_tudo(cliente_id)
        ia = _ia(d)
        ia["contexto"] = {"texto": str(texto or "").strip()[:20000], "por": por,
                          "em": dt.datetime.now().isoformat(timespec="seconds")}
        A._gravar_tudo(cliente_id, d)
    return ia["contexto"]


def salvar_edicao(cfg, cliente_id, resumo_id, texto, por):
    if cliente_id not in A._clientes(cfg):
        raise ValueError("Escolha o cliente.")
    with A._TRAVA:
        d = A._ler_tudo(cliente_id)
        r = next((x for x in _ia(d)["resumos"] if x.get("id") == resumo_id), None)
        if not r:
            raise ValueError("Resumo não encontrado.")
        r.update(texto=str(texto or "").strip()[:30000], editado_por=por, editado_em=dt.datetime.now().isoformat(timespec="seconds"))
        A._gravar_tudo(cliente_id, d)
    return r


# ---------------------------------------------------------------- dados que vão para a IA
def _r(v):
    return None if v is None else round(float(v), 2)


def _p(v):
    return None if v is None else "%.1f%%" % (float(v) * 100)


def _totais(t):
    t = t or {}
    return {"faturamento_bruto": _r(t.get("fat_bruto")), "faturamento_liquido": _r(t.get("fat_liquido")),
            "gm_margem": _r(t.get("margem")), "gm_pct": _p(t.get("margem_pct")), "publicidade": _r(abs(t.get("publicidade") or 0)),
            "tacos": _p(t.get("tacos")), "roas": t.get("roas"), "unidades": t.get("unidades"), "cancelamento_pct": _p(t.get("cancel_pct")),
            "afiliados": _r(abs(t.get("afiliados") or 0)), "ticket_medio_unidade": _r(t.get("ticket_un"))}


def montar_dados(cfg, cliente_id, mes):
    import painel_dados
    c = next((x for x in cfg["clientes"] if x["id"] == cliente_id), None)
    if not c:
        raise ValueError("Cliente não encontrado.")
    ano, m = int(mes[:4]), int(mes[5:7])
    d = A._ler_tudo(cliente_id)
    ia = _ia(d)
    ini = "%04d-%02d-01" % (ano, m)
    fim = "%04d-%02d-01" % ((ano + 1, 1) if m == 12 else (ano, m + 1))
    hoje = dt.date.today().isoformat()
    its = d["itens"]
    quando = lambda i: (i.get("resolvido_em") or "")[:10] or i.get("data") or ""
    enx = lambda i: {k: i.get(k) for k in ("data", "tipo", "assunto", "texto", "situacao", "prazo", "responsavel", "canal") if i.get(k)}
    abertos = [enx(i) for i in its if i.get("situacao") != "resolvido"]
    for i in abertos:
        if i.get("prazo") and i["prazo"] < hoje:
            i["vencido"] = True
    feitos = [enx(i) for i in its if i.get("situacao") == "resolvido" and ini <= quando(i) < fim]
    vendas = {"sem_dados": True}
    try:
        r = painel_dados.resumo(cfg, c, ano, m)
        if r.get("ok") and not r.get("sem_dados") and r.get("total"):
            vendas = {"mes": r.get("mes"), "vendas_ate": r.get("ate"), "totais": _totais(r["total"])}
            if r.get("anterior"):
                vendas["mes_anterior_ate_o_mesmo_dia"] = dict(_totais(r["anterior"]["total"]), mes=r["anterior"]["mes"],
                                                              ate_dia=r["anterior"].get("ate_dia"))
            fc = (r.get("forecast") or {}).get("ate")
            if fc:
                vendas["forecast_ate_o_mesmo_dia"] = {"faturamento_bruto": _r(fc.get("fat_bruto")), "gm_margem": _r(fc.get("margem"))}
            vendas["por_canal_e_conta"] = [dict(_totais(g), marketplace=g.get("marketplace"), conta=g.get("conta"))
                                           for g in (r.get("por_canal") or [])[:15]]
            vendas["produtos_que_mais_faturam"] = (r.get("top_produtos") or [])[:8]
            vendas["produtos_com_maior_gm_em_reais"] = (r.get("melhores_margens") or [])[:8]
            vendas["produtos_com_mais_unidades_vendidas"] = (r.get("mais_unidades") or [])[:8]
            vendas["produtos_com_margem_negativa"] = (r.get("piores_margens") or [])[:8]
            vendas["skus_vendidos"] = r.get("skus_vendidos")
            vendas["alertas_do_painel"] = [{"titulo": a.get("titulo"), "detalhe": a.get("detalhe")} for a in (r.get("alertas") or [])[:10]]
    except Exception as e:
        vendas = {"sem_dados": True, "erro": str(e)[:200]}
    return {"cliente": c["nome"], "mes": "%02d/%04d" % (m, ano), "hoje": dt.date.today().strftime("%d/%m/%Y"),
            "contexto_da_equipe": (ia["contexto"] or {}).get("texto") or "",
            "ficha_de_acompanhamento": {k: v for k, v in d["ficha"].items() if k in A.CAMPOS_FICHA and v},
            "vendas_do_mes": vendas,
            "atividades_em_aberto": abertos[:40], "atividades_realizadas_no_mes": feitos[:40],
            "direcionamento_geral_de_todos_os_clientes": (A.visao_geral() or {}).get("texto") or ""}


SISTEMA = """Você é um analista sênior de e-commerce e marketplaces (Mercado Livre, Shopee, Amazon, Magalu, TikTok Shop e outros) \
de uma consultoria que acompanha a rentabilidade de lojas online. Escreva em português do Brasil, em tom profissional e direto.

Regras:
- Use só os dados recebidos. Não invente números, nomes, causas ou fatos. Se faltar dado, diga que falta.
- Valores em reais no formato R$ 1.234; percentuais com uma casa (12,3%). GM = margem de contribuição; TACOS = publicidade / faturamento líquido.
- Compare com o mês anterior até o mesmo dia e com o forecast quando houver, dizendo se melhorou ou piorou.
- O contexto escrito pela equipe é a fonte mais importante para explicar o momento do cliente: use-o.
- Seja objetivo: frases curtas, sem floreio, sem repetir a mesma informação."""

PEDIDO = """Faça um RESUMO EXECUTIVO da empresa {cliente} referente a {mes} (hoje é {hoje}), para o dono ou gestor ler em 2 minutos. \
Use estas seções em Markdown, nesta ordem, com listas curtas e números:

## Principais pontos
3 a 5 frases curtas com o essencial: como está o mês e o que mais importa agora.

## Como está indo
Faturamento, GM e GM%, unidades e cancelamento, comparando com o mês anterior até o mesmo dia e com o forecast (se houver). \
Diga claramente se está melhor ou pior e quanto.

## Maior volume
Canais/contas e produtos que mais vendem (faturamento e unidades), com o GM% de cada um.

## Maiores margens
Produtos (e canais) que mais deixam margem (GM em R$ e GM%): onde vale a pena colocar mais força.

## Piores margens
Produtos e canais com GM baixo ou negativo, quanto estão tirando do resultado e a causa provável quando os dados mostrarem \
(publicidade alta, custo, cancelamento).

## Investimento
Publicidade do mês: valor, TACOS e ROAS, por canal quando houver; onde o investimento está trazendo resultado e onde não está.

## Onde está ruim
Lista objetiva dos problemas: margens negativas, TACOS alto, queda contra o mês anterior, cancelamento, pendências vencidas, \
alertas do painel e o que o contexto da equipe apontar.

## O que fazer agora
Até 5 ações práticas, em ordem de impacto, ligadas aos pontos acima; inclua o que já está programado nas atividades em aberto \
(com prazo e responsável) e o que foi realizado no mês, quando ajudar a explicar.

Dados (JSON):
{dados}"""


def gerar(cfg, cliente_id, mes, por):
    if not configurada():
        raise ValueError("A IA não está configurada: falta a chave ANTHROPIC_API_KEY nas variáveis do Render (Environment).")
    dados = montar_dados(cfg, cliente_id, mes)
    corpo = {"model": modelo(), "max_tokens": 16000, "system": SISTEMA,
             "messages": [{"role": "user", "content": PEDIDO.format(
                 cliente=dados["cliente"], mes=dados["mes"], hoje=dados["hoje"],
                 dados=json.dumps(dados, ensure_ascii=False, indent=1, default=str))}]}
    url = (os.environ.get("ANTHROPIC_BASE_URL") or "https://api.anthropic.com").rstrip("/") + "/v1/messages"
    req = urllib.request.Request(url, data=json.dumps(corpo).encode("utf-8"), method="POST",
                                 headers={"x-api-key": os.environ["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01",
                                          "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            resp = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
        except Exception:
            msg = ""
        comum.log("resumo IA: erro %s da API: %s" % (e.code, msg[:300]))
        if e.code in (401, 403):
            raise ValueError("A chave da IA (ANTHROPIC_API_KEY) foi recusada. Confira a chave no Render.")
        if e.code == 429:
            raise ValueError("A IA está com limite de uso no momento. Tente de novo em alguns minutos.")
        raise ValueError("A IA respondeu com erro %s: %s" % (e.code, msg[:200] or "tente de novo"))
    except (urllib.error.URLError, TimeoutError) as e:
        raise ValueError("Não consegui falar com a IA (%s). Tente de novo." % e)
    texto = "".join(b.get("text", "") for b in resp.get("content", []) if b.get("type") == "text").strip()
    parada = resp.get("stop_reason") or ""
    comum.log("resumo IA: modelo %s, parada %s, blocos %s, %d caracteres, uso %s" % (
        resp.get("model"), parada, [b.get("type") for b in resp.get("content", [])], len(texto), resp.get("usage")))
    if parada == "max_tokens":
        # o modelo gastou o limite (inclusive pensando) antes de terminar o texto: não guarda resumo cortado
        raise ValueError("A IA não terminou o resumo (limite de tamanho atingido). Tente de novo.")
    if parada == "refusal":
        raise ValueError("A IA recusou gerar este resumo. Tente de novo ou ajuste o contexto.")
    if not texto:
        raise ValueError("A IA não devolveu texto. Tente de novo.")
    uso = resp.get("usage") or {}
    novo = {"id": dt.date.today().strftime("%Y%m%d") + "-" + secrets.token_hex(4), "mes": mes, "texto": texto,
            "gerado_em": dt.datetime.now().isoformat(timespec="seconds"), "por": por, "modelo": resp.get("model") or modelo(),
            "tokens": (uso.get("input_tokens") or 0) + (uso.get("output_tokens") or 0)}
    with A._TRAVA:
        d = A._ler_tudo(cliente_id)
        ia = _ia(d)
        ia["resumos"].append(novo)
        ia["resumos"] = sorted(ia["resumos"], key=lambda r: r.get("gerado_em", ""))[-MAX_HIST:]
        A._gravar_tudo(cliente_id, d)
    comum.log("resumo IA de %s (%s) gerado por %s: %d tokens" % (dados["cliente"], mes, por, novo["tokens"]))
    return novo
