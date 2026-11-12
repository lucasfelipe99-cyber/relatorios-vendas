# -*- coding: utf-8 -*-
"""Acompanhamento dos clientes: o que já foi realizado, o que está programado e os pontos de atenção de cada um.

Cada cliente tem o seu arquivo em dados/acompanhamento/<cliente_id>.json com:
- "ficha": status, situação atual, contexto, prioridade estratégica, nível de prioridade e objetivo;
- "itens": registros (ações, próximos passos, reuniões, problemas...). A situação de cada registro diz se ele
  está "a fazer" (programado), "em andamento" ou "realizado". Quem cria ou altera fica registrado.
O direcionamento geral (de todos os clientes) fica em dados/acompanhamento/_visao_geral.json.
A página exporta um PDF no formato do relatório de acompanhamento.
"""
import datetime as dt
import io
import json
import os
import re
import secrets
import threading

import comum

PASTA = os.path.join(comum.PASTA_SISTEMA, "dados", "acompanhamento")
ARQ_GERAL = os.path.join(PASTA, "_visao_geral.json")
TIPOS = ["Ação", "Ponto de atenção", "Reunião", "Contato", "Problema", "Pendência", "Combinado", "Observação"]
SITUACOES = {"aberto": "A fazer", "andamento": "Em andamento", "resolvido": "Realizado"}
NIVEIS = {"alta": "Alta", "media": "Média", "baixa": "Baixa"}
CAMPOS_FICHA = ("status", "resumo", "contexto", "prioridade", "nivel", "objetivo")
_TRAVA = threading.RLock()


def _arq(cliente_id):
    if not re.fullmatch(r"[a-z0-9][a-z0-9._-]{0,79}", cliente_id or ""):
        raise ValueError("Cliente inválido.")
    return os.path.join(PASTA, cliente_id + ".json")


def _ler_tudo(cliente_id):
    try:
        with open(_arq(cliente_id), "r", encoding="utf-8") as f:
            d = json.load(f)
    except (OSError, ValueError):
        d = {}
    d.setdefault("itens", [])
    d.setdefault("ficha", {})
    return d


def _ler(cliente_id):
    return _ler_tudo(cliente_id)["itens"]


def _gravar_tudo(cliente_id, d):
    os.makedirs(PASTA, exist_ok=True)
    arq = _arq(cliente_id)
    with open(arq + ".tmp", "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    os.replace(arq + ".tmp", arq)


def _gravar(cliente_id, itens):
    d = _ler_tudo(cliente_id)
    d["itens"] = itens
    _gravar_tudo(cliente_id, d)


def _data(v, obrigatoria=False):
    v = str(v or "").strip()
    if not v:
        if obrigatoria:
            raise ValueError("Informe a data.")
        return ""
    try:
        return dt.date.fromisoformat(v[:10]).isoformat()
    except ValueError:
        raise ValueError("Data inválida: %s" % v)


def _clientes(cfg):
    return {c["id"]: c["nome"] for c in cfg.get("clientes", [])}


def _ordem_cliente(cfg):
    return {c["id"]: i for i, c in enumerate(cfg.get("clientes", []))}


# ---------------------------------------------------------------- leitura
def listar(cfg, cliente_id="", de="", ate="", tipo="", situacao="", busca=""):
    """Registros (mais novos primeiro). Sem cliente = todos os clientes."""
    nomes = _clientes(cfg)
    ids = [cliente_id] if cliente_id else list(nomes)
    de, ate = _data(de), _data(ate)
    busca = comum.sem_acento(busca or "").lower().strip()
    out = []
    for cid in ids:
        if cid not in nomes:
            continue
        for it in _ler(cid):
            if de and it["data"] < de or ate and it["data"] > ate:
                continue
            if tipo and it.get("tipo") != tipo:
                continue
            if situacao == "pendentes" and it.get("situacao") == "resolvido" or \
                    situacao in SITUACOES and it.get("situacao") != situacao:
                continue
            if busca and busca not in comum.sem_acento(" ".join(str(it.get(k) or "") for k in
                                                              ("assunto", "texto", "responsavel", "criado_por", "canal"))).lower():
                continue
            out.append(dict(it, cliente_id=cid, cliente=nomes[cid]))
    out.sort(key=lambda x: (x["data"], x.get("criado_em", "")), reverse=True)
    return out


def fichas(cfg, cliente_id=""):
    nomes = _clientes(cfg)
    ids = [cliente_id] if cliente_id else list(nomes)
    return {cid: _ler_tudo(cid)["ficha"] for cid in ids if cid in nomes}


def contagem(cfg):
    """Por cliente: total, a fazer/em andamento, pontos de atenção em aberto, prazos vencidos e o último registro."""
    hoje = dt.date.today().isoformat()
    res = {}
    for cid in _clientes(cfg):
        d = _ler_tudo(cid)
        its = d["itens"]
        pend = [i for i in its if i.get("situacao") != "resolvido"]
        res[cid] = {"total": len(its), "abertos": len(pend),
                    "atencao": sum(1 for i in pend if i.get("tipo") == "Ponto de atenção"),
                    "vencidos": sum(1 for i in pend if i.get("prazo") and i["prazo"] < hoje),
                    "realizados": len(its) - len(pend),
                    "ultimo": max((i["data"] for i in its), default=""),
                    "status": d["ficha"].get("status", ""), "nivel": d["ficha"].get("nivel", "")}
    return res


def visao_geral():
    try:
        with open(ARQ_GERAL, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


# ---------------------------------------------------------------- gravação
def _campos_item(dados):
    tipo = dados.get("tipo") if dados.get("tipo") in TIPOS else "Observação"
    situacao = dados.get("situacao") if dados.get("situacao") in SITUACOES else "aberto"
    assunto = str(dados.get("assunto") or "").strip()[:240]
    if not assunto:
        raise ValueError("Escreva o assunto.")
    return {"data": _data(dados.get("data"), True), "tipo": tipo, "assunto": assunto,
            "texto": str(dados.get("texto") or "").strip()[:8000], "situacao": situacao,
            "prazo": _data(dados.get("prazo")), "responsavel": str(dados.get("responsavel") or "").strip()[:60],
            "canal": str(dados.get("canal") or "").strip()[:60]}


def _novo_id():
    return dt.date.today().strftime("%Y%m%d") + "-" + secrets.token_hex(4)


def salvar(cfg, dados, por, login):
    cid = dados.get("cliente_id") or ""
    if cid not in _clientes(cfg):
        raise ValueError("Escolha o cliente.")
    campos = _campos_item(dados)
    agora = dt.datetime.now().isoformat(timespec="seconds")
    with _TRAVA:
        itens = _ler(cid)
        it = next((x for x in itens if x["id"] == dados.get("id")), None) if dados.get("id") else None
        if dados.get("id") and it is None:
            raise ValueError("Registro não encontrado (pode ter sido removido).")
        if it is None:
            it = {"id": _novo_id(), "criado_por": por, "criado_login": login, "criado_em": agora}
            itens.append(it)
        else:
            it.update(alterado_por=por, alterado_em=agora)
            if it.get("situacao") != campos["situacao"] and campos["situacao"] == "resolvido":
                it["resolvido_em"] = agora
        it.update(campos)
        _gravar(cid, itens)
    return dict(it, cliente_id=cid, cliente=_clientes(cfg)[cid])


def salvar_ficha(cfg, cliente_id, ficha, por):
    if cliente_id not in _clientes(cfg):
        raise ValueError("Escolha o cliente.")
    f = {k: str(ficha.get(k) or "").strip()[:4000 if k == "contexto" else 400] for k in CAMPOS_FICHA}
    if f["nivel"] not in NIVEIS:
        f["nivel"] = ""
    f.update(atualizado_por=por, atualizado_em=dt.datetime.now().isoformat(timespec="seconds"))
    with _TRAVA:
        d = _ler_tudo(cliente_id)
        d["ficha"] = f
        _gravar_tudo(cliente_id, d)
    return f


def salvar_visao_geral(texto, por):
    d = {"texto": str(texto or "").strip()[:12000], "atualizado_por": por,
         "atualizado_em": dt.datetime.now().isoformat(timespec="seconds")}
    with _TRAVA:
        os.makedirs(PASTA, exist_ok=True)
        with open(ARQ_GERAL + ".tmp", "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        os.replace(ARQ_GERAL + ".tmp", ARQ_GERAL)
    return d


def remover(cfg, cliente_id, item_id, login, admin):
    with _TRAVA:
        itens = _ler(cliente_id)
        it = next((x for x in itens if x["id"] == item_id), None)
        if it is None:
            return False
        if not admin and it.get("criado_login") != login:
            raise ValueError("Só quem criou o registro (ou um admin) pode removê-lo.")
        _gravar(cliente_id, [x for x in itens if x["id"] != item_id])
    return True


def importar(cfg, conteudo, por, login):
    """Importa um arquivo .json de acompanhamento (fichas, registros e visão geral).

    {"visao_geral": "...", "clientes": [{"cliente": "nome ou id", "ficha": {...}, "itens": [{"id": ..., ...}]}]}
    Registros com o mesmo "id" são atualizados (importar de novo não duplica)."""
    try:
        d = json.loads(conteudo) if isinstance(conteudo, str) else conteudo
    except ValueError:
        raise ValueError("O arquivo não é um .json válido.")
    if not isinstance(d, dict) or not isinstance(d.get("clientes", []), list):
        raise ValueError("Formato do arquivo não reconhecido.")
    nomes = _clientes(cfg)
    por_nome = {comum.sem_acento(n).lower().strip(): i for i, n in nomes.items()}
    agora = dt.datetime.now().isoformat(timespec="seconds")
    res = {"clientes": 0, "novos": 0, "atualizados": 0, "fichas": 0, "nao_encontrados": []}
    blocos = []
    for b in d.get("clientes", []):
        ref = str(b.get("cliente") or "").strip()
        cid = ref if ref in nomes else por_nome.get(comum.sem_acento(ref).lower())
        if not cid:
            res["nao_encontrados"].append(ref)
            continue
        itens = []
        for x in b.get("itens", []):
            campos = _campos_item(x)            # valida tudo antes de gravar qualquer coisa
            iid = str(x.get("id") or "").strip()[:60] or _novo_id()
            if not re.fullmatch(r"[A-Za-z0-9._-]+", iid):
                raise ValueError("Id de registro inválido: %s" % iid)
            itens.append((iid, campos))
        blocos.append((cid, b.get("ficha"), itens))
    with _TRAVA:
        for cid, ficha, itens in blocos:
            dd = _ler_tudo(cid)
            atuais = {x["id"]: x for x in dd["itens"]}
            for iid, campos in itens:
                if iid in atuais:
                    if any(atuais[iid].get(k) != v for k, v in campos.items()):     # só marca alterado se mudou algo
                        atuais[iid].update(campos, alterado_por=por, alterado_em=agora)
                        res["atualizados"] += 1
                else:
                    novo = dict(campos, id=iid, criado_por=por, criado_login=login, criado_em=agora, importado=True)
                    dd["itens"].append(novo)
                    atuais[iid] = novo
                    res["novos"] += 1
            _gravar_tudo(cid, dd)
            if isinstance(ficha, dict):
                salvar_ficha(cfg, cid, ficha, por)
                res["fichas"] += 1
            res["clientes"] += 1
        if d.get("visao_geral"):
            salvar_visao_geral(d["visao_geral"], por)
            res["visao_geral"] = True
    return res


# ---------------------------------------------------------------- PDF
def _br(d):
    return dt.date.fromisoformat(d).strftime("%d/%m/%Y") if d else ""


def _moeda(v):
    s = "{:,.2f}".format(v or 0).replace(",", "X").replace(".", ",").replace("X", ".")
    return "R$ " + s


def _numeros(cfg, cid):
    """Números do mês atual do cliente (do pacote final), para o PDF."""
    try:
        import painel_dados
        import consolidado
        c = next(x for x in cfg["clientes"] if x["id"] == cid)
        ano, mes = consolidado.mes_atual(dt.date.today())     # dia 1 = mês anterior (o mês que está sendo fechado)
        r = painel_dados.resumo(cfg, c, ano, mes)
        if not r.get("ok") or r.get("sem_dados") or not r.get("total"):
            return None
        t = r["total"]
        return {"mes": "%02d/%04d" % (mes, ano), "fat_bruto": t.get("fat_bruto"), "fat_liquido": t.get("fat_liquido"),
                "margem": t.get("margem"), "margem_pct": t.get("margem_pct"), "publicidade": t.get("publicidade")}
    except Exception as e:
        comum.log("acompanhamento: números do PDF não calculados (%s): %s" % (cid, e))
        return None


def pdf(cfg, filtro, por):
    """PDF no formato do relatório de acompanhamento. Devolve (bytes, nome do arquivo)."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import CondPageBreak, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    itens = listar(cfg, **{k: filtro.get(k, "") for k in ("cliente_id", "de", "ate", "tipo", "situacao", "busca")})
    nomes = _clientes(cfg)
    cid = filtro.get("cliente_id") or ""
    FI = fichas(cfg, cid)
    titulo = ("Acompanhamento — " + nomes[cid]) if cid in nomes else "Relatório de Acompanhamento – Clientes"
    hoje = dt.date.today().strftime("%d/%m/%Y")

    ss = getSampleStyleSheet()
    AZUL, CINZA, LINHA, ESC = (colors.HexColor("#1d4ed8"), colors.HexColor("#64748b"), colors.HexColor("#e2e8f0"),
                               colors.HexColor("#0f172a"))
    P = lambda nome, **kw: ParagraphStyle(nome, parent=ss["Normal"], **kw)
    st = {"tit": P("tit", fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=ESC, spaceAfter=2),
          "sub": P("sub", fontSize=8.5, leading=11, textColor=CINZA),
          "cli": P("cli", fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=AZUL, spaceBefore=4, spaceAfter=2),
          "status": P("status", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=ESC),
          "h3": P("h3", fontName="Helvetica-Bold", fontSize=10, leading=13, textColor=ESC, spaceBefore=8, spaceAfter=3),
          "txt": P("txt", fontSize=9, leading=12.8),
          "item": P("item", fontSize=9, leading=12.2),
          "meta": P("meta", fontSize=7.6, leading=9.6, textColor=CINZA),
          "dir": P("dir", fontSize=7.8, leading=10, textColor=CINZA, alignment=TA_RIGHT),
          "num": P("num", fontSize=7.8, leading=10, textColor=CINZA),
          "numv": P("numv", fontName="Helvetica-Bold", fontSize=10.5, leading=13),
          "th": P("th", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=colors.white),
          "td": P("td", fontSize=8.5, leading=11)}
    COR_SIT = {"aberto": colors.HexColor("#2563eb"), "andamento": colors.HexColor("#d97706"),
               "resolvido": colors.HexColor("#16a34a")}
    COR_ATENCAO = colors.HexColor("#dc2626")
    e = lambda t: escape(str(t or ""))
    txt = lambda t: e(t).replace("\n", "<br/>")

    filtros = []
    if filtro.get("de") or filtro.get("ate"):
        filtros.append("período %s a %s" % (_br(_data(filtro.get("de"))) or "início", _br(_data(filtro.get("ate"))) or "hoje"))
    if filtro.get("tipo"):
        filtros.append("tipo: %s" % filtro["tipo"])
    if filtro.get("situacao") in SITUACOES:
        filtros.append("situação: %s" % SITUACOES[filtro["situacao"]])
    elif filtro.get("situacao") == "pendentes":
        filtros.append("só o que falta fazer")
    if filtro.get("busca"):
        filtros.append("busca: “%s”" % filtro["busca"])

    corpo = [Paragraph(e(titulo), st["tit"]),
             Paragraph(e("Data de atualização: %s · gerado por %s%s" % (hoje, por, (" · " + "; ".join(filtros)) if filtros else "")),
                       st["sub"]), Spacer(1, 8)]

    ordem = _ordem_cliente(cfg)
    ids = [cid] if cid else sorted({i["cliente_id"] for i in itens} | {k for k, v in FI.items() if any(v.get(c) for c in CAMPOS_FICHA)},
                                   key=lambda x: ordem.get(x, 999))

    # resumo executivo (todos os clientes)
    if not cid and ids:
        corpo.append(Paragraph("Resumo executivo – visão consolidada", st["h3"]))
        linhas = [[Paragraph(x, st["th"]) for x in ("Cliente", "Situação atual", "Prioridade estratégica", "Nível")]]
        for c in ids:
            f = FI.get(c, {})
            linhas.append([Paragraph("<b>%s</b>" % e(nomes[c]), st["td"]), Paragraph(e(f.get("resumo") or f.get("status")), st["td"]),
                           Paragraph(e(f.get("prioridade")), st["td"]), Paragraph(e(NIVEIS.get(f.get("nivel"), "")), st["td"])])
        t = Table(linhas, colWidths=[34 * mm, 66 * mm, 62 * mm, 18 * mm], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), AZUL), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
                               ("LINEBELOW", (0, 0), (-1, -1), 0.4, LINHA), ("TOPPADDING", (0, 0), (-1, -1), 4),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        corpo += [t, Spacer(1, 6)]

    def item(r, cor):
        partes = [r.get("tipo", "")]
        if r.get("canal"):
            partes.append(r["canal"])
        if r.get("responsavel"):
            partes.append("responsável: " + r["responsavel"])
        if r.get("prazo"):
            venc = r.get("situacao") != "resolvido" and r["prazo"] < dt.date.today().isoformat()
            partes.append(("<font color='#dc2626'><b>prazo %s (vencido)</b></font>" if venc else "prazo %s") % _br(r["prazo"]))
        partes.append("registrado por %s" % e(r.get("criado_por", "?")))
        esq = [Paragraph("<b>%s</b>" % e(r["assunto"]), st["item"]),
               Paragraph(" · ".join(p if p.startswith("<font") else e(p) for p in partes[:-1]) + " · " + partes[-1], st["meta"])]
        if r.get("texto"):
            esq.append(Paragraph(txt(r["texto"]), st["txt"]))
        dirr = [Paragraph(_br(r["data"]), st["dir"])]
        if r.get("situacao") == "andamento":
            dirr.append(Paragraph("<font color='#d97706'><b>em andamento</b></font>", st["dir"]))
        t = Table([[esq, dirr]], colWidths=[152 * mm, 26 * mm])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBEFORE", (0, 0), (0, -1), 2, cor),
                               ("LEFTPADDING", (0, 0), (0, -1), 7), ("TOPPADDING", (0, 0), (-1, -1), 3),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        return t

    if not ids:
        corpo.append(Paragraph("Nenhum registro com esses filtros.", st["txt"]))
    for n_c, c in enumerate(ids):
        regs = [i for i in itens if i["cliente_id"] == c]
        f = FI.get(c, {})
        corpo.append(CondPageBreak(60 * mm))
        if n_c or not cid:
            corpo.append(Spacer(1, 6))
        cab = [] if cid else [Paragraph("%d. %s" % (n_c + 1, e(nomes[c])), st["cli"])]
        if f.get("status"):
            nivel = NIVEIS.get(f.get("nivel"))
            cab.append(Paragraph("Status: %s%s" % (e(f["status"]), (" · prioridade %s" % e(nivel).lower()) if nivel else ""), st["status"]))
        corpo.append(KeepTogether(cab))
        if f.get("contexto"):
            corpo += [Spacer(1, 3), Paragraph(txt(f["contexto"]), st["txt"])]
        num = _numeros(cfg, c)
        if num:
            cel = lambda rot, val: [Paragraph(e(rot), st["num"]), Paragraph(e(val), st["numv"])]
            dados = [cel("Fat. bruto " + num["mes"], _moeda(num["fat_bruto"])), cel("Fat. líquido", _moeda(num["fat_liquido"])),
                     cel("Publicidade", _moeda(abs(num["publicidade"] or 0))),
                     cel("Margem", _moeda(num["margem"]) + ((" (%.1f%%)" % (num["margem_pct"] * 100)).replace(".", ",")
                                                            if num["margem_pct"] is not None else ""))]
            t = Table([[x[0] for x in dados], [x[1] for x in dados]], colWidths=[44.5 * mm] * 4)
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f1f5f9")),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 7), ("TOPPADDING", (0, 0), (-1, 0), 5),
                                   ("BOTTOMPADDING", (0, -1), (-1, -1), 5), ("LINEAFTER", (0, 0), (-2, -1), 0.5, LINHA)]))
            corpo += [Spacer(1, 6), t]
        atencao = [r for r in regs if r.get("tipo") == "Ponto de atenção" and r.get("situacao") != "resolvido"]
        fazer = [r for r in regs if r.get("situacao") != "resolvido" and r not in atencao]
        feito = [r for r in regs if r.get("situacao") == "resolvido"]
        fazer.sort(key=lambda r: (r.get("prazo") or "9999", r["data"]))
        for nome_bloco, lst, cor in (("Pontos de atenção", atencao, COR_ATENCAO), ("Próximos passos (programado)", fazer, COR_SIT["aberto"]),
                                     ("Realizado", feito, COR_SIT["resolvido"])):
            if lst:
                corpo.append(KeepTogether([Paragraph("%s <font color='#64748b' size='8'>(%d)</font>" % (nome_bloco, len(lst)), st["h3"]),
                                           item(lst[0], cor)]))
                corpo += [item(r, cor) for r in lst[1:]]
        if f.get("objetivo"):
            corpo += [Spacer(1, 4), Paragraph("<b>Objetivo:</b> " + txt(f["objetivo"]), st["txt"])]
        if not regs and not any(f.get(k) for k in CAMPOS_FICHA):
            corpo.append(Paragraph("Nenhum registro com esses filtros.", st["meta"]))
        corpo.append(Spacer(1, 4))
        tl = Table([[""]], colWidths=[180 * mm], rowHeights=[1])
        tl.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.6, LINHA)]))
        corpo.append(tl)

    vg = visao_geral()
    if not cid and vg.get("texto"):
        corpo += [CondPageBreak(40 * mm), Paragraph("Direcionamento geral", st["cli"]), Paragraph(txt(vg["texto"]), st["txt"])]

    def rodape(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(CINZA)
        canvas.drawString(15 * mm, 9 * mm, "Relatórios de Vendas · " + titulo)
        canvas.drawRightString(A4[0] - 15 * mm, 9 * mm, "página %d" % doc.page)
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=14 * mm,
                            bottomMargin=16 * mm, title=titulo, author=por)
    doc.build(corpo, onFirstPage=rodape, onLaterPages=rodape)
    nome = "acompanhamento_%s_%s.pdf" % (re.sub(r"[^A-Za-z0-9]+", "-", comum.sem_acento(nomes.get(cid) or "todos")).strip("-"),
                                         dt.date.today().strftime("%Y-%m-%d"))
    return buf.getvalue(), nome
