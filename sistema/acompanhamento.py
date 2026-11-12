# -*- coding: utf-8 -*-
"""Acompanhamento dos clientes: registro do que está acontecendo com cada um (reuniões, problemas, combinados...).

Cada cliente tem o seu arquivo em dados/acompanhamento/<cliente_id>.json. Quem cria ou altera fica registrado.
A página exporta um PDF com os registros filtrados e os números do mês do cliente.
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
TIPOS = ["Reunião", "Contato", "Problema", "Pendência", "Combinado", "Observação"]
SITUACOES = {"aberto": "Aberto", "andamento": "Em andamento", "resolvido": "Resolvido"}
_TRAVA = threading.RLock()


def _arq(cliente_id):
    if not re.fullmatch(r"[a-z0-9._-]{1,80}", cliente_id or ""):
        raise ValueError("Cliente inválido.")
    return os.path.join(PASTA, cliente_id + ".json")


def _ler(cliente_id):
    try:
        with open(_arq(cliente_id), "r", encoding="utf-8") as f:
            return json.load(f).get("itens", [])
    except (OSError, ValueError):
        return []


def _gravar(cliente_id, itens):
    os.makedirs(PASTA, exist_ok=True)
    arq = _arq(cliente_id)
    with open(arq + ".tmp", "w", encoding="utf-8") as f:
        json.dump({"itens": itens}, f, ensure_ascii=False, indent=1)
    os.replace(arq + ".tmp", arq)


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
                                                              ("assunto", "texto", "responsavel", "criado_por"))).lower():
                continue
            out.append(dict(it, cliente_id=cid, cliente=nomes[cid]))
    out.sort(key=lambda x: (x["data"], x.get("criado_em", "")), reverse=True)
    return out


def contagem(cfg):
    """Registros em aberto por cliente (para o menu e a lista de clientes)."""
    res = {}
    for cid in _clientes(cfg):
        its = _ler(cid)
        res[cid] = {"total": len(its), "abertos": sum(1 for i in its if i.get("situacao") != "resolvido"),
                    "ultimo": max((i["data"] for i in its), default="")}
    return res


def salvar(cfg, dados, por, login):
    cid = dados.get("cliente_id") or ""
    if cid not in _clientes(cfg):
        raise ValueError("Escolha o cliente.")
    tipo = dados.get("tipo") if dados.get("tipo") in TIPOS else "Observação"
    situacao = dados.get("situacao") if dados.get("situacao") in SITUACOES else "aberto"
    assunto = str(dados.get("assunto") or "").strip()[:160]
    texto = str(dados.get("texto") or "").strip()[:8000]
    if not assunto:
        raise ValueError("Escreva o assunto.")
    agora = dt.datetime.now().isoformat(timespec="seconds")
    campos = {"data": _data(dados.get("data"), True), "tipo": tipo, "assunto": assunto, "texto": texto,
              "situacao": situacao, "prazo": _data(dados.get("prazo")),
              "responsavel": str(dados.get("responsavel") or "").strip()[:60]}
    with _TRAVA:
        itens = _ler(cid)
        it = next((x for x in itens if x["id"] == dados.get("id")), None) if dados.get("id") else None
        if dados.get("id") and it is None:
            raise ValueError("Registro não encontrado (pode ter sido removido).")
        if it is None:
            it = {"id": dt.date.today().strftime("%Y%m%d") + "-" + secrets.token_hex(4), "criado_por": por,
                  "criado_login": login, "criado_em": agora}
            itens.append(it)
        else:
            it.update(alterado_por=por, alterado_em=agora)
            if it.get("situacao") != situacao and situacao == "resolvido":
                it["resolvido_em"] = agora
        it.update(campos)
        _gravar(cid, itens)
    return dict(it, cliente_id=cid, cliente=_clientes(cfg)[cid])


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


# ---------------------------------------------------------------- PDF
def _br(d):
    return dt.date.fromisoformat(d).strftime("%d/%m/%Y") if d else ""


def _moeda(v):
    s = "{:,.2f}".format(v or 0).replace(",", "X").replace(".", ",").replace("X", ".")
    return "R$ " + s


def _numeros(cfg, cid):
    """Números do mês atual do cliente (do pacote final), para o topo do PDF."""
    try:
        import painel_dados
        c = next(x for x in cfg["clientes"] if x["id"] == cid)
        import consolidado
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
    """PDF dos registros filtrados. Devolve (bytes, nome do arquivo)."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    itens = listar(cfg, **{k: filtro.get(k, "") for k in ("cliente_id", "de", "ate", "tipo", "situacao", "busca")})
    nomes = _clientes(cfg)
    cid = filtro.get("cliente_id") or ""
    titulo = "Acompanhamento — " + (nomes.get(cid) or "todos os clientes")

    ss = getSampleStyleSheet()
    AZUL = colors.HexColor("#2563eb")
    CINZA = colors.HexColor("#64748b")
    LINHA = colors.HexColor("#e2e8f0")
    st = {
        "tit": ParagraphStyle("tit", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=17, leading=21, alignment=0,
                              spaceAfter=2, textColor=colors.HexColor("#0f172a")),
        "sub": ParagraphStyle("sub", parent=ss["Normal"], fontSize=8.5, textColor=CINZA, leading=11),
        "h2": ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=12, leading=15,
                             textColor=AZUL, spaceBefore=10, spaceAfter=4),
        "ass": ParagraphStyle("ass", parent=ss["Normal"], fontName="Helvetica-Bold", fontSize=10, leading=13),
        "txt": ParagraphStyle("txt", parent=ss["Normal"], fontSize=9, leading=12.5),
        "meta": ParagraphStyle("meta", parent=ss["Normal"], fontSize=7.8, leading=10, textColor=CINZA),
        "dir": ParagraphStyle("dir", parent=ss["Normal"], fontSize=8, leading=10, textColor=CINZA, alignment=TA_RIGHT),
        "num": ParagraphStyle("num", parent=ss["Normal"], fontSize=8, leading=10, textColor=CINZA),
        "numv": ParagraphStyle("numv", parent=ss["Normal"], fontName="Helvetica-Bold", fontSize=11, leading=14),
    }
    COR_SIT = {"aberto": colors.HexColor("#dc2626"), "andamento": colors.HexColor("#d97706"),
               "resolvido": colors.HexColor("#16a34a")}

    filtros = []
    if filtro.get("de") or filtro.get("ate"):
        filtros.append("período %s a %s" % (_br(_data(filtro.get("de"))) or "início", _br(_data(filtro.get("ate"))) or "hoje"))
    if filtro.get("tipo"):
        filtros.append("tipo: %s" % filtro["tipo"])
    if filtro.get("situacao") in SITUACOES:
        filtros.append("situação: %s" % SITUACOES[filtro["situacao"]])
    elif filtro.get("situacao") == "pendentes":
        filtros.append("só em aberto e em andamento")
    if filtro.get("busca"):
        filtros.append("busca: “%s”" % filtro["busca"])

    corpo = [Paragraph(escape(titulo), st["tit"]),
             Paragraph(escape("Gerado em %s por %s · %d registro(s)%s" % (
                 dt.datetime.now().strftime("%d/%m/%Y %H:%M"), por, len(itens),
                 (" · " + "; ".join(filtros)) if filtros else "")), st["sub"]),
             Spacer(1, 8)]

    ids = [cid] if cid else sorted({i["cliente_id"] for i in itens}, key=lambda x: nomes.get(x, x))
    if not itens:
        corpo.append(Paragraph("Nenhum registro com esses filtros.", st["txt"]))
    for c in ids:
        regs = [i for i in itens if i["cliente_id"] == c]
        if not cid:
            corpo.append(Paragraph(escape(nomes.get(c, c)), st["h2"]))
        n = _numeros(cfg, c)
        if n:
            celula = lambda rot, val: [Paragraph(escape(rot), st["num"]), Paragraph(escape(val), st["numv"])]
            dados = [celula("Fat. bruto " + n["mes"], _moeda(n["fat_bruto"])),
                     celula("Fat. líquido", _moeda(n["fat_liquido"])),
                     celula("Publicidade", _moeda(abs(n["publicidade"] or 0))),
                     celula("Margem", _moeda(n["margem"]) + ((" (%.1f%%)" % (n["margem_pct"] * 100)).replace(".", ",") if n["margem_pct"] is not None else ""))]
            t = Table([[x[0] for x in dados], [x[1] for x in dados]], colWidths=[45 * mm] * 4)
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f1f5f9")),
                                   ("LEFTPADDING", (0, 0), (-1, -1), 7), ("TOPPADDING", (0, 0), (-1, 0), 6),
                                   ("BOTTOMPADDING", (0, -1), (-1, -1), 6), ("LINEAFTER", (0, 0), (-2, -1), 0.5, LINHA)]))
            corpo += [t, Spacer(1, 6)]
        abertos = sum(1 for r in regs if r.get("situacao") != "resolvido")
        corpo.append(Paragraph("%d registro(s) · %d em aberto" % (len(regs), abertos), st["meta"]))
        corpo.append(Spacer(1, 4))
        for r in regs:
            sit = r.get("situacao", "aberto")
            meta = "%s · %s" % (r.get("tipo", ""), "registrado por %s" % r.get("criado_por", "?"))
            if r.get("responsavel"):
                meta += " · responsável: %s" % r["responsavel"]
            if r.get("prazo"):
                meta += " · prazo: %s" % _br(r["prazo"])
            if r.get("alterado_por"):
                meta += " · alterado por %s" % r["alterado_por"]
            esquerda = [Paragraph(escape(r["assunto"]), st["ass"]), Paragraph(escape(meta), st["meta"])]
            if r.get("texto"):
                esquerda.append(Spacer(1, 3))
                esquerda.append(Paragraph(escape(r["texto"]).replace("\n", "<br/>"), st["txt"]))
            direita = [Paragraph(_br(r["data"]), st["dir"]),
                       Paragraph('<font color="%s"><b>%s</b></font>' % (COR_SIT[sit].hexval().replace("0x", "#"), SITUACOES[sit]), st["dir"])]
            t = Table([[esquerda, direita]], colWidths=[148 * mm, 32 * mm])
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 0), (-1, -1), 0.5, LINHA),
                                   ("LINEBEFORE", (0, 0), (0, -1), 2.2, COR_SIT[sit]),
                                   ("LEFTPADDING", (0, 0), (0, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 6),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
            corpo.append(KeepTogether(t))

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
