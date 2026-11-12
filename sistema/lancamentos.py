# -*- coding: utf-8 -*-
"""ADS e Afiliados subidos manualmente (por cliente, marketplace e mês) para completar o consolidado.

O arquivo enviado pode ser o relatório do marketplace ou o modelo do app: as colunas são achadas por
nomes parecidos (ex.: "Investimento"/"Gasto"/"Despesas", "Cliques"/"Clicks", "ID do pedido"/"Order ID"),
com o cabeçalho em qualquer das 30 primeiras linhas e em qualquer aba.

ADS (por anúncio e dia)     -> 03.Despesas Marketplace/Publicidade (−gasto), Cliques, ADS F (receita de ADS)
Afiliados (por pedido)      -> as vendas do pedido viram Type "Afiliados" e a comissão entra em
                               03.Despesas Marketplace/Afiliados (−), no SKU/anúncio/dia da venda

Guardado padronizado em [raiz]\\_lancamentos\\[Cliente]\\[tipo]_[mk]_[AAAA-MM].xlsx; o arquivo enviado fica
em _originais e a versão anterior em _anteriores (nada é apagado).
"""
import calendar
import datetime as dt
import glob
import os
import re
import shutil
import threading
import unicodedata

import comum

TIPOS = {
    "ads": {
        "nome": "ADS",
        "campos": ["Data", "Id anúncio", "SKU", "Campanha", "Gasto", "Cliques", "Receita ADS"],
        "numericos": ["Gasto", "Cliques", "Receita ADS"],
        "obrigatorio_valor": ["Gasto"],
    },
    "afiliados": {
        "nome": "Afiliados",
        "campos": ["Data", "Id do pedido", "Id anúncio", "SKU", "Comissão afiliado"],
        "numericos": ["Comissão afiliado"],
        "obrigatorio_valor": ["Comissão afiliado"],
    },
    "visitas": {
        "nome": "Visitas",
        "campos": ["Data", "Id anúncio", "SKU", "Visitas", "Vendas", "Produto"],
        "numericos": ["Visitas", "Vendas"],
        "obrigatorio_valor": ["Visitas"],
    },
}
# Relatórios que NÃO são diários: um relatório só, com o número da venda, que liga a comissão à venda (margem).
# Cada envio substitui os pedidos que traz; os outros pedidos já enviados continuam.
RELATORIO_UNICO = {("afiliados", "shopee"), ("afiliados", "mercadolivre")}


def unico(tipo, mk):
    return (tipo, mk) in RELATORIO_UNICO


_APELIDOS = {
    "Data": ["data", "dia", "date", "desde", "datainicial", "inicio", "datadavenda", "datadopedido", "datadacompra", "datadeinicio",
             "datahora", "createdtime", "horadopedido", "datadeconclusao", "orderdate", "horariodopedido", "tempodopedido",
             "horariodacompra", "ordertime", "datadecriacao"],
    "Id anúncio": ["idanuncio", "iddoanuncio", "anuncio", "codigodoanuncio", "numerodoanuncio", "mlb", "itemid",
                   "iddoitem", "idproduto", "iddoproduto", "productid", "idprodutoshopee", "codigodoproduto",
                   "deanuncio", "asin", "skuid"],
    "SKU": ["sku", "skudoproduto", "sellersku", "numerodereferenciasku", "codigosku", "skuvendedor"],
    "Campanha": ["campanha", "nomedacampanha", "campaign", "campaignname", "nomedoanuncio", "titulodoanuncio"],
    "Gasto": ["gasto", "investimento", "despesas", "despesa", "custo", "valorinvestido", "spend", "cost",
              "custototal", "gastototal", "investimentototal", "custodapublicidade", "valorgasto"],
    "Cliques": ["cliques", "clicks", "clique", "numerodecliques"],
    "Receita ADS": ["receita", "receitaads", "receitadeads", "receitaporpublicidade",
                    "gmv", "receitadeanuncios", "revenue", "grossrevenue", "receitatotal",
                    "faturamentoads", "adsf", "receitadevendas", "vendasatribuidas", "receitaatribuida",
                    "vendas",    # Magalu (Product Report): "Vendas" = receita direta em R$
                    "receitabruta"],   # TikTok (relatório por produto)
    "Id do pedido": ["iddopedido", "pedido", "numerodopedido", "ndevenda", "numerodevenda", "orderid", "idpedido",
                     "codigodopedido", "venda", "nvenda", "amazonorderid", "ordersn", "numerodavenda"],
    "Visitas": ["visitasunicas", "visitas", "visits", "uniquevisits", "visitantesunicos", "visitasunicasdoanuncio"],
    "Produto": ["nomedoproduto", "titulodoanuncio", "titulo", "nomedoanuncio", "produto"],
    "Vendas": ["quantidadedevendas", "qtddevendas", "qtdvendas", "vendas", "numerodevendas", "pedidos", "orders"],
    "Comissão afiliado": ["comissaoafiliado", "comissaodoafiliado", "comissao", "comissaoestimada", "commission",
                          "valordacomissao", "despesadeafiliado", "despesaafiliados", "comissaototal",
                          "affiliatecommission", "estimatedcommission", "comissaodeafiliados", "taxadeafiliado",
                          "despesaestimada", "despesadecomissao", "despesasdeafiliados", "despesadoafiliado"],
}


def _chave(t):
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()
    return "".join(c for c in t if c.isalnum())


def _mapear(cab, tipo):
    campos = TIPOS[tipo]["campos"]
    achados = {}
    for j, nome in enumerate(cab):
        # tira moeda e explicações entre parênteses: "Receita\n(Moeda local)" -> receita; "Investimento (R$)" -> investimento
        txt = re.sub(r"\([^)]*\)", " ", str(nome or "").lower())
        k = _chave(re.sub(r"r\$|\bbrl\b|\breais\b|moeda local", "", txt)) or _chave(nome)
        if not k:
            continue
        for campo in campos:
            if campo in achados:
                continue
            ap = _APELIDOS[campo]
            if k.startswith(("cpc", "ctr", "roas", "acos", "custopor", "taxa", "percentual")) and campo != "Comissão afiliado":
                continue
            if k in ap or any(k.startswith(a) for a in ap if len(a) >= 8):
                achados[campo] = j
                break
    return achados


# ---------------------------------------------------------------- leitura
def _linhas(caminho):
    import tratador
    return tratador._linhas_da_planilha(caminho)


def _numero(v):
    import tratador
    n, ok = tratador._numero(v)
    if not ok and isinstance(v, str):
        t = v.strip().replace("%", "")
        n, ok = tratador._numero(t)
    return n, ok


_MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
          "feb": 2, "apr": 4, "may": 5, "aug": 8, "sep": 9, "oct": 10, "dec": 12}


def _data(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:   # número de série do Excel
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        a, mm, d = map(int, m.groups())
    else:
        m = re.match(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s)
        if m:
            d, mm, a = map(int, m.groups())
            if mm > 12 and d <= 12:      # formato americano (TikTok: MM/DD/AAAA)
                d, mm = mm, d
        else:
            m = re.match(r"(\d{1,2})[\s/.-]+(?:de\s+)?([a-zç]{3,})\.?[\s/.-]+(?:de\s+)?(\d{2,4})", s.lower())
            if not m:
                return None
            d, a = int(m.group(1)), int(m.group(3))
            mm = _MESES.get(m.group(2)[:3])
            if not mm:
                return None
        if a < 100:
            a += 2000
    try:
        return dt.date(a, mm, d)
    except ValueError:
        return None


def _texto(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if v.is_integer():
            return str(int(v))
    return str(v).strip()


_MESES_ABREV = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}


def _ler_afiliados_ml(linhas, aba, hoje=None):
    """Afiliados do Mercado Livre (página "Vendas com afiliados" copiada/exportada): cada venda ocupa 3 linhas.
    1ª: "#<n.º da venda (16 dígitos)><dia>/<mês> HH:MM hs" | ... | Tipo de campanha | Valor da venda | Tarifa por unidade |
        Taxa de venda ("R$ 29,47Tarifa 8,00%" = o que se paga ao afiliado pela venda; 0 = afiliado especial, sem custo)
    3ª: "#<anúncio> | SKU: #<sku+título> ..." (o anúncio vira MLB<número>)."""
    hoje = hoje or dt.date.today()
    saida, sem_valor = [], 0
    padrao = re.compile(r"^#?(\d{16})(\d{1,2})/([A-Za-zÀ-ú]{3})")
    for i, r in enumerate(linhas):
        a = str((r[0] if r else "") or "").strip()
        m = padrao.match(a)
        if not m:
            continue
        pedido, dia, mes_txt = m.group(1), int(m.group(2)), _chave(m.group(3))[:3]
        mes = _MESES_ABREV.get(mes_txt)
        if not mes:
            continue
        ano = hoje.year - (1 if mes > hoje.month else 0)
        try:
            data = dt.date(ano, mes, dia)
        except ValueError:
            continue
        taxa_txt = r[6] if len(r) > 6 else None
        mt = re.search(r"R\$\s*([\d.]+,\d{2})", str(taxa_txt or ""))
        taxa = _numero(mt.group(1))[0] if mt else (_numero(taxa_txt)[0] if isinstance(taxa_txt, (int, float)) else 0.0)
        anuncio = ""
        for rr in linhas[i + 1:i + 3]:
            mm = re.match(r"^#(\d{6,})\s*\|", str((rr[1] if len(rr) > 1 else "") or "").strip())
            if mm:
                anuncio = "MLB" + mm.group(1)
                break
        if not taxa:
            sem_valor += 1
            continue                        # afiliado especial / sem tarifa: não é custo
        saida.append({"Data": data, "Id do pedido": pedido, "Id anúncio": anuncio, "SKU": "", "Comissão afiliado": abs(taxa or 0.0)})
    info = {"aba": aba, "colunas": ["Comissão afiliado", "Data", "Id anúncio", "Id do pedido"], "sem_data": 0,
            "ignoradas_invalidas": 0, "formato": "afiliados Mercado Livre", "vendas_sem_tarifa": sem_valor}
    return saida, info


def _ler_visitas_shopee(linhas, aba):
    """Shopee > Dados do produto (parentskudetail...xlsx), aba "Produtos com Melhor Desempenho":
    visitas = "Cliques Por Produto"; vendas = "Unidades (Pedido pago)". Um item com variações: vale a linha do item (sem variação)."""
    cab = [str(c or "").strip() for c in linhas[0]]
    idx = {c: j for j, c in enumerate(cab)}
    jv, jq = idx.get("Cliques Por Produto"), idx.get("Unidades (Pedido pago)")
    ji, js, jsv, jvar = idx.get("ID do Item"), idx.get("SKU Principle"), idx.get("SKU da Variação"), idx.get("ID da Variação")
    if jv is None or jq is None or ji is None:
        raise ValueError("faltam as colunas ID do Item, Cliques Por Produto ou Unidades (Pedido pago).")
    por_item = {}
    for r in linhas[1:]:
        g = lambda j: (r[j] if j is not None and j < len(r) else None)
        item = _texto(g(ji))
        if not item or not item.strip("-"):
            continue
        var = _texto(g(jvar))
        eh_var = bool(var and var.strip("-"))
        sku = _texto(g(jsv)) if eh_var else ""
        if not sku or not sku.strip("-"):
            sku = _texto(g(js))
        sku = "" if sku.strip() in ("-", "--") else sku
        vis, _ = _numero(g(jv))
        qtd, _ = _numero(g(jq))
        d = por_item.setdefault(item, {"pai": None, "vars": []})
        linha = {"Data": None, "Id anúncio": item, "SKU": sku, "Visitas": float(vis or 0), "Vendas": float(qtd or 0)}
        if eh_var:
            d["vars"].append(linha)
        else:
            d["pai"] = linha
    saida = []
    for it, d in por_item.items():
        if d["pai"]:
            saida.append(d["pai"])          # a linha do item já soma as variações
        else:
            saida.extend(d["vars"])
    saida = [l for l in saida if l["Visitas"]]
    info = {"aba": aba, "colunas": ["Id anúncio", "SKU", "Vendas", "Visitas"], "sem_data": len(saida), "ignoradas_invalidas": 0,
            "formato": "Shopee (Dados do produto)"}
    return saida, info


def _ler_visitas_tiktok(linhas, aba):
    """TikTok Shop > Análise de produtos (product_list_AAAAMMDD.xlsx): linha 1 "Data de análise: dd/mm/aaaa~dd/mm/aaaa",
    cabeçalho na 4ª linha com grupos (Tudo / LIVE / Vídeo / Afiliado / Cartão). Vale o grupo "Tudo" (a 1ª ocorrência):
    visitas = "Cliques no produto"; vendas = "Itens vendidos". Não tem SKU: liga às vendas pelo nome do produto."""
    per = None
    for r in linhas[:3]:
        txt = " ".join(str(x) for x in r if x not in (None, ""))
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*~\s*(\d{1,2})/(\d{1,2})/(\d{4})", txt)
        if m:
            d1, m1, a1, d2, m2, a2 = map(int, m.groups())
            per = (dt.date(a1, m1, d1), dt.date(a2, m2, d2))
            break
    if per and per[0] != per[1]:
        raise ValueError("o relatório de visitas deve ser subido por dia: este arquivo é de %s a %s. "
                         "Exporte um arquivo para cada dia." % (per[0].strftime("%d/%m/%Y"), per[1].strftime("%d/%m/%Y")))
    n = next(i for i, r in enumerate(linhas[:10]) if "nomedoproduto" in {_chave(x) for x in r})
    ks = [_chave(c) for c in linhas[n]]
    pri = lambda k: next((j for j, x in enumerate(ks) if x == k), None)     # 1ª ocorrência = grupo "Tudo"
    jn, ji, jv, jq = pri("nomedoproduto"), pri("iddeproduto"), pri("cliquesnoproduto"), pri("itensvendidos")
    if None in (jn, ji, jv, jq):
        raise ValueError("faltam as colunas Nome do produto, ID de produto, Cliques no produto ou Itens vendidos.")
    saida = []
    for r in linhas[n + 1:]:
        g = lambda j: (r[j] if j < len(r) else None)
        pid, nome = _texto(g(ji)), _texto(g(jn))
        if not pid and not nome:
            continue
        vis, _ = _numero(g(jv))
        qtd, _ = _numero(g(jq))
        if vis:
            saida.append({"Data": per[0] if per else None, "Id anúncio": pid, "SKU": "", "Visitas": float(vis or 0),
                          "Vendas": float(qtd or 0), "Produto": nome[:200]})
    info = {"aba": aba, "colunas": ["Id anúncio", "Produto", "Vendas", "Visitas"], "sem_data": 0 if per else len(saida),
            "ignoradas_invalidas": 0, "formato": "TikTok Shop (Análise de produtos)"}
    if per:
        info["dia_do_arquivo"] = per[0].isoformat()
    return saida, info


def _ler_visitas_amazon(linhas, aba, n):
    """Amazon > Relatórios de negócios > "Detalhe da página de vendas e tráfego por item" (ASIN pai ou filho):
    ASIN, Título, Sessões - Total, Visualizações da página - Total, Unidades pedidas… Não traz data: vale o dia do envio
    (ou o dia no nome do arquivo). Visitas = "Sessões - Total" (visitantes únicos no dia; é a base da conversão da Amazon)."""
    ks = [_chave(c) for c in linhas[n]]
    j = lambda *k: next((ks.index(x) for x in k if x in ks), None)
    ja = j("asinfilho", "asinchild", "asinparent", "asinpai", "asin")
    jv, jt, ju = j("sessoestotal", "sessoes"), j("titulo", "title"), j("unidadespedidas")
    if ja is None or jv is None:
        raise ValueError("faltam as colunas ASIN ou Sessões - Total.")
    saida = []
    for r in linhas[n + 1:]:
        g = lambda x: (r[x] if x is not None and x < len(r) else None)
        asin = _texto(g(ja)).upper()
        vis, _ = _numero(g(jv))
        if not asin or not vis:
            continue
        un, _ = _numero(g(ju))
        saida.append({"Data": None, "Id anúncio": asin, "SKU": "", "Visitas": float(vis), "Vendas": float(un or 0),
                      "Produto": _texto(g(jt))[:200]})
    info = {"aba": aba, "colunas": ["Id anúncio (ASIN)", "Visitas (sessões)", "Produto"], "sem_data": 0, "ignoradas_invalidas": 0,
            "formato": "Amazon (vendas e tráfego por item)",
            "aviso": "Amazon: %d ASIN(s) com visita, %s sessões no total (visitas = Sessões - Total). O dia é o do envio." % (
                len(saida), ("%d" % sum(l["Visitas"] for l in saida)))}
    return saida, info


def _ler_visitas_magalu(linhas, aba):
    """Magalu > Desempenho dos produtos (CSV): "Data Análise" (01/10/2026 a 01/10/2026), SKU, "Quantidade de Visitas".
    Já vem com o SKU do vendedor; o ID do anúncio sai do "Link do Produto" (.../products/editar/<id>)."""
    ks = [_chave(c) for c in linhas[0]]
    j = lambda k: ks.index(k) if k in ks else None
    jd, js, jv, jn, jl = j("dataanalise"), j("sku"), j("quantidadedevisitas"), j("nomedoproduto"), j("linkdoproduto")
    if jv is None or js is None:
        raise ValueError("faltam as colunas SKU ou Quantidade de Visitas.")
    saida, per = [], None
    for r in linhas[1:]:
        g = lambda x: (r[x] if x is not None and x < len(r) else None)
        txt = str(g(jd) or "")
        m = re.findall(r"(\d{1,2})/(\d{1,2})/(\d{4})", txt)
        d1 = d2 = None
        if m:
            d1 = dt.date(int(m[0][2]), int(m[0][1]), int(m[0][0]))
            d2 = dt.date(int(m[-1][2]), int(m[-1][1]), int(m[-1][0]))
            if d1 != d2:
                raise ValueError("o relatório de visitas deve ser subido por dia: este arquivo é de %s a %s. "
                                 "Exporte um arquivo para cada dia." % (d1.strftime("%d/%m/%Y"), d2.strftime("%d/%m/%Y")))
            per = per or d1
        sku = _texto(g(js))
        vis, _ = _numero(g(jv))
        if not sku or not vis:
            continue
        ml = re.search(r"/editar/([0-9a-fA-F]+)", str(g(jl) or ""))
        saida.append({"Data": d1, "Id anúncio": ml.group(1) if ml else "", "SKU": sku, "Visitas": float(vis),
                      "Vendas": 0.0, "Produto": _texto(g(jn))[:200]})
    info = {"aba": aba, "colunas": ["Data", "SKU", "Visitas"], "sem_data": sum(1 for l in saida if not l["Data"]),
            "ignoradas_invalidas": 0, "formato": "Magalu (Desempenho dos produtos)"}
    if per:
        info["dia_do_arquivo"] = per.isoformat()
    return saida, info


def _eh_grupo_shopee(linhas):
    """Relatório "Grupo de Anúncio - Shopee" (Shopee-Ads-All-Ad-Group-Data-…): campanhas e, logo abaixo de cada uma,
    os anúncios dela. Devolve o índice da linha de cabeçalho ou None."""
    for n, r in enumerate(linhas[:15]):
        k = {_chave(x) for x in r}
        if "anuncionomedoproduto" in k and "iddoproduto" in k and "tiposdeanuncios" in k and "despesas" in k:
            return n
        # "GMV Max da Loja" (Dados-do-GMV-MAX-da-Loja-…): linha 1 = total da campanha (ID "-"), depois os produtos
        if "nomedoproduto" in k and "iddoproduto" in k and "despesas" in k and "nomedoanuncio" not in k:
            return n
    return None


def _rs(v):
    return ("R$ %.2f" % v).replace(".", ",")


def _ler_ads_grupo_shopee(linhas, aba, n):
    """Lê o relatório por grupo de anúncio da Shopee.
    - Linha de CAMPANHA: "ID do produto" = "-" (tem Status, Método de Lance e datas). Traz o TOTAL da campanha.
    - Linha de ANÚNCIO: "ID do produto" numérico, logo abaixo da sua campanha. É o detalhe — já está DENTRO do total.
    A campanha vira uma linha de campanha fechada (o investimento que vale) e os anúncios viram o detalhe dela
    (CAMPO_CAMP = "detalhe:<campanha>"), então efetivas_ads usa o total da campanha como teto: nunca soma os dois níveis.
    Anúncios sem campanha acima deles (não deveria acontecer) entram como anúncios avulsos."""
    cab = [_chave(x) for x in linhas[n]]
    gmv_loja = "tiposdeanuncios" not in cab      # relatório de uma campanha só (GMV Max da Loja)

    def j(*nomes):
        for nm in nomes:
            if nm in cab:
                return cab.index(nm)
        return None
    jn, jid, jst = j("anuncionomedoproduto", "nomedoproduto"), j("iddoproduto"), j("status")
    jg, jc, jr, ji = j("despesas"), j("cliques"), j("gmv"), j("impressoes")
    jp, jit = j("conversoes"), j("itensvendidos")

    def val(r, k):
        if k is None or k >= len(r):
            return 0.0
        v, ok = _numero(r[k])
        return abs(v or 0.0) if ok else 0.0
    campanhas, avulsos, atual, nomes = [], [], None, {}
    for r in linhas[n + 1:]:
        if not r or jn is None or jn >= len(r) or not _texto(r[jn]):
            continue
        nome = re.sub(r"\s+", " ", _texto(r[jn]))
        pid = _texto(r[jid]) if jid is not None and jid < len(r) else ""
        m = {"Gasto": val(r, jg), "Cliques": val(r, jc), "Receita ADS": val(r, jr), "Impressões": val(r, ji),
             "Pedidos": val(r, jp), "Itens": val(r, jit)}
        if not pid.isdigit():
            if nome.lower() in ("total", "totais"):
                continue
            # sem ID de campanha no arquivo: nome igual dentro do mesmo arquivo ganha um número para não juntar campanhas
            if nome.isdigit():
                nome = "Campanha " + nome        # nome só com números seria confundido com ID de produto
            nomes[nome] = nomes.get(nome, 0) + 1
            if nomes[nome] > 1:
                nome = "%s [%d]" % (nome, nomes[nome])
            atual = {"nome": nome, "status": _texto(r[jst]) if jst is not None and jst < len(r) else "", "m": m, "ads": []}
            campanhas.append(atual)
        elif atual is None:
            avulsos.append((pid, nome, m))
        else:
            atual["ads"].append((pid, nome, m))
    if not campanhas and not avulsos:
        raise ValueError("o relatório de grupo de anúncio da Shopee não tem nenhuma campanha.")
    saida, conc = [], []
    ADD = ("Gasto", "Cliques", "Receita ADS", "Impressões", "Pedidos")
    for c in campanhas:
        soma = {k: round(sum(a[2][k] for a in c["ads"]), 2) for k in ADD}
        dif = round(c["m"]["Gasto"] - soma["Gasto"], 2)
        sit = ("sem anúncios" if not c["ads"] else "igual" if abs(dif) < 0.005
               else "não detalhado" if dif > 0 else "acima (limitado ao total)")
        conc.append({"campanha": c["nome"], "status": c["status"], "anuncios": len(c["ads"]),
                     "campanha_valores": {k: round(c["m"][k], 2) for k in ADD}, "soma_anuncios": soma,
                     "diferenca_gasto": dif, "situacao": sit,
                     "diferencas": {k: round(c["m"][k] - soma[k], 2) for k in ADD if abs(c["m"][k] - soma[k]) >= 0.005}})
        if c["m"]["Gasto"] or c["m"]["Cliques"] or c["m"]["Receita ADS"]:
            saida.append({"Data": None, "Id anúncio": c["nome"], "SKU": "", "Campanha": c["nome"], "Gasto": c["m"]["Gasto"],
                          "Cliques": c["m"]["Cliques"], "Receita ADS": c["m"]["Receita ADS"], "_grupo": True, "_so_nome": gmv_loja, "_tag": ""})
        for pid, nm, m in c["ads"]:
            if m["Gasto"] or m["Cliques"] or m["Receita ADS"]:
                saida.append({"Data": None, "Id anúncio": pid, "SKU": "", "Campanha": nm, "Gasto": m["Gasto"], "Cliques": m["Cliques"],
                              "Receita ADS": m["Receita ADS"], "_grupo": True, "_so_nome": gmv_loja, "_tag": DETALHE + c["nome"]})
    for pid, nm, m in avulsos:
        if m["Gasto"] or m["Cliques"] or m["Receita ADS"]:
            saida.append({"Data": None, "Id anúncio": pid, "SKU": "", "Campanha": nm, "Gasto": m["Gasto"], "Cliques": m["Cliques"],
                          "Receita ADS": m["Receita ADS"], "_grupo": True, "_tag": ""})
    tot_c = round(sum(c["m"]["Gasto"] for c in campanhas), 2)
    tot_av = round(sum(a[2]["Gasto"] for a in avulsos), 2)
    nd = round(sum(x["diferenca_gasto"] for x in conc if x["diferenca_gasto"] > 0 and x["anuncios"]), 2)
    ac = [x for x in conc if x["situacao"].startswith("acima")]
    info = {"aba": aba, "formato": "gmv_max_loja_shopee" if gmv_loja else "grupo_shopee", "colunas": ["Data", "Id anúncio", "Campanha", "Gasto", "Cliques", "Receita ADS"],
            "sem_data": 0, "ignoradas_invalidas": 0, "conciliacao": conc,
            "campanhas": len(campanhas), "anuncios": sum(len(c["ads"]) for c in campanhas), "anuncios_sem_campanha": len(avulsos),
            "investimento_relatorio": round(tot_c + tot_av, 2)}
    if gmv_loja and len(campanhas) == 1:
        av = ("Relatório da campanha %s: total %s (linha 1) e %d anúncio(s) como detalhe dela, sem somar os dois. "
              "As outras campanhas do dia continuam como estavam." % (campanhas[0]["nome"], _rs(tot_c), info["anuncios"]))
    else:
        av = ("Relatório por grupo de anúncio: %d campanha(s) e %d anúncio(s). O investimento é o total das campanhas "
              "(%s), sem somar os anúncios de novo." % (len(campanhas), info["anuncios"], _rs(tot_c)))
    if nd:
        av += " %s das campanhas não veio detalhado nos anúncios e ficou na linha da campanha." % _rs(nd)
    if ac:
        av += " %d campanha(s) com anúncios somando MAIS que o total (%s): valeu o total da campanha." % (
            len(ac), ", ".join("%s +%s" % (x["campanha"], _rs(-x["diferenca_gasto"])) for x in ac[:4]))
    if avulsos:
        av += " %d anúncio(s) sem campanha acima deles entraram avulsos (%s)." % (len(avulsos), _rs(tot_av))
    info["aviso"] = av
    if not saida:
        marca = {c: (0.0 if c in TIPOS["ads"]["numericos"] else "") for c in TIPOS["ads"]["campos"]}
        marca["Data"] = None
        marca["Campanha"] = "(sem valor no dia)"
        info["sem_valor"] = True
        return [marca], info
    return saida, info


def _ler_afiliados_tiktok(linhas, aba, n):
    """Afiliados do TikTok (affiliate_orders_….xlsx): uma linha por item de pedido, vários dias no mesmo arquivo.
    - Data = "Horário de criação" do pedido; pedido = "ID do pedido"; anúncio = "ID do SKU" (liga ao "SKU ID" do tratado).
    - Comissão = paga de verdade (comissão real + Shop Ads real + bônus real) quando o TikTok já informou;
      senão a estimada (padrão + Shop Ads + bônus estimados).
    - Sem comissão: status "Inelegível" ou "Não pago pelo cliente", e item totalmente devolvido/reembolsado."""
    cab = [_chave(x) for x in linhas[n]]
    j = lambda k: cab.index(k) if k in cab else None
    cel = lambda r, k: r[j(k)] if j(k) is not None and j(k) < len(r) else None

    def num(v):
        if v is None or _texto(v) == "":
            return None
        x, ok = _numero(v)
        return abs(x or 0.0) if ok else None
    REAL = ("pagamentodecomissaoreal", "pagamentorealdecomissaodoshopads", "bonusrealdecriadorcofinanciado")
    EST = ("estimativadepagamentodecomissaopadrao", "estimativadepagamentodacomissaodoshopads", "bonusestimadodecriadorcofinanciado")
    saida, cont = [], {}
    tot = {"pendente": 0.0, "fora": 0.0}
    sem_data = 0
    for r in linhas[n + 1:]:
        ped = _texto(cel(r, "iddopedido"))
        if not ped:
            continue
        st = _texto(cel(r, "statusdopedido"))
        dev = _texto(cel(r, "totalmentedevolvidooureembolsado")).lower() in ("sim", "yes")
        reais = [num(cel(r, k)) for k in REAL]
        if any(v is not None for v in reais):
            v, origem = sum(x or 0.0 for x in reais), "real"
        else:
            v, origem = sum(num(cel(r, k)) or 0.0 for k in EST), "estimada"
        fora = st.lower().startswith(("inelegivel", "inelegível", "nao pago", "não pago")) or dev
        rot = (st or "sem status") + (" (devolvido)" if dev else "")
        c = cont.setdefault(rot, [0, 0.0])
        c[0] += 1
        c[1] += v
        if fora:
            tot["fora"] += v
            continue
        d = _data(cel(r, "horariodecriacao"))
        if d is None:
            sem_data += 1
            continue
        if v <= 0:
            continue
        tot["pendente"] += v
        saida.append({"Data": d, "Id do pedido": ped, "Id anúncio": _texto(cel(r, "iddosku")),
                      "SKU": _texto(cel(r, "skudovendedor")), "Comissão afiliado": round(v, 6), "_origem": origem})
    dias = sorted({l["Data"] for l in saida})
    info = {"aba": aba, "formato": "tiktok_afiliados", "colunas": ["Data", "Id do pedido", "Id anúncio", "SKU", "Comissão afiliado"],
            "sem_data": sem_data, "ignoradas_invalidas": 0, "por_status": {k: [v[0], round(v[1], 2)] for k, v in cont.items()}}
    reais = sum(1 for l in saida if l["_origem"] == "real")
    info["aviso"] = ("Afiliados TikTok: %d item(ns) com comissão em %d dia(s)%s, total %s (%s). Fora: %s (sem comissão)." % (
        len(saida), len(dias), (" (%s a %s)" % (dias[0].strftime("%d/%m"), dias[-1].strftime("%d/%m"))) if dias else "",
        _rs(tot["pendente"]), ("%d com comissão já paga, o resto estimada" % reais) if reais else "comissão estimada",
        ", ".join("%s %d" % (k, v[0]) for k, v in cont.items() if k.lower().startswith(("inel", "não pago", "nao pago")) or "devolvido" in k) or "nenhum"))
    if not saida:
        raise ValueError("nenhum pedido com comissão de afiliado neste relatório (%s)." %
                         ", ".join("%s: %d" % (k, v[0]) for k, v in cont.items()))
    return saida, info


def ler(tipo, caminho):
    """Devolve (linhas padronizadas, info). Linhas: dicts com os campos de TIPOS[tipo]."""
    t = TIPOS[tipo]
    if tipo == "visitas":
        abas = _linhas(caminho)
        for aba, linhas in abas:
            cab = {_chave(x) for x in (linhas[0] if linhas else [])}
            cabs = [{_chave(x) for x in l} for l in linhas[:10]]
            tiktok = any("nomedoproduto" in c and "cliquesnoproduto" in c and "iddeproduto" in c for c in cabs)
            n_amz = next((k for k, c in enumerate(cabs) if ("sessoestotal" in c or "sessoes" in c) and
                          any(x in c for x in ("asinparent", "asinpai", "asinfilho", "asinchild", "asin"))), None)
            if n_amz is not None:
                saida, info = _ler_visitas_amazon(linhas, aba, n_amz)
                if not saida:
                    marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
                    marca["Data"] = None
                    info["sem_valor"] = True
                    return [marca], info
                return saida, info
            if "dataanalise" in cab and "quantidadedevisitas" in cab:
                saida, info = _ler_visitas_magalu(linhas, aba)
                if not saida:
                    marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
                    marca["Data"] = None
                    info["sem_valor"] = True
                    return [marca], info
                return saida, info
            if tiktok or ("iddoitem" in cab and "cliquesporproduto" in cab):
                saida, info = (_ler_visitas_tiktok if tiktok else _ler_visitas_shopee)(linhas, aba)
                if not saida:
                    marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
                    marca["Data"] = None
                    info["sem_valor"] = True
                    return [marca], info
                return saida, info
    if tipo == "afiliados":
        for aba, linhas in _linhas(caminho):
            cab = {_chave(x) for x in (linhas[0] if linhas else [])}
            if "detalhesdavenda" in cab and "taxadevenda" in cab:
                saida, info = _ler_afiliados_ml(linhas, aba)
                if not saida and not info["vendas_sem_tarifa"]:
                    raise ValueError("não achei nenhuma venda no relatório de afiliados do Mercado Livre.")
                if not saida:
                    marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
                    marca["Data"] = None
                    info["sem_valor"] = True
                    return [marca], info
                return saida, info
    if tipo == "afiliados":
        for aba, linhas in _linhas(caminho):
            for n, r in enumerate(linhas[:15]):
                k = {_chave(x) for x in r}
                if "iddopedido" in k and "iddosku" in k and "statusdopedido" in k and "horariodecriacao" in k and (
                        "estimativadepagamentodecomissaopadrao" in k or "pagamentodecomissaoreal" in k):
                    return _ler_afiliados_tiktok(linhas, aba, n)
            break
    if tipo == "ads":
        for aba, linhas in _linhas(caminho):
            n = _eh_grupo_shopee(linhas)
            if n is not None:
                return _ler_ads_grupo_shopee(linhas, aba, n)
    melhor = None
    for aba, linhas in _linhas(caminho):
        for n, cab in enumerate(linhas[:30]):
            m = _mapear(cab, tipo)
            tem_valor = all(c in m for c in t["obrigatorio_valor"])
            tem_id = any(c in m for c in (("Id anúncio", "SKU", "Campanha") if tipo == "ads" else ("Id do pedido", "Id anúncio", "SKU")))
            pontos = len(m) + (5 if tem_valor and tem_id else 0)
            if melhor is None or pontos > melhor[0]:
                melhor = (pontos, aba, n, m, linhas, tem_valor and tem_id)
    if not melhor or not melhor[5]:
        if tipo == "ads":
            for aba, linhas in _linhas(caminho):
                cabs = [{_chave(x) for x in l} for l in linhas[:5]]
                if any("tempo" in c and "custo" in c for c in cabs):
                    raise ValueError("este relatório da TikTok é por HORA e não diz o produto, então o gasto não liga às vendas. "
                                     "Baixe o relatório por PRODUTO (um arquivo por dia; o dia vai no nome do arquivo, ex.: 5.xlsx).")
        exig = "o gasto e o anúncio (ou SKU/campanha)" if tipo == "ads" else "a comissão e o pedido (ou anúncio/SKU)"
        raise ValueError("não achei as colunas de %s. Use o modelo da planilha de %s." % (exig, t["nome"]))
    _, aba, n, m, linhas, _ = melhor
    chaves_cab = [_chave(x) for x in linhas[n]]
    # TikTok "Product data" de todas as campanhas (Nome do produto, ID da campanha, Por dia, Nome da campanha, ID do produto,
    # Custo, Custo (afiliado), Pedidos de SKU, Custo por pedido, Receita bruta, ROI): uma linha por produto x campanha x dia;
    # pode vir mais de um dia. Gasto = "Custo" (o "Custo (afiliado)" não é investimento em anúncio: fica de fora).
    tiktok_dias = tipo == "ads" and "pordia" in chaves_cab and "iddacampanha" in chaves_cab and "nomedoproduto" in chaves_cab
    if tiktok_dias:
        m["Data"] = chaves_cab.index("pordia")
        m["Campanha"] = chaves_cab.index("nomedoproduto")     # liga às vendas pelo nome, como o relatório por produto
        m["Id anúncio"] = chaves_cab.index("iddoproduto")
        m["Gasto"] = chaves_cab.index("custo")
        if "receitabruta" in chaves_cab:
            m["Receita ADS"] = chaves_cab.index("receitabruta")
    if tipo == "ads" and "Campanha" not in m:
        # TikTok (por produto): o nome do produto fica em Campanha para ligar às vendas pelo nome
        for j, nome in enumerate(linhas[n]):
            if _chave(nome) == "nomedoproduto":
                m["Campanha"] = j
                break
    # Shopee (SellerConversionReport): o custo do afiliado é "despesas" MENOS a "Taxa de serviço de Afiliados do Vendedor",
    # porque essa taxa já vem descontada no relatório de vendas (senão contaria 2 vezes)
    j_desp = j_taxa = None
    if tipo == "afiliados":
        for j, nome in enumerate(linhas[n]):
            k = _chave(re.sub(r"\([^)]*\)", " ", re.sub(r"r\$", "", str(nome or "").lower())))
            if k == "despesas" and j_desp is None:
                j_desp = j
            elif k.startswith("taxadeservicodeafiliados") and j_taxa is None:
                j_taxa = j
        if j_desp is None or j_taxa is None:
            j_desp = j_taxa = None
    saida, invalidos, sem_data, datas_zeradas = [], [], 0, []
    for i, r in enumerate(linhas[n + 1:]):
        def col(c):
            j = m.get(c)
            return r[j] if j is not None and j < len(r) else None
        item = {}
        for c in t["campos"]:
            if c in t["numericos"]:
                v, ok = _numero(col(c))
                if not ok:
                    invalidos.append("linha %d (%s): “%s”" % (n + 2 + i, c, col(c)))
                    v = None
                item[c] = v or 0.0
            elif c == "Data":
                item[c] = _data(col(c))
            else:
                item[c] = _texto(col(c))
        if j_desp is not None:
            desp, ok1 = _numero(r[j_desp] if j_desp < len(r) else None)
            taxa, ok2 = _numero(r[j_taxa] if j_taxa < len(r) else None)
            if ok1 and desp is not None:
                item["Comissão afiliado"] = max(abs(desp) - abs(taxa or 0.0), 0.0)
        ids = [item.get(k) for k in ("Id anúncio", "SKU", "Campanha", "Id do pedido") if k in item]
        if not any(ids) or not any(item[c] for c in t["numericos"]):
            if any(ids) and item.get("Data"):
                datas_zeradas.append(item["Data"])
            continue                      # linha vazia, de total ou sem valor
        if any(str(x).strip().lower() in ("total", "totais", "soma") for x in ids):
            continue
        for k in ("Id anúncio", "SKU"):
            if str(item.get(k) or "").strip() in ("-", "--", "N/A", "n/a"):
                item[k] = ""
        if not item.get("Id anúncio") and item.get("Campanha"):
            item["Id anúncio"] = item["Campanha"]
        if tipo == "afiliados":
            item["Comissão afiliado"] = abs(item["Comissão afiliado"])
        elif tipo == "ads":
            item["Gasto"] = abs(item["Gasto"])
        if item.get("Data") is None:
            sem_data += 1
        saida.append(item)
    if invalidos and len(invalidos) > max(3, len(saida) // 10):
        raise ValueError("valores que não são número em %d linha(s), ex.: %s" % (len(invalidos), "; ".join(invalidos[:3])))
    info = {"aba": aba, "colunas": sorted(m), "sem_data": sem_data, "ignoradas_invalidas": len(invalidos)}
    if tiktok_dias:
        # dias do relatório sem nenhum gasto: entram como dia enviado (linha zerada), para substituir o que havia
        com_valor = {l["Data"] for l in saida}
        for d in sorted(set(datas_zeradas) - com_valor):
            marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
            marca.update(Data=d, Campanha="(sem valor no dia)")
            saida.append(marca)
        dias = sorted({l["Data"] for l in saida if l.get("Data")})
        info["formato"] = "tiktok_produto_por_dia"
        if dias:
            info["aviso"] = ("TikTok por produto e dia: %d dia(s) (%s a %s), gasto total %s. "
                             "O \"Custo (afiliado)\" não entra no ADS." % (len(dias), dias[0].strftime("%d/%m"), dias[-1].strftime("%d/%m"),
                                                                       _rs(sum(l["Gasto"] for l in saida))))
    if not saida:
        if datas_zeradas or (tipo in ("ads", "visitas") and len(linhas) > n + 1):
            # relatório do dia sem investimento (tudo zero): vale como dia enviado, com uma linha zerada de marcação
            marca = {c: (0.0 if c in t["numericos"] else "") for c in t["campos"]}
            marca["Data"] = min(datas_zeradas) if datas_zeradas else None
            marca["Campanha"] = "(sem valor no dia)"
            info["sem_valor"] = True
            return [marca], info
        raise ValueError("a aba “%s” tem o cabeçalho, mas nenhuma linha com valor." % aba)
    return saida, info


# ---------------------------------------------------------------- guardar
def dono(unidade):
    """ADS e afiliados são por CONTA: [raiz]\\_lancamentos\\[Cliente]\\[Conta]."""
    return {"id": unidade["id"], "nome": os.path.join(unidade["cliente_nome"], unidade["conta_nome"]),
            "rotulo": unidade["nome"], "loja": unidade.get("loja") or unidade["conta_nome"]}


def migrar_legado(cfg, cliente):
    """Envios antigos ficavam por cliente ([raiz]\\_lancamentos\\[Cliente]); passam para a conta principal
    (a de mesmo nome do cliente, senão a primeira). Só move, nada é apagado."""
    raiz_cli = os.path.join(comum.raiz(cfg), "_lancamentos", cliente["nome"])
    if not os.path.isdir(raiz_cli):
        return False
    soltos = [f for f in os.listdir(raiz_cli)
              if re.match(r"^(ads|afiliados)_.+_\d{4}-\d{2}\.xlsx$", f) or f in ("_originais", "_anteriores", ".versao_leitor")]
    contas = cliente.get("contas") or []
    if not soltos or not contas:
        return False
    principal = next((c for c in contas if c["nome"] == cliente["nome"]), contas[0])
    destino = os.path.join(raiz_cli, principal["nome"])
    os.makedirs(destino, exist_ok=True)
    for f in soltos:
        o, d = os.path.join(raiz_cli, f), os.path.join(destino, f)
        if os.path.isdir(o) and os.path.isdir(d):
            for g in os.listdir(o):
                if not os.path.exists(os.path.join(d, g)):
                    shutil.move(os.path.join(o, g), os.path.join(d, g))
        elif not os.path.exists(d):
            shutil.move(o, d)
    comum.log("[%s] ADS/afiliados antigos (por cliente) passados para a conta %s" % (cliente["nome"], principal["nome"]))
    return True


def pasta(cfg, cliente):
    return os.path.join(comum.raiz(cfg), "_lancamentos", cliente["nome"])


def caminho(cfg, cliente, tipo, mk, ano, mes):
    return os.path.join(pasta(cfg, cliente), "%s_%s_%04d-%02d.xlsx" % (tipo, mk, ano, mes))


def _fim_do_mes(ano, mes, hoje):
    ult = dt.date(ano, mes, calendar.monthrange(ano, mes)[1])
    ontem = hoje - dt.timedelta(days=1)
    return min(ult, ontem) if (ano, mes) == (ontem.year, ontem.month) else ult


def _gravar(linhas, tipo, destino):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = TIPOS[tipo]["nome"]
    cols = TIPOS[tipo]["campos"] + ([CAMPO_CAMP] if tipo == "ads" else [])
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
    for l in linhas:
        ws.append([l.get(c) for c in cols])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "dd/mm/yyyy"
    tmp = destino + ".novo.xlsx"
    wb.save(tmp)
    os.replace(tmp, destino)


# ---------------------------------------------------------------- ADS Shopee: dia do arquivo + de-para (mass_update)
def _periodo_shopee(nome, caminho=None):
    """ADS da Shopee: o dia vem do nome (…-01_10_2026-01_10_2026) ou da linha "Período" do arquivo.
    A coluna "Data de Início" do relatório é a data em que a campanha começou, não serve.
    Devolve (de, ate) ou None."""
    base = os.path.splitext(os.path.basename(nome))[0]
    m = re.search(r"(\d{1,2})[_./-](\d{1,2})[_./-](\d{4})\D+(\d{1,2})[_./-](\d{1,2})[_./-](\d{4})", base)
    if not m and caminho:
        try:
            for _, linhas in _linhas(caminho):
                for r in linhas[:15]:
                    if r and _chave(r[0]) == "periodo" and len(r) > 1:
                        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\D+(\d{1,2})/(\d{1,2})/(\d{4})", " ".join(str(x) for x in r[1:]))
                        break
                break
        except Exception:
            m = None
    if not m:
        return None
    d1, m1, a1, d2, m2, a2 = map(int, m.groups())
    try:
        return dt.date(a1, m1, d1), dt.date(a2, m2, d2)
    except ValueError:
        return None


def _periodo_nome(nome):
    """Período no nome do arquivo: "..._2026_10_01-2026_10_01.xlsx" (Mercado Livre) ou "01_10_2026-01_10_2026"."""
    base = os.path.splitext(os.path.basename(nome))[0]
    m = re.search(r"(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})\D{1,3}(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})", base) \
        or re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})\D{1,3}(20\d{2})(\d{2})(\d{2})(?!\d)", base)     # Shopee: 20261001_20261001
    try:
        if m:
            a1, m1, d1, a2, m2, d2 = map(int, m.groups())
            return dt.date(a1, m1, d1), dt.date(a2, m2, d2)
        m = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2})\D{1,3}(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2})", base)
        if m:
            d1, m1, a1, d2, m2, a2 = map(int, m.groups())
            return dt.date(a1, m1, d1), dt.date(a2, m2, d2)
    except ValueError:
        return None
    return None


def _dia_forcado(tipo, mk, nome, caminho=None):
    """Dia que vale para TODAS as linhas do arquivo (ignora as datas das linhas), ou None."""
    if tipo == "visitas":
        per = _periodo_nome(nome)
        if per and per[0] != per[1]:
            raise ValueError("o relatório de visitas deve ser subido por dia: este arquivo é de %s a %s. "
                             "Exporte um arquivo para cada dia." % (per[0].strftime("%d/%m/%Y"), per[1].strftime("%d/%m/%Y")))
        return per[0] if per else None
    if (tipo, mk) != ("ads", "shopee"):
        return None
    per = _periodo_shopee(nome, caminho)
    if not per:
        return None
    if per[0] != per[1]:
        raise ValueError("o relatório de ADS da Shopee deve ser subido por dia: este arquivo é de %s a %s. "
                         "Exporte um arquivo para cada dia." % (per[0].strftime("%d/%m/%Y"), per[1].strftime("%d/%m/%Y")))
    return per[0]


# De-para do anúncio do ADS -> SKU, por conta e marketplace (planilha "mass_update" de cada marketplace)
DEPARA = {
    "shopee": {"nome": "mass_update da Shopee (informações básicas)", "id": "ID do Produto", "sku": "SKU de referência"},
    "magalu": {"nome": "mass_update da Magalu (planilha de produtos)", "id": "Código site Magalu", "sku": "SKU"},
}


def tem_depara(mk):
    return mk in DEPARA


def _depara_arq(cfg, dono_, mk="shopee"):
    return os.path.join(pasta(cfg, dono_), "depara_%s.json" % mk)


def ler_depara(mk, caminho):
    """{id do anúncio: {"sku", "nome"[, "variacoes"]}} lido da mass_update do marketplace.

    Shopee: a mass_update de informações básicas traz o SKU de referência (pai) de cada anúncio; a de informações
    de vendas traz uma linha por variação, com o SKU de cada uma. Todas as linhas do mesmo ID do anúncio são juntadas:
    o anúncio com variações aponta para todos os SKUs dele ("290, 289, 291") e fica marcado com "variacoes"."""
    var_col = lambda k: False
    if mk == "shopee":
        ok_id = lambda k: k == "iddoproduto"
        ok_sku = lambda k: k.startswith("skudereferencia") or k == "sku" or k.startswith("skupai") or k == "parentsku"
        var_col = lambda k: k in ("sku", "skudavariacao", "skudavariante", "variationsku", "skuvariacao")
        ok_nome = lambda k: k.startswith("nomedoproduto")
        id_valido = lambda v: v.isdigit()
    elif mk == "magalu":
        ok_id = lambda k: k == "codigositemagalu"
        ok_sku = lambda k: k == "sku"
        ok_nome = lambda k: k.startswith("tituloproduto") or k.startswith("titulodoproduto")
        id_valido = lambda v: bool(v) and " " not in v and not v.lower().startswith("exemplo")
    else:
        raise ValueError("Marketplace sem de-para.")
    melhor = {}
    for _, linhas in _linhas(caminho):
        for n, cab in enumerate(linhas[:15]):
            ks = [_chave(c) for c in cab]
            ji = next((j for j, k in enumerate(ks) if ok_id(k)), None)
            jv = next((j for j, k in enumerate(ks) if var_col(k)), None)              # SKU da variação (se houver)
            jp = next((j for j, k in enumerate(ks) if ok_sku(k) and j != jv), None)   # SKU de referência / pai
            if ji is None or (jv is None and jp is None):
                continue
            jn = next((j for j, k in enumerate(ks) if ok_nome(k)), None)
            tem_var = jv is not None and (jp is not None or any(k.startswith("iddavaria") or k.startswith("nomedavaria")
                                                               or k.startswith("variationid") for k in ks))
            cel = lambda r, j: _texto(r[j]) if j is not None and j < len(r) else ""
            saida = {}
            for r in linhas[n + 1:]:
                pid = cel(r, ji)
                if not id_valido(pid):
                    continue
                d = saida.setdefault(pid, {"sku": [], "nome": cel(r, jn)[:120]})
                if not d["nome"]:
                    d["nome"] = cel(r, jn)[:120]
                sku = cel(r, jv) if tem_var else ""
                sku = sku or cel(r, jp) or (cel(r, jv) if not tem_var else "")
                for x in skus_lista(sku):
                    if x not in d["sku"]:
                        d["sku"].append(x)
            for d in saida.values():
                if tem_var and len(d["sku"]) > 1:
                    d["variacoes"] = True
                d["sku"] = ", ".join(d["sku"])
            if len(saida) > len(melhor) or (len(saida) == len(melhor) and saida and jn is not None):
                melhor = saida
            break
    if melhor:
        return melhor
    d = DEPARA[mk]
    raise ValueError("não achei as colunas “%s” e “%s”. Use a %s." % (d["id"], d["sku"], d["nome"]))


def ler_depara_shopee(caminho):
    return ler_depara("shopee", caminho)


def guardar_depara(cfg, dono_, nome_original, conteudo, mk="shopee"):
    """Guarda o de-para da conta. Junta com o anterior: anúncios que saíram continuam (para os dias antigos)."""
    import json
    base = pasta(cfg, dono_)
    os.makedirs(os.path.join(base, "_originais"), exist_ok=True)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    nome_seguro = re.sub(r'[\\/:*?"<>|]+', "-", os.path.basename(nome_original))
    original = os.path.join(base, "_originais", "%s_depara_%s_%s" % (carimbo, mk, nome_seguro))
    with open(original, "wb") as f:
        f.write(conteudo)
    try:
        novo = ler_depara(mk, original)
    except ValueError:
        os.replace(original, original + ".recusado")
        raise
    atual = carregar_depara(cfg, dono_, mk, com_manual=False)
    # a planilha de informações básicas (só o SKU pai) não desfaz as variações já lidas da planilha de vendas
    for k, v in novo.items():
        if atual.get(k, {}).get("variacoes") and not v.get("variacoes"):
            v.update(sku=atual[k]["sku"], variacoes=True)
    novos = sum(1 for k in novo if k not in atual)
    mudaram = sum(1 for k, v in novo.items() if k in atual and atual[k].get("sku") != v["sku"])
    atual.update(novo)
    arq = _depara_arq(cfg, dono_, mk)
    with open(arq + ".novo", "w", encoding="utf-8") as f:
        json.dump({"atualizado": dt.datetime.now().isoformat(timespec="seconds"), "origem": os.path.basename(nome_original),
                   "anuncios": atual}, f, ensure_ascii=False)
    os.replace(arq + ".novo", arq)
    info = {"no_arquivo": len(novo), "total": len(atual), "novos": novos, "mudaram": mudaram,
            "sem_sku": sum(1 for v in novo.values() if not v["sku"]),
            "com_variacoes": sum(1 for v in novo.values() if v.get("variacoes"))}
    comum.log("[%s] de-para %s: %d anúncios no arquivo (%d novos, %d com SKU trocado), %d no total"
              % (dono_["nome"], mk, info["no_arquivo"], novos, mudaram, info["total"]))
    return info


def carregar_depara(cfg, dono_, mk="shopee", com_manual=True):
    """{id_anuncio: {"sku", "nome"}} da conta ({} se não houver). O que foi preenchido no app vale por cima."""
    import json
    try:
        with open(_depara_arq(cfg, dono_, mk), "r", encoding="utf-8") as f:
            base = json.load(f).get("anuncios", {})
    except (OSError, ValueError):
        base = {}
    if com_manual:
        for k, v in depara_manual(cfg, dono_, mk).items():
            if v.get("sku"):
                base[k] = {"sku": v["sku"], "nome": v.get("nome") or (base.get(k) or {}).get("nome", ""), "manual": True}
    return base


def _depara_manual_arq(cfg, dono_, mk):
    return os.path.join(pasta(cfg, dono_), "depara_manual_%s.json" % mk)


def depara_manual(cfg, dono_, mk):
    """De-para preenchido no app: {anúncio (ou nome do produto, no TikTok): {"sku", "nome", "em", "por"}}."""
    import json
    try:
        with open(_depara_manual_arq(cfg, dono_, mk), "r", encoding="utf-8") as f:
            return json.load(f).get("anuncios", {})
    except (OSError, ValueError):
        return {}


def skus_lista(texto):
    """"290, 289" -> ["290", "289"]: anúncio com variações aponta para mais de um SKU (separados por vírgula ou ;)."""
    vistos = []
    for x in re.split(r"[,;]", str(texto or "")):
        x = x.strip()
        if x and x not in vistos:
            vistos.append(x)
    return vistos


def salvar_depara_manual(cfg, dono_, mk, itens, por=""):
    """itens: {anúncio: {"sku", "nome"}}; SKU vazio tira o anúncio do de-para manual."""
    import json
    atual = depara_manual(cfg, dono_, mk)
    agora = dt.datetime.now().isoformat(timespec="seconds")
    feitos, tirados = 0, 0
    for k, v in (itens or {}).items():
        k = str(k or "").strip()[:200]
        if not k:
            continue
        sku = ", ".join(skus_lista((v or {}).get("sku")))[:300]
        if not sku:
            if atual.pop(k, None) is not None:
                tirados += 1
            continue
        atual[k] = {"sku": sku, "nome": str((v or {}).get("nome") or "")[:200], "em": agora, "por": por}
        feitos += 1
    os.makedirs(pasta(cfg, dono_), exist_ok=True)
    arq = _depara_manual_arq(cfg, dono_, mk)
    with open(arq + ".novo", "w", encoding="utf-8") as f:
        json.dump({"atualizado": agora, "anuncios": atual}, f, ensure_ascii=False, indent=1)
    os.replace(arq + ".novo", arq)
    comum.log("[%s] de-para %s preenchido no app: %d anúncio(s) ligados, %d tirados" % (dono_["nome"], mk, feitos, tirados))
    return {"ligados": feitos, "tirados": tirados, "total": len(atual)}


def info_depara(cfg, dono_, mk="shopee"):
    import json
    try:
        with open(_depara_arq(cfg, dono_, mk), "r", encoding="utf-8") as f:
            d = json.load(f)
        return {"anuncios": len(d.get("anuncios", {})), "atualizado": d.get("atualizado"), "origem": d.get("origem")}
    except (OSError, ValueError):
        return None


# ---------------------------------------------------------------- campanhas (ADS por campanha: Magalu e outros, menos ML)
CAMPO_CAMP = "Campanha (app)"     # id da campanha cadastrada no app (vazio = envio sem campanha)
SEM_CAMPANHAS = {"mercadolivre"}


def aceita_campanhas(mk):
    return mk not in SEM_CAMPANHAS


def _camp_arq(cfg, dono_):
    return os.path.join(pasta(cfg, dono_), "campanhas.json")


def campanhas(cfg, dono_, mk=None):
    import json
    try:
        with open(_camp_arq(cfg, dono_), "r", encoding="utf-8") as f:
            todas = json.load(f)
    except (OSError, ValueError):
        todas = {}
    return todas.get(mk, []) if mk else todas


def editar_campanha(cfg, dono_, mk, acao, nome=None, cid=None):
    """acao: criar | renomear | desativar | ativar. Campanha nunca é apagada (os dados enviados continuam)."""
    import json
    if not aceita_campanhas(mk):
        raise ValueError("Este marketplace não usa campanhas.")
    todas = campanhas(cfg, dono_)
    lista = todas.setdefault(mk, [])
    nome = re.sub(r"\s+", " ", str(nome or "")).strip()[:80]
    if acao in ("criar", "renomear"):
        if not nome:
            raise ValueError("Dê um nome para a campanha.")
        if any(c["nome"].lower() == nome.lower() and c["id"] != cid for c in lista):
            raise ValueError("Já existe uma campanha com esse nome.")
    if acao == "criar":
        n = 1 + max([int(c["id"][1:]) for c in lista if c["id"][1:].isdigit()] or [0])
        c = {"id": "c%d" % n, "nome": nome, "ativa": True, "criada": dt.datetime.now().isoformat(timespec="seconds")}
        lista.append(c)
    else:
        c = next((x for x in lista if x["id"] == cid), None)
        if not c:
            raise ValueError("Campanha não encontrada.")
        if acao == "renomear":
            c["nome"] = nome
        elif acao in ("desativar", "ativar"):
            c["ativa"] = acao == "ativar"
        else:
            raise ValueError("Ação inválida.")
    os.makedirs(pasta(cfg, dono_), exist_ok=True)
    with open(_camp_arq(cfg, dono_) + ".novo", "w", encoding="utf-8") as f:
        json.dump(todas, f, ensure_ascii=False, indent=1)
    os.replace(_camp_arq(cfg, dono_) + ".novo", _camp_arq(cfg, dono_))
    return c


# ---------------------------------------------------------------- OK manual (dia sem investimento / sem afiliado)
def _ok_arq(cfg, dono_):
    return os.path.join(pasta(cfg, dono_), "ok_manual.json")


def _ok_chave(tipo, mk, campanha=None):
    return "%s|%s|%s" % (tipo, mk, campanha or "")


def oks_manuais(cfg, dono_):
    """{"tipo|mk|campanha": {"AAAA-MM-DD": {"motivo", "quando"}}}"""
    import json
    try:
        with open(_ok_arq(cfg, dono_), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def marcar_ok(cfg, dono_, tipo, mk, dias, motivo="", campanha=None, remover=False):
    """Dá (ou tira) o OK manual nos dias: o dia conta como enviado mesmo sem arquivo."""
    import json
    if tipo not in TIPOS:
        raise ValueError("Tipo inválido.")
    todos = oks_manuais(cfg, dono_)
    k = _ok_chave(tipo, mk, campanha)
    atual = todos.setdefault(k, {})
    feitos = []
    for d in dias or []:
        d = dt.date.fromisoformat(str(d)[:10]).isoformat()
        if remover:
            if atual.pop(d, None) is not None:
                feitos.append(d)
        else:
            atual[d] = {"motivo": (motivo or "OK manual").strip()[:200], "quando": dt.datetime.now().isoformat(timespec="seconds")}
            feitos.append(d)
    if not atual:
        todos.pop(k, None)
    os.makedirs(pasta(cfg, dono_), exist_ok=True)
    with open(_ok_arq(cfg, dono_) + ".novo", "w", encoding="utf-8") as f:
        json.dump(todos, f, ensure_ascii=False, indent=1)
    os.replace(_ok_arq(cfg, dono_) + ".novo", _ok_arq(cfg, dono_))
    comum.log("[%s] OK manual %s %s%s: %s %s" % (dono_["nome"], tipo, mk, (" (" + campanha + ")") if campanha else "",
                                               "removido de" if remover else "dado em", ", ".join(feitos)))
    return feitos


def _aplicar_ok(r, oks_k, ano, mes):
    pref = "%04d-%02d" % (ano, mes)
    meus = sorted(d for d in (oks_k or {}) if d.startswith(pref))
    r["ok_manual"] = [{"dia": d, "motivo": oks_k[d].get("motivo", "")} for d in meus]
    r["faltam"] = [d for d in r.get("faltam", []) if d not in set(meus)]
    return r


# ---------------------------------------------------------------- campanhas fechadas (Shopee) e ajustes manuais
DETALHE = "detalhe:"     # CAMPO_CAMP das linhas do relatório de dentro de uma campanha fechada
AJUSTE = "ajuste:"       # CAMPO_CAMP das linhas de ajuste manual (campanha apagada)


def eh_fechada(mk, linha):
    """Linha de campanha fechada: sem anúncio/produto, só com o nome da campanha (ex.: "GMV Max da Loja" na Shopee)."""
    if mk != "shopee":
        return False
    tag = linha.get(CAMPO_CAMP) or ""
    an = str(linha.get("Id anúncio") or "")
    return (not tag.startswith((DETALHE, AJUSTE)) and not an.isdigit() and not str(linha.get("SKU") or "").strip()
            and bool(linha.get("Campanha") or an))


def nome_fechada(linha):
    return str(linha.get("Campanha") or linha.get("Id anúncio") or "").strip()


def _eh_total_detalhe(l):
    """Linha de TOTAL que vem dentro do relatório de detalhe (sem anúncio e sem SKU): é a soma dos anúncios, não conta."""
    return (l.get(CAMPO_CAMP) or "").startswith(DETALHE) and not str(l.get("Id anúncio") or "").strip() \
        and not str(l.get("SKU") or "").strip()


def efetivas_ads(mk, linhas):
    """Linhas que valem. Campanha fechada com o detalhe (por anúncio) subido no dia: o gasto passa a vir do detalhe,
    e o valor da campanha no dia é o TETO — detalhe acima da campanha é reduzido na proporção; abaixo, a diferença
    continua na linha da campanha. Assim o total do dia nunca passa nem fica abaixo do valor da campanha.
    A linha de total que vem dentro do relatório de detalhe é descartada (senão o gasto contaria 2 vezes)."""
    if not any((l.get(CAMPO_CAMP) or "").startswith(DETALHE) for l in linhas):
        return linhas
    det, camp, out = {}, {}, []
    for l in linhas:
        if _eh_total_detalhe(l):
            continue
        tag = l.get(CAMPO_CAMP) or ""
        if tag.startswith(DETALHE):
            det.setdefault((l.get("Data"), tag[len(DETALHE):]), []).append(l)
        elif eh_fechada(mk, l):
            camp.setdefault((l.get("Data"), nome_fechada(l)), []).append(l)
            out.append(l)
        else:
            out.append(l)
    if not det:
        return out
    tirar = set()
    for k, ds in det.items():
        cs = camp.get(k)
        if not cs:
            out.extend(ds)                 # sem a linha da campanha no dia: vale o detalhe
            continue
        tirar.update(id(c) for c in cs)
        # cada métrica somável (gasto, cliques, receita) respeita o total da campanha separadamente:
        # detalhe acima -> reduzido na proporção; abaixo -> a diferença fica numa linha da campanha ("não detalhado")
        novos = [dict(l) for l in ds]
        resto = dict(cs[0], Gasto=0.0, Cliques=0.0, **{"Receita ADS": 0.0})
        tem_resto = False
        for c_ in ("Gasto", "Cliques", "Receita ADS"):
            sc = sum(c.get(c_) or 0 for c in cs)
            sd = sum(l.get(c_) or 0 for l in ds)
            if sd > sc + 0.005 and sd > 0:
                f = sc / sd
                for l in novos:
                    l[c_] = round((l.get(c_) or 0) * f, 6)
            elif sc - sd > 0.005:
                resto[c_] = round(sc - sd, 6)
                tem_resto = True
        out.extend(novos)
        if tem_resto:
            out.append(resto)
    return [l for l in out if id(l) not in tirar]


def fechadas(mk, linhas):
    """{campanha: {dias, gasto, com_detalhe, gasto_detalhe, acima}} das campanhas fechadas no mês.
    acima: dias em que o detalhe passou do valor da campanha (foi reduzido ao teto)."""
    out, por_dia_c, por_dia_d = {}, {}, {}
    for l in linhas:
        if eh_fechada(mk, l):
            n = nome_fechada(l)
            f = out.setdefault(n, {"dias": set(), "gasto": 0.0, "com_detalhe": set(), "gasto_detalhe": 0.0})
            f["dias"].add(l["Data"].isoformat())
            f["gasto"] += l.get("Gasto") or 0
            por_dia_c[(n, l["Data"])] = por_dia_c.get((n, l["Data"]), 0.0) + (l.get("Gasto") or 0)
    for l in linhas:
        tag = l.get(CAMPO_CAMP) or ""
        if tag.startswith(DETALHE) and tag[len(DETALHE):] in out and not _eh_total_detalhe(l):
            n = tag[len(DETALHE):]
            f = out[n]
            f["com_detalhe"].add(l["Data"].isoformat())
            f["gasto_detalhe"] += l.get("Gasto") or 0
            por_dia_d[(n, l["Data"])] = por_dia_d.get((n, l["Data"]), 0.0) + (l.get("Gasto") or 0)
    acima = {}
    for (n, d), v in por_dia_d.items():
        if (n, d) in por_dia_c and v > por_dia_c[(n, d)] + 0.05:
            acima.setdefault(n, []).append(d.isoformat())
    return {n: {"dias": sorted(f["dias"]), "gasto": round(f["gasto"], 2), "com_detalhe": sorted(f["com_detalhe"]),
                "gasto_detalhe": round(f["gasto_detalhe"], 2), "acima": sorted(acima.get(n, []))} for n, f in out.items()}


def _aj_arq(cfg, dono_):
    return os.path.join(pasta(cfg, dono_), "ajustes_ads.json")


def ajustes(cfg, dono_):
    import json
    try:
        with open(_aj_arq(cfg, dono_), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return []


def salvar_ajuste(cfg, dono_, mk, nome, gasto, de, ate, receita=0.0, cliques=0.0, motivo="", aid=None, tipo="ads"):
    """Ajuste manual: o valor é dividido igualmente nos dias de/até.
    tipo "ads": campanha apagada que sumiu do relatório (gasto/receita/cliques);
    tipo "afiliados": comissão de afiliados lançada de uma vez, sem relatório (ex.: meses anteriores)."""
    import json, uuid
    tipo = "afiliados" if tipo == "afiliados" else "ads"
    nome = re.sub(r"\s+", " ", str(nome or "")).strip()[:80] or ("Ajuste afiliados" if tipo == "afiliados" else "")
    if not nome:
        raise ValueError("Informe o nome da campanha.")
    d1, d2 = dt.date.fromisoformat(str(de)[:10]), dt.date.fromisoformat(str(ate)[:10])
    if d2 < d1:
        raise ValueError("A data final é antes da inicial.")
    if (d2 - d1).days > 92:
        raise ValueError("Período longo demais (máx. 93 dias).")
    gasto, receita, cliques = abs(float(gasto or 0)), abs(float(receita or 0)), abs(float(cliques or 0))
    if tipo == "afiliados":
        receita = cliques = 0.0
        if not gasto:
            raise ValueError("Informe a comissão de afiliados.")
    if not gasto and not receita and not cliques:
        raise ValueError("Informe o gasto (ou receita/cliques).")
    lista = [a for a in ajustes(cfg, dono_) if a["id"] != aid]
    item = {"id": aid or uuid.uuid4().hex[:10], "tipo": tipo, "mk": mk, "nome": nome, "gasto": round(gasto, 2), "receita": round(receita, 2),
            "cliques": round(cliques, 0), "de": d1.isoformat(), "ate": d2.isoformat(), "motivo": (motivo or "").strip()[:200],
            "criado": dt.datetime.now().isoformat(timespec="seconds")}
    lista.append(item)
    os.makedirs(pasta(cfg, dono_), exist_ok=True)
    with open(_aj_arq(cfg, dono_) + ".novo", "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=1)
    os.replace(_aj_arq(cfg, dono_) + ".novo", _aj_arq(cfg, dono_))
    comum.log("[%s] ajuste manual de %s %s: %s R$ %.2f de %s a %s" % (dono_["nome"], "afiliados" if tipo == "afiliados" else "ADS",
                                                                    mk, nome, gasto, d1, d2))
    return item


def remover_ajuste(cfg, dono_, aid):
    import json
    lista = ajustes(cfg, dono_)
    item = next((a for a in lista if a["id"] == aid), None)
    if not item:
        raise ValueError("Ajuste não encontrado.")
    lista = [a for a in lista if a["id"] != aid]
    with open(_aj_arq(cfg, dono_) + ".novo", "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=1)
    os.replace(_aj_arq(cfg, dono_) + ".novo", _aj_arq(cfg, dono_))
    comum.log("[%s] ajuste manual de ADS removido: %s" % (dono_["nome"], item["nome"]))
    return item


def _dividir(total, n, casas=2):
    """Divide igualmente em n partes, o resto vai no último dia (a soma bate exatamente)."""
    if n <= 0:
        return []
    parte = round(total / n, casas)
    partes = [parte] * (n - 1)
    partes.append(round(total - parte * (n - 1), casas))
    return partes


def linhas_ajuste(cfg, dono_, ano, mes):
    """{mk: [linhas de ADS]} dos ajustes manuais, só os dias do mês pedido."""
    saida = {}
    for a in ajustes(cfg, dono_):
        if a.get("tipo") == "afiliados":
            continue
        d1, d2 = dt.date.fromisoformat(a["de"]), dt.date.fromisoformat(a["ate"])
        dias = [d1 + dt.timedelta(days=i) for i in range((d2 - d1).days + 1)]
        n = len(dias)
        g, r, c = _dividir(a["gasto"], n), _dividir(a["receita"], n), _dividir(a["cliques"], n, 0)
        for i, d in enumerate(dias):
            if (d.year, d.month) != (ano, mes):
                continue
            saida.setdefault(a["mk"], []).append({"Data": d, "Id anúncio": a["nome"], "SKU": a["nome"], "Campanha": a["nome"],
                                                  "Gasto": g[i], "Cliques": c[i], "Receita ADS": r[i], CAMPO_CAMP: AJUSTE + a["id"]})
    return saida


def linhas_ajuste_afiliados(cfg, dono_, ano, mes):
    """{mk: [linhas de afiliados]} dos ajustes manuais de afiliados (comissão sem relatório), só os dias do mês pedido."""
    saida = {}
    for a in ajustes(cfg, dono_):
        if a.get("tipo") != "afiliados":
            continue
        d1, d2 = dt.date.fromisoformat(a["de"]), dt.date.fromisoformat(a["ate"])
        dias = [d1 + dt.timedelta(days=i) for i in range((d2 - d1).days + 1)]
        partes = _dividir(a["gasto"], len(dias))
        for d, v in zip(dias, partes):
            if (d.year, d.month) == (ano, mes) and v:
                saida.setdefault(a["mk"], []).append({"Data": d, "Id do pedido": "", "Id anúncio": a["nome"], "SKU": a["nome"],
                                                      "Comissão afiliado": v, "ajuste": a["id"]})
    return saida


def _aplicar_tag(linhas, campanha, info=None):
    """CAMPO_CAMP de cada linha: a campanha/detalhe escolhido no envio. O relatório por grupo de anúncio da Shopee já
    traz a sua (linha da campanha = base do dia; anúncios = detalhe da campanha) e ignora a escolha."""
    grupo = any(l.get("_grupo") for l in linhas)
    for l in linhas:
        l[CAMPO_CAMP] = l.get("_tag", "") if grupo else (campanha or "")
    if grupo and campanha and info is not None and campanha not in {l.get("_tag") for l in linhas}:
        info["aviso"] = (info.get("aviso", "") + " Este relatório traz todas as campanhas da conta: a campanha/detalhe "
                         "escolhido no envio foi ignorado.").strip()


def _tirar_cobertas(antigas, novas, grupo):
    """ADS Shopee — evita contar o mesmo investimento por dois relatórios do mesmo dia.
    - Chegou o relatório por GRUPO de anúncio: sai do dia o que ele cobre (linha de campanha com o mesmo nome e
      anúncio com o mesmo ID do produto, de qualquer relatório anterior); o que ele não traz fica.
    - Chegou o relatório geral (antigo, por anúncio): ele substitui a base do dia (regra de sempre) e também o detalhe
      que ficaria órfão (campanha que não está no novo arquivo) cujos anúncios estão no novo arquivo — senão o
      detalhe valeria sozinho e somaria de novo o que o novo arquivo já trouxe."""
    dias = {l["Data"] for l in novas}
    if not dias:
        return antigas
    nomes, ids = {}, {}
    for l in novas:
        d, tag = l["Data"], l.get(CAMPO_CAMP) or ""
        if tag.startswith(DETALHE):
            nomes.setdefault(d, set()).add(tag[len(DETALHE):])
        elif eh_fechada("shopee", l):
            nomes.setdefault(d, set()).add(nome_fechada(l))
        an = str(l.get("Id anúncio") or "")
        if an.isdigit() and not tag.startswith(AJUSTE) and not l.get("_so_nome"):
            ids.setdefault(d, set()).add(an)
    if not grupo:
        base_nomes = {}
        for l in novas:
            if not (l.get(CAMPO_CAMP) or "") and eh_fechada("shopee", l):
                base_nomes.setdefault(l["Data"], set()).add(nome_fechada(l))
        if not any(not (l.get(CAMPO_CAMP) or "") for l in novas):
            return antigas                       # só detalhe/campanha do app: regra de sempre
        tirar_det = set()
        det_ids = {}
        for l in antigas:
            tag = l.get(CAMPO_CAMP) or ""
            if l.get("Data") in dias and tag.startswith(DETALHE):
                det_ids.setdefault((l["Data"], tag), set()).add(str(l.get("Id anúncio") or ""))
        for (d, tag), ids_d in det_ids.items():
            if tag[len(DETALHE):] not in base_nomes.get(d, set()) and (ids_d & ids.get(d, set())):
                tirar_det.add((d, tag))
        return [l for l in antigas if (l.get("Data"), l.get(CAMPO_CAMP) or "") not in tirar_det]
    out = []
    for l in antigas:
        d, tag = l.get("Data"), l.get(CAMPO_CAMP) or ""
        if d in dias and not tag:
            if eh_fechada("shopee", l) and nome_fechada(l) in nomes.get(d, set()):
                continue
            if str(l.get("Id anúncio") or "").isdigit() and str(l["Id anúncio"]) in ids.get(d, set()):
                continue
        out.append(l)
    return out


def _juntar(tipo, mk, antigas, novas):
    """Junta o que já estava guardado no mês com o que chegou.
    Diário: os dias que chegaram substituem os mesmos dias. Relatório único: os pedidos que chegaram
    substituem os mesmos pedidos (linhas sem pedido seguem a regra do dia)."""
    if unico(tipo, mk):
        pedidos = {l.get("Id do pedido") for l in novas if l.get("Id do pedido")}
        dias_sem = {l["Data"] for l in novas if not l.get("Id do pedido")}
        mantidas = [l for l in antigas if (l.get("Id do pedido") and l["Id do pedido"] not in pedidos)
                    or (not l.get("Id do pedido") and l.get("Data") not in dias_sem)]
    else:
        # diário: os dias que chegaram substituem os mesmos dias DA MESMA CAMPANHA (cada campanha forma uma parte do dia)
        grupo = any(l.get("_grupo") for l in novas)
        chaves = {(l["Data"], l.get(CAMPO_CAMP) or "") for l in novas if not (grupo and not (l.get(CAMPO_CAMP) or ""))}
        if (tipo, mk) == ("ads", "shopee"):
            antigas = _tirar_cobertas(antigas, novas, grupo)
        mantidas = [l for l in antigas if (l.get("Data"), l.get(CAMPO_CAMP) or "") not in chaves]
    return sorted(mantidas + novas, key=lambda l: (l["Data"], l.get("Id anúncio") or "", l.get("Id do pedido") or ""))


def dia_do_nome(nome, ano, mes):
    """Dia do relatório pelo nome do arquivo. Aceita data completa (2026-10-05, 05-10-2026, 05.10.2026, 05_10)
    ou só o dia (05.xlsx, "ads 5.csv") — nesse caso usa o mês/ano escolhido. Devolve date ou None."""
    base = os.path.splitext(os.path.basename(nome))[0]
    m = re.search(r"(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](20\d{2})(?!\d)", base)     # 01_10_2026 (Shopee), 05-10-2026
    m_ano = re.search(r"(20\d{2})[-_.](\d{1,2})[-_.](\d{1,2})", base) or re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)", base)
    if m and m_ano and m_ano.start() < m.start():
        m = None                       # "2026_10_01-..." é ano-mês-dia, não "10_01-2026"
    if m:
        d, mm, a = map(int, m.groups())
        m2 = None
    else:
        m2 = m_ano
    if m:
        pass
    elif m2:
        a, mm, d = map(int, m2.groups())
    else:
        m = re.search(r"(?<!\d)(\d{1,2})[-_.](\d{1,2})(?:[-_.](\d{2,4}))?(?!\d)", base)
        if m and 1 <= int(m.group(2)) <= 12:
            d, mm = int(m.group(1)), int(m.group(2))
            a = int(m.group(3)) if m.group(3) else ano
            a = a + 2000 if a < 100 else a
        else:
            nums = [int(x) for x in re.findall(r"(?<!\d)(\d{1,2})(?!\d)", base)]
            nums = [n for n in nums if 1 <= n <= 31]
            if not nums:
                return None
            d, mm, a = nums[0], mes, ano
    try:
        return dt.date(a, mm, d)
    except ValueError:
        return None


def tirar_total_campanha(tipo, mk, linhas, info=None):
    """ADS da Shopee (relatório de dentro da campanha): a primeira linha, sem anúncio e sem SKU, é o TOTAL da campanha
    (soma das outras linhas). Ela é desconsiderada, senão o gasto contaria 2 vezes."""
    if tipo != "ads" or mk != "shopee" or len(linhas) < 2 or any(l.get("_grupo") for l in linhas):
        return linhas
    tem_id = lambda l: str(l.get("Id anúncio") or "").strip() or str(l.get("SKU") or "").strip()
    sem = [l for l in linhas if not tem_id(l) and (l.get("Gasto") or l.get("Cliques") or l.get("Receita ADS"))]
    com = [l for l in linhas if tem_id(l)]
    if not com or not sem:
        return linhas
    tirar = []
    por_dia = {}
    for l in com:
        k = l.get("Data")
        d = por_dia.setdefault(k, [0.0, 0.0])
        d[0] += l.get("Gasto") or 0
        d[1] += l.get("Receita ADS") or 0
    for l in sem:
        g, r = por_dia.get(l.get("Data"), [None, None])
        if g is None and len(por_dia) == 1:
            g, r = list(por_dia.values())[0]
        if g is not None and abs((l.get("Gasto") or 0) - g) <= max(0.05, 0.005 * g) and abs((l.get("Receita ADS") or 0) - r) <= max(0.05, 0.005 * r):
            tirar.append(l)
    if tirar and info is not None:
        info["total_campanha_ignorado"] = len(tirar)
        info["aviso"] = (info.get("aviso", "") + " A linha de total da campanha (soma das outras linhas) foi desconsiderada.").strip()
    return [l for l in linhas if not any(l is t for t in tirar)]


def _so_do_mes(linhas, mes_alvo, info):
    """Só as linhas do mês escolhido na tela (AAAA-MM): as de outros meses ficam de fora e são avisadas."""
    if not mes_alvo:
        return linhas
    a, m = int(mes_alvo[:4]), int(mes_alvo[5:7])
    dentro = [l for l in linhas if l.get("Data") and (l["Data"].year, l["Data"].month) == (a, m)]
    fora = [l for l in linhas if l not in dentro]
    if fora:
        dias = sorted({l["Data"] for l in fora if l.get("Data")})
        info["fora_do_mes"] = len(fora)
        txt = "%d linha(s) de fora de %02d/%04d ficaram de fora (%s)." % (
            len(fora), m, a, ", ".join(d.strftime("%d/%m/%Y") for d in dias[:6]) + (" …" if len(dias) > 6 else ""))
        info["aviso"] = (info.get("aviso", "") + " " + txt).strip()
    if not dentro:
        raise ValueError("o arquivo não tem nenhuma linha de %02d/%04d (o mês escolhido na tela)%s." % (
            m, a, (": " + info["aviso"].rstrip(".")) if info.get("aviso") else ""))
    return dentro


def guardar(cfg, cliente, tipo, mk, nome_original, conteudo, dia_escolhido=None, hoje=None, campanha=None, mes_alvo=None):
    """Lê e junta ao que já existe no mês, DIA A DIA: os dias que vêm no arquivo substituem os mesmos dias
    já guardados; os outros dias do mês continuam. Linhas sem data vão para dia_escolhido (padrão: ontem)."""
    hoje = hoje or dt.date.today()
    if tipo not in TIPOS:
        raise ValueError("Tipo inválido.")
    base = pasta(cfg, cliente)
    os.makedirs(os.path.join(base, "_originais"), exist_ok=True)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    nome_seguro = re.sub(r'[\\/:*?"<>|]+', "-", os.path.basename(nome_original))
    original = os.path.join(base, "_originais", "%s_%s_%s_%s" % (carimbo, tipo, mk, nome_seguro))
    with open(original, "wb") as f:
        f.write(conteudo)
    try:
        linhas, info = ler(tipo, original)
        linhas = tirar_total_campanha(tipo, mk, linhas, info)
        forcado = _dia_forcado(tipo, mk, nome_original, original)
    except ValueError:
        os.replace(original, original + ".recusado")
        raise
    if forcado:
        for l in linhas:
            l["Data"] = forcado
        info["dia_do_arquivo"] = forcado.isoformat()
    elif (tipo, mk) == ("ads", "shopee"):
        for l in linhas:          # sem período no nome/arquivo: vale o dia escolhido no envio
            l["Data"] = None
    if isinstance(dia_escolhido, str) and dia_escolhido:
        dia_escolhido = dt.date.fromisoformat(dia_escolhido[:10])
    dia_escolhido = dia_escolhido or (hoje - dt.timedelta(days=1))
    _indice_salvar(base, os.path.basename(original), dia_escolhido, campanha)
    if tipo == "ads":
        _aplicar_tag(linhas, campanha, info)
    info["sem_data"] = 0
    for l in linhas:
        if l.get("Data") is None:
            l["Data"] = dia_escolhido
            info["sem_data"] += 1
    try:
        linhas = _so_do_mes(linhas, mes_alvo, info)
    except ValueError:
        os.replace(original, original + ".recusado")
        raise
    por_mes = {}
    for l in linhas:
        por_mes.setdefault((l["Data"].year, l["Data"].month), []).append(l)
    res = []
    for (a, m), novas in sorted(por_mes.items()):
        destino = caminho(cfg, cliente, tipo, mk, a, m)
        dias_novos = sorted({l["Data"] for l in novas})
        antigas = []
        if os.path.isfile(destino):
            antigas = _ler_padrao(destino, tipo)
            ant = os.path.join(base, "_anteriores", os.path.basename(destino).replace(".xlsx", "_%s.xlsx" % carimbo))
            os.makedirs(os.path.dirname(ant), exist_ok=True)
            shutil.copy2(destino, ant)
        juntas = _juntar(tipo, mk, antigas, novas)
        _gravar(juntas, tipo, destino)
        ef = (lambda x: efetivas_ads(mk, x)) if tipo == "ads" else (lambda x: x)   # sem somar campanha + detalhe
        r = dict(resumo(ef(juntas), tipo), mes="%04d-%02d" % (a, m), arquivo=destino,
                 dias_enviados=[d.isoformat() for d in dias_novos], linhas_enviadas=len(novas))
        r["enviado"] = resumo(ef(novas), tipo)
        if unico(tipo, mk):
            r["pedidos_enviados"] = len({l.get("Id do pedido") for l in novas if l.get("Id do pedido")})
        res.append(r)
    comum.registrar_evento({"tipo": "lancamento", "subtipo": tipo, "cliente": cliente["id"], "mk": mk,
                            "origem": os.path.basename(nome_original),
                            "dias": [d for r in res for d in r["dias_enviados"]]})
    comum.log("[%s] %s %s subido: %s" % (cliente["nome"], TIPOS[tipo]["nome"], mk,
                                        ", ".join("%s (%d linhas)" % (",".join(r["dias_enviados"]), r["linhas_enviadas"]) for r in res)))
    return res, info


# ---------------------------------------------------------------- visitas totais do dia (Mercado Livre > Métricas > Negócio)
_MESES_EXT = {"janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
              "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}


def ler_visitas_total(caminho):
    """Relatório de evolução do negócio do Mercado Livre (Relatorio_evolucao_negocio_...xlsx): uma linha por dia com as
    VISITAS TOTAIS da conta (e a quantidade de vendas). Devolve ({dia: {"visitas", "vendas"}}, info) ou None se não for ele.
    Dias depois do fim do período do relatório ("... até 26 de outubro de 2026") ficam de fora: o dia de hoje vem incompleto."""
    try:
        abas = _linhas(caminho)
    except Exception:
        return None
    for aba, linhas in abas:
        for n, cab in enumerate(linhas[:15]):
            ks = [_chave(c) for c in cab]
            if not ("data" in ks and "visitas" in ks and "compradoresunicos" in ks):
                continue
            jd, jv = ks.index("data"), ks.index("visitas")
            jq = ks.index("quantidadedevendas") if "quantidadedevendas" in ks else None
            de = ate = None
            for r in linhas[:n]:
                txt = unicodedata.normalize("NFKD", " ".join(str(x) for x in r if x)).encode("ascii", "ignore").decode().lower()
                ds = re.findall(r"(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{4})", txt)
                ds = [dt.date(int(a), _MESES_EXT[mm], int(d)) for d, mm, a in ds if mm in _MESES_EXT]
                if len(ds) >= 2:
                    de, ate = ds[0], ds[-1]
                    break
            if not ate:
                per = _periodo_nome(os.path.basename(caminho))
                if per:
                    de, ate = per
            dias, fora = {}, 0
            for r in linhas[n + 1:]:
                d = _data(r[jd] if jd < len(r) else None)
                if not d:
                    continue
                if ate and d > ate:
                    fora += 1
                    continue
                v, _ = _numero(r[jv] if jv < len(r) else None)
                q, _ = _numero(r[jq] if jq is not None and jq < len(r) else None)
                dias[d] = {"visitas": float(v or 0), "vendas": float(q or 0)}
            if not dias:
                raise ValueError("o relatório de evolução do negócio não tem nenhum dia.")
            return dias, {"aba": aba, "colunas": ["Data", "Visitas"] + (["Quantidade de vendas"] if jq is not None else []),
                          "formato": "Mercado Livre (visitas totais por dia)", "total": True, "sem_data": 0,
                          "ignoradas_invalidas": 0, "dias_depois_do_periodo": fora,
                          "periodo": [de.isoformat(), ate.isoformat()] if ate else None}
    return None


def _total_arq(cfg, dono_, mk):
    return os.path.join(pasta(cfg, dono_), "visitas_total_%s.json" % mk)


def _total_ler(cfg, dono_, mk):
    import json
    try:
        with open(_total_arq(cfg, dono_, mk), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"dias": {}}


def _total_gravar(cfg, dono_, mk, dados):
    import json
    os.makedirs(pasta(cfg, dono_), exist_ok=True)
    arq = _total_arq(cfg, dono_, mk)
    with open(arq + ".novo", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
    os.replace(arq + ".novo", arq)


def visitas_total(cfg, dono_, mk, ano, mes):
    """{date: {"visitas", "vendas"}} do mês, das visitas totais do dia subidas."""
    pre = "%04d-%02d" % (ano, mes)
    return {dt.date.fromisoformat(k): v for k, v in (_total_ler(cfg, dono_, mk).get("dias") or {}).items() if k.startswith(pre)}


def guardar_visitas_total(cfg, cliente, mk, nome_original, conteudo, mes_alvo=None):
    """Sobe o relatório com as visitas totais de cada dia: os dias que vêm substituem os mesmos dias; os outros continuam."""
    base = pasta(cfg, cliente)
    os.makedirs(os.path.join(base, "_originais"), exist_ok=True)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    nome_seguro = re.sub(r'[\\/:*?"<>|]+', "-", os.path.basename(nome_original))
    original = os.path.join(base, "_originais", "%s_visitastotal_%s_%s" % (carimbo, mk, nome_seguro))
    with open(original, "wb") as f:
        f.write(conteudo)
    try:
        lido = ler_visitas_total(original)
        if lido is None:
            raise ValueError("não é o relatório de evolução do negócio (Data, Visitas, Compradores únicos).")
    except ValueError:
        os.replace(original, original + ".recusado")
        raise
    dias, info = lido
    if mes_alvo:
        a_, m_ = int(mes_alvo[:4]), int(mes_alvo[5:7])
        fora = sorted(d for d in dias if (d.year, d.month) != (a_, m_))
        dias = {d: v for d, v in dias.items() if (d.year, d.month) == (a_, m_)}
        if fora:
            info["fora_do_mes"] = len(fora)
            info["aviso"] = "%d dia(s) de fora de %02d/%04d ficaram de fora." % (len(fora), m_, a_)
        if not dias:
            os.replace(original, original + ".recusado")
            raise ValueError("o relatório não tem nenhum dia de %02d/%04d (o mês escolhido na tela)." % (m_, a_))
    dados = _total_ler(cfg, cliente, mk)
    agora = dt.datetime.now().isoformat(timespec="seconds")
    for d, v in dias.items():
        dados.setdefault("dias", {})[d.isoformat()] = dict(v, em=agora)
    dados["atualizado"] = agora
    dados["arquivo"] = os.path.basename(nome_original)
    _total_gravar(cfg, cliente, mk, dados)
    por_mes = {}
    for d in sorted(dias):
        por_mes.setdefault((d.year, d.month), []).append(d)
    res = []
    for (a, m), ds in sorted(por_mes.items()):
        res.append({"mes": "%04d-%02d" % (a, m), "dias_enviados": [d.isoformat() for d in ds], "linhas_enviadas": len(ds),
                    "total": True, "enviado": {"Visitas": sum(dias[d]["visitas"] for d in ds), "Vendas": sum(dias[d]["vendas"] for d in ds)}})
    comum.registrar_evento({"tipo": "lancamento", "subtipo": "visitas_total", "cliente": cliente["id"], "mk": mk,
                            "origem": os.path.basename(nome_original), "dias": [d for r in res for d in r["dias_enviados"]]})
    comum.log("[%s] visitas totais %s subidas: %d dia(s)" % (cliente["nome"], mk, len(dias)))
    return res, info


def situacao_totais(cfg, dono_, ano, mes):
    """Para o painel: {mk: {"dias", "ate", "visitas", "vendas", "atualizado"}} das visitas totais do mês."""
    saida = {}
    for p in glob.glob(os.path.join(pasta(cfg, dono_), "visitas_total_*.json")):
        mk = os.path.basename(p)[len("visitas_total_"):-len(".json")]
        t = visitas_total(cfg, dono_, mk, ano, mes)
        if t:
            ds = sorted(t)
            saida[mk] = {"dias": [d.isoformat() for d in ds], "ate": ds[-1].isoformat(),
                         "visitas": sum(v["visitas"] for v in t.values()), "vendas": sum(v["vendas"] for v in t.values()),
                         "atualizado": _total_ler(cfg, dono_, mk).get("atualizado")}
    return saida


def _limpar_total(cfg, dono_, mk, chave_mes, dias):
    dados = _total_ler(cfg, dono_, mk)
    antes = len(dados.get("dias") or {})
    dados["dias"] = {k: v for k, v in (dados.get("dias") or {}).items()
                     if not (k.startswith(chave_mes) and (dias is None or k in dias))}
    tirados = antes - len(dados["dias"])
    if tirados:
        _total_gravar(cfg, dono_, mk, dados)
    return tirados


LEITOR_VERSAO = 17  # sobe quando a leitura dos arquivos muda: os enviados são relidos (reprocessar)


def _indice_arq(base):
    return os.path.join(base, "_originais", "indice.json")


def _indice_ler(base):
    import json
    try:
        with open(_indice_arq(base), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _indice_salvar(base, arquivo, dia, campanha=None):
    import json
    idx = _indice_ler(base)
    idx[arquivo] = {"dia": dia.isoformat() if dia else None, "campanha": campanha or ""}
    os.makedirs(os.path.dirname(_indice_arq(base)), exist_ok=True)
    with open(_indice_arq(base), "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=0)


def versao_dados(cfg, cliente):
    try:
        with open(os.path.join(pasta(cfg, cliente), ".versao_leitor"), "r") as f:
            return int(f.read().strip() or 0)
    except (OSError, ValueError):
        return 0


def _recuperar_recusados(base, idx):
    """Arquivos de ADS recusados antes por estarem zerados (dia sem investimento) voltam a valer.
    A campanha vem do envio aceito mais próximo no mesmo lote (até 20 s de diferença)."""
    orig = os.path.join(base, "_originais")
    padrao = re.compile(r"^(\d{8}-\d{6})_(ads)_([a-z0-9-]+?)_(.+)\.recusado$")
    aceitos = []
    for f, ent in idx.items():
        m = re.match(r"^(\d{8}-\d{6})_ads_([a-z0-9-]+?)_", f)
        if m and os.path.isfile(os.path.join(orig, f)):
            camp = ent.get("campanha", "") if isinstance(ent, dict) else ""
            aceitos.append((dt.datetime.strptime(m.group(1), "%Y%m%d-%H%M%S"), m.group(2), camp))
    recuperados = 0
    for f in sorted(os.listdir(orig)):
        m = padrao.match(f)
        if not m:
            continue
        ts, _, mk, nome = m.groups()
        quando = dt.datetime.strptime(ts, "%Y%m%d-%H%M%S")
        perto = sorted((abs((q - quando).total_seconds()), c) for q, k, c in aceitos if k == mk and abs((q - quando).total_seconds()) <= 20)
        if not perto:
            continue
        try:
            linhas, info = ler("ads", os.path.join(orig, f))
        except Exception:
            continue
        if not info.get("sem_valor"):
            continue
        novo = f[:-len(".recusado")]
        if os.path.exists(os.path.join(orig, novo)):
            continue
        os.replace(os.path.join(orig, f), os.path.join(orig, novo))
        idx[novo] = {"dia": None, "campanha": perto[0][1]}
        recuperados += 1
    if recuperados:
        import json
        with open(_indice_arq(base), "w", encoding="utf-8") as fh:
            json.dump(idx, fh, ensure_ascii=False, indent=0)
    return recuperados


def reprocessar(cfg, cliente, so_mes=None):
    """Relê todos os arquivos enviados (em _originais), na ordem em que foram enviados, e refaz os
    padronizados do zero. Usado quando a leitura melhora. Devolve {(tipo, mk, mês): linhas} e erros.
    so_mes=(ano, mês): só regrava os padronizados desse mês (os outros meses ficam como estão)."""
    base = pasta(cfg, cliente)
    orig = os.path.join(base, "_originais")
    if not os.path.isdir(orig):
        return {"meses": [], "erros": []}
    idx = _indice_ler(base)
    _recuperar_recusados(base, idx)
    estado, erros = {}, []
    lims = limpezas(base)
    padrao = re.compile(r"^(\d{8}-\d{6})_(ads|afiliados)_([a-z0-9-]+?)_(.+)$")
    arquivos = sorted(f for f in os.listdir(orig) if padrao.match(f) and not f.endswith(".recusado"))
    for f in arquivos:
        ts, tipo, mk, nome = padrao.match(f).groups()
        try:
            linhas, _ = ler(tipo, os.path.join(orig, f))
            linhas = tirar_total_campanha(tipo, mk, linhas)
            forcado = _dia_forcado(tipo, mk, nome, os.path.join(orig, f))
        except Exception as e:
            erros.append("%s: %s" % (nome, e))
            continue
        if forcado or (tipo, mk) == ("ads", "shopee"):
            for l in linhas:
                l["Data"] = forcado
        ent = idx.get(f)
        if isinstance(ent, dict):
            dia, camp = ent.get("dia"), ent.get("campanha") or ""
        else:
            dia, camp = ent, ""
        if tipo == "ads":
            _aplicar_tag(linhas, camp)
        if dia:
            dia = dt.date.fromisoformat(dia)
        else:
            enviado = dt.datetime.strptime(ts, "%Y%m%d-%H%M%S").date()
            a, m = comum.periodo(enviado)[0].year, comum.periodo(enviado)[0].month
            dia = dia_do_nome(nome, a, m) or (enviado - dt.timedelta(days=1))
        novas = {}
        for l in linhas:
            if l.get("Data") is None:
                l["Data"] = dia
            if lims and _limpo(lims, ts, tipo, mk, l["Data"], l.get(CAMPO_CAMP) or ""):
                continue            # limpo depois deste envio
            novas.setdefault((l["Data"].year, l["Data"].month), []).append(l)
        for (a, m), ls in novas.items():
            chave = (tipo, mk, a, m)
            estado[chave] = _juntar(tipo, mk, estado.get(chave, []), ls)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    feitos = []
    for (tipo, mk, a, m), ls in sorted(estado.items()):
        if so_mes and (a, m) != tuple(so_mes):
            continue
        destino = caminho(cfg, cliente, tipo, mk, a, m)
        if os.path.isfile(destino):
            ant = os.path.join(base, "_anteriores", os.path.basename(destino).replace(".xlsx", "_%s_antes-reprocessar.xlsx" % carimbo))
            os.makedirs(os.path.dirname(ant), exist_ok=True)
            shutil.copy2(destino, ant)
        juntas = sorted(ls, key=lambda l: (l["Data"], l.get("Id anúncio") or "", l.get("Id do pedido") or ""))
        _gravar(juntas, tipo, destino)
        feitos.append(dict(resumo(juntas, tipo), tipo=tipo, mk=mk, mes="%04d-%02d" % (a, m)))
    if not so_mes:
        with open(os.path.join(base, ".versao_leitor"), "w") as fh:
            fh.write(str(LEITOR_VERSAO))
    comum.log("[%s] ADS/afiliados reprocessados: %s%s" % (cliente["nome"], ", ".join("%s %s %s" % (x["tipo"], x["mk"], x["mes"]) for x in feitos),
                                                       (" · erros: " + "; ".join(erros)) if erros else ""))
    return {"meses": feitos, "erros": erros}


# ---------------------------------------------------------------- limpar (para enviar de novo)
def _limp_arq(base):
    return os.path.join(base, "limpezas.json")


def limpezas(base):
    import json
    try:
        with open(_limp_arq(base), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return []


def _limpo(lims, ts_arquivo, tipo, mk, data, campanha=""):
    """A linha (de um arquivo enviado em ts_arquivo) foi limpa depois? Então não volta no reprocessar."""
    d = data.isoformat() if hasattr(data, "isoformat") else str(data)
    for L in lims:
        if L["tipo"] == tipo and L["mk"] == mk and ts_arquivo < L["ts"] and d.startswith(L["mes"]) \
                and (L.get("dias") is None or d in L["dias"]) \
                and (L.get("campanha") is None or L["campanha"] == (campanha or "")):
            return True
    return False


def limpar(cfg, dono_, tipo, mk, ano, mes, dias=None, campanha=None):
    """Tira do mês os lançamentos dos dias escolhidos (dias=None: o mês inteiro) e os OK manuais desses dias.
    Nada é apagado: o arquivo do mês vai antes para _anteriores e os originais continuam em _originais
    (marcados como limpos, para não voltarem no reprocessar)."""
    import json
    if tipo not in TIPOS:
        raise ValueError("Tipo inválido.")
    base = pasta(cfg, dono_)
    chave_mes = "%04d-%02d" % (ano, mes)
    dias = sorted({dt.date.fromisoformat(str(d)[:10]).isoformat() for d in dias}) if dias else None
    if dias and any(not d.startswith(chave_mes) for d in dias):
        raise ValueError("Os dias precisam ser do mês %s." % chave_mes)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = caminho(cfg, dono_, tipo, mk, ano, mes)
    tirados, ficaram = 0, 0
    if os.path.isfile(destino):
        linhas = _ler_padrao(destino, tipo)
        sai = lambda l: l.get("Data") and (dias is None or l["Data"].isoformat() in dias) \
            and (campanha is None or (l.get(CAMPO_CAMP) or "") == campanha)
        fica = [l for l in linhas if not sai(l)]
        tirados, ficaram = len(linhas) - len(fica), len(fica)
        ant = os.path.join(base, "_anteriores", os.path.basename(destino).replace(".xlsx", "_%s_antes-limpar.xlsx" % carimbo))
        os.makedirs(os.path.dirname(ant), exist_ok=True)
        shutil.copy2(destino, ant)
        if fica:
            _gravar(fica, tipo, destino)
        else:
            os.remove(destino)          # a cópia ficou em _anteriores
    if tipo == "visitas" and campanha is None:
        tirados += _limpar_total(cfg, dono_, mk, chave_mes, dias)     # visitas totais do dia, dos mesmos dias
    # OK manuais dos dias limpos (do marketplace e das campanhas dele)
    oks = oks_manuais(cfg, dono_)
    tirou_ok = 0
    for k in list(oks):
        t_, m_, _c = (k.split("|") + ["", ""])[:3]
        if t_ != tipo or m_ != mk or (campanha is not None and _c != campanha):
            continue
        for d in list(oks[k]):
            if d.startswith(chave_mes) and (dias is None or d in dias):
                oks[k].pop(d)
                tirou_ok += 1
        if not oks[k]:
            oks.pop(k)
    if tirou_ok:
        with open(_ok_arq(cfg, dono_) + ".novo", "w", encoding="utf-8") as f:
            json.dump(oks, f, ensure_ascii=False, indent=1)
        os.replace(_ok_arq(cfg, dono_) + ".novo", _ok_arq(cfg, dono_))
    lims = limpezas(base)
    lims.append({"tipo": tipo, "mk": mk, "mes": chave_mes, "dias": dias, "campanha": campanha, "ts": carimbo,
                 "quando": dt.datetime.now().isoformat(timespec="seconds")})
    os.makedirs(base, exist_ok=True)
    with open(_limp_arq(base) + ".novo", "w", encoding="utf-8") as f:
        json.dump(lims, f, ensure_ascii=False, indent=1)
    os.replace(_limp_arq(base) + ".novo", _limp_arq(base))
    comum.log("[%s] %s %s%s limpo%s: %d linha(s) tiradas, %d OK manual(is) tirado(s)" % (
        dono_["nome"], TIPOS[tipo]["nome"], mk, (" campanha " + campanha) if campanha is not None else "", (" em " + ", ".join(dias)) if dias else (" (mês %s inteiro)" % chave_mes),
        tirados, tirou_ok))
    return {"linhas_tiradas": tirados, "linhas_ficaram": ficaram, "oks_tirados": tirou_ok, "dias": dias}


def resumo(linhas, tipo):
    r = {"linhas": len(linhas)}
    for c in TIPOS[tipo]["numericos"]:
        r[c] = round(sum(l.get(c) or 0 for l in linhas), 2)
    datas = sorted({l["Data"] for l in linhas if l.get("Data")})
    if datas:
        r["de"], r["ate"] = datas[0].isoformat(), datas[-1].isoformat()
    r["dias"] = [d.isoformat() for d in datas]
    return r


def dias_esperados(ano, mes, hoje=None):
    """Dias do mês que deveriam ter lançamento: até ontem no mês atual; o mês inteiro nos anteriores."""
    hoje = hoje or dt.date.today()
    ult = _fim_do_mes(ano, mes, hoje)
    if dt.date(ano, mes, 1) > ult:
        return []
    return [dt.date(ano, mes, d) for d in range(1, ult.day + 1)]


_SEM = object()                  # "coluna ausente nesta linha" (para devolver o dicionário igual ao lido)
_CACHE_PADRAO = {}               # (arquivo, tipo) -> (mtime_ns, tamanho, cabeçalho, linhas já convertidas)
_CACHE_MAX_LINHAS = 1000000
_TRAVA_PADRAO = threading.Lock()


def _ler_padrao(p, tipo):
    """Linhas do arquivo padronizado (lista de dicionários novos a cada chamada).

    A leitura do .xlsx fica em memória enquanto o arquivo não muda: a tela pede a situação do mês a cada minuto
    e, sem isso, todas as planilhas de ADS/afiliados/visitas eram abertas de novo a cada vez."""
    st = os.stat(p)
    chave = (os.path.abspath(p), tipo)
    with _TRAVA_PADRAO:
        c = _CACHE_PADRAO.get(chave)
    if not c or c[0] != st.st_mtime_ns or c[1] != st.st_size:
        cab, linhas = _ler_padrao_arquivo(p, tipo)
        c = (st.st_mtime_ns, st.st_size, cab, linhas)
        with _TRAVA_PADRAO:
            _CACHE_PADRAO.pop(chave, None)
            _CACHE_PADRAO[chave] = c
            total = sum(len(x[3]) for x in _CACHE_PADRAO.values())
            while total > _CACHE_MAX_LINHAS and len(_CACHE_PADRAO) > 1:
                velho = next(iter(_CACHE_PADRAO))
                total -= len(_CACHE_PADRAO.pop(velho)[3])
    else:
        with _TRAVA_PADRAO:          # usado agora: vai para o fim da fila (sai por último)
            _CACHE_PADRAO.pop(chave, None)
            _CACHE_PADRAO[chave] = c
    cab = c[2]
    return [{k: v for k, v in zip(cab, r) if v is not _SEM} for r in c[3]]


def _ler_padrao_arquivo(p, tipo):
    from openpyxl import load_workbook
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    cab = list(linhas[0]) if linhas else []
    itens = []
    for r in linhas[1:]:
        d = dict(zip(cab, r))
        if isinstance(d.get("Data"), dt.datetime):
            d["Data"] = d["Data"].date()
        for c in TIPOS[tipo]["numericos"]:
            d[c] = float(d.get(c) or 0)
        for c in ("Id anúncio", "SKU", "Campanha", "Id do pedido", "Produto", CAMPO_CAMP):
            if c in d:
                d[c] = _texto(d[c])
        itens.append(d)
    cab_final = list(dict.fromkeys(k for d in itens for k in d)) if itens else cab
    return cab_final, [tuple(d.get(k, _SEM) for k in cab_final) for d in itens]


def carregar(cfg, cliente, tipo, ano, mes):
    """{mk: [linhas]} do mês, lendo os arquivos padronizados."""
    saida = {}
    for p in glob.glob(os.path.join(pasta(cfg, cliente), "%s_*_%04d-%02d.xlsx" % (tipo, ano, mes))):
        mk = os.path.basename(p)[len(tipo) + 1:-len("_0000-00.xlsx")]
        saida[mk] = _ler_padrao(p, tipo)
    return saida


def _dias_ajuste_afiliados(cfg, dono_, ano, mes):
    """{mk: {dias (AAAA-MM-DD) do mês cobertos por ajuste manual de afiliados}}."""
    ini, fim = dt.date(ano, mes, 1), dt.date(ano, mes, calendar.monthrange(ano, mes)[1])
    out = {}
    for a in ajustes(cfg, dono_):
        if a.get("tipo") != "afiliados":
            continue
        d1, d2 = max(ini, dt.date.fromisoformat(a["de"])), min(fim, dt.date.fromisoformat(a["ate"]))
        while d1 <= d2:
            out.setdefault(a["mk"], set()).add(d1.isoformat())
            d1 += dt.timedelta(days=1)
    return out


def situacao(cfg, cliente, ano, mes):
    """Para o painel: o que já foi subido no mês, por tipo e marketplace."""
    s = {}
    esperados = dias_esperados(ano, mes)
    oks = oks_manuais(cfg, cliente)
    for tipo in TIPOS:
        s[tipo] = {}
        for mk, ls in carregar(cfg, cliente, tipo, ano, mes).items():
            p = caminho(cfg, cliente, tipo, mk, ano, mes)
            fch = fechadas(mk, ls) if tipo == "ads" else {}
            if tipo == "ads":
                ls = efetivas_ads(mk, ls)
            r = dict(resumo(ls, tipo), arquivo=p,
                     atualizado=dt.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec="seconds"))
            tem = set(r["dias"])
            if unico(tipo, mk):     # relatório único: não se cobra dia a dia
                r["unico"] = True
                r["pedidos"] = len({l.get("Id do pedido") for l in ls if l.get("Id do pedido")})
                r["faltam"] = []
            else:
                r["faltam"] = [d.isoformat() for d in esperados if d.isoformat() not in tem]
                _aplicar_ok(r, oks.get(_ok_chave(tipo, mk)), ano, mes)
            if fch:
                r["fechadas"] = fch
            s[tipo][mk] = r
        # marketplace sem nenhum arquivo no mês, mas com OK manual: aparece com os dias de OK
        # (vale também para o relatório único, ex.: afiliados da Shopee/ML num mês sem comissão)
        for k, dias in oks.items():
            t, mk, camp = k.split("|", 2)
            if t == tipo and not camp and mk not in s[tipo] and any(d.startswith("%04d-%02d" % (ano, mes)) for d in dias):
                r = dict(resumo([], tipo), arquivo=None, atualizado=None, so_ok=True,
                         faltam=[d.isoformat() for d in esperados])
                s[tipo][mk] = _aplicar_ok(r, dias, ano, mes)
        if tipo == "afiliados":
            # comissão lançada por ajuste manual (sem relatório): os dias do ajuste contam como enviados
            for mk, dias_aj in _dias_ajuste_afiliados(cfg, cliente, ano, mes).items():
                r = s[tipo].get(mk)
                if r is None:
                    r = dict(resumo([], tipo), arquivo=None, atualizado=None, so_ok=True,
                             faltam=[d.isoformat() for d in esperados])
                    s[tipo][mk] = _aplicar_ok(r, oks.get(_ok_chave(tipo, mk)), ano, mes)
                r["faltam"] = [d for d in r.get("faltam", []) if d not in dias_aj]
                r["dias_ajuste"] = sorted(dias_aj)
        s["dias_esperados"] = len(esperados)
    s["visitas_totais"] = situacao_totais(cfg, cliente, ano, mes)
    # campanhas: cada campanha ativa precisa ter todos os dias; o dia só fica completo com todas
    s["campanhas"] = {}
    ads = carregar(cfg, cliente, "ads", ano, mes)
    for mk, lista in campanhas(cfg, cliente).items():
        ls = ads.get(mk, [])
        saida = []
        for c in lista:
            dela = [l for l in ls if (l.get(CAMPO_CAMP) or "") == c["id"]]
            r = dict(c, **resumo(dela, "ads"))
            tem = set(r["dias"])
            r["faltam"] = [d.isoformat() for d in esperados if d.isoformat() not in tem] if c.get("ativa", True) else []
            _aplicar_ok(r, oks.get(_ok_chave("ads", mk, c["id"])), ano, mes)
            saida.append(r)
        sem = [l for l in ls if not l.get(CAMPO_CAMP)]
        if sem:
            saida.append(dict({"id": "", "nome": "Sem campanha (enviado antes)", "ativa": False}, **resumo(sem, "ads"), faltam=[]))
        s["campanhas"][mk] = saida
        ativas = [c for c in saida if c.get("ativa", True) and c["id"]]
        if ativas and mk not in s["ads"] and any(c.get("ok_manual") for c in ativas):
            s["ads"][mk] = dict(resumo([], "ads"), arquivo=None, atualizado=None, so_ok=True, faltam=[])
        if ativas and mk in s["ads"]:
            falt = sorted({d for c in ativas for d in c["faltam"]})
            s["ads"][mk]["faltam"] = falt
            s["ads"][mk]["por_campanha"] = True
    # ajustes manuais (campanha apagada) que tocam o mês
    ini, fim = dt.date(ano, mes, 1), dt.date(ano, mes, calendar.monthrange(ano, mes)[1])
    s["ajustes"] = {}
    for a in ajustes(cfg, cliente):
        if dt.date.fromisoformat(a["de"]) <= fim and dt.date.fromisoformat(a["ate"]) >= ini:
            s["ajustes"].setdefault(a["mk"], []).append(a)
    s["depara"] = {mk: info_depara(cfg, cliente, mk) for mk in DEPARA}
    s["depara_shopee"] = s["depara"]["shopee"]
    return s


def modelo(cfg, tipo):
    p = os.path.join(comum.PASTA_SISTEMA, "dados", "modelo-%s.xlsx" % tipo)
    if not os.path.isfile(p):
        os.makedirs(os.path.dirname(p), exist_ok=True)
        hoje = dt.date.today()
        if tipo == "ads":
            ex = [{"Data": hoje.replace(day=1), "Id anúncio": "MLB1234567890", "SKU": "ABC-001", "Campanha": "Campanha geral",
                   "Gasto": 12.5, "Cliques": 30, "Receita ADS": 180.0}]
        elif tipo == "visitas":
            ex = [{"Data": hoje.replace(day=1), "Id anúncio": "MLB1234567890", "SKU": "ABC-001", "Visitas": 120, "Vendas": 3}]
        else:
            ex = [{"Data": hoje.replace(day=1), "Id do pedido": "2000012345678901", "Id anúncio": "MLB1234567890",
                   "SKU": "ABC-001", "Comissão afiliado": 6.9}]
        _gravar(ex, tipo, p)
    return p
