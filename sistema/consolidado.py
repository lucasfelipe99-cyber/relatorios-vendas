# -*- coding: utf-8 -*-
"""Relatório consolidado (pacote final) — um arquivo por cliente e mês, para subir no site.

Formato igual ao modelo "10-Custo_MKP.xlsx": aba DB, formato longo, uma linha por valor:
Marketplace | Marketplace venda | SKU | Id anúncio | Data | Categoria | Sub Categoria | Valor completo |
Type | Descrição | Categoria2 | Datatype | Record Date | Full Data | Coluna Ponto

Fonte: os tratados do mês (de todas as contas do cliente), somados por SKU × anúncio × dia × Type.
Regras combinadas (19/10/2026):
- Marketplace venda = nome do cliente no app; um arquivo por cliente.
- Sem embalagem.
- Venda cancelada (ou devolvida): entra o faturamento (+) e sai em "Cancelamento" (−), dentro de
  13.Faturamento Bruto, somando zero; não leva imposto, comissão, frete nem custo. Custos de ADS e
  afiliados ligados a ela continuam (quando essas etapas existirem).
- Antes de liberar, confere os totais do consolidado com os tratados, por marketplace.
"""
import datetime as dt
import glob
import json
import os
import re
import shutil
import tempfile

import comum

COLUNAS = ["Marketplace", "Marketplace venda", "SKU", "Id anúncio", "Data", "Categoria", "Sub Categoria",
           "Valor completo", "Type", "Descrição", "Categoria2", "Datatype", "Record Date", "Full Data", "Coluna Ponto"]

# (categoria, sub categoria) de cada medida
QTD = ("14.Quantidade", "Quantidade")
FAT = ("13.Faturamento Bruto", "Preço total com desconto")
CANC = ("13.Faturamento Bruto", "Cancelamento")
IMP = ("02.Imposto", "Imposto")
COM = ("03.Despesas Marketplace", "Comissão")
FRE = ("03.Despesas Marketplace", "Tarifa de frete")
REB = ("03.Despesas Marketplace", "Rebate")
ANT = ("03.Despesas Marketplace", "Antecipação")
CUS = ("04.Materia Prima", "Produto")
PUB = ("03.Despesas Marketplace", "Publicidade")
AFI = ("03.Despesas Marketplace", "Afiliados")
CLQ = ("Cliques", "Cliques")
ADF = ("ADS F", "ADS F")
VIS = ("Visitas", "Visitas únicas")
QVV = ("Quantidade de vendas", "Quantidade de vendas")      # do relatório de visitas (conversão = vendas / visitas)

# como ler o tratado de cada marketplace
MAPA = {
    "mercadolivre": {"nome": "Mercado Livre", "pedido": "N.º de venda", "anuncio": "# de anúncio", "sku": "SKU", "qtd": "Unidades",
                     "fat": "Faturamento", "desc": "Desconto", "rebate": "rebate", "status": "Type.1"},
    "shopee": {"nome": "Shopee", "pedido": "ID do pedido", "anuncio": "ID do Produto", "sku": "Número de referência SKU", "qtd": "Quantidade",
               "fat": "Faturamento", "desc": "Desconto", "rebate": "Rebate", "status": "Status"},
    "magalu": {"nome": "Magalu", "pedido": "Número do pedido", "anuncio": "Codigo SKU seller", "sku": "SKU.1", "qtd": "Quantidade de itens",
               "fat": "Fat", "desc": "Desconto", "rebate": "Rebate", "status": "Status"},
    "tiktokshop": {"nome": "TikTok", "pedido": "Order ID", "anuncio": "SKU ID", "sku": "Seller SKU", "qtd": "Quantity",
                   "fat": "Faturamento", "desc": "Desconto", "rebate": "Rebate", "status": "Status"},
    "amazon": {"nome": "Amazon", "pedido": "amazon-order-id", "anuncio": "asin", "sku": "SKU.1", "qtd": "quantity",
               "fat": "Faturamento", "desc": "item-promotion-discount", "rebate": "Rebate", "status": "Status"},
    "viavarejo": {"nome": "Casas Bahia", "pedido": "N.º de venda", "anuncio": "# de anúncio", "sku": "SKU", "qtd": "Unidades",
                  "fat": "Faturamento", "desc": "Desconto", "rebate": "rebate", "status": "Type.1"},
    "temu": {"nome": "Temu", "pedido": "N.º de venda", "anuncio": "# de anúncio", "sku": "SKU", "qtd": "Unidades",
             "fat": "Faturamento", "desc": "Desconto", "rebate": "rebate", "status": "Type.1"},
    "madeiramadeira": {"nome": "MadeiraMadeira", "pedido": "N.º de venda", "anuncio": "# de anúncio", "sku": "SKU",
                       "qtd": "Unidades", "fat": "Faturamento", "desc": "Desconto", "rebate": "rebate", "status": "Type.1"},
}
VENDA = {"venda"}
CANCELADA = {"cancelada", "cancelado", "devolução", "devolucao"}
# o resto (Não pago, Amostra Grátis...) nunca foi venda: fica de fora
METRICAS = [("Quantidade", QTD), ("Faturamento bruto", FAT), ("Cancelamento", CANC), ("Imposto", IMP),
            ("Comissão", COM), ("Frete", FRE), ("Rebate", REB), ("Antecipação", ANT), ("Custo do produto", CUS)]
TOLERANCIA = 0.05


# ---------------------------------------------------------------- onde ficam os arquivos
def nome_arquivo(cliente, ano, mes):
    nome = re.sub(r'[\\/:*?"<>|]+', "-", cliente["nome"]).strip()
    return "relatorio_consolidado_%s_%02d-%04d.xlsx" % (nome, mes, ano)


def mes_atual(hoje=None):
    ini, _ = comum.periodo(hoje or dt.date.today())
    return ini.year, ini.month


def _pasta_dia_data(cfg, rotulo):
    return os.path.join(comum.raiz(cfg), rotulo)


def dia_fonte(cfg, ano, mes, hoje=None):
    """Data da pasta [DD-MM] cujos tratados cobrem o mês pedido.

    Mês atual: a pasta de hoje. Mês anterior: a pasta do dia 1 do mês seguinte (que tem o mês inteiro);
    sem ela, a última pasta existente dentro do mês (parcial)."""
    hoje = hoje or dt.date.today()
    if (ano, mes) == mes_atual(hoje):
        return hoje, False
    seguinte = dt.date(ano + (mes == 12), 1 if mes == 12 else mes + 1, 1)
    if seguinte <= hoje and os.path.isdir(_pasta_dia_data(cfg, comum.rotulo_dia(seguinte))):
        return seguinte, False
    ult = dt.date(seguinte.year, seguinte.month, 1) - dt.timedelta(days=1)
    d = min(ult, hoje)
    while d.month == mes and d.year == ano:
        if os.path.isdir(_pasta_dia_data(cfg, comum.rotulo_dia(d))):
            return d, True
        d -= dt.timedelta(days=1)
    return None, True


def meses_disponiveis(cfg, hoje=None, quantos=6):
    hoje = hoje or dt.date.today()
    ano, mes = mes_atual(hoje)
    saida = []
    for _ in range(quantos):
        d, parcial = dia_fonte(cfg, ano, mes, hoje)
        saida.append({"chave": "%04d-%02d" % (ano, mes), "rotulo": "%02d/%04d" % (mes, ano),
                      "disponivel": d is not None, "parcial": parcial and d is not None,
                      "fonte": d.isoformat() if d else None})
        ano, mes = (ano - 1, 12) if mes == 1 else (ano, mes - 1)
    return saida


# ---------------------------------------------------------------- leitura
def _num(serie):
    import pandas as pd
    return pd.to_numeric(serie, errors="coerce").fillna(0.0)


def _texto(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if v.is_integer():
            return str(int(v))
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _descricoes(cliente, ano=None, mes=None):
    """SKU -> (Descrição, Categoria) da tabela de custos que vale para o mês."""
    import tratador
    p = tratador.caminho_custos(cliente, ano, mes) if ano else tratador.caminho_custos(cliente)
    if not p:
        return {}
    try:
        linhas, _ = tratador.ler_custos(p)
    except Exception:
        return {}
    return {l["SKU"]: (l.get("Descrição") or "X", l.get("Categoria") or "X") for l in linhas}


def _ler_tratado(caminho, mk, ano, mes):
    import pandas as pd
    m = MAPA[mk]
    df = ler_excel(caminho, sheet_name="DB", dtype={m["anuncio"]: str, m["sku"]: str, m["pedido"]: str})
    for c in (m["qtd"], m["fat"], m["desc"], m["rebate"], "Comissão", "Frete", "Cancelamento", "Imposto",
              "Custo do produto", "Antecipa"):
        df[c] = _num(df[c]) if c in df.columns else 0.0
    df["_data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    df = df[(df["_data"].dt.year == ano) & (df["_data"].dt.month == mes)].copy()
    df["_status"] = df[m["status"]].astype(str).str.strip().str.lower()
    df["_sku"] = df[m["sku"]].map(_texto)
    df["_anuncio"] = df[m["anuncio"]].map(_texto)
    df["_pedido"] = df[m["pedido"]].map(_texto) if m["pedido"] in df.columns else ""
    return df


def _linhas_do_tratado(df, mk):
    """Gera as linhas longas (sem somar) e os totais do tratado para a conferência."""
    m = MAPA[mk]
    venda = df[df["_status"].isin(VENDA)]
    canc = df[df["_status"].isin(CANCELADA)]
    import pandas as pd
    partes = []

    def add(base, cat, valores):
        if len(base):
            partes.append(pd.DataFrame({"Pedido": base["_pedido"].values, "SKU": base["_sku"].values, "Id anúncio": base["_anuncio"].values,
                                        "Data": base["_data"].values, "Categoria": cat[0], "Sub Categoria": cat[1],
                                        "Valor completo": pd.to_numeric(valores, errors="coerce").fillna(0).values}))

    add(venda, QTD, venda[m["qtd"]])
    add(venda, FAT, venda[m["fat"]] + venda[m["desc"]])
    add(venda, CANC, venda["Cancelamento"])          # reembolso parcial em venda que seguiu
    add(venda, IMP, venda["Imposto"])
    add(venda, COM, venda["Comissão"])
    add(venda, FRE, venda["Frete"])
    add(venda, REB, venda[m["rebate"]])
    add(venda, ANT, venda["Antecipa"])
    add(venda, CUS, venda["Custo do produto"])
    fat_c = canc[m["fat"]] + canc[m["desc"]]
    add(canc, FAT, fat_c)
    add(canc, CANC, -fat_c)
    col_liq = "Liquido" if "Liquido" in df.columns else ("Líquido" if "Líquido" in df.columns else None)
    # cancelada/devolução com COBRANÇA (o tratador deixa no Líquido só o que foi cobrado, sempre ≤ 0):
    #   com estorno no relatório (coluna Cancelamento) -> entra como cancelamento; sem estorno -> comissão e frete
    #   (Líquido positivo = tratado antigo, antes dessa regra: fica de fora)
    cob_com = cob_fre = cob_can = 0.0
    if col_liq and len(canc):
        lq = _num(canc[col_liq])
        com_c = _num(canc["Comissão"]) if "Comissão" in canc.columns else lq * 0
        # Shopee/Magalu/TikTok: o Cancelamento do tratado é só o estorno do faturamento -> a cobrança é comissão/frete
        com_estorno = ((_num(canc["Cancelamento"]).abs() > 0.005) if "Cancelamento" in canc.columns else lq < -1e18) \
            & (mk not in ("shopee", "magalu", "tiktokshop"))
        cobra = lq < -0.005
        a, b = canc[cobra & com_estorno], canc[cobra & ~com_estorno]
        if len(a):
            add(a, CANC, lq[a.index])
            cob_can = float(lq[a.index].sum())
        if len(b):
            cb = com_c[b.index].clip(lower=lq[b.index], upper=0)      # comissão (nunca além da cobrança)
            add(b, COM, cb)
            add(b, FRE, lq[b.index] - cb)
            cob_com, cob_fre = float(cb.sum()), float((lq[b.index] - cb).sum())
    medidas = pd.concat(partes, ignore_index=True) if partes else None

    totais = {
        "Líquido (conferência)": float(_num(venda[col_liq]).sum()) + cob_can + cob_com + cob_fre if col_liq else None,
        "Quantidade": float(venda[m["qtd"]].sum()),
        "Faturamento bruto": float((venda[m["fat"]] + venda[m["desc"]]).sum() + fat_c.sum()),
        "Cancelamento": float(venda["Cancelamento"].sum() - fat_c.sum()) + cob_can,
        "Imposto": float(venda["Imposto"].sum()), "Comissão": float(venda["Comissão"].sum()) + cob_com,
        "Frete": float(venda["Frete"].sum()) + cob_fre, "Rebate": float(venda[m["rebate"]].sum()),
        "Antecipação": float(venda["Antecipa"].sum()), "Custo do produto": float(venda["Custo do produto"].sum()),
    }
    extra = {"vendas": int(len(venda)), "canceladas": int(len(canc)),
             "fora": int(len(df) - len(venda) - len(canc)),
             "sem_custo": int((venda["Custo do produto"] == 0).sum()),
             "cancelada_sem_valor": int((fat_c == 0).sum())}
    return medidas, totais, extra


# ---------------------------------------------------------------- anúncio com variações (mais de um SKU)
def _repartir(df, nome, loja, sku, cache, dia=None):
    """SKU do de-para pode ser "290, 289" (anúncio com variações). Devolve [(sku, peso)].

    ADS (com o dia): o gasto do anúncio no dia vai para os SKUs dele que venderam naquele dia, na mesma conta,
    proporcional ao faturamento. Dia sem venda de nenhum SKU do anúncio: pelo faturamento do mês (na conta;
    senão em todas); sem nada, em partes iguais. Só o faturamento é usado no ADS.
    Visitas (sem o dia): pelo faturamento do mês de cada SKU (na conta; senão em todas); sem venda, pelas unidades;
    sem nada, em partes iguais."""
    import lancamentos
    skus = lancamentos.skus_lista(sku)
    if len(skus) <= 1:
        return [(skus[0] if skus else sku, 1.0)]
    if dia is not None:
        kd = ("dia", nome, loja, tuple(skus))
        if kd not in cache:
            b = df[(df["Marketplace"] == nome) & (df["Marketplace venda"] == loja) & df["SKU"].astype(str).isin(skus)
                   & (df["Categoria"] == FAT[0]) & (df["Sub Categoria"] == FAT[1])]
            por_dia = {}
            if len(b):
                g = b.groupby([pd_datas(b["Data"]), b["SKU"].astype(str)])["Valor completo"].sum()
                for (d, x), v in g.items():
                    if v > 0:
                        por_dia.setdefault(d, {})[x] = float(v)
            cache[kd] = por_dia
        vend = cache[kd].get(pd_datas_um(dia))
        if vend:
            tot = sum(vend.values())
            return [(x, vend.get(x, 0.0) / tot) for x in skus]
    k = (nome, loja, tuple(skus), dia is not None)
    if k not in cache:
        pesos = None
        base = df[(df["Marketplace"] == nome) & df["SKU"].astype(str).isin(skus)]
        for filtro in (base["Marketplace venda"] == loja, None):
            b = base[filtro] if filtro is not None else base
            for cat in ((FAT,) if dia is not None else (FAT, QTD)):
                v = b[(b["Categoria"] == cat[0]) & (b["Sub Categoria"] == cat[1])].groupby(b["SKU"].astype(str))["Valor completo"].sum()
                tot = float(sum(max(0.0, float(v.get(x, 0.0))) for x in skus))
                if tot > 0:
                    pesos = [(x, max(0.0, float(v.get(x, 0.0))) / tot) for x in skus]
                    break
            if pesos:
                break
        cache[k] = pesos or [(x, 1.0 / len(skus)) for x in skus]
    return cache[k]


def _skus_como_vendas_shopee(sku):
    """Os SKUs das variações no de-para, do jeito que a venda da Shopee guarda o SKU (cortado no primeiro "-")."""
    import lancamentos
    vistos = []
    for x in lancamentos.skus_lista(sku):
        x = x.split("-")[0].strip() or x
        if x not in vistos:
            vistos.append(x)
    return ", ".join(vistos)


def pd_datas(serie):
    """Datas (sem hora) de uma coluna, como date."""
    import pandas as pd
    return pd.to_datetime(serie, errors="coerce").dt.date


def pd_datas_um(v):
    import pandas as pd
    try:
        return pd.Timestamp(v).date()
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- ADS e afiliados (subidos pelo app)
def _nome_mk(cfg, mk):
    return MAPA[mk]["nome"] if mk in MAPA else cfg["marketplaces"].get(mk, {}).get("nome", mk)


def _sku_norm(t):
    t = str(t or "").strip().upper().replace(" ", "")
    return t.lstrip("0") or t


def _nome_norm(t):
    import unicodedata
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()
    return " ".join("".join(c if c.isalnum() else " " for c in t).split()) if t not in ("nan", "none") else ""


def _aplicar_lancamentos(cfg, cliente, df, ano, mes, nomes_tiktok=None):
    """ADS e afiliados são subidos POR CONTA: cada linha sai com o nome da conta (Marketplace venda).
    Afiliados: as vendas do pedido (na mesma conta) viram Type "Afiliados" e a comissão é rateada pelos itens."""
    import pandas as pd
    import lancamentos
    novas, totais, resumo = [], {}, {"ads": {}, "afiliados": {}, "afiliados_sem_pedido": 0, "afiliados_sem_pedido_por": {},
                                     "afiliados_pedidos_achados": {}, "ads_sem_depara": {}, "visitas": {}, "ids_shopee": {}}
    pesos_sh = {}          # Shopee: {loja: {id do anúncio: [gasto ADS, visitas]}} para escolher o anúncio do SKU que está em mais de um
    fat = (df["Categoria"] == FAT[0]) & (df["Sub Categoria"] == FAT[1]) & (df["Valor completo"] > 0)
    for u in comum.contas_do_cliente(cfg, cliente):
        d = lancamentos.dono(u)
        loja = d["loja"]
        da_conta = df["Marketplace venda"] == loja

        afi_mes = lancamentos.carregar(cfg, d, "afiliados", ano, mes)
        aj_afi = lancamentos.linhas_ajuste_afiliados(cfg, d, ano, mes)     # comissão lançada de uma vez (sem relatório)
        for mk in aj_afi:
            afi_mes.setdefault(mk, [])
        for mk, itens in afi_mes.items():
            nome = _nome_mk(cfg, mk)
            t = totais.setdefault(nome, {})
            t["Afiliados (comissão)"] = t.get("Afiliados (comissão)", 0.0) + sum(i["Comissão afiliado"] for i in itens) \
                + sum(i["Comissão afiliado"] for i in aj_afi.get(mk, []))
            for i in aj_afi.get(mk, []):
                novas.append({"Pedido": "", "Marketplace venda": loja, "SKU": i["SKU"], "Id anúncio": i["Id anúncio"],
                              "Data": pd.Timestamp(i["Data"]), "Categoria": AFI[0], "Sub Categoria": AFI[1],
                              "Valor completo": -i["Comissão afiliado"], "Marketplace": nome, "Type": "Afiliados"})
            resumo["afiliados"]["%s|%s" % (u["id"], mk)] = len(itens)
            do_mk = (df["Marketplace"] == nome) & da_conta
            pedidos = {i["Id do pedido"] for i in itens if i.get("Id do pedido")}
            if pedidos:
                df.loc[do_mk & df["Pedido"].isin(pedidos), "Type"] = "Afiliados"
            vendas = df[do_mk & fat]
            por_pedido = {p: g for p, g in vendas.groupby("Pedido")} if pedidos else {}
            chave_k = "%s|%s" % (u["id"], mk)
            achados = set()
            com_sku = df[(df["Marketplace"] == nome) & (df["SKU"] != "") & (df["Id anúncio"] != "")]
            sku_do_anuncio = com_sku.groupby(com_sku["Id anúncio"].astype(str))["SKU"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(com_sku) else {}
            usados = set()
            dp_afi = lancamentos.carregar_depara(cfg, d, mk) if lancamentos.tem_depara(mk) else None
            rep_afi = {}
            for i in itens:
                g = por_pedido.get(i.get("Id do pedido"))
                if g is not None and len(g) and i.get("Id anúncio"):
                    so_item = g[g["Id anúncio"].astype(str) == str(i["Id anúncio"])]   # relatório por item: vai para o item
                    if len(so_item):
                        g = so_item
                if g is not None and len(g):
                    achados.add(i.get("Id do pedido"))
                    usados.add(str(g["Pedido"].iloc[0]))
                    pesos = g["Valor completo"] / g["Valor completo"].sum()
                    for (_, r), w in zip(g.iterrows(), pesos):
                        novas.append({"Pedido": r["Pedido"], "Marketplace venda": loja, "SKU": r["SKU"], "Id anúncio": r["Id anúncio"],
                                      "Data": r["Data"], "Categoria": AFI[0], "Sub Categoria": AFI[1],
                                      "Valor completo": -i["Comissão afiliado"] * w, "Marketplace": nome, "Type": "Afiliados"})
                else:
                    resumo["afiliados_sem_pedido"] += 1
                    resumo["afiliados_sem_pedido_por"][chave_k] = resumo["afiliados_sem_pedido_por"].get(chave_k, 0) + 1
                    an_i = str(i.get("Id anúncio") or "")
                    sku_i = i.get("SKU") or sku_do_anuncio.get(an_i) or ""
                    if not sku_i and an_i and dp_afi is not None:
                        # pedido fora das vendas do mês: o SKU vem do de-para do anúncio (com variações, rateado pelo faturamento)
                        sku_i = (dp_afi.get(an_i) or {}).get("sku") or ""
                        if mk == "shopee" and sku_i:
                            sku_i = _skus_como_vendas_shopee(sku_i)
                    sku_i = sku_i or "0"
                    rp = [(x, w) for x, w in _repartir(df, nome, loja, sku_i, rep_afi, dia=i["Data"]) if w > 0] or [(sku_i, 1.0)]
                    for sku_p, w in rp:
                        novas.append({"Pedido": i.get("Id do pedido", ""), "Marketplace venda": loja,
                                      "SKU": sku_p, "Id anúncio": i.get("Id anúncio", ""), "Data": pd.Timestamp(i["Data"]),
                                      "Categoria": AFI[0], "Sub Categoria": AFI[1], "Valor completo": -i["Comissão afiliado"] * w,
                                      "Marketplace": nome, "Type": "Afiliados"})
            resumo["afiliados_pedidos_achados"][chave_k] = [len(achados), len(pedidos)]

        ads_mes = lancamentos.carregar(cfg, d, "ads", ano, mes)
        aj = lancamentos.linhas_ajuste(cfg, d, ano, mes)
        for mk in set(ads_mes) | set(aj):
            # campanha fechada com detalhe subido: vale o detalhe; ajustes manuais (campanha apagada) entram somados
            itens = lancamentos.efetivas_ads(mk, ads_mes.get(mk, [])) + aj.get(mk, [])
            nome = _nome_mk(cfg, mk)
            t = totais.setdefault(nome, {})
            for rot, campo in (("ADS gasto", "Gasto"), ("Cliques", "Cliques"), ("ADS F (receita)", "Receita ADS")):
                t[rot] = t.get(rot, 0.0) + sum(i[campo] for i in itens)
            resumo["ads"]["%s|%s" % (u["id"], mk)] = len(itens)
            base_sku = df[(df["Marketplace"] == nome) & (df["SKU"] != "") & (df["Id anúncio"] != "")]
            sub = base_sku[base_sku["Marketplace venda"] == loja]
            sku_conta = sub.groupby("Id anúncio")["SKU"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(sub) else {}
            sku_geral = base_sku.groupby("Id anúncio")["SKU"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(base_sku) else {}
            depara = lancamentos.carregar_depara(cfg, d, mk) if mk in lancamentos.DEPARA_ADS_POR_PLANILHA else None
            # demais canais (ML, Amazon, TikTok...): planilha de anúncios (de-para) — usada quando o anúncio não vendeu
            # no mês ou tem variações; com venda e um SKU só, vale o SKU das vendas
            dp_ml = lancamentos.carregar_depara(cfg, d, mk) if (lancamentos.tem_depara(mk) and depara is None) else {}
            # TikTok: o anúncio das vendas é o SKU ID de cada variação
            an_tt = sub.groupby("SKU")["Id anúncio"].agg(lambda x: x.value_counts().index[0]).to_dict() \
                if (mk == "tiktokshop" and len(sub)) else {}
            skus_vendas = set(base_sku["SKU"].astype(str))
            sem_depara, nomes_sem = {}, {}
            manual_tt = lancamentos.depara_manual(cfg, d, mk)
            rep_cache = {}
            for i in itens:
                an = i.get("Id anúncio", "")
                if (i.get(lancamentos.CAMPO_CAMP) or "").startswith(lancamentos.AJUSTE):
                    base = {"Pedido": "", "Marketplace venda": loja, "SKU": i["SKU"], "Id anúncio": i["SKU"], "Data": pd.Timestamp(i["Data"]),
                            "Marketplace": nome, "Type": "Organico"}
                    for (cat, sc), v in ((PUB, -i["Gasto"]), (CLQ, i["Cliques"]), (ADF, i["Receita ADS"])):
                        if v:
                            novas.append(dict(base, Categoria=cat, **{"Sub Categoria": sc, "Valor completo": v}))
                    continue
                if depara is not None:
                    # Shopee: ID do produto -> SKU pela mass_update; Id anúncio = SKU (igual às vendas da Shopee).
                    # Linha da loja (GMV Max da Loja, sem produto) e anúncio sem de-para: SKU 0.
                    # Shopee: ID numérico; linha da loja ("GMV Max da Loja") não tem ID. Magalu: Código site Magalu.
                    # A coluna "SKU" do relatório da Magalu traz o Código site Magalu: se o "SKU" do arquivo
                    # for um código do de-para, ele é o anúncio.
                    sku_arq = str(i.get("SKU") or "")
                    if sku_arq and sku_arq in depara and sku_arq not in skus_vendas:
                        an, sku_arq = sku_arq, ""
                    eh_id = an.isdigit() if mk == "shopee" else bool(an)
                    if mk == "shopee" and eh_id:
                        pesos_sh.setdefault(loja, {}).setdefault(an, [0.0, 0.0])[0] += abs(i.get("Gasto") or 0)
                    if eh_id and (depara.get(an) or {}).get("manual"):
                        sku_arq = ""                    # preenchido à mão no app vale por cima do SKU do relatório
                    if eh_id:
                        sku = (depara.get(an) or {}).get("sku") or (an if an in skus_vendas else "")
                        if mk == "shopee" and "," in sku:
                            sku = _skus_como_vendas_shopee(sku)
                        if not sku and not sku_arq:
                            sem_depara[an] = sem_depara.get(an, 0.0) + i["Gasto"]
                            nomes_sem.setdefault(an, str(i.get("Campanha") or "")[:150])
                    else:
                        # campanha fechada (sem produto, ex.: "GMV Max da Loja"): o SKU leva o nome da campanha
                        sku = an if (mk == "shopee" and an) else ""
                    if not an and sku_arq and mk == "magalu" and sku_arq not in skus_vendas:
                        an, sku_arq = sku_arq, ""        # código Magalu sem de-para: aparece no aviso
                        sem_depara[an] = sem_depara.get(an, 0.0) + i["Gasto"]
                        nomes_sem.setdefault(an, str(i.get("Campanha") or "")[:150])
                    sku = sku_arq or sku or "0"
                    rp = [(x, w) for x, w in _repartir(df, nome, loja, sku, rep_cache, dia=i["Data"]) if w > 0]
                    for sku_p, w in rp:
                        an_p = sku_p if (sku_p != "0" or not eh_id) else an
                        base = {"Pedido": "", "Marketplace venda": loja, "SKU": sku_p, "Id anúncio": an_p, "Data": pd.Timestamp(i["Data"]),
                                "Marketplace": nome, "Type": "Organico"}
                        for (cat, sc), v in ((PUB, -i["Gasto"]), (CLQ, i["Cliques"]), (ADF, i["Receita ADS"])):
                            if v:
                                novas.append(dict(base, Categoria=cat, **{"Sub Categoria": sc, "Valor completo": v * w}))
                    continue
                sku = i.get("SKU") or ""
                if not sku and dp_ml:
                    # ML: anúncio com variações na planilha (e alguma delas vendendo no mês): reparte entre as variações;
                    # sem venda no mês: o SKU da planilha. Com venda e um SKU só: o SKU das vendas, como antes.
                    sk_dp = (dp_ml.get(lancamentos.id_depara(mk, an)) or dp_ml.get(str(an)) or {}).get("sku") or ""
                    lst = lancamentos.skus_lista(sk_dp)
                    vend = sku_conta.get(an) or sku_geral.get(an)
                    if len(lst) > 1 and any(x in skus_vendas for x in lst):
                        sku = sk_dp
                    elif not vend and lst:
                        sku = sk_dp
                sku = sku or sku_conta.get(an) or sku_geral.get(an) or ""
                if not sku and mk != "tiktokshop" and (manual_tt.get(an) or {}).get("sku"):
                    sku = manual_tt[an]["sku"]          # apontado à mão no app
                if not sku and mk == "tiktokshop":
                    nm = _nome_norm(i.get("Campanha"))
                    man = manual_tt.get(an) or (manual_tt.get(nm) if nm else None)   # preenchido no app (por ID ou nome)
                    par = (nomes_tiktok or {}).get(loja, {}).get(nm) or (nomes_tiktok or {}).get("*", {}).get(nm) if nm else None
                    if man and man.get("sku"):
                        sku = man["sku"]
                        sku_id = next((p_[1] for p_ in ((nomes_tiktok or {}).get("*") or {}).values() if p_[0] == sku and p_[1]), "")
                        an = sku_id or an
                    elif par and (par[0] or par[1]):
                        sku, an = par[0] or par[1], par[1] or an     # mesmo anúncio (SKU ID) das vendas
                    else:
                        sem_depara[an] = sem_depara.get(an, 0.0) + i["Gasto"]
                        nomes_sem.setdefault(an, str(i.get("Campanha") or "")[:150])
                elif not sku and an:
                    # anúncio com ADS sem venda no mês (e sem SKU no relatório): entra no SKU 0 e pede o SKU no sistema
                    sem_depara[an] = sem_depara.get(an, 0.0) + i["Gasto"]
                    nomes_sem.setdefault(an, str(i.get("Campanha") or "")[:150])
                sku = sku or "0"
                varios = len(lancamentos.skus_lista(sku)) > 1
                rp = [(x, w) for x, w in _repartir(df, nome, loja, sku, rep_cache, dia=i["Data"]) if w > 0]
                for sku_p, w in rp:
                    base = {"Pedido": "", "Marketplace venda": loja, "SKU": sku_p, "Id anúncio": (an_tt.get(sku_p) or sku_p) if (varios and mk == "tiktokshop") else an,
                            "Data": pd.Timestamp(i["Data"]), "Marketplace": nome, "Type": "Organico"}
                    for (cat, sc), v in ((PUB, -i["Gasto"]), (CLQ, i["Cliques"]), (ADF, i["Receita ADS"])):
                        if v:
                            novas.append(dict(base, Categoria=cat, **{"Sub Categoria": sc, "Valor completo": v * w}))
            if depara is not None or sem_depara:
                resumo["ads_sem_depara"]["%s|%s" % (u["id"], mk)] = {
                    "tem_depara": bool(depara) or bool(dp_ml) or mk == "tiktokshop",
                    "ids": {k: round(v, 2) for k, v in sorted(sem_depara.items(), key=lambda x: -x[1])},
                    "nomes": {k: nomes_sem.get(k, "") for k in sem_depara}}
        # visitas & conversão: visitas únicas e quantidade de vendas por anúncio e dia
        vis_mks = lancamentos.carregar(cfg, d, "visitas", ano, mes)
        totais_vis = {mk: lancamentos.visitas_total(cfg, d, mk, ano, mes) for mk in (u.get("marketplaces") or []) + list(vis_mks)}
        for mk, tv in totais_vis.items():
            if tv:
                vis_mks.setdefault(mk, [])
        for mk, itens in vis_mks.items():
            nome = _nome_mk(cfg, mk)
            t = totais.setdefault(nome, {})
            t["Visitas"] = t.get("Visitas", 0.0) + sum(i["Visitas"] for i in itens)
            base_sku = df[(df["Marketplace"] == nome) & (df["SKU"] != "") & (df["Id anúncio"] != "")]
            sub = base_sku[base_sku["Marketplace venda"] == loja]
            sku_conta = sub.groupby("Id anúncio")["SKU"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(sub) else {}
            sku_geral = base_sku.groupby("Id anúncio")["SKU"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(base_sku) else {}
            vis = vds = 0.0
            dp_vis = lancamentos.carregar_depara(cfg, d, mk) if mk in lancamentos.DEPARA else {}
            man_vis = lancamentos.depara_manual(cfg, d, mk)          # SKU apontado à mão no app (vale primeiro)
            sem_ids, sem_nomes = {}, {}
            rep_vis = {}
            # para juntar com as vendas: o SKU decide; o anúncio passa a ser o mesmo que o SKU tem nas vendas da conta
            an_conta = sub.groupby("SKU")["Id anúncio"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(sub) else {}
            an_geral = base_sku.groupby("SKU")["Id anúncio"].agg(lambda x: x.value_counts().index[0]).to_dict() if len(base_sku) else {}
            sku_norm = {_sku_norm(k): k for k in list(an_geral)}
            pares = set(zip(base_sku["SKU"].astype(str), base_sku["Id anúncio"].astype(str)))
            sem_sku = 0
            for i in itens:
                if not i.get("Visitas"):
                    continue                        # só as visitas: as vendas vêm do relatório de vendas
                an = str(i.get("Id anúncio") or "")
                if mk == "shopee" and an.isdigit():
                    pesos_sh.setdefault(loja, {}).setdefault(an, [0.0, 0.0])[1] += i.get("Visitas") or 0
                if mk == "mercadolivre" and an.isdigit():
                    an = "MLB" + an                    # igual ao "# de anúncio" das vendas
                sku = str(i.get("SKU") or "").strip()
                nm = _nome_norm(i.get("Produto")) if mk == "tiktokshop" else ""
                m_ = man_vis.get(an) or man_vis.get(str(i.get("Id anúncio") or "")) or (man_vis.get(nm) if nm else None)
                if m_ and m_.get("sku"):
                    sku = m_["sku"]
                elif not sku and nm:
                    # TikTok: o relatório vem por "ID de produto" (não existe nas vendas): liga pelo nome do produto
                    par = (nomes_tiktok or {}).get(loja, {}).get(nm) or (nomes_tiktok or {}).get("*", {}).get(nm)
                    if par and (par[0] or par[1]):
                        sku, an = par[0] or par[1], par[1] or an
                if not sku:
                    sku = sku_conta.get(an) or sku_geral.get(an) or (dp_vis.get(an) or {}).get("sku") or \
                          (dp_vis.get(str(i.get("Id anúncio") or "")) or {}).get("sku") or ""
                if sku and sku not in an_geral:        # mesmo SKU escrito diferente (maiúsc./espaços/zeros à esquerda)
                    sku = sku_norm.get(_sku_norm(sku), sku)
                if sku:
                    if (sku, an) not in pares:             # anúncio que não vende esse SKU: usa o anúncio do SKU nas vendas
                        an = an_conta.get(sku) or an_geral.get(sku) or (sku if mk in lancamentos.DEPARA_ADS_POR_PLANILHA else an)
                else:
                    sem_sku += 1
                    sem_ids[an] = sem_ids.get(an, 0.0) + (i.get("Visitas") or 0)
                    if i.get("Produto"):
                        sem_nomes.setdefault(an, str(i["Produto"])[:150])
                sku = sku or "0"
                rp = _repartir(df, nome, loja, sku, rep_vis)
                for sku_p, w in rp:
                    an_p = an
                    if len(rp) > 1:
                        an_p = an_conta.get(sku_p) or an_geral.get(sku_p) or sku_p
                    base = {"Pedido": "", "Marketplace venda": loja, "SKU": sku_p, "Id anúncio": an_p, "Data": pd.Timestamp(i["Data"]),
                            "Marketplace": nome, "Type": "Organico"}
                    novas.append(dict(base, Categoria=VIS[0], **{"Sub Categoria": VIS[1], "Valor completo": i["Visitas"] * w}))
                vis += i.get("Visitas") or 0
                vds += i.get("Vendas") or 0
            # visitas TOTAIS do dia (Mercado Livre > Negócio): a soma do dia fica igual ao total; a diferença para os anúncios
            # entra como "Ajuste visitas totais" (SKU 0)
            vis_det, ajuste = vis, 0.0
            tv = totais_vis.get(mk) or {}
            if tv:
                soma_dia = {}
                for i in itens:
                    if i.get("Data"):
                        soma_dia[i["Data"]] = soma_dia.get(i["Data"], 0.0) + (i.get("Visitas") or 0)
                for dd, x in sorted(tv.items()):
                    dif = round(float(x.get("visitas") or 0) - soma_dia.get(dd, 0.0), 6)
                    if dif:
                        novas.append({"Pedido": "", "Marketplace venda": loja, "SKU": "0", "Id anúncio": "Ajuste visitas totais",
                                      "Data": pd.Timestamp(dd), "Marketplace": nome, "Type": "Organico",
                                      "Categoria": VIS[0], "Sub Categoria": VIS[1], "Valor completo": dif})
                        ajuste += dif
                vis += ajuste
                t["Visitas"] = t.get("Visitas", 0.0) + ajuste
            resumo["visitas"]["%s|%s" % (u["id"], mk)] = {"visitas": vis, "sem_sku": sem_sku,
                                                          "visitas_detalhe": vis_det, "ajuste_total": round(ajuste, 2),
                                                          "dias_total": len(tv),
                                                          "sem_sku_ids": {k: v for k, v in sorted(sem_ids.items(), key=lambda x: -x[1])},
                                                          "sem_sku_nomes": sem_nomes}
    if novas:
        df = pd.concat([df, pd.DataFrame(novas)], ignore_index=True)
    _ids_anuncio_shopee(cfg, cliente, df, pesos_sh, resumo)
    return df, {"totais": totais, "resumo": resumo}


def _ids_shopee_do_depara(depara):
    """De-para da Shopee ao contrário: {SKU como nas vendas (cortado no "-"): {ids de anúncio}}."""
    import lancamentos
    r = {}
    for an, v in (depara or {}).items():
        an = str(an).strip()
        if not an.isdigit():
            continue
        for s in lancamentos.skus_lista(_skus_como_vendas_shopee((v or {}).get("sku") or "")):
            s = str(s).strip()
            if s and s != "0":
                r.setdefault(s, set()).add(an)
    return r


def _ids_anuncio_shopee(cfg, cliente, df, pesos_sh, resumo):
    """Shopee: o "Id anúncio" das linhas que estão com o próprio SKU (vendas, ADS, afiliados e visitas) passa a ser o
    ID do anúncio do de-para da conta. Só muda essa coluna — SKU e valores ficam iguais.
    SKU em um anúncio só: esse ID. Em mais de um: o que teve mais gasto de ADS no mês na conta, depois mais visitas;
    sem nenhum dos dois, fica o SKU (aparece na conferência). SKU fora do de-para (ou SKU 0) fica como está."""
    import lancamentos
    if not len(df):
        return
    nome_sh = _nome_mk(cfg, "shopee")
    do_mk = df["Marketplace"] == nome_sh
    if not do_mk.any():
        return
    for u in comum.contas_do_cliente(cfg, cliente):
        d = lancamentos.dono(u)
        loja = d["loja"]
        depara = lancamentos.carregar_depara(cfg, d, "shopee")
        if not depara:
            continue
        por_sku = _ids_shopee_do_depara(depara)
        norm = {}
        for s in por_sku:
            norm.setdefault(_sku_norm(s), s)
        sku_txt = df["SKU"].astype(str)
        alvo = do_mk & (df["Marketplace venda"] == loja) & (sku_txt != "") & (sku_txt != "0") \
            & (df["Id anúncio"].astype(str) == sku_txt)
        if not alvo.any():
            continue
        pesos = pesos_sh.get(loja, {})
        escolha, sem, ambiguos, decididos = {}, [], {}, {}
        for s in sorted(set(sku_txt[alvo])):
            ids = por_sku.get(s) or por_sku.get(norm.get(_sku_norm(s), ""), set())
            if not ids:
                sem.append(s)
                continue
            if len(ids) == 1:
                escolha[s] = next(iter(ids))
                continue
            ordem = sorted(ids, key=lambda a: (-(pesos.get(a) or [0, 0])[0], -(pesos.get(a) or [0, 0])[1], a))
            p0 = pesos.get(ordem[0]) or [0, 0]
            p1 = pesos.get(ordem[1]) or [0, 0]
            if (p0[0] or p0[1]) and (p0[0], p0[1]) != (p1[0], p1[1]):
                escolha[s] = ordem[0]
                decididos[s] = ordem[0]
            else:
                ambiguos[s] = sorted(ids)
        if escolha:
            df.loc[alvo, "Id anúncio"] = sku_txt[alvo].map(lambda x: escolha.get(x, x))
        resumo["ids_shopee"]["%s|shopee" % u["id"]] = {
            "com_id": len(escolha), "sem_anuncio": len(sem), "sem_anuncio_ex": sem[:8],
            "varios_escolhidos": len(decididos), "ambiguos": len(ambiguos),
            "ambiguos_ex": {k: v for k, v in list(ambiguos.items())[:8]}}


# ---------------------------------------------------------------- geração
def _tem_lancamentos(cfg, cliente, ano, mes):
    """Algo subido no app no mês (ADS, afiliados ou visitas) em alguma conta do cliente."""
    import lancamentos
    for u in comum.contas_do_cliente(cfg, cliente):
        d = lancamentos.dono(u)
        for tipo in lancamentos.TIPOS:
            if any(ls for ls in lancamentos.carregar(cfg, d, tipo, ano, mes).values()):
                return True
        if any(lancamentos.visitas_total(cfg, d, mk, ano, mes) for mk in u.get("marketplaces") or []):
            return True
        if lancamentos.linhas_ajuste(cfg, d, ano, mes) or lancamentos.linhas_ajuste_afiliados(cfg, d, ano, mes):
            return True
    return False


def _tratado_da_conta(cfg, unidade, mk, dia):
    import tratador
    return tratador.caminho_saida(cfg, unidade, mk, dia)


def _tratado_anterior(cfg, unidade, mk, ano, mes, dia):
    """Último tratado do mês antes de `dia` (os tratados são acumulados do mês: o de ontem já tem tudo até anteontem).
    Usado quando o relatório do dia recebeu OK manual (ex.: sem vendas hoje). Devolve (caminho, data) ou (None, None)."""
    d = dia - dt.timedelta(days=1)
    limite = dt.date(ano, mes, 2)     # a pasta do dia 1 do mês tem o mês anterior
    while d >= limite:
        p = _tratado_da_conta(cfg, unidade, mk, d)
        if os.path.isfile(p):
            return p, d
        d -= dt.timedelta(days=1)
    return None, None


_TRAVA_GERAR = __import__("threading").RLock()


def gerar(cfg, cliente, ano=None, mes=None, hoje=None, origem="manual"):
    """Um pacote de cada vez (vários computadores/telas podem pedir ao mesmo tempo)."""
    with _TRAVA_GERAR:
        try:
            return _gerar(cfg, cliente, ano, mes, hoje, origem)
        finally:
            comum.devolver_memoria()


def _gerar(cfg, cliente, ano=None, mes=None, hoje=None, origem="manual"):
    """Gera o consolidado de um cliente (grupo). Nunca levanta exceção: devolve o resultado."""
    import pandas as pd
    hoje = hoje or dt.date.today()
    if not ano:
        ano, mes = mes_atual(hoje)
    res = {"cliente_id": cliente["id"], "cliente": cliente["nome"], "mes": "%04d-%02d" % (ano, mes),
           "gerado": dt.datetime.now().isoformat(timespec="seconds"), "origem": origem, "ok": False}
    if os.path.isfile(pronto_arq(cliente["id"], ano, mes)):
        return _gerar_de_pronto(cfg, cliente, ano, mes, hoje, res)
    try:
        dia, parcial = dia_fonte(cfg, ano, mes, hoje)
        tem_lanc = _tem_lancamentos(cfg, cliente, ano, mes)
        if not dia and not tem_lanc:
            res["erro"] = "Não há pasta de relatórios para %02d/%04d." % (mes, ano)
            return _registrar(res)
        res["fonte"] = comum.rotulo_dia(dia) if dia else None
        res["parcial"] = parcial
        medidas_mk, totais, extras, faltando, fontes, anteriores = {}, {}, {}, [], [], []
        nomes_tiktok = {}
        for u in (comum.contas_do_cliente(cfg, cliente) if dia else []):
            for mk in u.get("marketplaces", []):
                if mk not in MAPA:
                    continue
                p = _tratado_da_conta(cfg, u, mk, dia)
                if not os.path.isfile(p):
                    # sem tratado no dia: entra o último tratado do mês (OK manual = dia sem vendas; senão = ainda não chegou)
                    ok_manual = bool(comum.ok_vendas(u["id"], mk, dia))
                    p, d_ant = _tratado_anterior(cfg, u, mk, ano, mes, dia)
                    if not p:
                        if not ok_manual:
                            faltando.append("%s (%s)" % (MAPA[mk]["nome"], u["conta_nome"]))
                        continue
                    anteriores.append({"marketplace": MAPA[mk]["nome"], "conta": u["conta_nome"], "dia": comum.rotulo_dia(d_ant),
                                       "data": d_ant.isoformat(), "motivo": "ok_manual" if ok_manual else "sem_relatorio"})
                df = _ler_tratado(p, mk, ano, mes)
                if mk == "tiktokshop" and "Product Name" in df.columns:
                    # o ADS da TikTok vem por "ID do produto" (não existe nas vendas): liga pelo nome do produto
                    loja_u = u.get("loja") or cliente["nome"]
                    for nm, g in df.groupby(df["Product Name"].astype(str).map(_nome_norm)):
                        if not nm:
                            continue
                        sku_v = g[MAPA[mk]["sku"]].astype(str).replace("nan", "")
                        id_v = g[MAPA[mk]["anuncio"]].astype(str).replace("nan", "")
                        par = (sku_v[sku_v != ""].value_counts().index[0] if (sku_v != "").any() else "",
                               id_v[id_v != ""].value_counts().index[0] if (id_v != "").any() else "")
                        nomes_tiktok.setdefault(loja_u, {}).setdefault(nm, par)
                        nomes_tiktok.setdefault("*", {}).setdefault(nm, par)
                med, tot, ext = _linhas_do_tratado(df, mk)
                nome = MAPA[mk]["nome"]
                if med is not None:
                    # cada conta sai com o nome dela (ex.: "KLM" e "KLM 2"); com uma conta só, é o nome do cliente
                    medidas_mk.setdefault(nome, []).append(med.assign(**{"Marketplace venda": u.get("loja") or cliente["nome"]}))
                t = totais.setdefault(nome, {})
                for k, v in tot.items():
                    if v is None or (k in t and t[k] is None):
                        t[k] = None
                    else:
                        t[k] = t.get(k, 0.0) + v
                e = extras.setdefault(nome, {"vendas": 0, "canceladas": 0, "fora": 0, "sem_custo": 0, "cancelada_sem_valor": 0})
                for k, v in ext.items():
                    e[k] += v
                fontes.append(os.path.relpath(p, comum.raiz(cfg)))
        res["faltando"] = faltando
        res["fontes"] = fontes
        res["anteriores"] = anteriores
        sem_rel = [a for a in anteriores if a["motivo"] == "sem_relatorio"]
        if sem_rel:
            res["aviso_anteriores"] = "Sem relatório de %s: usada a última versão — %s." % (
                comum.rotulo_dia(dia), "; ".join("%s (%s) de %s" % (a["marketplace"], a["conta"], a["dia"]) for a in sem_rel))
            if not any(os.path.isfile(_tratado_da_conta(cfg, u, mk, dia)) for u in comum.contas_do_cliente(cfg, cliente)
                       for mk in u.get("marketplaces", []) if mk in MAPA):
                ds = sorted({dt.date.fromisoformat(a["data"]) for a in anteriores})
                res["dados_de"] = comum.rotulo_dia(ds[0]) if len(ds) == 1 else "%s a %s" % (comum.rotulo_dia(ds[0]), comum.rotulo_dia(ds[-1]))
        if not medidas_mk and not tem_lanc:
            res["erro"] = "Nenhum tratado encontrado para %02d/%04d%s." % (
                mes, ano, (" (faltam: %s)" % ", ".join(faltando)) if faltando else "")
            return _registrar(res)

        desc = _descricoes(cliente, ano, mes)
        frames = []
        for nome, lst in medidas_mk.items():
            for f in lst:
                frames.append(f.assign(Marketplace=nome))
        if frames:
            df = pd.concat(frames, ignore_index=True)
        else:
            # sem vendas tratadas: o pacote sai só com o que foi subido no app (ADS, afiliados, visitas & conversão)
            df = pd.DataFrame({"Pedido": pd.Series(dtype=object), "SKU": pd.Series(dtype=object),
                               "Id anúncio": pd.Series(dtype=object), "Data": pd.Series(dtype="datetime64[ns]"),
                               "Categoria": pd.Series(dtype=object), "Sub Categoria": pd.Series(dtype=object),
                               "Valor completo": pd.Series(dtype=float), "Marketplace venda": pd.Series(dtype=object),
                               "Marketplace": pd.Series(dtype=object)})
            res["sem_vendas"] = True
            res["aviso_sem_vendas"] = ("Pacote montado sem as vendas tratadas%s: só com o que foi subido no sistema "
                                       "(ADS, afiliados, visitas & conversão)." % (
                                           (" (faltam: %s)" % ", ".join(faltando)) if faltando else ""))
        df["Type"] = "Organico"
        df, lanc = _aplicar_lancamentos(cfg, cliente, df, ano, mes, nomes_tiktok)
        res["lancamentos"] = lanc["resumo"]
        df = df.drop(columns=["Pedido"])
        chave = ["Marketplace", "Marketplace venda", "SKU", "Id anúncio", "Data", "Categoria", "Sub Categoria", "Type"]
        df = df.groupby(chave, as_index=False, sort=False)["Valor completo"].sum()
        df = df[df["Valor completo"].round(6) != 0]
        df["Datatype"] = "Actual"
        # forecast definido do mês (travado): entra na mesma base com Datatype "Forecast"
        try:
            import forecast
            fc = forecast.linhas_consolidado(cliente["id"], ano, mes)
        except Exception as e:
            fc = []
            comum.log("[%s] forecast não incluído no consolidado: %s" % (cliente["nome"], e))
        if fc:
            dff = pd.DataFrame(fc, columns=["Marketplace", "SKU", "Data", "Categoria", "Sub Categoria", "Valor completo"])
            dff["Data"] = pd.to_datetime(dff["Data"])
            dff["Id anúncio"], dff["Type"], dff["Datatype"] = "", "Organico", "Forecast"
            dff["Marketplace venda"] = cliente["nome"]
            df = pd.concat([df, dff[df.columns]], ignore_index=True)
        res["forecast_linhas"] = len(fc)
        ordem = {cs: i for i, cs in enumerate([cs for _, cs in METRICAS] + [AFI, PUB, CLQ, ADF, VIS, QVV])}
        df["_o"] = [ordem.get((c, s), 99) for c, s in zip(df["Categoria"], df["Sub Categoria"])]
        ordem_mk = {v["nome"]: i for i, v in enumerate(MAPA.values())}
        df["_m"] = df["Marketplace"].map(ordem_mk).fillna(99)
        df["_d"] = (df["Datatype"] != "Actual").astype(int)
        df = df.sort_values(["_d", "_m", "Data", "SKU", "Id anúncio", "_o"]).drop(columns=["_o", "_m", "_d"])

        sem_desc = sorted({s for s in df.loc[df["Datatype"] == "Actual", "SKU"] if s and s not in desc})
        res["sem_descricao"] = len(sem_desc)
        res["exemplos_sem_descricao"] = sem_desc[:8]
        destino = os.path.join(comum.raiz(cfg), comum.rotulo_dia(hoje), cliente["nome"], nome_arquivo(cliente, ano, mes))
        _gravar(df, cliente, desc, destino, hoje)

        # conferência: relê o arquivo gravado e compara com os totais tirados direto dos tratados
        lido = ler_excel(destino, sheet_name="DB")
        lido = lido[lido["Datatype"] == "Actual"]
        conf, ok_conf = [], True
        for nome, tot in totais.items():
            sub = lido[lido["Marketplace"] == nome]
            soma = lambda cat, sc: float(sub[(sub["Categoria"] == cat) & (sub["Sub Categoria"] == sc)]["Valor completo"].sum())
            linhas_conf = [(rot, tot.get(rot, 0.0), soma(cat, sc)) for rot, (cat, sc) in METRICAS]
            if tot.get("Líquido (conferência)") is not None:
                # independente: o Líquido que o tratador calculou x faturamento − despesas do consolidado
                # (as canceladas somam zero: entram e saem com o mesmo valor)
                liq = sum(soma(*k) for k in (FAT, CANC, COM, FRE, REB))
                linhas_conf.append(("Líquido (conferência)", tot["Líquido (conferência)"], liq))
            for rot, v_trat, v_cons in linhas_conf:
                dif = round(v_cons - v_trat, 2)
                if abs(dif) > TOLERANCIA:
                    ok_conf = False
                if v_trat or v_cons:
                    conf.append({"marketplace": nome, "metrica": rot, "tratado": round(v_trat, 2),
                                 "consolidado": round(v_cons, 2), "diferenca": dif})
        # ADS e afiliados: o que foi subido x o que entrou no consolidado
        sinal = {"ADS gasto": (PUB, -1), "Cliques": (CLQ, 1), "ADS F (receita)": (ADF, 1), "Afiliados (comissão)": (AFI, -1),
                 "Visitas": (VIS, 1), "Vendas (visitas)": (QVV, 1)}
        for nome, tot in lanc["totais"].items():
            sub = lido[lido["Marketplace"] == nome]
            for rot, v_sub in tot.items():
                (cat, sc), sg = sinal[rot]
                v_cons = sg * float(sub[(sub["Categoria"] == cat) & (sub["Sub Categoria"] == sc)]["Valor completo"].sum())
                dif = round(v_cons - v_sub, 2)
                if abs(dif) > TOLERANCIA:
                    ok_conf = False
                if v_sub or v_cons:
                    conf.append({"marketplace": nome, "metrica": rot, "tratado": round(v_sub, 2), "consolidado": round(v_cons, 2),
                                 "diferenca": dif, "fonte": "subido"})
        res["conferencia"] = conf
        res["conferido"] = ok_conf
        res["resumo"] = {n: dict(extras[n], faturamento=round(totais[n]["Faturamento bruto"] + totais[n]["Cancelamento"], 2))
                         for n in totais}

        res.update(ok=True, arquivo=destino, linhas=int((df["Datatype"] == "Actual").sum()))
        if fc and (ano, mes) == (hoje.year, hoje.month):
            try:   # sinal de atenção: realizado x forecast definido
                _registrar(res)
                cmp = forecast.comparar(cfg, cliente, hoje)
                res["forecast_alertas"] = cmp["alertas"] if cmp else []
                res["forecast_versao"] = cmp["versao"] if cmp else None
            except Exception as e:
                comum.log("[%s] comparação com o forecast falhou: %s" % (cliente["nome"], e))
        return _registrar(res)
    except PermissionError:
        res["erro"] = "Não consegui gravar: feche o relatório consolidado no Excel e gere de novo."
        return _registrar(res)
    except Exception as e:
        res["erro"] = "Falha ao gerar: %s" % str(e)[:500]
        return _registrar(res)


# ---------------------------------------------------------------- pacote final pronto (subido no sistema)
OBRIG_PRONTO = ["Marketplace", "Data", "Categoria", "Sub Categoria", "Valor completo"]


def _pasta_prontos(cliente_id):
    return os.path.join(comum.PASTA_SISTEMA, "dados", "pacotes_prontos", re.sub(r"[^\w.-]+", "_", str(cliente_id)))


def pronto_arq(cliente_id, ano, mes):
    return os.path.join(_pasta_prontos(cliente_id), "%04d-%02d.xlsx" % (ano, mes))


def pronto_info(cliente_id, ano, mes):
    try:
        with open(pronto_arq(cliente_id, ano, mes)[:-5] + ".json", "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def _ler_pronto(caminho, ano, mes):
    """Lê o pacote pronto (aba DB ou a primeira aba com as colunas do pacote). Devolve (df, desc_do_arquivo, avisos)."""
    import pandas as pd
    abas = ler_excel(caminho, sheet_name=None)
    df = None
    for nome in (["DB"] if "DB" in abas else []) + [n for n in abas if n != "DB"]:
        d = abas[nome]
        d.columns = [str(c).strip() for c in d.columns]
        if all(c in d.columns for c in OBRIG_PRONTO):
            df = d
            break
    if df is None:
        raise ValueError("não achei a aba do pacote final: precisa das colunas %s (como o pacote que o sistema gera, aba DB)."
                         % ", ".join(OBRIG_PRONTO))
    avisos = []
    df = df.copy()
    for c, padrao in (("Marketplace venda", ""), ("SKU", ""), ("Id anúncio", ""), ("Type", "Organico"), ("Datatype", "Actual")):
        if c not in df.columns:
            df[c] = padrao
    def txt(v):
        if v is None or (isinstance(v, float) and v != v):
            return ""
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).strip()
    for c in ("Marketplace", "Marketplace venda", "SKU", "Id anúncio", "Categoria", "Sub Categoria", "Type", "Datatype"):
        df[c] = df[c].map(txt)
    df.loc[df["SKU"] == "0", "SKU"] = ""
    df.loc[df["Type"] == "", "Type"] = "Organico"
    df.loc[df["Datatype"] == "", "Datatype"] = "Actual"
    df["Valor completo"] = pd.to_numeric(df["Valor completo"], errors="coerce")
    if not pd.api.types.is_datetime64_any_dtype(df["Data"]):
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    ruins = df["Data"].isna() | df["Valor completo"].isna() | (df["Marketplace"] == "") | (df["Categoria"] == "")
    if ruins.any():
        avisos.append("%d linha(s) sem data, valor, marketplace ou categoria ficaram de fora." % int(ruins.sum()))
        df = df[~ruins]
    fora = (df["Datatype"] == "Actual") & ((df["Data"].dt.year != ano) | (df["Data"].dt.month != mes))
    if fora.any():
        avisos.append("%d linha(s) de fora de %02d/%04d ficaram de fora." % (int(fora.sum()), mes, ano))
        df = df[~fora]
    if not (df["Datatype"] == "Actual").any():
        raise ValueError("o pacote não tem nenhuma linha de %02d/%04d." % (mes, ano))
    desc = {}
    if "Descrição" in df.columns:
        c2 = df["Categoria2"] if "Categoria2" in df.columns else None
        for i, (sku, d) in enumerate(zip(df["SKU"], df["Descrição"])):
            d = txt(d)
            if sku and d and d != "X" and sku not in desc:
                desc[sku] = (d, (txt(c2.iloc[i]) if c2 is not None else "") or "X")
    cols = ["Marketplace", "Marketplace venda", "SKU", "Id anúncio", "Data", "Categoria", "Sub Categoria", "Type", "Valor completo", "Datatype"]
    return df[cols].reset_index(drop=True), desc, avisos


def guardar_pronto(cfg, cliente, ano, mes, nome, conteudo, por=""):
    """Guarda o pacote final pronto do cliente no mês (substitui o anterior) e refaz o pacote do sistema a partir dele."""
    os.makedirs(_pasta_prontos(cliente["id"]), exist_ok=True)
    destino = pronto_arq(cliente["id"], ano, mes)
    tmp = destino + ".novo.xlsx"
    with open(tmp, "wb") as f:
        f.write(conteudo)
    try:
        _ler_pronto(tmp, ano, mes)
    except ValueError:
        os.remove(tmp)
        raise
    except Exception as e:
        os.remove(tmp)
        raise ValueError("não consegui ler o arquivo como Excel (%s)." % str(e)[:120])
    if os.path.isfile(destino):
        hist = os.path.join(_pasta_prontos(cliente["id"]), "_anteriores")
        os.makedirs(hist, exist_ok=True)
        shutil.copy2(destino, os.path.join(hist, "%04d-%02d_%s.xlsx" % (ano, mes, dt.datetime.now().strftime("%Y%m%d-%H%M%S"))))
    os.replace(tmp, destino)
    with open(destino[:-5] + ".json", "w", encoding="utf-8") as f:
        json.dump({"nome": os.path.basename(nome or ""), "enviado": dt.datetime.now().isoformat(timespec="seconds"), "por": por or ""},
                  f, ensure_ascii=False)
    comum.log("[%s] PACOTE PRONTO %02d/%04d subido: %s" % (cliente["nome"], mes, ano, nome))
    return gerar(cfg, cliente, ano, mes, origem="pacote pronto subido")


def remover_pronto(cfg, cliente, ano, mes):
    """Tira o pacote pronto (vai para _anteriores) e volta a montar o pacote pelo sistema."""
    destino = pronto_arq(cliente["id"], ano, mes)
    if not os.path.isfile(destino):
        raise ValueError("Não há pacote pronto neste mês.")
    hist = os.path.join(_pasta_prontos(cliente["id"]), "_anteriores")
    os.makedirs(hist, exist_ok=True)
    shutil.move(destino, os.path.join(hist, "%04d-%02d_%s_removido.xlsx" % (ano, mes, dt.datetime.now().strftime("%Y%m%d-%H%M%S"))))
    try:
        os.remove(destino[:-5] + ".json")
    except OSError:
        pass
    comum.log("[%s] PACOTE PRONTO %02d/%04d removido" % (cliente["nome"], mes, ano))
    return gerar(cfg, cliente, ano, mes, origem="pacote pronto removido")


def _gerar_de_pronto(cfg, cliente, ano, mes, hoje, res):
    """Pacote final a partir do pacote pronto subido: o arquivo vale como está (vendas, ADS, afiliados…); só as VISITAS
    alimentadas no sistema entram por cima (substituem as visitas do arquivo nos marketplaces que têm visitas no sistema)."""
    import pandas as pd
    info = pronto_info(cliente["id"], ano, mes) or {}
    res.update(pronto=True, pronto_info=info, fonte="pacote pronto", parcial=False, faltando=[], anteriores=[])
    try:
        df, desc_arq, avisos = _ler_pronto(pronto_arq(cliente["id"], ano, mes), ano, mes)
        original = df.copy()
        # visitas alimentadas no sistema
        vis_sis, lanc = None, {"totais": {}, "resumo": {}}
        try:
            base = df[(df["Datatype"] == "Actual") & (df["Categoria"] != VIS[0])].copy()
            base["Pedido"] = ""
            depois, lanc = _aplicar_lancamentos(cfg, cliente, base, ano, mes)
            vis_sis = depois[(depois["Categoria"] == VIS[0]) & (depois["Sub Categoria"] == VIS[1])]   # a base não tem visitas
        except Exception as e:
            avisos.append("Visitas do sistema não entraram: %s" % str(e)[:200])
        mks_vis = sorted(set(vis_sis["Marketplace"])) if vis_sis is not None and len(vis_sis) else []
        if mks_vis:
            tirar = (df["Categoria"] == VIS[0]) & df["Marketplace"].isin(mks_vis) & (df["Datatype"] == "Actual")
            if tirar.any():
                avisos.append("Visitas do arquivo trocadas pelas do sistema em: %s." % ", ".join(mks_vis))
            v = vis_sis.drop(columns=[c for c in vis_sis.columns if c not in df.columns]).assign(Datatype="Actual")
            df = pd.concat([df[~tirar], v[df.columns]], ignore_index=True)
        res["visitas_do_sistema"] = mks_vis
        res["lancamentos"] = {"visitas": (lanc.get("resumo") or {}).get("visitas", {})}
        # forecast definido do mês: entra se o arquivo não trouxer o seu
        if not (df["Datatype"] == "Forecast").any():
            try:
                import forecast
                fc = forecast.linhas_consolidado(cliente["id"], ano, mes)
            except Exception:
                fc = []
            if fc:
                dff = pd.DataFrame(fc, columns=["Marketplace", "SKU", "Data", "Categoria", "Sub Categoria", "Valor completo"])
                dff["Data"] = pd.to_datetime(dff["Data"])
                dff["Id anúncio"], dff["Type"], dff["Datatype"], dff["Marketplace venda"] = "", "Organico", "Forecast", cliente["nome"]
                df = pd.concat([df, dff[df.columns]], ignore_index=True)
        desc = dict(_descricoes(cliente, ano, mes))
        desc.update(desc_arq)
        destino = os.path.join(comum.raiz(cfg), comum.rotulo_dia(hoje), cliente["nome"], nome_arquivo(cliente, ano, mes))
        _gravar(df, cliente, desc, destino, hoje)
        # conferência: arquivo final x pacote pronto (todas as medidas) e visitas x o que foi alimentado
        lido = ler_excel(destino, sheet_name="DB")
        lido = lido[lido["Datatype"] == "Actual"]
        org = original[original["Datatype"] == "Actual"]
        conf, ok_conf = [], True
        for nome in sorted(set(org["Marketplace"]) | set(lido["Marketplace"])):
            a, b = org[org["Marketplace"] == nome], lido[lido["Marketplace"] == nome]
            pares = sorted(set(zip(a["Categoria"], a["Sub Categoria"])) | set(zip(b["Categoria"], b["Sub Categoria"])))
            for cat, sc in pares:
                v_b = float(b[(b["Categoria"] == cat) & (b["Sub Categoria"] == sc)]["Valor completo"].sum())
                if cat == VIS[0] and nome in mks_vis:
                    v_a = float(vis_sis[vis_sis["Marketplace"] == nome]["Valor completo"].sum())
                    rot, fonte = "Visitas (sistema)", "subido"
                else:
                    v_a = float(a[(a["Categoria"] == cat) & (a["Sub Categoria"] == sc)]["Valor completo"].sum())
                    rot, fonte = (sc if sc == cat else "%s · %s" % (cat, sc)), "pronto"
                dif = round(v_b - v_a, 2)
                if abs(dif) > TOLERANCIA:
                    ok_conf = False
                if v_a or v_b:
                    conf.append({"marketplace": nome, "metrica": rot, "tratado": round(v_a, 2), "consolidado": round(v_b, 2),
                                 "diferenca": dif, "fonte": fonte})
        resumo = {}
        for nome in sorted(set(lido["Marketplace"])):
            b = lido[lido["Marketplace"] == nome]
            s_ = lambda k: float(b[(b["Categoria"] == k[0]) & (b["Sub Categoria"] == k[1])]["Valor completo"].sum())
            if s_(FAT) or s_(QTD):
                resumo[nome] = {"vendas": int(round(s_(QTD))), "canceladas": 0, "fora": 0, "sem_custo": 0, "cancelada_sem_valor": 0,
                                "faturamento": round(s_(FAT) + s_(CANC), 2)}
        res.update(conferencia=conf, conferido=ok_conf, resumo=resumo, sem_descricao=0, exemplos_sem_descricao=[],
                   ok=True, arquivo=destino, linhas=int((df["Datatype"] == "Actual").sum()),
                   aviso_pronto=" ".join(avisos) if avisos else "")
        return _registrar(res)
    except PermissionError:
        res["erro"] = "Não consegui gravar: feche o relatório consolidado no Excel e gere de novo."
        return _registrar(res)
    except Exception as e:
        res["erro"] = "Falha ao montar a partir do pacote pronto: %s" % str(e)[:500]
        return _registrar(res)


def _gravar(df, cliente, desc, destino, hoje):
    """Grava o arquivo (xlsxwriter, bem mais rápido; sem ele, openpyxl). Pode ter dezenas de milhares de linhas."""
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        return _gravar_openpyxl(df, cliente, desc, destino, hoje)
    import xlsxwriter
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    tmp = destino + ".novo.xlsx"
    wb = xlsxwriter.Workbook(tmp, {"constant_memory": True})
    ws = wb.add_worksheet("DB")
    f_cab = wb.add_format({"bold": True})
    f_data = wb.add_format({"num_format": "dd/mm/yyyy"})
    f_num = wb.add_format({"num_format": "#,##0.00"})
    for j, larg in enumerate((14, 22, 16, 20, 11, 24, 24, 14, 10, 40, 18, 9, 11, 13, 12)):
        ws.set_column(j, j, larg)
    ws.freeze_panes(1, 0)
    ws.write_row(0, 0, COLUNAS, f_cab)
    hoje_dt = dt.datetime(hoje.year, hoje.month, hoje.day)
    cols = ["Marketplace", "SKU", "Id anúncio", "Data", "Categoria", "Sub Categoria", "Type", "Valor completo", "Datatype", "Marketplace venda"]
    if "Datatype" not in df.columns:
        df = df.assign(Datatype="Actual")
    if "Marketplace venda" not in df.columns:
        df = df.assign(**{"Marketplace venda": cliente["nome"]})
    for i, (mk, sku, anuncio, data, cat, sub, tipo, valor, datatype, nome) in enumerate(df[cols].itertuples(index=False, name=None), 1):
        d, c2 = desc.get(sku, ("X", "X"))
        if hasattr(data, "to_pydatetime"):
            data = data.to_pydatetime()
        ws.write_string(i, 0, mk)
        ws.write_string(i, 1, str(nome or cliente["nome"]))
        if sku:
            ws.write_string(i, 2, str(sku))
        else:
            ws.write_number(i, 2, 0)
        if anuncio:
            ws.write_string(i, 3, str(anuncio))
        ws.write_datetime(i, 4, data, f_data)
        ws.write_string(i, 5, cat)
        ws.write_string(i, 6, sub)
        ws.write_number(i, 7, round(float(valor), 6), f_num)
        ws.write_row(i, 8, [tipo, d, c2, datatype])
        ws.write_datetime(i, 12, hoje_dt, f_data)
        ws.write_string(i, 13, "Coluna Ponto")
    wb.close()
    os.replace(tmp, destino)


def _gravar_openpyxl(df, cliente, desc, destino, hoje):
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("DB")
    for col, larg in zip("ABCDEFGHIJKLMNO", (14, 22, 16, 20, 11, 24, 24, 14, 10, 40, 18, 9, 11, 13, 12)):
        ws.column_dimensions[col].width = larg
    ws.freeze_panes = "A2"
    negrito = Font(bold=True)
    cab = []
    for nome in COLUNAS:
        c = WriteOnlyCell(ws, value=nome)
        c.font = negrito
        cab.append(c)
    ws.append(cab)
    hoje_dt = dt.datetime(hoje.year, hoje.month, hoje.day)
    c_hoje = lambda: _celula(ws, hoje_dt, "dd/mm/yyyy")
    cols = ["Marketplace", "SKU", "Id anúncio", "Data", "Categoria", "Sub Categoria", "Type", "Valor completo", "Datatype", "Marketplace venda"]
    if "Datatype" not in df.columns:
        df = df.assign(Datatype="Actual")
    if "Marketplace venda" not in df.columns:
        df = df.assign(**{"Marketplace venda": cliente["nome"]})
    for mk, sku, anuncio, data, cat, sub, tipo, valor, datatype, venda in df[cols].itertuples(index=False, name=None):
        d, c2 = desc.get(sku, ("X", "X"))
        if hasattr(data, "to_pydatetime"):
            data = data.to_pydatetime()
        ws.append([mk, venda or cliente["nome"], sku or 0, anuncio or None, _celula(ws, data, "dd/mm/yyyy"), cat, sub,
                   _celula(ws, round(float(valor), 6), "#,##0.00"), tipo, d, c2, datatype, c_hoje(), "Coluna Ponto", None])
    tmp = destino + ".novo.xlsx"
    wb.save(tmp)
    os.replace(tmp, destino)


def _celula(ws, valor, formato):
    from openpyxl.cell import WriteOnlyCell
    c = WriteOnlyCell(ws, value=valor)
    c.number_format = formato
    return c


def ler_excel(caminho, **kw):
    """pandas.read_excel com o leitor rápido (calamine) quando disponível."""
    import pandas as pd
    try:
        return pd.read_excel(caminho, engine="calamine", **kw)
    except (ImportError, ValueError):
        return pd.read_excel(caminho, **kw)


def cobertura_vendas(cfg, cliente, ano, mes, hoje=None):
    """Para a visão mês a mês: quais tratados de vendas existem na pasta que cobre o mês."""
    dia, parcial = dia_fonte(cfg, ano, mes, hoje)
    itens = []
    for u in comum.contas_do_cliente(cfg, cliente):
        if not u.get("ativo", True):
            continue
        for mk in u.get("marketplaces", []):
            if mk not in MAPA:
                continue
            tem = bool(dia) and os.path.isfile(_tratado_da_conta(cfg, u, mk, dia))
            ok = bool(dia) and not tem and bool(comum.ok_vendas(u["id"], mk, dia))
            itens.append({"conta": u["conta_nome"], "conta_id": u["id"], "mk": mk, "nome": MAPA[mk]["nome"], "tem": tem, "ok_manual": ok})
    return {"fonte": dia.isoformat() if dia else None, "parcial": parcial, "itens": itens}


# ---------------------------------------------------------------- estado
def _arq_estado():
    return os.path.join(comum.PASTA_SISTEMA, "dados", "consolidados.json")


def ler_estado():
    try:
        with open(_arq_estado(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _registrar(res):
    est = ler_estado()
    ant = (est.get(res["cliente_id"]) or {}).get(res["mes"])
    if not res.get("ok") and ant:
        # geração falhou (ex.: hoje ainda sem tratados): guarda o último consolidado bom para o painel continuar mostrando
        bom = ant if ant.get("ok") else ant.get("ultimo_ok")
        if bom and bom.get("arquivo"):
            res["ultimo_ok"] = {k: v for k, v in bom.items() if k != "ultimo_ok"}
    est.setdefault(res["cliente_id"], {})[res["mes"]] = res
    os.makedirs(os.path.dirname(_arq_estado()), exist_ok=True)
    tmp = _arq_estado() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(est, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _arq_estado())
    comum.registrar_evento({"tipo": "consolidado", "cliente": res["cliente_id"], "mes": res["mes"], "ok": res.get("ok"),
                            "conferido": res.get("conferido"), "erro": res.get("erro"), "origem": res.get("origem")})
    if res.get("ok"):
        comum.log("[%s] CONSOLIDADO %s: %d linhas, conferência %s" % (
            res["cliente"], res["mes"], res.get("linhas", 0), "ok" if res.get("conferido") else "COM DIFERENÇAS"))
    else:
        comum.log("[%s] CONSOLIDADO %s não gerado: %s" % (res["cliente"], res["mes"], res.get("erro")))
    return res


def gerar_clientes(cfg, cliente_ids=None, ano=None, mes=None, origem="automático"):
    saida = []
    for c in cfg.get("clientes", []):
        if cliente_ids and c["id"] not in cliente_ids:
            continue
        if not c.get("ativo", True) and not cliente_ids:
            continue
        saida.append(gerar(cfg, c, ano, mes, origem=origem))
    return saida


def automatico(cfg, cliente_ids, origem="automático"):
    """Chamado depois de tratar: refaz o consolidado do mês atual dos clientes afetados."""
    if not cfg.get("consolidado", {}).get("automatico", True) or not cliente_ids:
        return []
    try:
        return gerar_clientes(cfg, sorted(set(cliente_ids)), origem=origem)
    except Exception as e:
        comum.log("CONSOLIDADO automático falhou: %s" % e)
        return []
