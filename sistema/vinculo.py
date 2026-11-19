# -*- coding: utf-8 -*-
"""Vínculo forçado por conta, para clientes que NÃO cadastram SKU nas plataformas (cadastro: cliente.sem_sku = true).

Roda na montagem do pacote final, depois de ADS, afiliados e visitas entrarem e antes de juntar as linhas.
Toda linha de ADS / afiliados / visitas / cliques que ficaria sem produto (SKU fora da tabela de custos: Descrição "X")
é ligada a um produto VENDIDO da mesma conta (Marketplace + Marketplace venda), pela cascata abaixo. Para no
primeiro nível que casar e grava o nível em "Metodo_vinculo":

  N1 ID do anúncio      ID do anúncio (normalizado) igual ao das vendas
  N2 De-para            ID do anúncio no de-para do canal (Shopee mass_update, TikTok product_id -> SKUs, ...)
  N3 SKU texto          SKU (ou o texto no lugar do ID) igual ao SKU das vendas, normalizado (sem acento, pontuação, caixa)
  N4 Anúncio+variação   SKU é só o nome da variação ("Brilho", "2kg"): vendas cujo SKU termina com essa variação
  N5 Título             título igual, por prefixo ou parecido (>= 0,90) ao das vendas
  N6 Campanha           ID com nome de campanha/grupo: rateio pelo faturamento da conta no dia
  N7 Rateio conta       o que sobrar: rateio pelo faturamento da conta no dia (afiliados: faturamento Type Afiliados)

Rateio: valores em R$ são divididos pelo faturamento de cada produto no dia (sem venda no dia: no mês; sem nada: partes
iguais) e a soma fica igual ao original (resíduo no maior). Contagens (visitas, cliques, quantidade de vendas) não são
divididas: as frações de um mesmo anúncio e dia são somadas de novo e o total inteiro vai para o produto de maior peso.

Conferência obrigatória: soma por Marketplace x Conta x Categoria x Sub Categoria x Dia igual à de antes (R$ 0,01)
e nenhuma linha sem produto depois da cascata. Se falhar, levanta erro e o pacote não é gravado.
"""
import difflib
import re
import unicodedata

METODOS = ["Venda", "Original", "N1 ID do anúncio", "N2 De-para", "N3 SKU texto", "N4 Anúncio + variação", "N5 Título",
           "N6 Campanha", "N7 Rateio conta"]
_VAZIO = {"", "0", "x", "nan", "none", "null", "-"}
_CAMPANHA = re.compile(r"(curva|grupo|campanha|anuncios|gmv|loja|todos|\d{1,2}/\d{1,2}/\d{2,4})", re.I)


def norm_id(v):
    t = "" if v is None else str(v).strip()
    if t.lower() in _VAZIO:
        return ""
    if re.fullmatch(r"\d+\.0+", t):
        t = t.split(".")[0]
    if re.fullmatch(r"\d+", t):
        t = t.lstrip("0") or "0"
        return "" if t == "0" else t
    return t


def norm_txt(v):
    t = "" if v is None else str(v)
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"[^a-z0-9]+", " ", t).strip()
    return "" if t in _VAZIO else t


def _numerico(t):
    return bool(re.fullmatch(r"\d{5,}", t or ""))


class _Conta:
    """Produtos vendidos de uma conta e os mapas para achar o produto."""

    def __init__(self, vendas, desc, depara):
        # vendas: linhas de faturamento (SKU vinculado) da conta
        self.fat_dia, self.fat_mes, self.fat_afi_dia = {}, {}, {}
        self.id_de_sku, ids = {}, {}
        for sku, an, dia, val, tipo in vendas:
            self.fat_mes[sku] = self.fat_mes.get(sku, 0.0) + val
            self.fat_dia[(sku, dia)] = self.fat_dia.get((sku, dia), 0.0) + val
            if tipo == "Afiliados":
                self.fat_afi_dia[(sku, dia)] = self.fat_afi_dia.get((sku, dia), 0.0) + val
            n = norm_id(an)
            if n:
                ids.setdefault(n, {}).setdefault(sku, 0.0)
                ids[n][sku] += val
                if _numerico(n):
                    cont = self.id_de_sku.setdefault(sku, {})
                    cont[an] = cont.get(an, 0) + 1
        self.prods = sorted(self.fat_mes, key=lambda s: -self.fat_mes[s])
        self.por_id = {k: sorted(v, key=lambda s: -v[s]) for k, v in ids.items()}
        self.por_txt, self.por_titulo = {}, {}
        for s in self.prods:
            self.por_txt.setdefault(norm_txt(s), []).append(s)
            tit = norm_txt((desc.get(s) or ("",))[0])
            if tit:
                self.por_titulo.setdefault(tit, []).append(s)
        self.depara = {}
        for an, v in (depara or {}).items():
            n = norm_id(an)
            skus = [x.strip() for x in str((v or {}).get("sku") or "").split(",") if x.strip()]
            achou = []
            for x in skus:
                if x in self.fat_mes:
                    achou.append(x)
                else:
                    achou += [y for y in self.por_txt.get(norm_txt(x), []) if y not in achou]
            if n and achou:
                self.depara[n] = achou
            tit = norm_txt((v or {}).get("nome"))
            if n and tit and n not in self.depara:
                self.depara_titulo = getattr(self, "depara_titulo", {})
                self.depara_titulo[n] = tit
        self._cache = {}
        self.dica = {}

    def id_principal(self, sku):
        c = self.id_de_sku.get(sku)
        return max(c, key=c.get) if c else ""

    def candidatos(self, sku, an):
        k = (sku, an)
        if k not in self._cache:
            self._cache[k] = self._achar(sku, an)
        return self._cache[k]

    def _achar(self, sku, an):
        nid, tsku, tan = norm_id(an), norm_txt(sku), norm_txt(an)
        if nid and nid in self.por_id:
            return "N1 ID do anúncio", self.por_id[nid]
        if nid and nid in self.depara:
            return "N2 De-para", self.depara[nid]
        for t in (tsku, tan):
            if t and t in self.por_txt:
                return "N3 SKU texto", self.por_txt[t]
        if tsku and len(tsku.split()) <= 4:
            cands = [s for s in self.prods if norm_txt(s).endswith(" " + tsku)]
            # título do anúncio: o do de-para, senão o maior texto que veio com o mesmo ID (as outras variações)
            tit = getattr(self, "depara_titulo", {}).get(nid) or self.dica.get(nid, "")
            if len(cands) > 1 and tit:
                notas = [(difflib.SequenceMatcher(None, tit, norm_txt(s)[:len(tit) + 10]).ratio(), s) for s in cands]
                melhor = max(n for n, _ in notas)
                if melhor >= 0.6:
                    cands = [s_ for n, s_ in notas if n >= melhor - 0.05]
            if cands:
                return "N4 Anúncio + variação", cands
        for t in (tsku, tan, getattr(self, "depara_titulo", {}).get(nid, "")):
            if not t or len(t) < 8 or _numerico(t):
                continue
            if t in self.por_titulo:
                return "N5 Título", self.por_titulo[t]
            pref = [s for tt, ss in self.por_titulo.items() if len(min(tt, t, key=len)) >= 15
                    and (tt.startswith(t) or t.startswith(tt)) for s in ss]
            pref += [s for s in self.prods if len(t) >= 15 and norm_txt(s).startswith(t) and s not in pref]
            if pref:
                return "N5 Título", pref
            melhor, cands = 0.0, []
            for tt, ss in list(self.por_titulo.items()) + [(norm_txt(s), [s]) for s in self.prods]:
                r = difflib.SequenceMatcher(None, t, tt).ratio()
                if r >= 0.90:
                    if r > melhor + 1e-9:
                        melhor, cands = r, list(ss)
                    elif abs(r - melhor) <= 1e-9:
                        cands += [s for s in ss if s not in cands]
            if cands:
                return "N5 Título", cands
        if an and not nid.isdigit() and (_CAMPANHA.search(str(an)) or len(str(an)) > 25):
            return "N6 Campanha", list(self.prods)
        return "N7 Rateio conta", list(self.prods)

    def pesos(self, cands, dia, afiliados=False):
        if afiliados:
            w = [self.fat_afi_dia.get((s, dia), 0.0) for s in cands]
            if sum(w) > 0:
                return w
        w = [max(self.fat_dia.get((s, dia), 0.0), 0.0) for s in cands]
        if sum(w) > 0:
            return w
        w = [max(self.fat_mes.get(s, 0.0), 0.0) for s in cands]
        if sum(w) > 0:
            return w
        return [1.0] * len(cands)


def forcar(df, desc, cats, depara_por_conta, contagens, afi):
    """df: linhas do pacote (antes de juntar), desc: {SKU: (Descrição, Categoria)}, cats: {(Categoria, Sub)} de
    lançamentos (ADS, afiliados, visitas...), depara_por_conta: {(Marketplace, Marketplace venda): de-para},
    contagens: {(Categoria, Sub)} que são contagem, afi: (Categoria, Sub) dos afiliados.
    Devolve (df novo, auditoria). Levanta ValueError se a conferência falhar."""
    import pandas as pd
    from consolidado import FAT
    df = df.reset_index(drop=True).copy()
    an_txt = df["Id anúncio"].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v).strip())
    df["Id_anuncio_original"] = an_txt
    df["Id_anuncio_normalizado"] = an_txt.map(norm_id)
    df["Fator_rateio"] = 1.0
    cs = list(zip(df["Categoria"], df["Sub Categoria"]))
    eh_lanc = pd.Series([c in cats for c in cs], index=df.index)
    sku_txt = df["SKU"].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v).strip())
    ligado = sku_txt.map(lambda s: bool(s) and s in desc)
    df["Metodo_vinculo"] = ""
    df.loc[~eh_lanc, "Metodo_vinculo"] = "Venda"
    df.loc[eh_lanc & ligado, "Metodo_vinculo"] = "Original"
    alvo = eh_lanc & ~ligado
    chave_conf = ["Marketplace", "Marketplace venda", "Categoria", "Sub Categoria", "Data"]
    antes = df.groupby(chave_conf, dropna=False)["Valor completo"].sum()
    if not alvo.any():
        return df, _auditoria(df)

    # produtos vendidos de cada conta (e do cliente todo por marketplace, para conta sem venda no mês)
    eh_fat = pd.Series([c == FAT for c in cs], index=df.index) & ligado
    vend = df[eh_fat]
    contas, por_mk = {}, {}
    for (mk, loja), g in vend.groupby(["Marketplace", "Marketplace venda"], sort=False):
        linhas = list(zip(g["SKU"], g["Id anúncio"], g["Data"], g["Valor completo"].astype(float), g["Type"]))
        contas[(mk, loja)] = _Conta(linhas, desc, depara_por_conta.get((mk, loja)))
        por_mk.setdefault(mk, []).extend(linhas)
    todas = [x for v in por_mk.values() for x in v]
    reserva = {mk: _Conta(l, desc, None) for mk, l in por_mk.items()}
    geral = _Conta(todas, desc, None) if todas else None

    novas = []
    sai = alvo.copy()
    # contagens (visitas, cliques...): o total do anúncio no dia (somando as frações já ligadas e as soltas) vai inteiro
    # para um produto só — nunca dividido entre variações
    eh_cont = pd.Series([c in contagens for c in cs], index=df.index) & eh_lanc
    gk = ["Marketplace", "Marketplace venda", "Id_anuncio_original", "Data", "Categoria", "Sub Categoria", "Type"]
    cont_grupos = []
    if (eh_cont & alvo).any():
        com_solta = df[eh_cont & alvo][gk].drop_duplicates()
        todos = df[eh_cont].reset_index().merge(com_solta, on=gk, how="inner")
        sai.loc[todos["index"]] = True
        for k, g in todos.groupby(gk, sort=False, dropna=False):
            cont_grupos.append((dict(zip(gk, k)), g))
    # dica de título por anúncio: o maior texto de SKU que veio com o mesmo ID (ajuda o N4 a achar o produto da variação)
    for (mk, loja, nid), g in df[alvo].groupby(["Marketplace", "Marketplace venda", "Id_anuncio_normalizado"], sort=False):
        c = contas.get((mk, loja))
        if c is not None and nid:
            textos = [norm_txt(x) for x in g["SKU"]]
            maior = max(textos, key=len, default="")
            if len(maior.split()) > 4:
                c.dica[nid] = maior
    rateio_conta = {}   # N6/N7: junta por conta/dia/métrica antes de dividir (não explode o número de linhas)
    for r in df[alvo & ~eh_cont].to_dict("records"):
        conta = _conta_de(r, contas, reserva, geral)
        metodo, cands = conta.candidatos(str(r["SKU"] or ""), r["Id_anuncio_original"])
        if metodo in ("N6 Campanha", "N7 Rateio conta"):
            k = (r["Marketplace"], r["Marketplace venda"], r["Data"], r["Categoria"], r["Sub Categoria"], r["Type"], metodo)
            if k in rateio_conta:
                rateio_conta[k]["Valor completo"] += float(r["Valor completo"])
                ids = rateio_conta[k]["_ids"]
                if r["Id_anuncio_original"] and r["Id_anuncio_original"] not in ids:
                    ids.append(r["Id_anuncio_original"])
            else:
                rateio_conta[k] = dict(r, **{"Valor completo": float(r["Valor completo"]), "_ids": [r["Id_anuncio_original"]] if r["Id_anuncio_original"] else []})
            continue
        novas += _dividir(r, conta, metodo, cands, afi)
    for k, r in rateio_conta.items():
        conta = _conta_de(r, contas, reserva, geral)
        ids = r.pop("_ids")
        r["Id_anuncio_original"] = ids[0] if len(ids) == 1 else ("%d anúncios" % len(ids) if ids else "")
        r["Id_anuncio_normalizado"] = norm_id(ids[0]) if len(ids) == 1 else ""
        novas += _dividir(r, conta, k[-1], list(conta.prods), afi)
    for chave, g in cont_grupos:
        conta = _conta_de(chave, contas, reserva, geral)
        soltas = g[~g["SKU"].isin(set(desc))]
        presos = [s for s in g["SKU"] if s in desc]
        r0 = g.loc[g["Valor completo"].abs().idxmax()].to_dict()
        r0.pop("index", None)
        if presos:
            metodo, cands = "Original", list(dict.fromkeys(presos))
        else:
            rs = soltas.loc[soltas["Valor completo"].abs().idxmax()]
            metodo, cands = conta.candidatos(str(rs["SKU"] or ""), chave["Id_anuncio_original"])
        w = conta.pesos(cands, chave["Data"])
        s = cands[max(range(len(cands)), key=lambda i: w[i])]
        tot = float(g["Valor completo"].sum())
        if abs(tot - round(tot)) < 1e-3:      # frações de 1/3 gravadas com 6 casas: volta ao inteiro
            tot = float(round(tot))
        r0.update({"SKU": s, "Valor completo": tot, "Fator_rateio": 1.0, "Metodo_vinculo": metodo,
                   "Id anúncio": r0["Id anúncio"] if metodo == "Original" else (conta.id_principal(s) or r0["Id anúncio"])})
        novas.append(r0)
    alvo = sai
    novo = pd.concat([df[~alvo], pd.DataFrame(novas)[df.columns]], ignore_index=True)

    depois = novo.groupby(chave_conf, dropna=False)["Valor completo"].sum()
    dif = depois.sub(antes, fill_value=0).abs()
    if (dif > 0.01).any():
        k = dif.idxmax()
        raise ValueError("Vínculo forçado: totais não batem em %s (diferença %.2f). Pacote não gravado." % (" / ".join(map(str, k)), dif.max()))
    resta = novo[novo["Metodo_vinculo"].isin(METODOS[2:]) & ~novo["SKU"].isin(set(desc))]
    if len(resta):
        raise ValueError("Vínculo forçado: %d linha(s) continuaram sem produto. Pacote não gravado." % len(resta))
    return novo, _auditoria(novo)


def _conta_de(r, contas, reserva, geral):
    conta = contas.get((r["Marketplace"], r["Marketplace venda"])) or reserva.get(r["Marketplace"]) or geral
    if conta is None or not conta.prods:
        raise ValueError("Vínculo forçado: o cliente não tem nenhuma venda no mês para ligar ADS/afiliados/visitas.")
    return conta


def _dividir(r, conta, metodo, cands, afi):
    """Divide um valor em R$ entre os candidatos pelo faturamento (soma igual ao original; resíduo no maior)."""
    if not cands:
        metodo, cands = "N7 Rateio conta", list(conta.prods)
    eh_afi = (r["Categoria"], r["Sub Categoria"]) == afi
    w = conta.pesos(cands, r["Data"], afiliados=eh_afi and metodo in ("N6 Campanha", "N7 Rateio conta"))
    tot = float(sum(w))
    v = float(r["Valor completo"])
    fat = [x / tot for x in w]
    vals = [v * f for f in fat]
    i_max = max(range(len(vals)), key=lambda i: abs(vals[i]))
    vals[i_max] += v - sum(vals)
    out = []
    for s, val, f in zip(cands, vals, fat):
        if f <= 0:
            continue
        nr = dict(r)
        nr.update({"SKU": s, "Valor completo": val, "Fator_rateio": round(f, 6), "Metodo_vinculo": metodo,
                   "Id anúncio": conta.id_principal(s) or r["Id anúncio"]})
        out.append(nr)
    return out


def _auditoria(df):
    import pandas as pd
    d = df[df["Metodo_vinculo"] != "Venda"]
    if not len(d):
        return pd.DataFrame(columns=["Marketplace", "Marketplace venda", "Sub Categoria", "Metodo_vinculo", "Valor", "Linhas", "%"])
    a = (d.assign(Valor=d["Valor completo"].abs())
         .groupby(["Marketplace", "Marketplace venda", "Sub Categoria", "Metodo_vinculo"], as_index=False)
         .agg(Valor=("Valor", "sum"), Linhas=("Valor", "size")))
    tot = a.groupby(["Marketplace", "Marketplace venda", "Sub Categoria"])["Valor"].transform("sum")
    a["%"] = (a["Valor"] / tot.replace(0, float("nan"))).fillna(0).round(4)
    ordem = {m: i for i, m in enumerate(METODOS)}
    a["_o"] = a["Metodo_vinculo"].map(ordem)
    return a.sort_values(["Marketplace", "Marketplace venda", "Sub Categoria", "_o"]).drop(columns="_o").reset_index(drop=True)
