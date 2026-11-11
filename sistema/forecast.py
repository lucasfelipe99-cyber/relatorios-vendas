# -*- coding: utf-8 -*-
"""Forecast de vendas: fechamento do mês atual, por cliente, marketplace e SKU.

Base (de onde saem ritmo e custos médios):
- com 7 dias ou mais realizados no mês: 50% o ritmo dos últimos 7 dias + 50% a média do mês até agora;
- no começo do mês (menos de 7 dias): o mês anterior inteiro.
Ritmo = unidades por dia de cada SKU em cada marketplace. Custos médios de venda, tirados da base:
comissão, rebate e antecipação em % do faturamento; frete e custo do produto em R$ por unidade;
afiliados em % do faturamento do marketplace; ADS em R$ por dia do marketplace (não sobe com o preço);
imposto = alíquota do cliente.

O ticket médio de cada SKU × marketplace pode ser trocado (e o crescimento % do volume por marketplace):
o faturamento passa a ser unidades × ticket novo, e os custos em % acompanham.

Saída: na tela (DRE realizado + projetado = fechamento) e arquivo no layout do consolidado com
Datatype = "Forecast", só para os dias que faltam do mês (o realizado já vai como "Actual" no consolidado).
"""
import calendar
import datetime as dt
import json
import os

import comum

DIAS_RECENTES = 7
SUB = {  # Sub Categoria do consolidado -> nome curto
    "Quantidade": "qtd", "Preço total com desconto": "fat", "Cancelamento": "canc", "Imposto": "imp",
    "Comissão": "com", "Tarifa de frete": "fre", "Rebate": "reb", "Antecipação": "ant", "Produto": "cus",
    "Publicidade": "ads", "Afiliados": "afi",
}
LINHAS_DRE = [("qtd", "Quantidade"), ("fat", "Faturamento"), ("imp", "Imposto"), ("com", "Comissão"),
              ("fre", "Frete"), ("reb", "Rebate"), ("ant", "Antecipação"), ("afi", "Afiliados"), ("ads", "ADS"),
              ("cus", "Custo do produto"), ("margem", "Margem")]


# ---------------------------------------------------------------- dados
def _arq_ajustes(cliente_id, ano, mes):
    p = os.path.join(comum.PASTA_SISTEMA, "dados", "forecast")
    os.makedirs(p, exist_ok=True)
    return os.path.join(p, "%s_%04d-%02d.json" % (cliente_id, ano, mes))


def ler_ajustes(cliente_id, ano, mes):
    try:
        with open(_arq_ajustes(cliente_id, ano, mes), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"tickets": {}, "crescimento": {}}


def salvar_ajustes(cliente_id, ano, mes, ajustes):
    limpo = {"tickets": {}, "crescimento": {}}
    for k, v in (ajustes.get("tickets") or {}).items():
        try:
            v = float(str(v).replace(",", "."))
        except ValueError:
            continue
        if v > 0:
            limpo["tickets"][k] = round(v, 4)
    for k, v in (ajustes.get("crescimento") or {}).items():
        try:
            v = float(str(v).replace(",", "."))
        except ValueError:
            continue
        if -100 < v <= 1000 and v != 0:
            limpo["crescimento"][k] = v
    with open(_arq_ajustes(cliente_id, ano, mes), "w", encoding="utf-8") as f:
        json.dump(limpo, f, ensure_ascii=False, indent=1)
    return limpo


def _consolidado_df(cfg, cliente, ano, mes):
    """DataFrame do consolidado do mês (gera se ainda não existir)."""
    import pandas as pd
    import consolidado
    r = (consolidado.ler_estado().get(cliente["id"]) or {}).get("%04d-%02d" % (ano, mes))
    if not r or not r.get("ok") or not os.path.isfile(r.get("arquivo", "")):
        d, _ = consolidado.dia_fonte(cfg, ano, mes)
        if not d:
            return None
        r = consolidado.gerar(cfg, cliente, ano, mes, origem="forecast")
        if not r.get("ok"):
            return None
    df = consolidado.ler_excel(r["arquivo"], sheet_name="DB", dtype={"SKU": str, "Id anúncio": str})
    if "Datatype" in df.columns:
        df = df[df["Datatype"] == "Actual"].copy()      # o próprio forecast não entra na base
    df["k"] = df["Sub Categoria"].map(SUB)
    df = df[df["k"].notna()].copy()
    df["Data"] = pd.to_datetime(df["Data"]).dt.date
    df["SKU"] = df["SKU"].fillna("0").astype(str)
    return df


def _agrega(df, chaves):
    import pandas as pd
    if df is None or not len(df):
        return pd.DataFrame(columns=chaves + list(SUB.values()))
    p = df.pivot_table(index=chaves, columns="k", values="Valor completo", aggfunc="sum", fill_value=0.0)
    for c in SUB.values():
        if c not in p.columns:
            p[c] = 0.0
    return p.reset_index()


def _base(cfg, cliente, hoje):
    """Escolhe a base e devolve (df_base, dias_base, descrição, realizado_df, ultimo_dia_realizado)."""
    ano, mes = hoje.year, hoje.month
    atual = _consolidado_df(cfg, cliente, ano, mes)
    ontem = hoje - dt.timedelta(days=1)
    if atual is not None and len(atual):
        atual = atual[(atual["Data"] >= dt.date(ano, mes, 1)) & (atual["Data"] <= ontem)]
    dias_real = 0
    ultimo = None
    if atual is not None and len(atual):
        ultimo = max(atual["Data"])
        dias_real = (ultimo - dt.date(ano, mes, 1)).days + 1
    if dias_real >= DIAS_RECENTES:
        ini7 = ultimo - dt.timedelta(days=DIAS_RECENTES - 1)
        recente = atual[atual["Data"] >= ini7]
        desc = ("50%% últimos %d dias (%s a %s) + 50%% mês até agora (01 a %s)"
                % (DIAS_RECENTES, ini7.strftime("%d/%m"), ultimo.strftime("%d/%m"), ultimo.strftime("%d/%m")))
        return [(recente, DIAS_RECENTES, 0.5), (atual, dias_real, 0.5)], desc, atual, ultimo
    pa, pm = (ano - 1, 12) if mes == 1 else (ano, mes - 1)
    ant = _consolidado_df(cfg, cliente, pa, pm)
    if ant is None or not len(ant):
        if dias_real:
            return [(atual, dias_real, 1.0)], "mês até agora (%d dia(s); sem mês anterior)" % dias_real, atual, ultimo
        return None, "sem dados: gere o consolidado primeiro", atual, ultimo
    dias_ant = calendar.monthrange(pa, pm)[1]
    ult_ant = max(ant["Data"])
    dias_ant = min(dias_ant, (ult_ant - dt.date(pa, pm, 1)).days + 1)
    desc = "mês anterior (%02d/%04d, %d dias) — início do mês" % (pm, pa, dias_ant)
    return [(ant, dias_ant, 1.0)], desc, atual, ultimo


CATS = {"qtd": ("14.Quantidade", "Quantidade"), "fat": ("13.Faturamento Bruto", "Preço total com desconto"),
        "imp": ("02.Imposto", "Imposto"), "com": ("03.Despesas Marketplace", "Comissão"),
        "fre": ("03.Despesas Marketplace", "Tarifa de frete"), "reb": ("03.Despesas Marketplace", "Rebate"),
        "ant": ("03.Despesas Marketplace", "Antecipação"), "afi": ("03.Despesas Marketplace", "Afiliados"),
        "ads": ("03.Despesas Marketplace", "Publicidade"), "cus": ("04.Materia Prima", "Produto")}


# ---------------------------------------------------------------- cálculo
def _custos_tabela(cliente, ano=None, mes=None):
    import tratador
    p = tratador.caminho_custos(cliente, ano, mes) if ano else tratador.caminho_custos(cliente)
    if not p:
        return {}
    try:
        return {l["SKU"]: l["Custo"] for l in tratador.ler_custos(p)[0] if l.get("Custo")}
    except Exception:
        return {}


def parametros(cfg, cliente, hoje=None, ajustes=None):
    """Ritmo diário e custos médios por SKU × marketplace (a 'receita' do forecast), já com os ajustes."""
    import tratador
    hoje = hoje or dt.date.today()
    ano, mes = hoje.year, hoje.month
    ajustes = ajustes if ajustes is not None else ler_ajustes(cliente["id"], ano, mes)
    bases, desc_base, realizado, ultimo = _base(cfg, cliente, hoje)
    info = {"cliente_id": cliente["id"], "cliente": cliente["nome"], "mes": "%04d-%02d" % (ano, mes),
            "base": desc_base, "dias_mes": calendar.monthrange(ano, mes)[1],
            "dias_realizados": (ultimo.day if ultimo else 0), "ultimo_realizado": ultimo.isoformat() if ultimo else None,
            "imposto": tratador.imposto(cliente), "ajustes": ajustes}
    if not bases:
        info["erro"] = desc_base
        return info, [], realizado
    acum, acum_mk = {}, {}
    for df, dias, peso in bases:
        for _, r in _agrega(df[df["k"] != "ads"], ["Marketplace", "SKU"]).iterrows():
            a = acum.setdefault((r["Marketplace"], r["SKU"]), {c: 0.0 for c in SUB.values()})
            for c in SUB.values():
                a[c] += peso * r[c] / dias
        for _, r in _agrega(df, ["Marketplace"]).iterrows():
            a = acum_mk.setdefault(r["Marketplace"], {c: 0.0 for c in SUB.values()})
            for c in SUB.values():
                a[c] += peso * r[c] / dias
    tab = _custos_tabela(cliente, ano, mes)
    aliq = info["imposto"]
    itens = []
    for (mk, sku), a in sorted(acum.items()):
        qd, fat_dia = a["qtd"], a["fat"] + a["canc"]
        if qd <= 0 or fat_dia <= 0:
            continue
        m = acum_mk.get(mk, {})
        fat_mk = (m.get("fat", 0) + m.get("canc", 0)) or 0
        ref, ref_fat = (a, fat_dia) if a["fat"] > 0 else (m, fat_mk)
        chave = "%s|%s" % (mk, sku)
        cresc = float(ajustes.get("crescimento", {}).get(mk) or 0) / 100.0
        custo_un = (-a["cus"] / qd) if a["cus"] else float(tab.get(sku, 0) or 0)
        itens.append({
            "mk": mk, "sku": sku, "chave": chave,
            "qtd_dia_base": qd, "qtd_dia": qd * (1 + cresc), "ticket_real": fat_dia / qd,
            "ticket": float(ajustes.get("tickets", {}).get(chave) or fat_dia / qd),
            "ticket_alterado": chave in ajustes.get("tickets", {}),
            # % do faturamento (negativos = despesa) e R$ por unidade (positivos = custo)
            "p_imp": -aliq, "p_com": (ref["com"] / ref_fat) if ref_fat else 0, "p_reb": (ref["reb"] / ref_fat) if ref_fat else 0,
            "p_ant": (ref["ant"] / ref_fat) if ref_fat else 0, "p_afi": (m.get("afi", 0) / fat_mk) if fat_mk else 0,
            "u_fre": -a["fre"] / qd, "u_cus": custo_un, "sem_custo": not custo_un,
            "_ads_dia_mk": m.get("ads", 0.0) * (1 + cresc),
        })
    # ADS em R$/dia do marketplace, rateado pela participação do SKU no faturamento projetado do marketplace
    por_mk = {}
    for i in itens:
        por_mk.setdefault(i["mk"], []).append(i)
    for mk, ls in por_mk.items():
        tot = sum(i["qtd_dia"] * i["ticket"] for i in ls)
        for i in ls:
            i["ads_dia"] = i.pop("_ads_dia_mk") * ((i["qtd_dia"] * i["ticket"]) / tot if tot else 0)   # negativo (despesa)
    for i in itens:
        _economia(i)
    return info, itens, realizado


def _economia(i):
    """Margem por unidade e % com o ticket atual; e o ticket que daria a margem alvo."""
    t = i["ticket"]
    p = i["p_imp"] + i["p_com"] + i["p_reb"] + i["p_ant"] + i["p_afi"]      # soma das % (despesa negativa)
    ads_un = (i["ads_dia"] / i["qtd_dia"]) if i["qtd_dia"] else 0           # negativo
    i["margem_un"] = t * (1 + p) - i["u_fre"] - i["u_cus"] + ads_un
    i["margem_pct"] = i["margem_un"] / t if t else 0
    i["margem_un_real"] = i["ticket_real"] * (1 + p) - i["u_fre"] - i["u_cus"] + ads_un
    i["margem_pct_real"] = i["margem_un_real"] / i["ticket_real"] if i["ticket_real"] else 0
    i["_p"], i["_fixo_un"] = p, i["u_fre"] + i["u_cus"] - ads_un


def ticket_para_margem(i, alvo):
    """Ticket que dá a margem % alvo: m = 1 + p − fixo/t  ->  t = fixo / (1 + p − m)."""
    den = 1 + i["_p"] - alvo
    return (i["_fixo_un"] / den) if den > 0 and i["_fixo_un"] > 0 else None


def projetar(itens, dias):
    """Valores (R$) de cada item para N dias."""
    saida = []
    for i in itens:
        q = i["qtd_dia"] * dias
        fat = q * i["ticket"]
        v = {"qtd": q, "fat": fat, "imp": i["p_imp"] * fat, "com": i["p_com"] * fat, "reb": i["p_reb"] * fat,
             "ant": i["p_ant"] * fat, "afi": i["p_afi"] * fat, "fre": -i["u_fre"] * q, "cus": -i["u_cus"] * q,
             "ads": i["ads_dia"] * dias}
        v["margem"] = sum(v[c] for c in ("fat", "imp", "com", "reb", "ant", "afi", "fre", "cus", "ads"))
        saida.append(v)
    return saida


def _soma_por_mk(itens, valores):
    d = {}
    for i, v in zip(itens, valores):
        t = d.setdefault(i["mk"], {k: 0.0 for k, _ in LINHAS_DRE})
        for k in t:
            t[k] += v.get(k, 0.0)
    return d


def _com_pct(d):
    d = {k: round(v, 2) for k, v in d.items()}
    d["margem_pct"] = round(d["margem"] / d["fat"], 4) if d.get("fat") else 0
    return d


def _realizado_por_mk(realizado, ate=None):
    if realizado is None or not len(realizado):
        return {}
    df = realizado if ate is None else realizado[realizado["Data"] <= ate]
    out = {}
    for _, r in _agrega(df, ["Marketplace"]).iterrows():
        v = {k: float(r.get(k, 0) or 0) for k in SUB.values()}
        v["fat"] += v["canc"]
        v["margem"] = sum(v[c] for c in ("fat", "imp", "com", "reb", "ant", "afi", "fre", "cus", "ads"))
        out[r["Marketplace"]] = {k: v.get(k, 0.0) for k, _ in LINHAS_DRE}
    return out


def simular(cfg, cliente, hoje=None, ajustes=None, alvo=None):
    """Forecast do mês inteiro com os ajustes (sem definir), para a tela."""
    info, itens, realizado = parametros(cfg, cliente, hoje, ajustes)
    if info.get("erro"):
        return info
    vals = projetar(itens, info["dias_mes"])
    por_mk = _soma_por_mk(itens, vals)
    total = {k: sum(d[k] for d in por_mk.values()) for k, _ in LINHAS_DRE}
    skus = []
    for i, v in zip(itens, vals):
        x = {k: (round(val, 4) if isinstance(val, float) else val) for k, val in i.items() if not k.startswith("_")}
        x.update({"fat_mes": round(v["fat"], 2), "qtd_mes": round(v["qtd"], 2), "margem_mes": round(v["margem"], 2)})
        if alvo is not None:
            t = ticket_para_margem(i, alvo) if i["margem_pct"] < alvo else None   # só sugere para quem está abaixo da meta
            x["ticket_sugerido"] = round(t, 2) if t else None
        skus.append(x)
    skus.sort(key=lambda x: -x["fat_mes"])
    info.update(dre=[{"mk": mk, **_com_pct(d)} for mk, d in sorted(por_mk.items())], total=_com_pct(total),
                skus=skus, linhas_dre=LINHAS_DRE)
    return info


# ---------------------------------------------------------------- forecast definido (travado)
def _arq_definido(cliente_id, ano, mes):
    return _arq_ajustes(cliente_id, ano, mes).replace(".json", "_definido.json")


def ler_definido(cliente_id, ano, mes):
    try:
        with open(_arq_definido(cliente_id, ano, mes), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def definir(cfg, cliente, hoje=None, motivo=""):
    """Trava o forecast do mês (mês inteiro, por dia). Só muda se for definido de novo; a versão anterior fica no histórico."""
    hoje = hoje or dt.date.today()
    ano, mes = hoje.year, hoje.month
    info, itens, _ = parametros(cfg, cliente, hoje)
    if info.get("erro"):
        return {"ok": False, "erro": info["erro"]}
    dias = info["dias_mes"]
    vals = projetar(itens, 1)        # valores de 1 dia (iguais em todos os dias do mês)
    linhas = [{"mk": i["mk"], "sku": i["sku"], "ticket": round(i["ticket"], 4), "ticket_real": round(i["ticket_real"], 4),
               "dia": {k: round(v[k], 6) for k in CATS}} for i, v in zip(itens, vals)]
    por_mk = _soma_por_mk(itens, projetar(itens, dias))
    total = {k: sum(d[k] for d in por_mk.values()) for k, _ in LINHAS_DRE}
    anterior = ler_definido(cliente["id"], ano, mes)
    versao = (anterior or {}).get("versao", 0) + 1
    hist = (anterior or {}).get("historico", [])
    if anterior:
        hist.append({k: anterior.get(k) for k in ("versao", "definido_em", "base", "total", "motivo")})
    doc = {"cliente_id": cliente["id"], "mes": info["mes"], "versao": versao,
           "definido_em": dt.datetime.now().isoformat(timespec="seconds"), "motivo": motivo or "",
           "base": info["base"], "dias_mes": dias, "imposto": info["imposto"], "ajustes": info["ajustes"],
           "linhas": linhas, "por_mk": {mk: _com_pct(d) for mk, d in por_mk.items()}, "total": _com_pct(total),
           "historico": hist}
    with open(_arq_definido(cliente["id"], ano, mes), "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    comum.registrar_evento({"tipo": "forecast", "cliente": cliente["id"], "mes": info["mes"], "versao": versao,
                            "faturamento": round(total["fat"], 2), "margem": round(total["margem"], 2), "motivo": motivo})
    comum.log("[%s] FORECAST %s definido (versão %d): faturamento %.2f, margem %.2f" % (
        cliente["nome"], info["mes"], versao, total["fat"], total["margem"]))
    import consolidado
    cons = consolidado.gerar(cfg, cliente, ano, mes, hoje, origem="forecast definido")
    return {"ok": True, "versao": versao, "total": doc["total"], "consolidado": {k: cons.get(k) for k in ("ok", "conferido", "erro")}}


def linhas_consolidado(cliente_id, ano, mes):
    """Linhas do forecast definido para o consolidado (Datatype "Forecast"), uma por SKU × dia × medida."""
    doc = ler_definido(cliente_id, ano, mes)
    if not doc:
        return []
    out = []
    for d in range(1, doc["dias_mes"] + 1):
        dia = dt.date(ano, mes, d)
        for l in doc["linhas"]:
            for k, (cat, sub) in CATS.items():
                v = l["dia"].get(k) or 0
                if abs(v) > 1e-9:
                    out.append((l["mk"], l["sku"], dia, cat, sub, v))
    return out


TOLERANCIA_ATENCAO = 0.10   # 10% abaixo do forecast = sinal de atenção
MINIMO_ALERTA = 100.0       # não alerta marketplace com valores muito pequenos


def comparar(cfg, cliente, hoje=None):
    """Forecast definido x realizado: até ontem e tendência de fechamento, por marketplace, com sinais de atenção."""
    hoje = hoje or dt.date.today()
    ano, mes = hoje.year, hoje.month
    doc = ler_definido(cliente["id"], ano, mes)
    if not doc:
        return None
    # tendência = ritmo real, sem os ajustes de ticket/crescimento (esses só valem para o forecast)
    info, itens, realizado = parametros(cfg, cliente, hoje, {"tickets": {}, "crescimento": {}})
    ult = dt.date.fromisoformat(info["ultimo_realizado"]) if info.get("ultimo_realizado") else None
    dias_real = ult.day if ult and ult.month == mes else 0
    real = _realizado_por_mk(realizado, ult) if dias_real else {}
    resto = doc["dias_mes"] - dias_real
    tend_mk = _soma_por_mk(itens, projetar(itens, resto)) if itens and resto > 0 else {}
    linhas, alertas = [], []
    for mk in sorted(set(doc["por_mk"]) | set(real)):
        f_mes = doc["por_mk"].get(mk, {})
        f_ate = {k: (f_mes.get(k, 0) or 0) * dias_real / doc["dias_mes"] for k in ("qtd", "fat", "margem")}
        r = real.get(mk, {})
        tend = {k: (r.get(k, 0) or 0) + (tend_mk.get(mk, {}).get(k, 0) or 0) for k in ("qtd", "fat", "margem")}
        lin = {"mk": mk, "forecast_mes": {k: round(f_mes.get(k, 0) or 0, 2) for k in ("qtd", "fat", "margem")},
               "forecast_ate": {k: round(v, 2) for k, v in f_ate.items()},
               "realizado_ate": {k: round(r.get(k, 0) or 0, 2) for k in ("qtd", "fat", "margem")},
               "tendencia": {k: round(v, 2) for k, v in tend.items()}}
        for k, nome in (("fat", "faturamento"), ("margem", "margem")):
            base = f_ate[k]
            dif = ((lin["realizado_ate"][k] - base) / abs(base)) if base else 0
            lin["dif_%s" % k] = round(dif, 4)
            if dias_real and abs(base) >= MINIMO_ALERTA and dif < -TOLERANCIA_ATENCAO:
                alertas.append("%s: %s até ontem %.0f%% abaixo do forecast" % (mk, nome, -dif * 100))
        linhas.append(lin)
    tot = {p: {k: round(sum(l[p][k] for l in linhas), 2) for k in ("qtd", "fat", "margem")}
           for p in ("forecast_mes", "forecast_ate", "realizado_ate", "tendencia")}
    for k in ("fat", "margem"):
        b = tot["forecast_mes"][k]
        if b and (tot["tendencia"][k] - b) / abs(b) < -TOLERANCIA_ATENCAO:
            alertas.append("Tendência de fechamento: %s %.0f%% abaixo do forecast do mês"
                           % ("faturamento" if k == "fat" else "margem", -(tot["tendencia"][k] - b) / abs(b) * 100))
    return {"versao": doc["versao"], "definido_em": doc["definido_em"], "base": doc["base"], "motivo": doc.get("motivo"),
            "dias_realizados": dias_real, "dias_mes": doc["dias_mes"], "linhas": linhas, "total": tot,
            "alertas": alertas, "historico": doc.get("historico", []), "total_definido": doc["total"],
            "por_mk": doc["por_mk"]}
