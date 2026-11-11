# -*- coding: utf-8 -*-
"""Números do painel de controle, lidos do relatório consolidado (a base já conferida contra os tratados).

Nada aqui altera arquivos: só lê o consolidado do mês, soma e devolve para a tela.
"""
import datetime as dt
import json
import os
import threading

import comum
import consolidado as C

# medidas em R$ (o sinal já vem certo no consolidado: receita +, despesa −)
MEDIDAS = [
    ("fat_bruto", "Faturamento bruto", C.FAT),
    ("cancelamento", "Cancelamentos e devoluções", C.CANC),
    ("imposto", "Imposto", C.IMP),
    ("comissao", "Comissão", C.COM),
    ("frete", "Frete", C.FRE),
    ("rebate", "Rebate", C.REB),
    ("antecipacao", "Antecipação", C.ANT),
    ("publicidade", "Publicidade (ADS)", C.PUB),
    ("afiliados", "Afiliados", C.AFI),
    ("custo", "Custo do produto", C.CUS),
]
OUTRAS = [("unidades", C.QTD), ("cliques", C.CLQ), ("receita_ads", C.ADF), ("visitas", C.VIS)]
_CACHE = {}
_TRAVA = threading.Lock()


_LEITURA = threading.Lock()        # um consolidado lido por vez: ler o .xlsx é o que mais usa memória
_MAX_CACHE = 60
_GRANDE = 2 * 1024 * 1024           # acima disso o .xlsx é lido em fluxo
_PASTA_CACHE = os.path.join(comum.PASTA_SISTEMA, "dados", "cache_painel")
_VERSAO_CACHE = 1


def _devolver_memoria():
    """Depois de ler um .xlsx grande, devolve ao sistema a memória que o pandas já soltou (Linux/Render)."""
    import gc
    gc.collect()
    try:
        import ctypes
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except (OSError, AttributeError):
        pass


def _vazio(v):
    return None if v is None or (isinstance(v, str) and not v.strip()) or (isinstance(v, float) and v != v) else v


def _sku(v):
    v = _vazio(v)
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if v == v else 0.0
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return 0.0


def _ler_somado(caminho):
    """Lê a aba DB linha a linha (sem montar a planilha inteira no pandas) já somando por
    tipo/marketplace/conta/data/SKU/medida. Devolve (DataFrame pequeno, {SKU: descrição})."""
    import pandas as pd
    chave = {cs: k for k, _, cs in MEDIDAS}
    chave.update({cs: k for k, cs in OUTRAS})
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        CalamineWorkbook = None
    grande = os.path.getsize(caminho) > _GRANDE
    if grande:                              # arquivo grande: lê em fluxo (pouca memória, mais lento; o resultado fica em disco)
        import openpyxl
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        it = wb["DB"].iter_rows(values_only=True)
        cab = [str(c).strip() if c is not None else "" for c in next(it, [])]
        linhas = it
    elif CalamineWorkbook is None:          # sem o leitor rápido: caminho antigo (pandas)
        df = C.ler_excel(caminho, sheet_name="DB", dtype={"SKU": str},
                         usecols=["Marketplace", "Marketplace venda", "SKU", "Data", "Categoria", "Sub Categoria",
                                  "Valor completo", "Descrição", "Datatype"])
        linhas = df.itertuples(index=False, name=None)
        cab = list(df.columns)
    else:
        wb = CalamineWorkbook.from_path(caminho)
        it = wb.get_sheet_by_name("DB").iter_rows()
        cab = [str(c).strip() for c in next(it, [])]
        linhas = it
    i = {n: cab.index(n) for n in ("Marketplace", "Marketplace venda", "SKU", "Data", "Categoria", "Sub Categoria",
                                   "Valor completo", "Descrição", "Datatype")}
    somas, desc = {}, {}
    for r in linhas:
        tipo = r[i["Datatype"]]
        if tipo not in ("Actual", "Forecast"):
            continue
        k = chave.get((_vazio(r[i["Categoria"]]), _vazio(r[i["Sub Categoria"]])))
        if k is None:
            continue
        sku = _sku(r[i["SKU"]])
        ch = (tipo, _vazio(r[i["Marketplace"]]), _vazio(r[i["Marketplace venda"]]), _vazio(r[i["Data"]]), sku, k)
        somas[ch] = somas.get(ch, 0.0) + _num(r[i["Valor completo"]])
        if tipo == "Actual" and sku not in desc:
            d = _vazio(r[i["Descrição"]])
            if d is not None:
                desc[sku] = str(d)[:80]
    linhas = None
    if grande:
        wb.close()
    df = pd.DataFrame([ch + (v,) for ch, v in somas.items()],
                      columns=["Datatype", "Marketplace", "Marketplace venda", "Data", "SKU", "k", "Valor completo"])
    if df.empty:
        df["Valor completo"] = df["Valor completo"].astype(float)
    return df, desc


def _arq_cache(caminho):
    import hashlib
    return os.path.join(_PASTA_CACHE, hashlib.sha1(os.path.abspath(caminho).encode("utf-8")).hexdigest() + ".pkl")


def _do_disco(caminho, mt):
    """Resumo já calculado deste consolidado (gravado pelo próprio sistema), se o arquivo não mudou."""
    import pickle
    try:
        with open(_arq_cache(caminho), "rb") as f:
            d = pickle.load(f)
        if d.get("versao") == _VERSAO_CACHE and d.get("mt") == mt and d.get("caminho") == os.path.abspath(caminho):
            d["real"].attrs["desc"] = d.get("desc") or {}
            return d["real"], d["fc"]
    except Exception:
        pass
    return None


def _para_disco(caminho, mt, real, fc):
    import pickle
    try:
        os.makedirs(_PASTA_CACHE, exist_ok=True)
        arq = _arq_cache(caminho)
        with open(arq + ".tmp", "wb") as f:
            pickle.dump({"versao": _VERSAO_CACHE, "mt": mt, "caminho": os.path.abspath(caminho),
                         "real": real, "fc": fc, "desc": real.attrs.get("desc") or {}}, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(arq + ".tmp", arq)
    except Exception as e:
        comum.log("painel: não gravei o resumo em disco: %s" % e)


def aquecer(cfg=None):
    """Prepara (em segundo plano) o resumo dos consolidados do mês atual e do anterior de todos os clientes."""
    try:
        cfg = cfg or comum.carregar_config()
        hoje = dt.date.today()
        meses = ["%04d-%02d" % (hoje.year, hoje.month),
                 "%04d-%02d" % ((hoje.year - 1, 12) if hoje.month == 1 else (hoje.year, hoje.month - 1))]
        for c in cfg.get("clientes", []):
            if c.get("ativo") is False:
                continue
            for m in meses:
                est = _estado_mes(c["id"], m) or {}
                for e in (est, est.get("ultimo_ok") or {}):
                    arq = e.get("arquivo")
                    if e.get("ok") and arq and os.path.isfile(arq):
                        try:
                            _ler(arq)
                        except Exception as ex:
                            comum.log("painel: não preparei %s: %s" % (os.path.basename(arq), ex))
    except Exception as e:
        comum.log("painel: preparo em segundo plano falhou: %s" % e)


def _ler(caminho):
    """Realizado (Actual) e forecast do consolidado, já somados por dia/marketplace/conta/SKU/medida.

    Guarda em memória só essa tabela resumida (bem menor que a planilha) enquanto o arquivo não muda;
    as descrições dos SKUs ficam em real.attrs["desc"]."""
    mt = os.path.getmtime(caminho)
    with _TRAVA:
        c = _CACHE.get(caminho)
        if c and c[0] == mt:
            return c[1], c[2]
    with _LEITURA:
        with _TRAVA:                       # outra requisição pode ter lido enquanto esta esperava
            c = _CACHE.get(caminho)
            if c and c[0] == mt:
                return c[1], c[2]
        import pandas as pd
        disco = _do_disco(caminho, mt)
        if disco is not None:
            with _TRAVA:
                _CACHE[caminho] = (mt,) + disco
            return disco
        df, desc = _ler_somado(caminho)
        chaves = ["Datatype", "Marketplace", "Marketplace venda", "Data", "SKU", "k"]
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.normalize()
        soma = df.groupby(chaves, dropna=False, sort=False)["Valor completo"].sum().reset_index()
        del df
        for col in ("Datatype", "Marketplace", "Marketplace venda", "SKU", "k"):
            soma[col] = soma[col].astype("category")
        real = soma[soma["Datatype"] == "Actual"].drop(columns=["Datatype"]).reset_index(drop=True)
        fc = soma[soma["Datatype"] == "Forecast"].drop(columns=["Datatype"]).reset_index(drop=True)
        for t in (real, fc):
            for col in ("Marketplace", "Marketplace venda", "SKU", "k"):
                t[col] = t[col].cat.remove_unused_categories()
        real.attrs["desc"] = desc
        del soma
        _devolver_memoria()
        _para_disco(caminho, mt, real, fc)
        with _TRAVA:
            _CACHE[caminho] = (mt, real, fc)
            while len(_CACHE) > _MAX_CACHE:
                _CACHE.pop(next(iter(_CACHE)))
        return real, fc


def _soma(df):
    s = df.groupby("k", observed=True)["Valor completo"].sum().to_dict()
    r = {k: round(float(s.get(k, 0.0)), 2) for k, _, _ in MEDIDAS}
    r.update({k: round(float(s.get(k, 0.0)), 2) for k, _ in OUTRAS})
    return _derivados(r)


def _derivados(r):
    r["fat_liquido"] = round(r["fat_bruto"] + r["cancelamento"], 2)
    r["taxas"] = round(r["comissao"] + r["frete"] + r["antecipacao"] + r["rebate"], 2)
    r["margem"] = round(sum(r[k] for k, _, _ in MEDIDAS), 2)
    base = r["fat_liquido"]
    r["margem_pct"] = round(r["margem"] / base, 4) if base else None
    r["cancel_pct"] = round(-r["cancelamento"] / r["fat_bruto"], 4) if r["fat_bruto"] else None
    r["roas"] = round(r["receita_ads"] / -r["publicidade"], 2) if r["publicidade"] else None
    r["tacos"] = round(-r["publicidade"] / base, 4) if base else None
    r["ticket_un"] = round(r["fat_bruto"] / r["unidades"], 2) if r["unidades"] else None
    return r


def _estado_mes(cliente_id, mes):
    return (C.ler_estado().get(cliente_id) or {}).get(mes)


# ---------------------------------------------------------------- filtros (período, marketplace, conta)
def _filtro(f):
    f = f or {}
    def d(x):
        try:
            return dt.date.fromisoformat(str(x)[:10]) if x else None
        except ValueError:
            return None
    return {"de": d(f.get("de")), "ate": d(f.get("ate")),
            "mks": [str(x) for x in (f.get("mks") or []) if x], "contas": [str(x) for x in (f.get("contas") or []) if x]}


def _aplicar(df, f, dias=True):
    if df is None or df.empty:
        return df
    if dias and f["de"]:
        df = df[df["Data"].dt.date >= f["de"]]
    if dias and f["ate"]:
        df = df[df["Data"].dt.date <= f["ate"]]
    if f["mks"]:
        df = df[df["Marketplace"].isin(f["mks"])]
    if f["contas"]:
        df = df[df["Marketplace venda"].isin(f["contas"])]
    return df


# ---------------------------------------------------------------- conferência manual por marketplace
_ARQ_CONF = os.path.join(comum.PASTA_SISTEMA, "dados", "conferencias_painel.json")
CAMPOS_CONF = ["unidades", "fat_bruto", "cancelamento", "fat_liquido", "comissao", "frete", "rebate", "antecipacao", "imposto", "publicidade",
               "receita_ads", "afiliados", "custo", "visitas", "margem"]


def _ler_conf():
    try:
        with open(_ARQ_CONF, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def chave_conf(mes, mk, conta, de=None, ate=None):
    return "|".join([mes, mk or "", conta or "", str(de or ""), str(ate or "")])


def salvar_conferencia(cliente_id, mes, mk, conta, de, ate, valores, informados=None, nota="", por="", remover=False):
    """Marca (ou desmarca) um marketplace/conta como conferido no período, guardando os números vistos naquele momento."""
    with _TRAVA:
        tudo = _ler_conf()
        cli = tudo.setdefault(cliente_id, {})
        k = chave_conf(mes, mk, conta, de, ate)
        if remover:
            cli.pop(k, None)
        else:
            cli[k] = {"quando": dt.datetime.now().isoformat(timespec="seconds"), "por": por or "", "nota": (nota or "")[:300],
                      "valores": {c: round(float((valores or {}).get(c) or 0), 2) for c in CAMPOS_CONF},
                      "informados": {c: round(float(v), 2) for c, v in (informados or {}).items() if c in CAMPOS_CONF and v not in (None, "")}}
        os.makedirs(os.path.dirname(_ARQ_CONF), exist_ok=True)
        tmp = _ARQ_CONF + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(tudo, f, ensure_ascii=False, indent=1)
        os.replace(tmp, _ARQ_CONF)
    return True


def _marcar_conferencias(cliente_id, mes, grupos, f):
    conf = _ler_conf().get(cliente_id) or {}
    for g in grupos:
        c = conf.get(chave_conf(mes, g["marketplace"], g["conta"], f["de"], f["ate"]))
        if c:
            mudou = [k for k in CAMPOS_CONF if abs((c["valores"].get(k) or 0) - (g.get(k) or 0)) > 0.01]
            g["conferencia"] = dict(c, mudou=mudou)


def resumo(cfg, cliente, ano, mes, hoje=None, filtro=None):
    """Tudo o que o painel do cliente mostra para o mês (ou para o período/marketplace/conta do filtro)."""
    f = _filtro(filtro)
    hoje = hoje or dt.date.today()
    chave = "%04d-%02d" % (ano, mes)
    est = _estado_mes(cliente["id"], chave)
    saida = {"ok": True, "mes": chave, "cliente": cliente["nome"], "cliente_id": cliente["id"],
             "consolidado": {k: (est or {}).get(k) for k in ("ok", "conferido", "gerado", "erro", "fonte", "parcial", "faltando")}}
    if est and not est.get("ok") and (est.get("ultimo_ok") or {}).get("arquivo") and os.path.isfile(est["ultimo_ok"]["arquivo"]):
        # a última geração falhou: mostra o último consolidado bom e avisa
        saida["usando_anterior"] = {"erro": est.get("erro"), "gerado_falha": est.get("gerado")}
        est = est["ultimo_ok"]
        saida["consolidado"] = {k: est.get(k) for k in ("ok", "conferido", "gerado", "erro", "fonte", "parcial", "faltando")}
    arq = (est or {}).get("arquivo")
    if not est or not est.get("ok") or not arq or not os.path.isfile(arq):
        saida["sem_dados"] = True
        return saida
    real, fc = _ler(arq)
    desc_skus = real.attrs.get("desc") or {}
    if real.empty:
        saida["sem_dados"] = True
        return saida
    import pandas as pd
    datas = real["Data"].dropna()
    saida["opcoes"] = {"marketplaces": sorted(real["Marketplace"].dropna().unique().tolist()),
                       "contas": sorted(real["Marketplace venda"].dropna().astype(str).unique().tolist()),
                       "de": datas.min().date().isoformat() if len(datas) else None,
                       "ate": datas.max().date().isoformat() if len(datas) else None}
    saida["filtro"] = {"de": f["de"].isoformat() if f["de"] else None, "ate": f["ate"].isoformat() if f["ate"] else None,
                       "mks": f["mks"], "contas": f["contas"]}
    filtrado = bool(f["de"] or f["ate"] or f["mks"] or f["contas"])
    real = _aplicar(real, f)
    fc = _aplicar(fc, f, dias=False)
    if real.empty:
        saida["sem_dados_filtro"] = True
        return saida
    ult = real["Data"].max()
    saida["ate"] = ult.date().isoformat() if pd.notna(ult) else None
    saida["total"] = _soma(real)

    # por marketplace e por conta (Marketplace venda)
    grupos = []
    for (mk, conta), g in real.groupby(["Marketplace", "Marketplace venda"], observed=True):
        grupos.append(dict(_soma(g), marketplace=mk, conta=conta))
    grupos.sort(key=lambda x: -x["fat_bruto"])
    _marcar_conferencias(cliente["id"], chave, grupos, f)
    saida["por_canal"] = grupos
    saida["contas"] = sorted({g["conta"] for g in grupos})

    # dia a dia
    dias = []
    for d, g in real.groupby(real["Data"].dt.date):
        s = _soma(g)
        mks = g[g["k"] == "fat_bruto"].groupby("Marketplace", observed=True)["Valor completo"].sum().round(2).to_dict()
        dias.append({"dia": d.isoformat(), "fat_bruto": s["fat_bruto"], "fat_liquido": s["fat_liquido"],
                     "publicidade": s["publicidade"], "margem": s["margem"], "margem_pct": s["margem_pct"],
                     "unidades": s["unidades"], "por_mk": mks})
    saida["dias"] = dias

    # produtos: os que mais faturam e os de pior margem (só SKUs com venda)
    vend = real[(real["SKU"].astype(str) != "0") & (real["SKU"].astype(str) != "")]
    desc = desc_skus
    porsku = vend.pivot_table(index="SKU", columns="k", values="Valor completo", aggfunc="sum", fill_value=0.0,
                              observed=True)
    porsku.columns = [str(x) for x in porsku.columns]
    porsku.index = porsku.index.astype(str)
    for k, _, _ in MEDIDAS:
        if k not in porsku.columns:
            porsku[k] = 0.0
    if "unidades" not in porsku.columns:
        porsku["unidades"] = 0.0
    porsku["fat_liquido"] = porsku["fat_bruto"] + porsku["cancelamento"]
    porsku["margem"] = sum(porsku[k] for k, _, _ in MEDIDAS)
    porsku = porsku[porsku["fat_bruto"] > 0]

    def linha(sku, r):
        fl = float(r["fat_liquido"])
        return {"sku": sku, "descricao": str(desc.get(sku, "") or "")[:80], "unidades": round(float(r["unidades"]), 0),
                "fat_liquido": round(fl, 2), "publicidade": round(float(r["publicidade"]), 2),
                "margem": round(float(r["margem"]), 2), "margem_pct": round(float(r["margem"]) / fl, 4) if fl else None}
    top = porsku.sort_values("fat_liquido", ascending=False).head(10)
    pior = porsku[porsku["fat_liquido"] > 0].sort_values("margem").head(8)
    saida["top_produtos"] = [linha(s, r) for s, r in top.iterrows()]
    saida["piores_margens"] = [linha(s, r) for s, r in pior.iterrows() if r["margem"] < 0]
    saida["skus_vendidos"] = int(len(porsku))

    # alertas sobre os dados (o que pode distorcer os números)
    alertas = []
    sc = porsku[(porsku["custo"] == 0)]
    if len(sc):
        alertas.append({"nivel": "atencao", "titulo": "%d SKU(s) vendidos sem custo do produto" % len(sc),
                        "detalhe": "Faturamento líquido desses SKUs: margem fica superestimada. Ex.: " + ", ".join(list(sc.sort_values("fat_liquido", ascending=False).index[:6])),
                        "valor": round(float(sc["fat_liquido"].sum()), 2), "acao": "custos"})
    lanc = (est or {}).get("lancamentos") or {}
    for k, v in (lanc.get("ads_sem_depara") or {}).items():
        ids = v.get("ids") or {}
        if ids:
            alertas.append({"nivel": "atencao", "titulo": "%d anúncio(s) de ADS sem SKU no de-para (%s)" % (len(ids), k.split("|")[1]),
                            "detalhe": "Entraram com SKU 0. Ex.: " + ", ".join(list(ids)[:5]), "valor": round(-sum(ids.values()), 2), "acao": "ads"})
    if lanc.get("afiliados_sem_pedido"):
        alertas.append({"nivel": "info", "titulo": "%d comissão(ões) de afiliado sem pedido nos tratados" % lanc["afiliados_sem_pedido"],
                        "detalhe": "Entraram só com a comissão, sem ligar à venda", "valor": None, "acao": "afiliados"})
    if (est or {}).get("sem_descricao"):
        alertas.append({"nivel": "info", "titulo": "%d SKU(s) sem descrição na tabela de custos" % est["sem_descricao"],
                        "detalhe": "Ex.: " + ", ".join((est.get("exemplos_sem_descricao") or [])[:6]), "valor": None, "acao": "custos"})
    if not (est or {}).get("conferido"):
        alertas.insert(0, {"nivel": "erro", "titulo": "Consolidado com diferenças na conferência",
                           "detalhe": "Os números abaixo podem não bater com os tratados; veja o Pacote final", "valor": None, "acao": "pacote"})
    if (est or {}).get("faltando"):
        alertas.insert(0, {"nivel": "atencao", "titulo": "Tratados faltando no consolidado",
                           "detalhe": ", ".join(est["faltando"]), "valor": None, "acao": "tratamento"})
    if (est or {}).get("aviso_sem_vendas"):
        alertas.append({"nivel": "atencao", "titulo": "Pacote sem as vendas tratadas",
                        "detalhe": est["aviso_sem_vendas"], "valor": None, "acao": "vendas"})
    if (est or {}).get("aviso_anteriores"):
        alertas.append({"nivel": "atencao", "titulo": "Números com dados de dia anterior" + ((" (%s)" % est["dados_de"]) if est.get("dados_de") else ""),
                        "detalhe": est["aviso_anteriores"], "valor": None, "acao": "vendas"})
    if saida.get("usando_anterior"):
        alertas.insert(0, {"nivel": "atencao", "titulo": "Mostrando o último consolidado gerado com sucesso",
                           "detalhe": "A geração de %s falhou: %s" % ((saida["usando_anterior"].get("gerado_falha") or "")[:16].replace("T", " "),
                                                                     saida["usando_anterior"].get("erro") or ""), "valor": None, "acao": "pacote"})
    if (est or {}).get("parcial"):
        alertas.insert(0, {"nivel": "info", "titulo": "Mês parcial", "detalhe": "A pasta do dia 1 do mês seguinte ainda não existe: os números vão até a última pasta do mês.", "valor": None})
    saida["alertas"] = alertas

    # forecast definido (proporcional aos dias realizados)
    if not fc.empty:
        sf = _soma(fc)
        ndias = (pd.Timestamp(ult).day if pd.notna(ult) else 0)
        d0 = f["de"].day if f["de"] else 1
        fc_ate = fc[(fc["Data"].dt.day <= ndias) & (fc["Data"].dt.day >= d0)] if ndias else fc.iloc[0:0]
        saida["forecast"] = {"mes": sf, "ate": _soma(fc_ate) if len(fc_ate) else None}

    # mês anterior até o mesmo dia (comparação justa)
    a0, m0 = (ano - 1, 12) if mes == 1 else (ano, mes - 1)
    est0 = _estado_mes(cliente["id"], "%04d-%02d" % (a0, m0))
    if est0 and not est0.get("ok") and est0.get("ultimo_ok"):
        est0 = est0["ultimo_ok"]
    if est0 and est0.get("ok") and est0.get("arquivo") and os.path.isfile(est0["arquivo"]) and pd.notna(ult):
        try:
            r0, _ = _ler(est0["arquivo"])
            r0 = _aplicar(r0, f, dias=False)
            d0 = f["de"].day if f["de"] else 1
            r0d = r0[(r0["Data"].dt.day <= ult.day) & (r0["Data"].dt.day >= d0)]
            if len(r0d):
                saida["anterior"] = {"mes": "%04d-%02d" % (a0, m0), "ate_dia": int(ult.day), "de_dia": int(d0), "total": _soma(r0d)}
        except Exception as e:
            comum.log("painel: mês anterior de %s não lido: %s" % (cliente["nome"], e))
    return saida


def resumo_todos(cfg, ano, mes, filtro=None):
    """Uma linha por cliente (para a visão "Todos os clientes")."""
    linhas = []
    filtro = dict(filtro or {}, contas=[])          # contas são de cada cliente: aqui só período e marketplace
    for c in cfg.get("clientes", []):
        if c.get("ativo") is False:
            continue
        try:
            r = resumo(cfg, c, ano, mes, filtro=filtro)
        except Exception as e:
            r = {"cliente": c["nome"], "cliente_id": c["id"], "erro": str(e)[:200], "sem_dados": True}
        linhas.append({k: r.get(k) for k in ("cliente", "cliente_id", "total", "ate", "consolidado", "sem_dados", "erro", "anterior",
                                             "sem_dados_filtro", "opcoes")})
    return linhas


def cargas(cfg, cliente_id=None, limite=60):
    """Histórico de cargas: arquivos que entraram e o que foi gerado com eles."""
    tipos = {"organizado": "Relatório de vendas organizado", "upload_manual": "Relatório subido pelo app",
             "lancamento": "ADS / afiliados", "tratamento": "Tratamento", "consolidado": "Consolidado", "forecast": "Forecast definido"}
    nomes = {u["id"]: u["nome"] for u in comum.contas(cfg)}
    nomes.update({c["id"]: c["nome"] for c in cfg.get("clientes", [])})
    saida = []
    for e in reversed(comum.ler_historico(6000)):
        if e.get("tipo") not in tipos:
            continue
        cid = str(e.get("cliente") or "")
        if cliente_id and not (cid == cliente_id or cid.startswith(cliente_id + comum.SEP_CONTA)):
            continue
        mk = e.get("mk")
        item = {"ts": e.get("ts"), "tipo": e.get("tipo"), "rotulo": tipos[e["tipo"]], "quem": nomes.get(cid, cid),
                "mk": cfg["marketplaces"].get(mk, {}).get("nome", mk) if mk else None, "arquivo": e.get("origem")}
        if e["tipo"] == "lancamento":
            item["rotulo"] = "ADS" if e.get("subtipo") == "ads" else "Afiliados"
            dias = e.get("dias") or []
            item["detalhe"] = ("%d dia(s): %s" % (len(dias), ", ".join(d[8:10] + "/" + d[5:7] for d in dias[:6]) + (" …" if len(dias) > 6 else ""))) if dias else ""
        elif e["tipo"] == "tratamento":
            if e.get("ok"):
                d = e.get("diag") or {}
                item["detalhe"] = "%s linhas%s" % (d.get("linhas", "?"), (" · %s sem custo" % d["semCusto"]) if d.get("semCusto") else "")
            else:
                item["detalhe"], item["falha"] = (e.get("erro") or "")[:160], not e.get("pendente")
        elif e["tipo"] == "consolidado":
            item["detalhe"] = "%s · %s" % (e.get("mes"), "conferido" if e.get("conferido") else ("erro: " + (e.get("erro") or "")[:120] if not e.get("ok") else "com diferenças"))
            item["falha"] = not e.get("ok")
            item["arquivo"] = None
            item["origem_evento"] = e.get("origem")
        elif e["tipo"] == "forecast":
            item["detalhe"] = "%s · versão %s · %s" % (e.get("mes"), e.get("versao"), e.get("motivo") or "")
        elif e["tipo"] == "organizado":
            item["detalhe"] = e.get("identificacao") or ""
        saida.append(item)
        if len(saida) >= limite:
            break
    return saida


def resumo_tratado(caminho, mk=None):
    """Raio-x de um tratado: o que tem no arquivo, por status e por dia, e checagens."""
    import pandas as pd
    if not mk:
        base = os.path.basename(caminho)
        mk = next((k for k in C.MAPA if base.startswith("vendas_%s_" % k)), None)
    if mk not in C.MAPA:
        raise ValueError("Não sei ler este tratado (marketplace desconhecido).")
    m = C.MAPA[mk]
    df = C.ler_excel(caminho, sheet_name="DB", dtype={m["sku"]: str, m["pedido"]: str})
    num = lambda c: pd.to_numeric(df[c], errors="coerce").fillna(0.0) if c in df.columns else pd.Series(0.0, index=df.index)
    col_m = "Gross margin" if "Gross margin" in df.columns else ("Gross margen" if "Gross margen" in df.columns else None)
    col_l = "Liquido" if "Liquido" in df.columns else ("Líquido" if "Líquido" in df.columns else None)
    df["_st"] = df[m["status"]].astype(str).str.strip()
    df["_d"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
    medidas = {"Faturamento": num(m["fat"]), "Desconto": num(m["desc"]), "Rebate": num(m["rebate"]), "Comissão": num("Comissão"),
               "Frete": num("Frete"), "Cancelamento": num("Cancelamento"), "Antecipação": num("Antecipa"), "Imposto": num("Imposto"),
               "Custo do produto": num("Custo do produto"), "Líquido": num(col_l) if col_l else None,
               "Margem": num(col_m) if col_m else None, "Unidades": num(m["qtd"])}
    medidas = {k: v for k, v in medidas.items() if v is not None}
    por_status = []
    for st, idx in df.groupby("_st").groups.items():
        linha = {"status": st, "linhas": int(len(idx)), "pedidos": int(df.loc[idx, m["pedido"]].nunique()) if m["pedido"] in df else None}
        linha.update({k: round(float(v.loc[idx].sum()), 2) for k, v in medidas.items()})
        por_status.append(linha)
    por_status.sort(key=lambda x: -x["linhas"])
    total = {k: round(float(v.sum()), 2) for k, v in medidas.items()}
    venda = df["_st"].str.lower().isin(C.VENDA)
    dias = []
    for d, idx in df[venda].groupby(df.loc[venda, "_d"].dt.date).groups.items():
        dias.append({"dia": d.isoformat(), "faturamento": round(float(medidas["Faturamento"].loc[idx].sum()), 2),
                     "pedidos": int(df.loc[idx, m["pedido"]].nunique()) if m["pedido"] in df else None,
                     "margem": round(float(medidas["Margem"].loc[idx].sum()), 2) if "Margem" in medidas else None})
    dias.sort(key=lambda x: x["dia"])
    # checagens
    chk = []
    sem_custo = df[venda & (medidas["Custo do produto"] == 0)]
    if len(sem_custo):
        skus = sem_custo[m["sku"]].astype(str).value_counts().head(8)
        chk.append({"nivel": "atencao", "titulo": "%d venda(s) sem custo do produto" % len(sem_custo),
                    "detalhe": "SKUs: " + ", ".join("%s (%d)" % (s, n) for s, n in skus.items()),
                    "valor": round(float(medidas["Faturamento"][venda & (medidas["Custo do produto"] == 0)].sum()), 2)})
    canc = df["_st"].str.lower().isin(C.CANCELADA)
    if "Margem" in medidas:
        cpos = canc & (medidas["Margem"] > 0.005)
        if cpos.any():
            chk.append({"nivel": "atencao", "titulo": "%d cancelada(s)/devolução(ões) com margem positiva" % int(cpos.sum()),
                        "detalhe": "Margem somada dessas linhas", "valor": round(float(medidas["Margem"][cpos].sum()), 2)})
        neg = venda & (medidas["Margem"] < 0)
        if neg.any():
            chk.append({"nivel": "info", "titulo": "%d venda(s) com margem negativa" % int(neg.sum()),
                        "detalhe": "Margem somada dessas vendas", "valor": round(float(medidas["Margem"][neg].sum()), 2)})
    sem_data = int(df["_d"].isna().sum())
    if sem_data:
        chk.append({"nivel": "atencao", "titulo": "%d linha(s) sem data válida" % sem_data, "detalhe": "", "valor": None})
    if m["pedido"] in df and m["sku"] in df:
        dup = df.duplicated(subset=[m["pedido"], m["sku"], m["status"], m["fat"]], keep=False) & venda
        if dup.any():
            chk.append({"nivel": "info", "titulo": "%d linha(s) repetidas (mesmo pedido, SKU e valor)" % int(dup.sum()),
                        "detalhe": "Pode ser kit ou pacote com itens iguais; confira se for inesperado", "valor": None})
    if not chk:
        chk.append({"nivel": "ok", "titulo": "Nenhum problema encontrado", "detalhe": "", "valor": None})
    datas = df["_d"].dropna()
    return {"ok": True, "mk": mk, "marketplace": m["nome"], "arquivo": caminho, "linhas": int(len(df)),
            "de": datas.min().date().isoformat() if len(datas) else None, "ate": datas.max().date().isoformat() if len(datas) else None,
            "pedidos": int(df[m["pedido"]].nunique()) if m["pedido"] in df else None,
            "skus": int(df.loc[venda, m["sku"]].nunique()) if m["sku"] in df else None,
            "total": total, "por_status": por_status, "dias": dias, "checagens": chk}
