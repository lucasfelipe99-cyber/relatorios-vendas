# -*- coding: utf-8 -*-
"""Ponte entre a rotina de download e o tratador de vendas (projeto automacao-vendas-local).

As regras de cálculo ficam nos módulos tratamento_*.py, copiados sem alteração do projeto
original. Aqui só mudamos: loja = nome da conta, imposto por cliente e custos por cliente
(tabela SKU, Descrição, Custo, Categoria, convertida para o formato que os tratamentos leem).
Saída: [raiz]\\[DD-MM]\\[Cliente]\\tratados\\vendas_[mk]_[DD-MM]_tratado.xlsx
"""
import datetime as dt
import glob
import json
import os
import re
import shutil
import tempfile

import comum

PASTA_TRATADOS = "tratados"
IMPOSTO_PADRAO = 0.14
EXTS_CUSTOS = (".xlsm", ".xlsx", ".csv")
EXTS_IDS = (".xlsx", ".xlsm", ".csv")


# ---------------------------------------------------------------- arquivos por cliente
def pasta_custos():
    p = os.path.join(comum.PASTA_SISTEMA, "dados", "custos")
    os.makedirs(p, exist_ok=True)
    return p


def mes_chave(ano, mes):
    return "%04d-%02d" % (ano, mes)


def mes_do_periodo(hoje=None):
    """Mês a que os relatórios de hoje se referem (dia 1 = mês anterior)."""
    ini, _ = comum.periodo(hoje or dt.date.today())
    return ini.year, ini.month


def escolher_custos(cliente, ano=None, mes=None):
    """Tabela de custos que vale para o mês: a última versão enviada para aquele mês; sem tabela no mês,
    a do mês anterior mais recente que tiver; sem nenhuma anterior, a mais antiga. Devolve (info, mes_usado)."""
    if ano is None:
        ano, mes = mes_do_periodo()
    alvo = mes_chave(ano, mes)
    meses = {k: v for k, v in (cliente.get("custos_meses") or {}).items()
             if os.path.isfile(os.path.join(comum.PASTA_SISTEMA, "dados", (v or {}).get("arquivo", "") or "-"))}
    if alvo in meses:
        return meses[alvo], alvo
    antes = sorted(k for k in meses if k < alvo)
    if antes:
        return meses[antes[-1]], antes[-1]
    legado = cliente.get("custos") or {}
    if legado.get("arquivo") and not legado.get("mes") and os.path.isfile(os.path.join(comum.PASTA_SISTEMA, "dados", legado["arquivo"])):
        return legado, None
    if meses:
        k = sorted(meses)[0]
        return meses[k], k
    return None, None


def caminho_custos(cliente, ano=None, mes=None):
    info, _ = escolher_custos(cliente, ano, mes)
    if info and info.get("arquivo"):
        p = os.path.join(comum.PASTA_SISTEMA, "dados", info["arquivo"])
        if os.path.isfile(p):
            return p
    return None


def caminho_ids_shopee(cliente):
    arq = (cliente.get("ids_shopee") or {}).get("arquivo")
    if arq:
        p = os.path.join(comum.PASTA_SISTEMA, "dados", arq)
        if os.path.isfile(p):
            return p
    return None


def caminho_ids_magalu():
    p = os.path.join(comum.PASTA_SISTEMA, "magalu-ids.csv")
    return p if os.path.isfile(p) else None


def imposto(cliente):
    try:
        v = float(cliente.get("imposto", IMPOSTO_PADRAO))
        return v if 0 <= v < 1 else IMPOSTO_PADRAO
    except (TypeError, ValueError):
        return IMPOSTO_PADRAO


# ---------------------------------------------------------------- tabela de custos (SKU, Descrição, Custo, Categoria)
COLS_CUSTOS = ["SKU", "Descrição", "Custo", "Categoria"]
_APELIDOS = {
    "SKU": {"sku", "codigosku", "skuproduto", "codigo", "codigodoproduto", "cod", "codsku", "skuseller"},
    "Descrição": {"descricao", "descricaodoproduto", "produto", "nomedoproduto", "nome", "titulo"},
    "Custo": {"custo", "custounitario", "custodoproduto", "custoproduto", "valordecusto", "valorcusto", "custototal", "custorz", "custor"},
    "Categoria": {"categoria", "categoriadoproduto", "categoriaproduto"},
}
OBRIGATORIAS = ["SKU", "Descrição", "Custo"]


def _chave(txt):
    import unicodedata
    t = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode().lower()
    return "".join(ch for ch in t if ch.isalnum())


def _mapear(cabecalho):
    """{coluna padrão: índice} para uma linha de cabeçalho."""
    achadas = {}
    for i, nome in enumerate(cabecalho):
        k = _chave(nome)
        for padrao, apelidos in _APELIDOS.items():
            if k in apelidos and padrao not in achadas:
                achadas[padrao] = i
    return achadas


def _texto_sku(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if v.is_integer():
            return str(int(v))
    return str(v).strip()


def _numero(v):
    """Aceita 12,5 · 12.50 · 1.234,56 · R$ 1.234,56 · número do Excel. Devolve (valor|None, ok)."""
    if v is None:
        return None, True
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (None, True) if v != v else (float(v), True)
    t = str(v).strip().replace("R$", "").replace("\u00a0", "").replace(" ", "")
    if t in ("", "-"):
        return None, True
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t), True
    except ValueError:
        return None, False


def _linhas_da_planilha(caminho):
    """Devolve [(nome_aba, [linhas])] lendo CSV ou Excel como valores brutos."""
    if caminho.lower().endswith(".csv"):
        import csv
        bruto = open(caminho, "rb").read()
        for cod in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                txt = bruto.decode(cod)
                break
            except UnicodeDecodeError:
                continue
        try:
            dial = csv.Sniffer().sniff(txt[:5000], delimiters=";,\t|")
        except csv.Error:
            class dial(csv.excel):
                delimiter = ";" if txt[:2000].count(";") > txt[:2000].count(",") else ","
        return [("CSV", [r for r in csv.reader(txt.splitlines(), dial)])]
    from openpyxl import load_workbook
    try:
        wb = load_workbook(caminho, read_only=True, data_only=True)
        try:
            abas = [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
        finally:
            wb.close()
        # planilhas geradas por alguns sistemas (ex.: TikTok) sem a referência das células: o modo rápido
        # do openpyxl enxerga só a 1ª coluna. Nesse caso relê com o calamine.
        if abas and all(max((len(r) for r in linhas), default=0) <= 1 for _, linhas in abas) and os.path.getsize(caminho) > 3000:
            raise ValueError("só uma coluna lida")
        return abas
    except Exception:
        # planilhas que o openpyxl não abre (ex.: mass_update da Shopee): lê com o calamine
        from python_calamine import CalamineWorkbook
        wbc = CalamineWorkbook.from_path(caminho)
        return [(n, [list(r) for r in wbc.get_sheet_by_name(n).to_python()]) for n in wbc.sheet_names]


def ler_custos(caminho):
    """Lê a tabela de custos e devolve (linhas padronizadas, info).

    Formato: SKU, Descrição, Custo e Categoria (opcional), em qualquer aba e com o cabeçalho
    em qualquer uma das primeiras 20 linhas. Também aceita a tabela antiga do projeto de
    tratamento (SKU, Sub Categoria, Valor), somando os componentes de cada SKU como antes.
    """
    abas = _linhas_da_planilha(caminho)
    melhor_faltando = None
    for aba, linhas in abas:
        for n, cab in enumerate(linhas[:20]):
            m = _mapear(cab)
            faltam = [c for c in OBRIGATORIAS if c not in m]
            if not faltam:
                return _padronizar(linhas[n + 1:], m, n + 2, aba)
            if "SKU" in m and len(faltam) < len(OBRIGATORIAS) - 1 and (melhor_faltando is None or len(faltam) < len(melhor_faltando[1])):
                melhor_faltando = (aba, faltam)
    # tabela antiga (1-Custo total.xlsm)
    try:
        from tratamento_mercado_livre import _carregar_custos
        df = _carregar_custos(caminho)
        linhas = [{"SKU": _texto_sku(r["SKU"]), "Descrição": "", "Custo": None if r["Valor total"] != r["Valor total"] else float(r["Valor total"]), "Categoria": ""}
                  for _, r in df.iterrows() if _texto_sku(r["SKU"])]
        return linhas, _resumo(linhas, {"formato": "antigo (SKU, Sub Categoria, Valor)"})
    except Exception:
        pass
    if melhor_faltando:
        raise ValueError("na aba \u201c%s\u201d faltam as colunas: %s. A planilha precisa ter SKU, Descrição e Custo (Categoria é opcional)."
                         % (melhor_faltando[0], ", ".join(melhor_faltando[1])))
    raise ValueError("não encontrei o cabeçalho com SKU, Descrição e Custo (Categoria é opcional). Use o modelo da planilha.")


def _padronizar(linhas, m, primeira_linha, aba):
    saida, invalidos, negativos, vistos, conflitos = [], [], [], {}, []
    for i, r in enumerate(linhas):
        def col(c):
            j = m.get(c)
            return r[j] if j is not None and j < len(r) else None
        sku = _texto_sku(col("SKU"))
        if not sku:
            continue
        custo, ok = _numero(col("Custo"))
        n_linha = primeira_linha + i
        if not ok:
            invalidos.append("linha %d (%s): \u201c%s\u201d" % (n_linha, sku, col("Custo")))
            continue
        if custo is not None and custo < 0:
            negativos.append("linha %d (%s)" % (n_linha, sku))
            continue
        item = {"SKU": sku, "Descrição": str(col("Descrição") or "").strip(), "Custo": custo,
                "Categoria": str(col("Categoria") or "").strip()}
        if sku in vistos:
            ant = vistos[sku]
            if ant["Custo"] is None and custo is not None:
                ant.update(item)
            elif custo is not None and abs((ant["Custo"] or 0) - custo) > 1e-9:
                # SKU repetido com custos diferentes: fica o MAIOR custo (avisado no envio)
                conflitos.append("%s (%s e %s → %s)" % (sku, _br(ant["Custo"]), _br(custo), _br(max(ant["Custo"], custo))))
                if custo > ant["Custo"]:
                    ant.update({k: v for k, v in item.items() if v not in ("", None)})
            continue
        vistos[sku] = item
        saida.append(item)
    erros = []
    if invalidos:
        erros.append("Custo que não é número em %d linha(s), ex.: %s" % (len(invalidos), "; ".join(invalidos[:3])))
    if negativos:
        erros.append("Custo negativo em %d linha(s), ex.: %s" % (len(negativos), "; ".join(negativos[:3])))
    if erros:
        raise ValueError(" · ".join(erros) + ". Corrija e envie de novo.")
    if not saida:
        raise ValueError("a aba \u201c%s\u201d tem o cabeçalho, mas nenhuma linha com SKU." % aba)
    return saida, _resumo(saida, {"formato": "SKU, Descrição, Custo, Categoria", "aba": aba,
                                  "tem_categoria": "Categoria" in m,
                                  "repetidos_maior": len({c.split(" (")[0] for c in conflitos}), "repetidos_ex": conflitos[:5]})


def _br(v):
    return "vazio" if v is None else ("%.2f" % v).replace(".", ",")


def _resumo(linhas, extra):
    info = {"skus": len(linhas), "sem_valor": sum(1 for l in linhas if l["Custo"] is None),
            "com_categoria": sum(1 for l in linhas if l["Categoria"])}
    info.update(extra)
    return info


def gravar_custos_padrao(linhas, destino):
    """Grava a tabela padronizada (o que fica guardado por cliente). SKU sempre como texto."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "Custos"
    ws.append(COLS_CUSTOS)
    for l in linhas:
        ws.append([l["SKU"], l["Descrição"], l["Custo"], l["Categoria"]])
    for c in ws[1]:
        c.font = Font(bold=True)
    for linha in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        linha[0].number_format = "@"
    for linha in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        linha[0].number_format = '#,##0.00'
    for col, larg in zip("ABCD", (22, 50, 14, 24)):
        ws.column_dimensions[col].width = larg
    ws.freeze_panes = "A2"
    wb.save(destino)


def custos_para_tratador(caminho, temp):
    """Converte a tabela do cliente para o formato que os tratamentos leem (SKU, Sub Categoria, Valor).

    SKU só de dígitos (sem zero à esquerda) vira número, como o Excel guardaria na tabela original,
    para casar com o SKU lido dos relatórios; os demais ficam como texto. Sub Categoria fica vazia
    (a categoria do produto não deve cair no filtro de Energia/Mão de Obra do tratamento)."""
    linhas, _ = ler_custos(caminho)
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Sub Categoria", "Valor"])
    for l in linhas:
        s = l["SKU"]
        sku = int(s) if s.isdigit() and not (len(s) > 1 and s[0] == "0") and len(s) < 16 else s
        ws.append([sku, "", l["Custo"]])
    destino = os.path.join(temp, "custos-tratador.xlsx")
    wb.save(destino)
    return destino


def validar_custos(caminho):
    return ler_custos(caminho)[1]


def caminho_modelo_custos():
    p = os.path.join(pasta_custos(), "modelo-custos.xlsx")
    if not os.path.isfile(p):
        gravar_custos_padrao([
            {"SKU": "ABC-001", "Descrição": "Exemplo: Bomba 1/2 CV", "Custo": 189.9, "Categoria": "Bombas"},
            {"SKU": "7891234567890", "Descrição": "Exemplo: Cloro 10 kg", "Custo": 74.5, "Categoria": ""},
        ], p)
    return p


def validar_ids_shopee(caminho):
    from tratamento_shopee import _ler_ids
    return {"linhas": int(len(_ler_ids(caminho)))}


def _ext_real(nome_original, dados_bytes, permitidas):
    ext = os.path.splitext(nome_original)[1].lower()
    if dados_bytes[:2] == b"PK" and ext not in (".xlsx", ".xlsm"):
        ext = ".xlsx"   # conteúdo é Excel (ex.: cópia da tabela guardada de outro cliente)
    return ext


def guardar_arquivo(cliente_id, tipo, nome_original, dados_bytes, mes=None):
    """Valida e guarda o arquivo do cliente; o anterior vai para backups. Devolve info para o config.

    Custos: o arquivo enviado é lido, validado e guardado padronizado (SKU, Descrição, Custo, Categoria)
    em custos/<cliente>__custos.xlsx; o original enviado vai para backups."""
    permitidas = EXTS_CUSTOS if tipo == "custos" else EXTS_IDS
    ext = _ext_real(nome_original, dados_bytes, permitidas)
    if ext not in permitidas:
        raise ValueError("Formato não aceito (%s). Use %s." % (ext or "sem extensão", ", ".join(permitidas)))
    pasta = pasta_custos()
    prefixo = "%s__%s" % (cliente_id, tipo) + (("_" + mes) if (tipo == "custos" and mes) else "")
    tmp = os.path.join(pasta, ".%s.novo%s" % (prefixo, ext))
    with open(tmp, "wb") as f:
        f.write(dados_bytes)
    try:
        if tipo == "custos":
            linhas, info = ler_custos(tmp)
            padrao = os.path.join(pasta, ".%s.padrao.xlsx" % prefixo)
            gravar_custos_padrao(linhas, padrao)
        else:
            info = validar_ids_shopee(tmp)
    except Exception as e:
        os.replace(tmp, tmp + ".recusado")
        msg = str(e).replace(tmp, os.path.basename(nome_original))
        raise ValueError("Arquivo não aceito: %s" % msg)
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    bkps = comum.pasta_dados(None, "backups")
    for antigo in glob.glob(os.path.join(pasta, prefixo + ".*")):
        shutil.move(antigo, os.path.join(bkps, "%s_%s%s" % (prefixo, carimbo, os.path.splitext(antigo)[1])))
    if tipo == "custos":
        shutil.move(tmp, os.path.join(bkps, "%s_enviado_%s%s" % (prefixo, carimbo, ext)))
        final = os.path.join(pasta, prefixo + ".xlsx")
        os.replace(padrao, final)
    else:
        final = os.path.join(pasta, prefixo + ext)
        os.replace(tmp, final)
    info.update({"arquivo": "custos/" + os.path.basename(final), "nome_original": os.path.basename(nome_original),
                 "atualizado": dt.datetime.now().isoformat(timespec="seconds")})
    if tipo == "custos" and mes:
        info["mes"] = mes
    return info


# ---------------------------------------------------------------- SKUs vendidos sem custo entram na tabela
# coluna do SKU (a mesma usada para buscar o custo) e do nome do produto em cada tratado
SKU_TRATADO = {"mercadolivre": ("SKU", "Título do anúncio"), "shopee": ("Número de referência SKU", "Nome do Produto"),
               "magalu": ("SKU.1", "Título do produto"), "tiktokshop": ("Seller SKU", "Product Name"),
               "amazon": ("SKU.1", "product-name"), "viavarejo": ("SKU", "Título do anúncio"),
               "temu": ("SKU", "Título do anúncio"), "madeiramadeira": ("SKU", "Título do anúncio")}
LIMITE_AUTO = 3000


def _info_para_salvar(info, linhas):
    r = _resumo(linhas, {})
    info.update(skus=r["skus"], sem_valor=r["sem_valor"], com_categoria=r["com_categoria"])
    return info


def completar_custos(cliente, mk, tratado, hoje=None):
    """Acrescenta à tabela de custos do mês (com custo vazio) os SKUs vendidos que ainda não estão nela.
    Sem tabela nenhuma: cria a tabela do mês só com esses SKUs. Devolve os SKUs acrescentados."""
    if mk not in SKU_TRATADO or not os.path.isfile(tratado):
        return []
    col_sku, col_nome = SKU_TRATADO[mk]
    try:
        import pandas as pd
        df = pd.read_excel(tratado, dtype=str, engine="calamine")
    except Exception:
        try:
            import pandas as pd
            df = pd.read_excel(tratado, dtype=str)
        except Exception:
            return []
    if col_sku not in df.columns:
        return []
    vendidos = {}
    for sku, nome in zip(df[col_sku], df[col_nome] if col_nome in df.columns else [""] * len(df)):
        s = _texto_sku(sku)
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        if s and s.lower() not in ("nan", "none", "-") and s not in vendidos:
            vendidos[s] = "" if str(nome) in ("nan", "None") else str(nome or "").strip()
    if not vendidos:
        return []
    cid = cliente.get("cliente_id") or cliente["id"]
    ano, mes = mes_do_periodo(hoje)
    with comum.TRAVA_CFG:
        cfg = comum.carregar_config()
        grupo = next((c for c in cfg["clientes"] if c["id"] == cid), None)
        if not grupo:
            return []
        info, mes_usado = escolher_custos(grupo, ano, mes)
        caminho = caminho_custos(grupo, ano, mes)
        linhas = []
        if caminho:
            try:
                linhas, _ = ler_custos(caminho)
            except Exception:
                return []          # tabela com problema: não mexe
        existentes = set()
        for l in linhas:
            existentes.add(l["SKU"])
            existentes.add(l["SKU"].lstrip("0") or l["SKU"])
        novos = [k for k in vendidos if k not in existentes and (k.lstrip("0") or k) not in existentes][:LIMITE_AUTO]
        if not novos:
            return []
        nome_mk = cfg.get("marketplaces", {}).get(mk, {}).get("nome", mk)
        linhas += [{"SKU": k, "Descrição": vendidos[k][:200], "Custo": None, "Categoria": ""} for k in novos]
        agora = dt.datetime.now().isoformat(timespec="seconds")
        if caminho and info is not None:
            tmp = caminho + ".completando.xlsx"
            gravar_custos_padrao(linhas, tmp)
            try:
                os.replace(tmp, caminho)
            except OSError:          # aberto no Excel: tenta no próximo tratamento
                os.remove(tmp)
                return []
            alvo = info
        else:
            chave = mes_chave(ano, mes)
            padrao = os.path.join(tempfile.mkdtemp(prefix="custos_"), "custos.xlsx")
            gravar_custos_padrao(linhas, padrao)
            with open(padrao, "rb") as f:
                alvo = guardar_arquivo(cid, "custos", "criada pelo sistema (SKUs vendidos sem custo).xlsx", f.read(), chave)
            alvo["versao"] = 1
            grupo.setdefault("custos_meses", {})[chave] = alvo
            grupo["custos"] = alvo
            mes_usado = chave
        _info_para_salvar(alvo, linhas)
        auto = alvo.setdefault("adicionados_auto", {})
        for k in novos:
            auto[k] = {"mk": nome_mk, "em": agora}
        if len(auto) > LIMITE_AUTO:
            for k in list(auto)[:len(auto) - LIMITE_AUTO]:
                auto.pop(k)
        alvo["atualizado_auto"] = agora
        comum.salvar_config(cfg)
    comum.log("[%s] %d SKU(s) vendidos sem custo acrescentados à tabela de custos de %s (%s): %s" % (
        cliente.get("cliente_nome") or cliente["nome"], len(novos), mes_usado or "?", nome_mk, ", ".join(novos[:10])))
    return novos


def ler_tabela_mes(cliente_grupo, ano, mes):
    """Linhas da tabela que vale para o mês, para editar no app."""
    info, mes_usado = escolher_custos(cliente_grupo, ano, mes)
    caminho = caminho_custos(cliente_grupo, ano, mes)
    linhas = ler_custos(caminho)[0] if caminho else []
    return linhas, info, mes_usado


def salvar_tabela_mes(cliente_id, mes, linhas_in):
    """Valida as linhas editadas no app e grava como nova versão da tabela do mês. Devolve (linhas, info)."""
    saida, vistos, erros = [], set(), []
    for i, l in enumerate(linhas_in or [], 1):
        sku = _texto_sku(l.get("SKU"))
        if not sku:
            continue
        if sku in vistos:
            erros.append("SKU repetido: %s" % sku)
            continue
        vistos.add(sku)
        c = l.get("Custo")
        if c in ("", None):
            custo = None
        else:
            custo, ok = _numero(c)
            if not ok or (custo is not None and custo < 0):
                erros.append("custo inválido em %s: %s" % (sku, c))
                continue
        saida.append({"SKU": sku, "Descrição": str(l.get("Descrição") or "").strip()[:300], "Custo": custo,
                      "Categoria": str(l.get("Categoria") or "").strip()[:100]})
    if erros:
        raise ValueError("; ".join(erros[:5]) + ("…" if len(erros) > 5 else ""))
    if not saida:
        raise ValueError("A tabela ficou sem nenhum SKU.")
    padrao = os.path.join(tempfile.mkdtemp(prefix="custos_"), "custos.xlsx")
    gravar_custos_padrao(saida, padrao)
    with open(padrao, "rb") as f:
        info = guardar_arquivo(cliente_id, "custos", "editada no app.xlsx", f.read(), mes)
    return saida, info


def ler_planilha_custos_parcial(caminho):
    """Planilha para completar a tabela no editor: basta SKU e Custo (Descrição e Categoria opcionais).
    Devolve (linhas, avisos). Não grava nada."""
    for aba, linhas in _linhas_da_planilha(caminho):
        for n, cab in enumerate(linhas[:20]):
            m = _mapear(cab)
            if "SKU" in m and "Custo" in m:
                saida, invalidos, vistos = [], [], {}
                for i, r in enumerate(linhas[n + 1:]):
                    col = lambda c: (r[m[c]] if c in m and m[c] < len(r) else None)
                    sku = _texto_sku(col("SKU"))
                    if not sku:
                        continue
                    custo, ok = _numero(col("Custo"))
                    if not ok or (custo is not None and custo < 0):
                        invalidos.append("linha %d (%s): “%s”" % (n + 2 + i, sku, col("Custo")))
                        continue
                    item = {"SKU": sku, "Descrição": str(col("Descrição") or "").strip() if "Descrição" in m else None,
                            "Custo": custo, "Categoria": str(col("Categoria") or "").strip() if "Categoria" in m else None}
                    if sku in vistos:   # repetido: fica o maior custo, como no envio da tabela
                        ant = vistos[sku]
                        if custo is not None and (ant["Custo"] is None or custo > ant["Custo"]):
                            ant.update(item)
                        continue
                    vistos[sku] = item
                    saida.append(item)
                return saida, {"aba": aba, "invalidos": invalidos}
    raise ValueError("não encontrei o cabeçalho com as colunas SKU e Custo (Descrição e Categoria são opcionais).")


def exportar_custos(cliente, mes, linhas, filtro=""):
    """Grava as linhas (como estão no editor) num Excel em Relatórios/_exportados para abrir e preencher."""
    cfg = comum.carregar_config()
    pasta = os.path.join(comum.raiz(cfg), "_exportados")
    os.makedirs(pasta, exist_ok=True)
    nome = "custos_%s_%s%s_%s.xlsx" % (_seguro(cliente["nome"]), mes, ("_" + _seguro(filtro)) if filtro else "",
                                         dt.datetime.now().strftime("%H%M%S"))
    destino = os.path.join(pasta, nome)
    padrao = []
    for l in linhas or []:
        sku = _texto_sku(l.get("SKU"))
        if not sku:
            continue
        custo, ok = _numero(l.get("Custo"))
        padrao.append({"SKU": sku, "Descrição": str(l.get("Descrição") or ""), "Custo": custo if ok else None,
                       "Categoria": str(l.get("Categoria") or "")})
    gravar_custos_padrao(padrao, destino)
    return destino, len(padrao)


def _seguro(t):
    import re
    return re.sub(r'[\\/:*?"<>|\s]+', "-", str(t or "")).strip("-")[:60]


# ---------------------------------------------------------------- tratamento
def pasta_saida(cfg, cliente, hoje):
    return os.path.join(comum.pasta_dia(cfg, comum.pasta_de(cliente), hoje), PASTA_TRATADOS)


def caminho_saida(cfg, cliente, mk, hoje):
    return os.path.join(pasta_saida(cfg, cliente, hoje), "vendas_%s_%s_tratado.xlsx" % (mk, comum.rotulo_dia(hoje)))


def _custos_ou_vazio(cliente, temp, hoje=None):
    ano, mes = mes_do_periodo(hoje)
    c = caminho_custos(cliente, ano, mes)
    if c:
        return custos_para_tratador(c, temp), False
    vazio = os.path.join(temp, "custos-vazios.csv")
    with open(vazio, "w", encoding="utf-8") as f:
        f.write("SKU,Sub Categoria,Valor\n")
    return vazio, True


def _amazon_txt(pasta, dia, bruto, temp):
    """O tratador da Amazon lê o TXT original (tabulado). O organizador guarda o original em _originais."""
    candidatos = sorted(glob.glob(os.path.join(pasta, "_originais", "vendas_amazon_%s*.txt" % dia)), key=os.path.getmtime)
    if candidatos:
        return candidatos[-1]
    if os.path.isfile(bruto) and bruto.lower().endswith(".xlsx"):
        import pandas as pd
        df = pd.read_excel(bruto, dtype=str).fillna("")
        txt = os.path.join(temp, "amazon.txt")
        df.to_csv(txt, sep="\t", index=False, encoding="utf-8")
        return txt
    return None


def _mercado_livre(bruto, custos, saida, aliquota, loja):
    # Igual a processar_vendas do projeto original, mas devolvendo um diagnóstico.
    import tratamento_mercado_livre as ml
    resultado = ml.transformar_vendas(ml._ler_vendas(bruto), ml._carregar_custos(custos), aliquota,
                                      None, marketplace_venda=loja)
    resultado.to_excel(saida, sheet_name="DB", index=False, engine="openpyxl")
    ml._formatar_saida(__import__("pathlib").Path(saida))
    return {"linhas": int(len(resultado)), "semCusto": int(resultado["Custo do produto"].isna().sum())}


NAO_PAGO = {"não pago", "nao pago"}


def zerar_canceladas(caminho, mk):
    """Linhas que não são venda (todas as plataformas): o faturamento continua aparecendo, mas
      - Não pago: sem cobrança nenhuma (comissão, frete e taxa de envio reversa = 0) e Líquido 0;
      - Cancelada / Devolução: Líquido = só o que foi cobrado (0 se não houve cobrança).
          Com coluna Cancelamento preenchida (Mercado Livre, Madeira Madeira, Temu...): o relatório já traz o estorno,
          então a cobrança é o Líquido do próprio relatório quando ele é negativo (positivo = dinheiro que ainda vai voltar: 0).
          Sem estorno no relatório (Shopee, Magalu): cobrança = comissão + frete (+ taxa de envio reversa na Shopee;
          TikTok: comissão e frete são só estimativa e não são cobrados no cancelado -> 0;
          o frete da Shopee é o valor fixo do cadastro, não é cobrança, então sai);
      - sem estorno no relatório: o faturamento (com desconto e rebate) sai na coluna Cancelamento, então
        Faturamento + Desconto + Rebate + Comissão + Frete + Cancelamento (+ taxa reversa) = Líquido;
      - Antecipa, Imposto e Custo do produto = 0; Gross margin = Líquido; Gross margin % em branco.
    Devolve quantas linhas foram ajustadas."""
    import consolidado
    from openpyxl import load_workbook
    m = consolidado.MAPA.get(mk, {})
    if not m.get("status"):
        return 0
    wb = load_workbook(caminho)
    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
    cab = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    js = cab.get(m["status"])
    if not js:
        wb.close()
        return 0
    jl = cab.get("Liquido") or cab.get("Líquido")
    jm = cab.get("Gross margin") or cab.get("Gross margen")
    jp = cab.get("Gross margin %") or cab.get("Gross margen %")
    jcom, jfre, jcan, jrev = cab.get("Comissão"), cab.get("Frete"), cab.get("Cancelamento"), cab.get("Taxa de Envio Reversa")
    jfat, jdesc, jreb = cab.get(m.get("fat")), cab.get(m.get("desc")), cab.get(m.get("rebate"))
    zerar = [cab.get(c) for c in ("Antecipa", "Imposto", "Custo do produto")]

    def num(v):
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0
    n = 0
    for r in range(2, ws.max_row + 1):
        st = str(ws.cell(r, js).value or "").strip().lower()
        if st not in consolidado.CANCELADA and st not in NAO_PAGO:
            continue
        g = lambda j: num(ws.cell(r, j).value) if j else 0.0
        # TikTok: comissão e frete são estimados pela regra fixa e não são cobrados em pedido cancelado
        if st in NAO_PAGO or mk == "tiktokshop":
            for j in (jcom, jfre, jrev):
                if j:
                    ws.cell(r, j).value = 0
            cobranca = 0.0
        estorno = mk == "mercadolivre" or (mk != "tiktokshop" and jcan and abs(g(jcan)) > 0.005)
        if not estorno and jcan:
            # sem estorno no relatório (Shopee, Magalu, TikTok): o faturamento entra e sai em Cancelamento
            ws.cell(r, jcan).value = round(-(g(jfat) + g(jdesc) + g(jreb)), 6)
        if st in NAO_PAGO:
            pass
        elif estorno:
            cobranca = min(0.0, g(jl))
            if jcan and abs(cobranca - g(jl)) > 0.005:
                # dinheiro que ainda vai voltar (ex.: devolução em andamento): o ajuste vai para o Cancelamento,
                # para as colunas continuarem somando o Líquido
                ws.cell(r, jcan).value = round(g(jcan) + cobranca - g(jl), 6)
        else:
            rev = -abs(g(jrev))
            if mk == "shopee":
                if jfre:
                    ws.cell(r, jfre).value = 0
                cobranca = g(jcom) + rev
            else:
                cobranca = g(jcom) + g(jfre) + rev
            cobranca = min(0.0, cobranca)
        cobranca = round(cobranca, 6)
        if jl:
            ws.cell(r, jl).value = cobranca
        for j in zerar:
            if j:
                ws.cell(r, j).value = 0
        if jm:
            ws.cell(r, jm).value = cobranca
        if jp:
            ws.cell(r, jp).value = None
        n += 1
    if n:
        wb.save(caminho)
    wb.close()
    return n


def _norm_sku(t):
    return re.sub(r"\s+", " ", str(t or "")).strip().lower()


def sku_por_titulo(caminho, mk, custos):
    """Venda sem SKU no relatório da plataforma: o título do anúncio vira o SKU (fica na tabela de custos para preencher).
    Se a tabela já tem esse título com custo, o custo e a margem da linha são calculados. Devolve quantas linhas mudaram."""
    import consolidado
    from openpyxl import load_workbook
    if mk not in SKU_TRATADO:
        return 0
    col_sku, col_nome = SKU_TRATADO[mk]
    m = consolidado.MAPA.get(mk, {})
    valores = {}
    try:
        for _, linhas in _linhas_da_planilha(custos):
            ks = [str(c or "").strip() for c in (linhas[0] if linhas else [])]
            if "SKU" in ks and "Valor" in ks:
                js, jv = ks.index("SKU"), ks.index("Valor")
                for r in linhas[1:]:
                    v = _numero(r[jv] if jv < len(r) else None)[0]
                    if r and js < len(r) and r[js] not in (None, "") and v:
                        valores[_norm_sku(r[js])] = v
            break
    except Exception:
        valores = {}
    wb = load_workbook(caminho)
    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
    cab = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    js, jn = cab.get(col_sku), cab.get(col_nome)
    if not js or not jn:
        wb.close()
        return 0
    jst, jq, jf = cab.get(m.get("status")), cab.get(m.get("qtd")), cab.get(m.get("fat"))
    jc, jl = cab.get("Custo do produto"), cab.get("Liquido") or cab.get("Líquido")
    ji, ja = cab.get("Imposto"), cab.get("Antecipa")
    jm = cab.get("Gross margin") or cab.get("Gross margen")
    jp = cab.get("Gross margin %") or cab.get("Gross margen %")

    def num(v):
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0
    n = 0
    for r in range(2, ws.max_row + 1):
        atual = str(ws.cell(r, js).value or "").strip()
        if atual and atual.lower() not in ("nan", "none", "-", "--"):
            continue
        titulo = re.sub(r"\s+", " ", str(ws.cell(r, jn).value or "")).strip()[:150]
        if not titulo:
            continue
        ws.cell(r, js).value = titulo
        n += 1
        st = str(ws.cell(r, jst).value or "").strip().lower() if jst else "venda"
        valor = valores.get(_norm_sku(titulo))
        if not valor or not jc or st != "venda":
            continue
        custo = -valor * (num(ws.cell(r, jq).value) if jq else 1)
        ws.cell(r, jc).value = custo
        if jm:
            marg = sum(num(ws.cell(r, j).value) for j in (jl, ji, ja) if j) + custo
            ws.cell(r, jm).value = marg
            if jp:
                fat = num(ws.cell(r, jf).value) if jf else 0
                ws.cell(r, jp).value = marg / fat if fat else None
    if n:
        wb.save(caminho)
    wb.close()
    return n


def _chave_sku(t):
    t = re.sub(r"\s+", " ", str(t if t is not None else "")).strip().lower()
    return re.sub(r"^(\d+)\.0+$", r"\1", t)        # 54523.0 -> 54523


# ---------------------------------------------------------------- tradução de anúncio/SKU (preenchida no app)
def _arq_traducoes(grupo_id):
    return os.path.join(comum.PASTA_SISTEMA, "dados", "traducao_sku", re.sub(r"[^\w.-]+", "_", str(grupo_id)) + ".json")


def _chave_an(t):
    t = str(t if t is not None else "").strip()
    return re.sub(r"^(\d+)\.0+$", r"\1", t)


def chave_traducao(mk, anuncio, sku_original):
    """anuncio "*" = vale para todo anúncio com esse SKU."""
    return "%s|%s|%s" % (mk, _chave_an(anuncio) or "*", _chave_sku(sku_original))


def ler_traducoes(grupo_id):
    try:
        with open(_arq_traducoes(grupo_id), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def salvar_traducoes(grupo_id, itens, por=""):
    """itens: [{mk, anuncio, sku_original, sku, escopo ("anuncio"|"sku"), antes (chave da regra que valia na linha)}].
    sku vazio (ou igual ao original) tira a tradução. Devolve quantas regras mudaram."""
    regras = ler_traducoes(grupo_id)
    n = 0
    for it in itens or []:
        mk = str(it.get("mk") or "")
        orig = str(it.get("sku_original") if it.get("sku_original") is not None else "").strip()
        if not mk:
            continue
        an = "*" if it.get("escopo") == "sku" else (_chave_an(it.get("anuncio")) or "*")
        k = chave_traducao(mk, an, orig)
        antes = it.get("antes")
        if antes and antes != k and antes in regras:
            regras.pop(antes)
            n += 1
        novo = re.sub(r"\s+", " ", str(it.get("sku") or "")).strip()[:150]
        if not novo or _chave_sku(novo) == _chave_sku(orig):
            if regras.pop(k, None) is not None:
                n += 1
            continue
        if (regras.get(k) or {}).get("sku") != novo:
            regras[k] = {"mk": mk, "anuncio": an, "sku_original": orig, "sku": novo, "titulo": str(it.get("titulo") or "")[:150],
                         "quando": dt.datetime.now().isoformat(timespec="seconds"), "por": por or ""}
            n += 1
    os.makedirs(os.path.dirname(_arq_traducoes(grupo_id)), exist_ok=True)
    tmp = _arq_traducoes(grupo_id) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(regras, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _arq_traducoes(grupo_id))
    return n


def _valores_custos(custos):
    exato, sem_zero = {}, {}
    try:
        for _, linhas in _linhas_da_planilha(custos):
            ks = [str(c or "").strip() for c in (linhas[0] if linhas else [])]
            if "SKU" in ks and "Valor" in ks:
                js, jv = ks.index("SKU"), ks.index("Valor")
                for r in linhas[1:]:
                    v = _numero(r[jv] if jv < len(r) else None)[0]
                    k = _chave_sku(r[js] if js < len(r) else "")
                    if k and v:
                        exato.setdefault(k, v)
                        sem_zero.setdefault(k.lstrip("0") or "0", v)
            break
    except Exception:
        pass
    return exato, sem_zero


def traduzir_sku(caminho, mk, regras, custos):
    """Troca o SKU das vendas conforme a tradução do app (anúncio + SKU do relatório -> SKU correto; ou SKU -> SKU para
    todos os anúncios) e refaz custo e margem com o SKU novo. Devolve quantas linhas mudaram."""
    import consolidado
    from openpyxl import load_workbook
    if mk not in SKU_TRATADO or not regras:
        return 0
    m = consolidado.MAPA.get(mk, {})
    exato, sem_zero = _valores_custos(custos)
    wb = load_workbook(caminho)
    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
    cab = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    js, jan = cab.get(SKU_TRATADO[mk][0]), cab.get(m.get("anuncio"))
    if not js:
        wb.close()
        return 0
    jst, jq, jf = cab.get(m.get("status")), cab.get(m.get("qtd")), cab.get(m.get("fat"))
    jc, jl = cab.get("Custo do produto"), cab.get("Liquido") or cab.get("Líquido")
    ji, ja = cab.get("Imposto"), cab.get("Antecipa")
    jm = cab.get("Gross margin") or cab.get("Gross margen")
    jp = cab.get("Gross margin %") or cab.get("Gross margen %")

    def num(v):
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0
    n = 0
    for r in range(2, ws.max_row + 1):
        sku = ws.cell(r, js).value
        an = ws.cell(r, jan).value if jan else ""
        g = regras.get(chave_traducao(mk, an, sku)) or regras.get(chave_traducao(mk, "*", sku))
        if not g or _chave_sku(g["sku"]) == _chave_sku(sku):
            continue
        ws.cell(r, js).value = g["sku"]
        n += 1
        st = str(ws.cell(r, jst).value or "").strip().lower() if jst else "venda"
        if not jc or st != "venda":
            continue
        k = _chave_sku(g["sku"])
        valor = exato.get(k) or sem_zero.get(k.lstrip("0") or "0") or 0.0
        custo = -valor * (num(ws.cell(r, jq).value) if jq else 1)
        ws.cell(r, jc).value = custo
        if jm:
            marg = sum(num(ws.cell(r, j).value) for j in (jl, ji, ja) if j) + custo
            ws.cell(r, jm).value = marg
            if jp:
                fat = num(ws.cell(r, jf).value) if jf else 0
                ws.cell(r, jp).value = marg / fat if fat else None
    if n:
        wb.save(caminho)
    wb.close()
    return n


def completar_custo_faltante(caminho, mk, custos):
    """Venda sem custo cujo SKU está na tabela, mas em outro formato (ex.: "54523" texto no relatório x 54523 número na
    tabela): compara como texto (sem espaços nas pontas e sem ".0") e, se não achar, sem os zeros à esquerda.
    Preenche o custo (valor x quantidade) e refaz a margem. Devolve quantas linhas ganharam custo."""
    import consolidado
    from openpyxl import load_workbook
    if mk not in SKU_TRATADO:
        return 0
    exato, sem_zero = {}, {}
    try:
        for _, linhas in _linhas_da_planilha(custos):
            ks = [str(c or "").strip() for c in (linhas[0] if linhas else [])]
            if "SKU" in ks and "Valor" in ks:
                js, jv = ks.index("SKU"), ks.index("Valor")
                for r in linhas[1:]:
                    v = _numero(r[jv] if jv < len(r) else None)[0]
                    k = _chave_sku(r[js] if js < len(r) else "")
                    if k and v:
                        exato.setdefault(k, v)
                        sem_zero.setdefault(k.lstrip("0") or "0", v)
            break
    except Exception:
        return 0
    if not exato:
        return 0
    col_sku = SKU_TRATADO[mk][0]
    m = consolidado.MAPA.get(mk, {})
    wb = load_workbook(caminho)
    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
    cab = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    js, jc = cab.get(col_sku), cab.get("Custo do produto")
    if not js or not jc:
        wb.close()
        return 0
    jst, jq, jf = cab.get(m.get("status")), cab.get(m.get("qtd")), cab.get(m.get("fat"))
    jl, ji, ja = cab.get("Liquido") or cab.get("Líquido"), cab.get("Imposto"), cab.get("Antecipa")
    jm = cab.get("Gross margin") or cab.get("Gross margen")
    jp = cab.get("Gross margin %") or cab.get("Gross margen %")

    def num(v):
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0
    n = 0
    for r in range(2, ws.max_row + 1):
        if jst and str(ws.cell(r, jst).value or "").strip().lower() != "venda":
            continue
        if num(ws.cell(r, jc).value):
            continue                                  # já tem custo
        k = _chave_sku(ws.cell(r, js).value)
        if not k or k in ("nan", "none"):
            continue
        valor = exato.get(k) or sem_zero.get(k.lstrip("0") or "0")
        if not valor:
            continue
        custo = -valor * (num(ws.cell(r, jq).value) if jq else 1)
        ws.cell(r, jc).value = custo
        if jm:
            marg = sum(num(ws.cell(r, j).value) for j in (jl, ji, ja) if j) + custo
            ws.cell(r, jm).value = marg
            if jp:
                fat = num(ws.cell(r, jf).value) if jf else 0
                ws.cell(r, jp).value = marg / fat if fat else None
        n += 1
    if n:
        wb.save(caminho)
    wb.close()
    return n


def contar_sem_custo(caminho, mk):
    """Quantas VENDAS (não canceladas/devolvidas/não pagas) ficaram sem custo no tratado final."""
    import consolidado
    from openpyxl import load_workbook
    m = consolidado.MAPA.get(mk, {})
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        cab = {str(c).strip(): j for j, c in enumerate(next(it)) if c is not None}
        jc, jst = cab.get("Custo do produto"), cab.get(m.get("status"))
        if jc is None:
            return None
        n = 0
        for r in it:
            if jst is not None and str(r[jst] or "").strip().lower() != "venda":
                continue
            try:
                vazio = not float(r[jc] or 0)
            except (TypeError, ValueError):
                vazio = True
            n += vazio
        return n
    finally:
        wb.close()


def filtrar_mes(caminho, ano, mes):
    """Tratado de um mês passado: fica só com as linhas do mês (coluna Data). Devolve quantas saíram."""
    from openpyxl import load_workbook
    wb = load_workbook(caminho)
    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
    cab = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    jd = cab.get("Data")
    if not jd:
        wb.close()
        return 0

    def mes_de(v):
        if isinstance(v, (dt.datetime, dt.date)):
            return v.year, v.month
        t = str(v or "").strip()
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", t) or None
        if m:
            return int(m.group(1)), int(m.group(2))
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
        if m:
            return int(m.group(3)), int(m.group(2))
        return None
    tirar = [r for r in range(2, ws.max_row + 1) if (lambda x: x is not None and x != (ano, mes))(mes_de(ws.cell(r, jd).value))]
    for r in reversed(tirar):
        ws.delete_rows(r)
    if tirar:
        wb.save(caminho)
    wb.close()
    return len(tirar)


def tratar(cfg, cliente, mk, hoje=None, origem="automático"):
    """Trata um marketplace de um cliente no dia. Nunca levanta exceção: devolve e registra o resultado."""
    hoje = hoje or dt.date.today()
    dia = comum.rotulo_dia(hoje)
    pasta = comum.pasta_dia(cfg, comum.pasta_de(cliente), hoje)
    bruto = os.path.join(pasta, comum.nome_arquivo(mk, cfg, hoje))
    saida = caminho_saida(cfg, cliente, mk, hoje)
    aliquota, loja = imposto(cliente), cliente.get("loja") or cliente["nome"]
    ev = {"tipo": "tratamento", "cliente": cliente["id"], "mk": mk, "data": hoje.isoformat(), "origem": origem,
          "imposto": aliquota, "loja": loja}
    if cfg["marketplaces"].get(mk, {}).get("sem_tratador"):
        ev.update(ok=False, pendente=True, sem_tratador=True,
                  erro="Tratador de %s ainda não definido: o relatório fica guardado sem tratamento." % cfg["marketplaces"][mk]["nome"])
        return ev
    temp = tempfile.mkdtemp(prefix="tratar_")
    try:
        if not os.path.isfile(bruto):
            ev.update(ok=False, pendente=True, erro="O relatório bruto ainda não está na pasta do dia.")
            return ev
        custos, sem_custos = _custos_ou_vazio(cliente, temp, hoje)
        ev["custos_mes"] = escolher_custos(cliente, *mes_do_periodo(hoje))[1]
        os.makedirs(os.path.dirname(saida), exist_ok=True)
        parcial = os.path.join(temp, "saida.xlsx")
        if mk == "mercadolivre":
            diag = _mercado_livre(bruto, custos, parcial, aliquota, loja)
        elif mk == "shopee":
            from tratamento_shopee import processar_shopee
            antecipa = cliente.get("shopee_antecipa", comum.SHOPEE_ANTECIPA_PADRAO)
            ev["shopee_antecipa"] = antecipa
            _, diag = processar_shopee(bruto, custos, caminho_ids_shopee(cliente), parcial,
                                       aliquota_imposto=aliquota, antecipacao=float(antecipa or 0), marketplace_venda=loja)
        elif mk == "magalu":
            from tratamento_magalu import processar_magalu
            import compactados
            if not compactados.ja_plano(bruto):   # .rar, zip dentro de zip, CSVs em subpasta...
                plano = os.path.join(temp, "magalu.zip")
                compactados.para_zip_plano(bruto, plano)
                bruto_mg = plano
            else:
                bruto_mg = bruto
            _, diag = processar_magalu(bruto_mg, custos, parcial, caminho_ids=caminho_ids_magalu(),
                                       aliquota_imposto=aliquota, marketplace_venda=loja)
        elif mk == "tiktokshop":
            from tratamento_tiktok import processar_tiktok
            _, diag = processar_tiktok(bruto, custos, parcial, aliquota_imposto=aliquota, marketplace_venda=loja)
        elif mk == "viavarejo":
            from tratamento_casas_bahia import processar_casas_bahia
            _, diag = processar_casas_bahia(bruto, custos, parcial, aliquota_imposto=aliquota, marketplace_venda=loja)
        elif mk == "temu":
            from tratamento_temu import processar_temu
            _, diag = processar_temu(bruto, custos, parcial, aliquota_imposto=aliquota, marketplace_venda=loja)
        elif mk == "madeiramadeira":
            from tratamento_madeira_madeira import processar_madeira_madeira
            _, diag = processar_madeira_madeira(bruto, custos, parcial, aliquota_imposto=aliquota, marketplace_venda=loja)
        elif mk == "amazon":
            from tratamento_amazon import processar_amazon
            unif = os.path.join(pasta, comum.nome_arquivo("amazon-unificado", cfg, hoje))
            receb = os.path.join(pasta, comum.nome_arquivo("amazon-receber", cfg, hoje))
            txt = _amazon_txt(pasta, dia, bruto, temp)
            faltam = [n for n, p in (("relatório unificado", unif), ("transações a receber", receb)) if not os.path.isfile(p)]
            if not txt:
                faltam.insert(0, "TXT original de pedidos")
            if faltam:
                ev.update(ok=False, pendente=True, erro="Aguardando: " + ", ".join(faltam) + ".")
                return ev
            _, diag = processar_amazon(txt, unif, receb, custos, parcial, aliquota_imposto=aliquota, marketplace_venda=loja)
        else:
            ev.update(ok=False, erro="Marketplace sem tratador.")
            return ev
        diag["canceladas_ajustadas"] = zerar_canceladas(parcial, mk)   # não pago/cancelada/devolução: só a cobrança
        try:
            diag["sku_pelo_titulo"] = sku_por_titulo(parcial, mk, custos)   # sem SKU na plataforma: o título vira o SKU
        except Exception as e:
            comum.log("[%s] SKU pelo título não aplicado: %s" % (cliente["nome"], e))
        try:   # tradução de anúncio/SKU feita no app (o cliente cadastrou o SKU errado na plataforma)
            regras = {k: v for k, v in ler_traducoes(cliente.get("cliente_id") or cliente["id"]).items() if k.startswith(mk + "|")}
            if regras:
                diag["sku_traduzido"] = traduzir_sku(parcial, mk, regras, custos)
        except Exception as e:
            comum.log("[%s] tradução de SKU não aplicada: %s" % (cliente["nome"], e))
        try:
            diag["custo_formato_sku"] = completar_custo_faltante(parcial, mk, custos)   # "54523" texto x 54523 número
        except Exception as e:
            comum.log("[%s] custo pelo formato do SKU não aplicado: %s" % (cliente["nome"], e))
        a_ref, m_ref = mes_do_periodo(hoje)
        if (a_ref, m_ref) != mes_do_periodo():          # mês passado (escolhido no topo): só as vendas daquele mês
            diag["fora_do_mes"] = filtrar_mes(parcial, a_ref, m_ref)
            if diag["fora_do_mes"]:
                diag["linhas"] = max(0, int(diag.get("linhas") or 0) - diag["fora_do_mes"])
        try:   # "sem custo" = vendas do tratado final sem custo (antes contava também canceladas e o que foi completado depois)
            sc = contar_sem_custo(parcial, mk)
            if sc is not None:
                diag["semCusto"] = sc
        except Exception:
            pass
        try:
            shutil.move(parcial, saida + ".novo")
            os.replace(saida + ".novo", saida)
        except PermissionError:
            ev.update(ok=False, erro="Não consegui gravar o tratado: feche o arquivo no Excel e clique em Tratar de novo.")
            return ev
        try:
            novos = completar_custos(cliente, mk, saida, hoje)
        except Exception as e:   # nunca derruba o tratamento
            novos = []
            comum.log("[%s] não consegui completar a tabela de custos: %s" % (cliente["nome"], e))
        if novos:
            diag["skus_novos_custos"] = len(novos)
            diag["skus_novos_ex"] = novos[:8]
        if sem_custos:
            diag["aviso"] = ("Cliente sem tabela de custos: custo e margem ficaram vazios. " +
                             ("A tabela do mês foi criada com os %d SKU(s) vendidos: preencha os custos em Custos e imposto." % len(novos)
                              if novos else "Envie em Custos e imposto."))
        ev.update(ok=True, saida=saida, diag=diag)
        return ev
    except Exception as e:  # inclui ErroDeEntrada (ex.: Mercado Livre com linhas em vermelho)
        ev.update(ok=False, erro=str(e)[:600])
        return ev
    finally:
        shutil.rmtree(temp, ignore_errors=True)
        if not ev.get("ok") and not ev.get("pendente") and os.path.isfile(saida):
            # o tratado antigo não corresponde mais ao bruto atual: guarda (não apaga) para não confundir
            try:
                ant = os.path.join(os.path.dirname(saida), "_anteriores",
                                   os.path.basename(saida).replace(".xlsx", "_%s.xlsx" % dt.datetime.now().strftime("%H%M%S")))
                os.makedirs(os.path.dirname(ant), exist_ok=True)
                shutil.move(saida, ant)
                ev["anterior_guardado"] = ant
            except OSError:
                pass
        comum.registrar_evento(ev)
        comum.log("[%s] TRATAMENTO %s %s%s" % (cliente["nome"], cfg["marketplaces"].get(mk, {}).get("nome", mk),
                                             "ok" if ev.get("ok") else ("pendente" if ev.get("pendente") else "ERRO"),
                                             "" if ev.get("ok") else " — " + str(ev.get("erro", ""))))


def tratar_cliente(cfg, cliente, hoje=None, mks=None, origem="manual"):
    hoje = hoje or dt.date.today()
    mks = mks or [m for m in cliente.get("marketplaces", []) if m in cfg["marketplaces"]]
    return [tratar(cfg, cliente, m, hoje, origem) for m in mks]


def pai(cfg, mk):
    """Relatórios auxiliares (ex.: amazon-unificado) disparam o tratamento do marketplace principal."""
    return cfg["marketplaces"].get(mk, {}).get("pai") or mk
