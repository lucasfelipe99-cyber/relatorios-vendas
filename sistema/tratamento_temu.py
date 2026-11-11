# -*- coding: utf-8 -*-
"""Tratamento de vendas — Temu (marketplace "temu" do app).

Entrada: .xlsx "order_export-....xlsx" exportado da Central do Vendedor Temu
(aba "Order report"; cabeçalho na linha que contém "ID do pedido", hoje a linha 6).

Saída: aba DB com as MESMAS colunas do tratado do Mercado Livre (tratamento_mercado_livre.COLUNAS_SAIDA).

Regras (proposta de 26/10/2026):
- A Temu não cobra taxas: o "Preço-base dos produtos (impostos incluídos)" já é o valor líquido do vendedor.
  Faturamento = Preço-base x quantidade comprada; Comissão = 0; Frete = 0; Liquido = Faturamento.
- Preço total de varejo, Desconto da Temu, frete pago pelo cliente e "Cliente pagou" não entram
  (são valores do cliente/Temu, não do vendedor).
- Data = "data da compra" (texto "25 de Out. de 2026, 11:46 BRT(UTC-3)").
- # de anúncio = "ID do SKU" (id da Temu); SKU = "sku de contribuição" (SKU do vendedor, usado no custo).
- Status (Type.1) pelo "status do item do pedido": Cancelado -> Cancelada; Devolvido/Reembolsado -> Devolução;
  Não pago/Aguardando pagamento -> Não pago; o resto (Enviado, Entregues, A enviar...) -> Venda.
  Cancelada/Devolução/Não pago: Cancelamento = -Faturamento -> Liquido 0; imposto e custo 0 -> margem 0.
- Imposto = -(Faturamento + Desconto) x alíquota do cliente; Custo do produto = -custo x Unidades.
- Gross margen = Liquido + Imposto + Custo do produto; Gross margen % = margem / Faturamento.
"""
import re
import unicodedata

import pandas as pd

try:
    from tratamento_mercado_livre import COLUNAS_SAIDA, _carregar_custos
except ImportError:  # fora do app (prévia): mesma ordem do tratado do Mercado Livre
    COLUNAS_SAIDA = ["Marketplace", "Marketplace venda", "N.º de venda", "Data da venda", "Data", "Estado",
                     "Type.1", "Forma de entrega", "# de anúncio", "Título do anúncio", "SKU",
                     "Preço unitário de venda do anúncio (BRL)", "Unidades", "Faturamento", "Desconto",
                     "rebate", "Comissão", "Frete", "Cancelamento", "Liquido", "Antecipa", "Imposto",
                     "Custo do produto", "Gross margen", "Gross margen %", "Id mercado Pago"]
    _carregar_custos = None

MESES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
         "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}


def _norm(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().strip().lower()


def _sku(v):
    s = str(v if v is not None else "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return "" if s.lower() in ("nan", "none") else s


def _reais(v):
    """'R$1.234,56' / '-R$54,26' / número -> float."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    neg = s.startswith("-")
    s = re.sub(r"[^\d,.]", "", s).replace(".", "").replace(",", ".")
    try:
        x = float(s) if s else 0.0
    except ValueError:
        return 0.0
    return -x if neg else x


def _data(v):
    """'25 de Out. de 2026, 11:46 BRT(UTC-3)' -> Timestamp (hora de Brasília)."""
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zçÇ]{3})\w*\.?\s+de\s+(\d{4}),?\s*(\d{1,2}):(\d{2})", str(v or ""))
    if not m:
        return pd.to_datetime(v, errors="coerce", dayfirst=True)
    mes = MESES.get(_norm(m.group(2))[:3])
    if not mes:
        return pd.NaT
    return pd.Timestamp(int(m.group(3)), mes, int(m.group(1)), int(m.group(4)), int(m.group(5)))


def status_padrao(st):
    s = _norm(st)
    if "devol" in s or "reembols" in s:
        return "Devolução"
    if "cancel" in s:
        return "Cancelada"
    if "nao pago" in s or "aguardando pagamento" in s:
        return "Não pago"
    return "Venda"


def ler_bruto(caminho):
    xl = pd.ExcelFile(caminho)
    aba = next((a for a in xl.sheet_names if "order" in a.lower() or "pedido" in a.lower()), xl.sheet_names[0])
    cru = pd.read_excel(caminho, sheet_name=aba, header=None, dtype=str)
    linha = next((i for i in range(min(30, len(cru)))
                  if (cru.iloc[i].astype(str).str.strip() == "ID do pedido").any()), None)
    if linha is None:
        raise ValueError("Não achei o cabeçalho “ID do pedido” no relatório da Temu "
                         "(use a exportação de pedidos em .xlsx).")
    df = pd.read_excel(caminho, sheet_name=aba, header=linha, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    obrig = ("ID do pedido", "status do item do pedido", "data da compra", "nome do produto",
             "sku de contribuição", "ID do SKU", "quantidade comprada",
             "Preço-base dos produtos (impostos incluídos)")
    faltam = [c for c in obrig if c not in df.columns]
    if faltam:
        raise ValueError("Relatório da Temu sem as colunas: %s." % ", ".join(faltam))
    return df[df["ID do pedido"].fillna("").str.strip() != ""].reset_index(drop=True)


def transformar(b, custos, aliquota, loja):
    d = pd.DataFrame(index=b.index)
    d["Marketplace"] = "Temu"
    d["Marketplace venda"] = loja
    d["N.º de venda"] = b["ID do pedido"].str.strip()
    d["Data da venda"] = b["data da compra"].map(_data)
    d["Data"] = pd.to_datetime(d["Data da venda"]).dt.normalize()
    d["Estado"] = b["status do item do pedido"].fillna(b.get("status do pedido", "")).fillna("").str.strip()
    d["Type.1"] = d["Estado"].map(status_padrao)
    d["Forma de entrega"] = b["transportadora"].fillna("") if "transportadora" in b.columns else ""
    d["# de anúncio"] = b["ID do SKU"].map(_sku)
    d["Título do anúncio"] = b["nome do produto"]
    d["SKU"] = b["sku de contribuição"].map(_sku)
    preco = b["Preço-base dos produtos (impostos incluídos)"].map(_reais)
    d["Preço unitário de venda do anúncio (BRL)"] = preco
    d["Unidades"] = pd.to_numeric(b["quantidade comprada"], errors="coerce").fillna(1).astype(int)
    d["Faturamento"] = preco * d["Unidades"]
    d["Desconto"] = 0.0
    d["rebate"] = 0.0
    d["Comissão"] = 0.0          # a Temu não cobra comissão
    d["Frete"] = 0.0             # nem frete
    d["Cancelamento"] = 0.0
    fora = d["Type.1"].isin(["Cancelada", "Devolução", "Não pago"])
    d.loc[fora, "Cancelamento"] = -d.loc[fora, "Faturamento"]
    d["Liquido"] = d[["Faturamento", "Comissão", "Frete", "rebate", "Desconto", "Cancelamento"]].sum(axis=1)
    d["Antecipa"] = 0
    d["Imposto"] = (-d["Faturamento"] - d["Desconto"]) * float(aliquota)
    if custos is not None and len(custos):
        custo = {_sku(k): v for k, v in zip(custos["SKU"], custos["Valor total"])}
        d["Custo do produto"] = -d["SKU"].map(custo) * d["Unidades"]
    else:
        d["Custo do produto"] = float("nan")
    d.loc[fora, ["Imposto", "Custo do produto"]] = 0.0
    d["Gross margen"] = d["Liquido"] + d["Imposto"] + d["Custo do produto"].fillna(0)
    d["Gross margen %"] = d["Gross margen"] / d["Faturamento"].where(d["Faturamento"] != 0)
    d["Id mercado Pago"] = d["N.º de venda"]
    for c in ("Faturamento", "Cancelamento", "Liquido", "Imposto", "Custo do produto", "Gross margen"):
        d[c] = d[c].round(2) + 0.0
    return d[COLUNAS_SAIDA].sort_values("Data da venda").reset_index(drop=True)


def gravar(df, caminho_saida):
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    df.to_excel(caminho_saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(caminho_saida)
    ws = wb["DB"]
    for c in ws[1]:
        c.font = Font(bold=True)
    cab = [c.value for c in ws[1]]
    ini = cab.index("Preço unitário de venda do anúncio (BRL)") + 1
    for j, nome in enumerate(cab, 1):
        fmt = ("dd/mm/yyyy hh:mm" if nome == "Data da venda" else "yyyy-mm-dd" if nome == "Data" else
               "0.00%" if nome == "Gross margen %" else
               "#,##0.00" if j >= ini and nome not in ("Unidades", "Id mercado Pago") else None)
        if fmt:
            for (cel,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cel.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho_saida)


def processar_temu(caminho_vendas, caminho_custos, caminho_saida, aliquota_imposto=0.14,
                   marketplace_venda="Temu"):
    """Gera o tratado (aba DB, formato do Mercado Livre). Devolve (caminho, diagnóstico)."""
    custos = _carregar_custos(caminho_custos) if (_carregar_custos and caminho_custos) else None
    df = transformar(ler_bruto(caminho_vendas), custos, aliquota_imposto, marketplace_venda)
    gravar(df, caminho_saida)
    return caminho_saida, {"linhas": int(len(df)), "semCusto": int(df["Custo do produto"].isna().sum()),
                           "status": df["Type.1"].value_counts().to_dict()}
