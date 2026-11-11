# -*- coding: utf-8 -*-
"""Tratamento de vendas — Casas Bahia / Via Varejo (marketplace "viavarejo" do app).

Entrada: .xlsx "Pedidos_CasasBahia_Marketplace_..." (aba "Via - Relatório de pedidos";
cabeçalho na linha que contém "Número do pedido").

Saída: aba DB com as MESMAS colunas do tratado do Mercado Livre (tratamento_mercado_livre.COLUNAS_SAIDA).

Regras (combinadas em 26/10/2026):
- Faturamento = "Valor do produto" (sem o frete pago pelo cliente).
- Comissão = -17% do Faturamento (sem o frete).
- Frete = 0 (a Casas Bahia não cobra frete do vendedor; o frete pago pelo cliente não entra).
- Imposto = -(Faturamento + Desconto) x alíquota do cliente; Custo do produto = -custo x Unidades.
- Gross margen = Liquido + Imposto + Custo do produto; Gross margen % = margem / Faturamento.
- Status (Type.1): Venda / Cancelada / Devolução / Não pago.
  Cancelada, Devolução e Não pago: Cancelamento = -Faturamento e comissão 0 (margem 0).
  Imposto e custo das canceladas/devoluções são zerados depois pelo tratador.zerar_canceladas.
- O relatório não traz quantidade: Unidades = 1 por linha.
"""
import unicodedata

import pandas as pd

from tratamento_mercado_livre import COLUNAS_SAIDA, _carregar_custos

COMISSAO = 0.17


def _norm(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().strip().lower()


def _sku(v):
    s = str(v if v is not None else "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return "" if s.lower() in ("nan", "none") else s


def status_padrao(st):
    s = _norm(st)
    if "devol" in s or "reembols" in s:
        return "Devolução"
    if "cancel" in s or "recusad" in s:
        return "Cancelada"
    if "aguardando pagamento" in s or "pagamento pendente" in s or "nao pago" in s:
        return "Não pago"
    return "Venda"   # Novo, Aprovado, Faturado, Enviado, Entregue...


def ler_bruto(caminho):
    xl = pd.ExcelFile(caminho)
    aba = next((a for a in xl.sheet_names if "pedido" in a.lower()), xl.sheet_names[0])
    cru = pd.read_excel(caminho, sheet_name=aba, header=None, dtype=str)
    linha = next((i for i in range(min(30, len(cru)))
                  if (cru.iloc[i].astype(str).str.strip() == "Número do pedido").any()), None)
    if linha is None:
        raise ValueError("Não achei o cabeçalho “Número do pedido” no relatório da Casas Bahia "
                         "(use o Relatório de pedidos em .xlsx).")
    df = pd.read_excel(caminho, sheet_name=aba, header=linha, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in ("Número do pedido", "Status do pedido", "Data de aprovação", "Nome do produto",
                          "SKU da loja", "SKU no site da venda", "Valor do produto") if c not in df.columns]
    if faltam:
        raise ValueError("Relatório da Casas Bahia sem as colunas: %s." % ", ".join(faltam))
    return df[df["Número do pedido"].fillna("").str.strip() != ""].reset_index(drop=True)


def transformar(b, custos, aliquota, loja, comissao=COMISSAO):
    num = lambda c: pd.to_numeric(b[c], errors="coerce").fillna(0.0) if c in b.columns else pd.Series(0.0, index=b.index)
    d = pd.DataFrame(index=b.index)
    d["Marketplace"] = "Casas Bahia"
    d["Marketplace venda"] = loja
    d["N.º de venda"] = b["Número do pedido"].str.strip()
    d["Data da venda"] = pd.to_datetime(b["Data de aprovação"], errors="coerce")
    sem_aprov = d["Data da venda"].isna()
    if sem_aprov.any() and "Data da compra" in b.columns:
        d.loc[sem_aprov, "Data da venda"] = pd.to_datetime(b.loc[sem_aprov, "Data da compra"], errors="coerce")
    d["Data"] = d["Data da venda"].dt.normalize()
    d["Estado"] = b["Status do pedido"].fillna("").str.strip()
    d["Type.1"] = d["Estado"].map(status_padrao)
    d["Forma de entrega"] = b["Transportadora"].fillna("") if "Transportadora" in b.columns else ""
    d["# de anúncio"] = b["SKU no site da venda"].map(_sku)
    d["Título do anúncio"] = b["Nome do produto"]
    d["SKU"] = b["SKU da loja"].map(_sku)
    d["Preço unitário de venda do anúncio (BRL)"] = num("Valor do produto")
    d["Unidades"] = 1
    d["Faturamento"] = num("Valor do produto")
    d["Desconto"] = 0.0
    d["rebate"] = 0.0
    d["Comissão"] = -(d["Faturamento"] * comissao)
    d["Frete"] = 0.0
    d["Cancelamento"] = 0.0
    canc = d["Type.1"].isin(["Cancelada", "Não pago"])
    devol = d["Type.1"] == "Devolução"
    d.loc[canc | devol, "Cancelamento"] = -d.loc[canc | devol, "Faturamento"]
    d.loc[canc | devol, "Comissão"] = 0.0
    d["Liquido"] = d[["Faturamento", "Comissão", "Frete", "rebate", "Desconto", "Cancelamento"]].sum(axis=1)
    d["Antecipa"] = 0
    d["Imposto"] = (-d["Faturamento"] - d["Desconto"]) * float(aliquota)
    custo = {_sku(k): v for k, v in zip(custos["SKU"], custos["Valor total"])}
    d["Custo do produto"] = -d["SKU"].map(custo) * d["Unidades"]
    d.loc[d["Type.1"] == "Não pago", ["Imposto", "Custo do produto"]] = 0.0
    d["Gross margen"] = d["Liquido"] + d["Imposto"] + d["Custo do produto"].fillna(0)
    d["Gross margen %"] = d["Gross margen"] / d["Faturamento"].where(d["Faturamento"] != 0)
    d["Id mercado Pago"] = d["N.º de venda"]
    for c in ("Comissão", "Frete", "Liquido", "Imposto", "Custo do produto", "Gross margen"):
        d[c] = d[c].round(2) + 0.0
    return d[COLUNAS_SAIDA].reset_index(drop=True)


def processar_casas_bahia(caminho_vendas, caminho_custos, caminho_saida, aliquota_imposto=0.14,
                          marketplace_venda="Casas Bahia"):
    """Gera o tratado (aba DB, formato do Mercado Livre). Devolve (caminho, diagnóstico)."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    df = transformar(ler_bruto(caminho_vendas), _carregar_custos(caminho_custos), aliquota_imposto,
                     marketplace_venda)
    df.to_excel(caminho_saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(caminho_saida)
    ws = wb["DB"]
    for c in ws[1]:
        c.font = Font(bold=True)
    cab = [c.value for c in ws[1]]
    for j, nome in enumerate(cab, 1):
        fmt = ("dd/mm/yyyy hh:mm" if nome == "Data da venda" else "yyyy-mm-dd" if nome == "Data" else
               "0.00%" if nome == "Gross margen %" else
               "#,##0.00" if j >= cab.index("Preço unitário de venda do anúncio (BRL)") + 1 and nome not in ("Unidades", "Id mercado Pago") else None)
        if fmt:
            for (cel,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cel.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho_saida)
    return caminho_saida, {"linhas": int(len(df)), "semCusto": int(df["Custo do produto"].isna().sum()),
                           "status": df["Type.1"].value_counts().to_dict()}
