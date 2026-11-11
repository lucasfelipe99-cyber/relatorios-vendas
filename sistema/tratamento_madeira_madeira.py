# -*- coding: utf-8 -*-
"""Tratamento de vendas — MadeiraMadeira (marketplace "madeiramadeira" do app).

Entrada: .csv "report_pedido_<hash>.csv" do Portal do Seller MadeiraMadeira
(separador ";", codificação Windows-1252, uma linha por item do pedido). O organizador do app
converte o .csv em .xlsx (vendas_madeiramadeira_DD-MM.xlsx); os dois formatos são aceitos.

Saída: aba DB com as MESMAS colunas do tratado do Mercado Livre (tratamento_mercado_livre.COLUNAS_SAIDA).

Regras (proposta de 26/10/2026, para teste):
- Uma linha por item. N.º de venda = Id mercado Pago = "Pedido" (nº do pedido no seller).
- Data da venda = "Data Pedido" (com hora); Data = dia da venda.
- # de anúncio e SKU = "SKU" (o relatório não traz id de anúncio). Título = "Produto".
- Unidades = "Quantidade". Faturamento = "Valor" (total do item, já com o desconto de/por da loja);
  Preço unitário = Valor / Quantidade. Desconto e rebate = 0.
- Comissão = −"Comissão" do relatório. A comissão vem por pedido (repetida em cada item):
  é rateada entre os itens do pedido pelo Valor.
- Frete = 0. O frete pago pelo cliente (Valor Pedido − soma dos itens) não entra.
- Status ("Status"): Cancelado → Cancelada; Devolvido/Estornado → Devolução;
  Aguardando pagamento/Pendente → Não pago; o resto (Novo, Aprovado, Faturado, Enviado, Entregue) → Venda.
  Fora de venda: Cancelamento = −Faturamento, comissão 0, imposto e custo 0 → margem 0.
- Imposto = −(Faturamento + Desconto) × alíquota do cliente; Custo do produto = −custo × Unidades.
- Gross margen = Liquido + Imposto + Custo do produto; Gross margen % = margem / Faturamento.
- Forma de entrega = "Nome da Transportadora"; Estado = status original.
"""
import unicodedata

import pandas as pd

from tratamento_mercado_livre import COLUNAS_SAIDA, _carregar_custos

OBRIGATORIAS = ("Pedido", "Data Pedido", "Status", "SKU", "Produto", "Quantidade", "Valor", "Comissão", "Valor Pedido")


def _norm(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().strip().lower()


def _sku(v):
    s = str(v if v is not None else "").strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return "" if s.lower() in ("nan", "none") else s


def status_padrao(st):
    s = _norm(st)
    if "devol" in s or "estorn" in s or "reembols" in s:
        return "Devolução"
    if "cancel" in s:
        return "Cancelada"
    if "aguardando pagamento" in s or "pendente" in s or "nao pago" in s:
        return "Não pago"
    return "Venda"


def ler_bruto(caminho):
    """Aceita o .csv original ou o .xlsx em que o organizador do app converte o .csv (tudo como texto)."""
    df = None
    if str(caminho).lower().endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(caminho, dtype=str)
    else:
        for enc in ("utf-8-sig", "cp1252"):
            try:
                df = pd.read_csv(caminho, sep=";", encoding=enc, dtype=str)
                break
            except UnicodeDecodeError:
                continue
    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in OBRIGATORIAS if c not in df.columns]
    if faltam:
        raise ValueError("Relatório da MadeiraMadeira sem as colunas: %s "
                         "(use o relatório de pedidos em .csv do Portal do Seller)." % ", ".join(faltam))
    return df[df["Pedido"].fillna("").str.strip() != ""].reset_index(drop=True)


def _num(serie):
    # valores em R$ vêm com ponto decimal (447.03); aceita também vírgula
    s = serie.fillna("").astype(str).str.strip()
    s = s.where(~s.str.contains(","), s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def transformar(b, custos, aliquota, loja):
    d = pd.DataFrame(index=b.index)
    d["Marketplace"] = "MadeiraMadeira"
    d["Marketplace venda"] = loja
    d["N.º de venda"] = b["Pedido"].str.strip()
    d["Data da venda"] = pd.to_datetime(b["Data Pedido"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    d["Data"] = d["Data da venda"].dt.normalize()
    d["Estado"] = b["Status"].fillna("").str.strip()
    d["Type.1"] = d["Estado"].map(status_padrao)
    d["Forma de entrega"] = b["Nome da Transportadora"].fillna("") if "Nome da Transportadora" in b.columns else ""
    d["# de anúncio"] = b["SKU"].map(_sku)
    d["Título do anúncio"] = b["Produto"]
    d["SKU"] = b["SKU"].map(_sku)
    unid = pd.to_numeric(b["Quantidade"], errors="coerce").fillna(1).astype(int)
    fat = _num(b["Valor"])
    d["Preço unitário de venda do anúncio (BRL)"] = (fat / unid.where(unid != 0, 1)).round(2)
    d["Unidades"] = unid
    d["Faturamento"] = fat
    d["Desconto"] = 0.0
    d["rebate"] = 0.0
    # comissão do pedido (repetida em cada item) rateada pelo valor dos itens
    com_ped = _num(b["Comissão"]).groupby(d["N.º de venda"]).transform("first")
    fat_ped = fat.groupby(d["N.º de venda"]).transform("sum")
    itens = fat.groupby(d["N.º de venda"]).transform("size")
    parte = (fat / fat_ped.where(fat_ped != 0)).fillna(1 / itens)
    d["Comissão"] = -(com_ped * parte)
    d["Frete"] = 0.0
    d["Cancelamento"] = 0.0
    fora = d["Type.1"] != "Venda"
    d.loc[fora, "Cancelamento"] = -d.loc[fora, "Faturamento"]
    d.loc[fora, "Comissão"] = 0.0
    d["Liquido"] = d[["Faturamento", "Comissão", "Frete", "rebate", "Desconto", "Cancelamento"]].sum(axis=1)
    d["Antecipa"] = 0
    d["Imposto"] = (-d["Faturamento"] - d["Desconto"]) * float(aliquota)
    custo = {_sku(k): v for k, v in zip(custos["SKU"], custos["Valor total"])}
    d["Custo do produto"] = -d["SKU"].map(custo) * d["Unidades"]
    d.loc[fora, ["Imposto", "Custo do produto"]] = 0.0
    d["Gross margen"] = d["Liquido"] + d["Imposto"] + d["Custo do produto"].fillna(0)
    d["Gross margen %"] = d["Gross margen"] / d["Faturamento"].where(d["Faturamento"] != 0)
    d["Id mercado Pago"] = d["N.º de venda"]
    for c in ("Comissão", "Frete", "Liquido", "Imposto", "Custo do produto", "Gross margen"):
        d[c] = d[c].round(2) + 0.0
    return d[COLUNAS_SAIDA].reset_index(drop=True)


def processar_madeira_madeira(caminho_vendas, caminho_custos, caminho_saida, aliquota_imposto=0.14,
                              marketplace_venda="MadeiraMadeira"):
    """Gera o tratado (aba DB, formato do Mercado Livre). Devolve (caminho, diagnóstico)."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    df = transformar(ler_bruto(caminho_vendas), _carregar_custos(caminho_custos), aliquota_imposto, marketplace_venda)
    df.to_excel(caminho_saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(caminho_saida)
    ws = wb["DB"]
    for c in ws[1]:
        c.font = Font(bold=True)
    cab = [c.value for c in ws[1]]
    for j, nome in enumerate(cab, 1):
        fmt = ("dd/mm/yyyy hh:mm" if nome == "Data da venda" else "yyyy-mm-dd" if nome == "Data" else
               "0.00%" if nome == "Gross margen %" else
               "#,##0.00" if j >= cab.index("Preço unitário de venda do anúncio (BRL)") + 1
               and nome not in ("Unidades", "Id mercado Pago", "Antecipa") else None)
        if fmt:
            for (cel,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                cel.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho_saida)
    return caminho_saida, {"linhas": int(len(df)), "semCusto": int(df["Custo do produto"].isna().sum()),
                           "status": df["Type.1"].value_counts().to_dict()}
