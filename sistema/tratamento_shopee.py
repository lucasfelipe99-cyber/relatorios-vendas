"""Tratamento do relatório bruto de pedidos da Shopee."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from tratamento_mercado_livre import ErroDeEntrada, _carregar_custos

COLUNAS_SAIDA = [
    "Marketplace", "Marketplace venda", "ID do pedido", "Data Completa", "Data",
    "Status do pedido", "Status", "Opção de envio", "ID do Produto", "Nome do Produto",
    "Número de referência SKU", "Preço acordado", "Quantidade", "Faturamento",
    "Desconto", "Rebate", "Comissão", "Frete", "Taxa de Envio Reversa", "Cancelamento", "Líquido",
    "Antecipa", "Imposto", "Custo do produto", "Gross margin", "Gross margin %",
]
NUMERICAS = [
    "Preço acordado", "Subtotal do produto", "Incentivo Shopee para ação comercial",
    "Ajuste por participação em ação comercial", "Cupom do vendedor",
    "Coin Cashback Voucher Amount Sponsored by Seller",
    "Desconto Shopee da Leve Mais por Menos",
    "Desconto da Leve Mais por Menos do vendedor", "Taxa de Envio Reversa",
    "Taxa de comissão bruta", "Taxa de serviço bruta", "Desconto do vendedor",
    "Desconto do vendedor_2",
]
OBRIGATORIAS = [
    "ID do pedido", "Status do pedido", "Status da Devolução / Reembolso",
    "Hora do pagamento do pedido", "Data de criação do pedido", "Opção de envio",
    "Nome do Produto", "Número de referência SKU", "Preço acordado", "Quantidade",
    "Subtotal do produto", "Desconto do vendedor", "Desconto do vendedor_2",
    "Incentivo Shopee para ação comercial", "Ajuste por participação em ação comercial",
    "Cupom do vendedor", "Coin Cashback Voucher Amount Sponsored by Seller",
    "Desconto Shopee da Leve Mais por Menos",
    "Desconto da Leve Mais por Menos do vendedor", "Taxa de Envio Reversa",
    "Taxa de comissão bruta", "Taxa de serviço bruta",
]


def _texto(valor: Any) -> str:
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))          # SKU numérico lido como número: 696.0 -> "696"
    return str(valor).strip()


def _numero(serie: pd.Series, nome: str) -> pd.Series:
    limpa = serie.map(lambda v: pd.NA if _texto(v) in {"", "-"} else v)
    limpa = limpa.map(lambda v: v.replace(",", "") if isinstance(v, str) else v)
    resultado = pd.to_numeric(limpa, errors="coerce")
    invalidos = limpa.notna() & resultado.isna()
    if invalidos.any():
        raise ErroDeEntrada(
            f"A coluna '{nome}' contém números inválidos: "
            f"{limpa[invalidos].astype(str).head(3).tolist()}."
        )
    return resultado.fillna(0).astype(float)


def _ler_pedidos(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if caminho.suffix.lower() == ".csv":
        df = pd.read_csv(caminho, sep=None, engine="python")
    else:
        try:
            df = pd.read_excel(caminho, sheet_name="orders")
        except ValueError as exc:
            raise ErroDeEntrada("O arquivo da Shopee precisa conter a aba 'orders'.") from exc
    if "Desconto do vendedor.1" in df.columns:
        df = df.rename(columns={"Desconto do vendedor.1": "Desconto do vendedor_2"})
    ausentes = [c for c in OBRIGATORIAS if c not in df.columns]
    if ausentes:
        raise ErroDeEntrada("Colunas ausentes na base Shopee: " + ", ".join(ausentes))
    return df


def _ler_ids(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    if caminho.suffix.lower() == ".csv":
        ids = pd.read_csv(caminho, sep=None, engine="python")
    else:
        ids = None
        for aba in pd.ExcelFile(caminho).sheet_names:
            candidato = pd.read_excel(caminho, sheet_name=aba)
            if {"sku", "Nome do Produto", "ID do Produto"}.issubset(candidato.columns):
                ids = candidato
                break
        if ids is None:
            raise ErroDeEntrada("A tabela de IDs precisa ter sku, Nome do Produto e ID do Produto.")
    faltam = {"sku", "Nome do Produto", "ID do Produto"} - set(ids.columns)
    if faltam:
        raise ErroDeEntrada("Colunas ausentes na tabela de IDs: " + ", ".join(sorted(faltam)))
    ids = ids[["sku", "Nome do Produto", "ID do Produto"]].copy()
    ids["sku"] = ids["sku"].map(_texto)
    ids["Nome do Produto"] = ids["Nome do Produto"].map(_texto)
    return ids.drop_duplicates(["sku", "Nome do Produto"], keep="first")


def transformar_shopee(pedidos: pd.DataFrame, custos: pd.DataFrame, ids: pd.DataFrame | None,
                       aliquota_imposto: float = .14, antecipacao: float = .025,
                       frete: float = 0, marketplace_venda: str = "Box fan") -> pd.DataFrame:
    df = pedidos.copy()
    for coluna in NUMERICAS:
        df[coluna] = _numero(df[coluna], coluna)
    df["Quantidade"] = _numero(df["Quantidade"], "Quantidade").astype("Int64")
    df["ID do pedido"] = df["ID do pedido"].map(_texto)
    conjunta = pd.DataFrame({
        "ID do pedido": df["ID do pedido"],
        "Preço total": df["Subtotal do produto"],
        "Desconto A vista": df["Desconto do vendedor"] - df["Desconto do vendedor_2"],
    }).groupby("ID do pedido", as_index=False, sort=False).sum()
    df = df.merge(conjunta, on="ID do pedido", how="left", validate="many_to_one")

    devolucao = df["Status da Devolução / Reembolso"].map(_texto).ne("")
    status = df["Status do pedido"].map(_texto)
    venda = status.isin(["Concluído", "Enviado", "A Enviar", "Entregue"]) | status.str.startswith(
        "O comprador pode pedir uma devolução", na=False)
    recebido = status.eq("Order Received")
    nao_pago = ((status.eq("Cancelado") & df["Hora do pagamento do pedido"].map(_texto).eq("-"))
                | status.eq("Não pago"))
    df["Status"] = "Cancelado"
    df.loc[nao_pago, "Status"] = "Não pago"
    df.loc[recebido, "Status"] = "Devolução"
    df.loc[venda, "Status"] = "Venda"
    df.loc[devolucao, "Status"] = "Devolução"

    pagamento = df["Hora do pagamento do pedido"].map(_texto)
    criacao = df["Data de criação do pedido"].map(_texto)
    df["Data Completa"] = pagamento.where(pagamento.ne("-") & pagamento.ne(""), criacao)
    df["Data"] = pd.to_datetime(df["Data Completa"].str.split().str[0], errors="coerce")
    share = (df["Subtotal do produto"] / df["Preço total"].replace(0, pd.NA)).fillna(0)
    df["Comissão"] = -(df["Taxa de comissão bruta"] + df["Taxa de serviço bruta"]) * share
    desc = df["Cupom do vendedor"] + df["Coin Cashback Voucher Amount Sponsored by Seller"]
    base = df["Desconto A vista"] - df["Ajuste por participação em ação comercial"]
    avista = (-base).where(base <= 0, 0)
    df["Desconto"] = -(desc + avista) * share
    df["Rebate"] = (df["Incentivo Shopee para ação comercial"]
                    + df["Ajuste por participação em ação comercial"]
                    + df["Desconto Shopee da Leve Mais por Menos"]) * share
    df["Faturamento"] = df["Desconto Shopee da Leve Mais por Menos"] * share + df["Subtotal do produto"]
    df["Frete"] = float(frete)
    df["Cancelamento"] = 0.0      # cancelado/não pago/devolução: o tratador (zerar_canceladas) estorna o faturamento aqui
    df["Líquido"] = df["Faturamento"] + df["Rebate"] + df["Desconto"] + df["Comissão"] + df["Frete"]
    df["Antecipa"] = df["Líquido"] * -float(antecipacao)

    df["_sku_completo"] = df["Número de referência SKU"].map(_texto)
    df["Nome do Produto"] = df["Nome do Produto"].map(_texto)
    if ids is None or ids.empty:
        df["ID do Produto"] = df["_sku_completo"]
    else:
        df = df.merge(ids.rename(columns={"sku": "_sku_completo"}),
                      on=["_sku_completo", "Nome do Produto"], how="left", validate="many_to_one")
        df["ID do Produto"] = df["ID do Produto"].where(
            df["ID do Produto"].notna(), df["_sku_completo"]
        )
    df["Imposto"] = float(aliquota_imposto) * -(df["Faturamento"] + df["Desconto"])
    df["Número de referência SKU"] = df["_sku_completo"].str.split("-").str[0]
    custo = custos[["SKU", "Valor total"]].copy()
    custo["SKU"] = custo["SKU"].map(_texto)
    df = df.merge(custo, left_on="Número de referência SKU", right_on="SKU",
                  how="left", validate="many_to_one")
    df["Custo do produto"] = -df["Valor total"] * df["Quantidade"]
    df["Gross margin"] = df["Líquido"] + df["Imposto"] + df["Custo do produto"] + df["Antecipa"]
    df["Gross margin %"] = df["Gross margin"] / df["Faturamento"].replace(0, pd.NA)
    df["Marketplace"] = "Shopee"
    df["Marketplace venda"] = marketplace_venda
    return df[COLUNAS_SAIDA].reset_index(drop=True)


def processar_shopee(caminho_pedidos: str | Path, caminho_custos: str | Path,
                     caminho_ids: str | Path | None, caminho_saida: str | Path,
                     aliquota_imposto: float = .14, antecipacao: float = .025,
                     frete: float = 0, marketplace_venda: str = "Box fan") -> tuple[str, dict]:
    ids = _ler_ids(caminho_ids) if caminho_ids is not None else None
    resultado = transformar_shopee(_ler_pedidos(caminho_pedidos), _carregar_custos(caminho_custos),
                                   ids, aliquota_imposto, antecipacao,
                                   frete, marketplace_venda)
    caminho_saida = Path(caminho_saida).resolve()
    resultado.to_excel(caminho_saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(caminho_saida)
    ws = wb["DB"]
    for celula in ws[1]:
        celula.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho_saida)
    return str(caminho_saida), {
        "linhas": len(resultado),
        "semCusto": int(resultado["Custo do produto"].isna().sum()),
        "semIdProduto": int(resultado["ID do Produto"].isna().sum()),
    }
