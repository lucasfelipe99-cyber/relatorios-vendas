"""Tratamento dos relatórios de vendas do Magalu entregues em arquivo ZIP."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from tratamento_mercado_livre import ErroDeEntrada, _carregar_custos

COLUNAS_SAIDA = [
    "Marketplace", "Marketplace venda", "Número do pedido", "Data do Pedido", "Data",
    "Status pacote no momento que o relatório foi solicitado", "Status",
    "Modalidade de entrega", "Codigo SKU seller", "Título do produto", "SKU.1",
    "Valor Total do Item", "Quantidade de itens", "Fat", "Desconto", "Rebate",
    "Comissão", "Frete", "Cancelamento", "Liquido", "Antecipa", "Imposto",
    "Custo do produto", "Gross margin", "Gross margin %",
]

MOEDAS = [
    "Valor Total do Item", "Valor bruto do pedido", "Valor total dos itens do pedido",
    "Coparticipação de Fretes estimada",
    "Pago pelo Parceiro (Coparticipação de Desconto à Vista)",
    "Pago pelo Magalu (Coparticipação de Desconto à Vista)",
    "Pago pelo Magalu (Coparticipação de Preço Promocional)",
    "Pago pelo Parceiro (Coparticipação de Preço Promocional)",
    "Pago pelo Magalu (Valor subsídio Cupom)",
    "Pago pelo Parceiro (Valor subsídio Cupom)",
    "Tarifa fixa (Forma de pagamento 1)", "Tarifa fixa (Forma de pagamento 2)",
    "Serviços do marketplace (1+2+3+4) (Forma de pagamento 1)",
]

OBRIGATORIAS_PEDIDOS = [
    "Data do Pedido", "Número do pedido", "Codigo SKU seller", "Título do produto",
    "Quantidade de itens", "Valor Total do Item", "Valor bruto do pedido",
    "Coparticipação de Fretes estimada",
    "Pago pelo Parceiro (Coparticipação de Desconto à Vista)",
    "Pago pelo Magalu (Coparticipação de Desconto à Vista)",
    "Pago pelo Magalu (Coparticipação de Preço Promocional)",
    "Pago pelo Parceiro (Coparticipação de Preço Promocional)",
    "Pago pelo Magalu (Valor subsídio Cupom)", "Pago pelo Parceiro (Valor subsídio Cupom)",
    "Tarifa fixa (Forma de pagamento 1)", "Tarifa fixa (Forma de pagamento 2)",
    "Serviços do marketplace (1+2+3+4) (Forma de pagamento 1)",
]
OBRIGATORIAS_PACOTES = [
    "Número do pedido", "Status pacote no momento que o relatório foi solicitado",
    "Modalidade de entrega",
]


def _validar(df: pd.DataFrame, colunas: list[str], origem: str) -> None:
    ausentes = [c for c in colunas if c not in df.columns]
    if ausentes:
        raise ErroDeEntrada(f"Colunas ausentes em {origem}: {', '.join(ausentes)}.")


def _moeda(valor: Any) -> float:
    if pd.isna(valor):
        return 0.0
    texto = str(valor).strip()
    if not texto or texto.lower() in {"não se aplica", "nao se aplica", "-"}:
        return 0.0
    texto = texto.replace("R$", "").replace(" ", "")
    if "," in texto and "." in texto:
        texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError as exc:
        raise ErroDeEntrada(f"Valor monetário inválido no relatório Magalu: {valor!r}.") from exc


def _ler_zip(caminho: str | Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    caminho = Path(caminho)
    try:
        with zipfile.ZipFile(caminho) as arquivo:
            nomes = arquivo.namelist()
            pedidos = [n for n in nomes if "pedidos" in Path(n).name.lower() and n.lower().endswith(".csv")]
            pacotes = [n for n in nomes if "pacotes" in Path(n).name.lower() and n.lower().endswith(".csv")]
            if len(pedidos) != 1 or len(pacotes) != 1:
                raise ErroDeEntrada(
                    "O ZIP do Magalu deve conter exatamente um CSV de pedidos e um CSV de pacotes."
                )
            df_pedidos = pd.read_csv(io.BytesIO(arquivo.read(pedidos[0])), encoding="utf-8-sig")
            df_pacotes = pd.read_csv(io.BytesIO(arquivo.read(pacotes[0])), encoding="utf-8-sig")
    except zipfile.BadZipFile as exc:
        raise ErroDeEntrada("O arquivo selecionado não é um ZIP válido do Magalu.") from exc
    _validar(df_pedidos, OBRIGATORIAS_PEDIDOS, "CSV de pedidos")
    _validar(df_pacotes, OBRIGATORIAS_PACOTES, "CSV de pacotes")
    return df_pedidos, df_pacotes


def _normalizar_sku(valor: Any) -> str:
    texto = "" if pd.isna(valor) else str(valor).strip()
    if texto.startswith("GA"):
        base = texto.split("-", 1)[0]
        if base.endswith("VA"):
            base = base[:-2]
        elif base.endswith("V"):
            base = base[:-1]
        return base
    return texto


def transformar_magalu(
    pedidos: pd.DataFrame,
    pacotes: pd.DataFrame,
    custos: pd.DataFrame,
    ids: pd.DataFrame | None = None,
    aliquota_imposto: float = .14,
    marketplace_venda: str = "Box fan",
) -> pd.DataFrame:
    _validar(pedidos, OBRIGATORIAS_PEDIDOS, "pedidos")
    _validar(pacotes, OBRIGATORIAS_PACOTES, "pacotes")
    df = pedidos.copy()
    for coluna in MOEDAS:
        df[coluna] = df[coluna].map(_moeda)
    df["Quantidade de itens"] = pd.to_numeric(df["Quantidade de itens"], errors="coerce").astype("Int64")
    if df["Quantidade de itens"].isna().any():
        raise ErroDeEntrada("A coluna 'Quantidade de itens' contém valores inválidos.")

    status = pacotes[OBRIGATORIAS_PACOTES].drop_duplicates("Número do pedido", keep="first")
    df = df.merge(status, on="Número do pedido", how="left", validate="many_to_one")
    vendidos = {"Nota fiscal gerada e aprovada", "Pedido despachado", "Pedido entregue"}
    df["Status"] = df["Status pacote no momento que o relatório foi solicitado"].map(
        lambda valor: "Venda" if valor in vendidos else "Cancelado"
    )

    tarifa1 = df["Tarifa fixa (Forma de pagamento 1)"]
    forma2_na = pedidos["Tarifa fixa (Forma de pagamento 2)"].astype(str).str.strip().eq("Não se aplica")
    df["Tarifa fixa"] = tarifa1.where(forma2_na, tarifa1 + df["Tarifa fixa (Forma de pagamento 2)"])
    df["Serviços marketplace"] = df["Serviços do marketplace (1+2+3+4) (Forma de pagamento 1)"]
    df["Desconto visiveis"] = (
        df["Pago pelo Parceiro (Coparticipação de Desconto à Vista)"]
        - df["Pago pelo Magalu (Coparticipação de Desconto à Vista)"]
        - df["Pago pelo Magalu (Valor subsídio Cupom)"]
        + df["Pago pelo Parceiro (Valor subsídio Cupom)"]
        - df["Pago pelo Magalu (Coparticipação de Preço Promocional)"]
    )
    df["Fat"] = df["Quantidade de itens"] * df["Valor Total do Item"]
    df["Elegível"] = (df["Valor Total do Item"] > 78.99).astype(int)
    resumo = df.groupby("Número do pedido", as_index=False, sort=False).agg(
        desconto_promocional=("Pago pelo Parceiro (Coparticipação de Preço Promocional)", "sum"),
        descontos_visiveis=("Desconto visiveis", "sum"),
        preco_total=("Fat", "sum"), contagem=("Número do pedido", "size"),
        elegiveis=("Elegível", "sum"),
    )
    df = df.merge(resumo, on="Número do pedido", how="left", validate="many_to_one")
    diferenca = df["Valor bruto do pedido"] - (df["preco_total"] + df["descontos_visiveis"])
    rateio_oculto = (
        df["Pago pelo Parceiro (Coparticipação de Preço Promocional)"]
        / df["desconto_promocional"].replace(0, pd.NA) * diferenca
    ).fillna(0)
    df["Desconto"] = df["Desconto visiveis"] + rateio_oculto.where(diferenca <= 0, 0)
    df["Data"] = df["Data do Pedido"].astype(str).str.split().str[0]
    df["Rebate"] = (
        df["Pago pelo Magalu (Coparticipação de Desconto à Vista)"]
        + df["Pago pelo Magalu (Coparticipação de Preço Promocional)"]
        + df["Pago pelo Magalu (Valor subsídio Cupom)"]
    )
    # float: com Int64, a divisão da tarifa entre os itens (ex.: -5/3) não cabia na coluna e dava erro
    tarifa_rateada = (-5 * df["Quantidade de itens"].astype("float64")).where(
        df["Tarifa fixa"].eq(0), df["Tarifa fixa"] / df["contagem"]
    )
    comissao_variavel = (df["Fat"] / df["preco_total"].replace(0, pd.NA)).fillna(0) * df["Serviços marketplace"]
    df["Comissão"] = tarifa_rateada + comissao_variavel
    df["Frete"] = 0.0
    frete_valido = df["Coparticipação de Fretes estimada"].ne(0) & df["Elegível"].eq(1) & df["elegiveis"].gt(0)
    df.loc[frete_valido, "Frete"] = (
        df.loc[frete_valido, "Coparticipação de Fretes estimada"] / df.loc[frete_valido, "elegiveis"]
    )
    df["Liquido"] = df["Fat"] + df["Desconto"] + df["Rebate"] + df["Comissão"] + df["Frete"]
    df["Imposto"] = float(aliquota_imposto) * -(df["Fat"] + df["Desconto"])

    df["_codigo"] = df["Codigo SKU seller"].astype(str).str.strip()
    df["SKU.1"] = df["_codigo"].map(_normalizar_sku)
    if ids is not None and not ids.empty:
        mapa = ids[["ID MAGALU", "SKU"]].copy()
        mapa["ID MAGALU"] = mapa["ID MAGALU"].astype(str).str.strip()
        mapa = mapa.drop_duplicates("ID MAGALU", keep="first")
        df = df.merge(mapa.rename(columns={"ID MAGALU": "_codigo", "SKU": "_sku_mapeado"}),
                      on="_codigo", how="left", validate="many_to_one")
        df["SKU.1"] = df["_sku_mapeado"].where(df["_sku_mapeado"].notna(), df["SKU.1"])
    custo = custos[["SKU", "Valor total"]].copy()
    custo["SKU"] = custo["SKU"].astype(str).str.strip()
    df = df.merge(custo, left_on="SKU.1", right_on="SKU", how="left", validate="many_to_one")
    df["Custo do produto"] = -df["Valor total"] * df["Quantidade de itens"]
    df["Gross margin"] = df["Liquido"] + df["Imposto"] + df["Custo do produto"]
    df["Gross margin %"] = df["Gross margin"] / df["Fat"].replace(0, pd.NA)
    df["Marketplace"] = "Magalu"
    df["Marketplace venda"] = marketplace_venda
    df["Cancelamento"] = 0.0
    df["Antecipa"] = 0.0
    return df[COLUNAS_SAIDA].reset_index(drop=True)


def processar_magalu(
    caminho_zip: str | Path, caminho_custos: str | Path, caminho_saida: str | Path,
    caminho_ids: str | Path | None = None, aliquota_imposto: float = .14,
    marketplace_venda: str = "Box fan",
) -> tuple[str, dict[str, Any]]:
    pedidos, pacotes = _ler_zip(caminho_zip)
    custos = _carregar_custos(caminho_custos)
    ids = pd.read_csv(caminho_ids, dtype=str) if caminho_ids and Path(caminho_ids).is_file() else None
    resultado = transformar_magalu(pedidos, pacotes, custos, ids, aliquota_imposto, marketplace_venda)
    saida = Path(caminho_saida).resolve()
    resultado.to_excel(saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(saida)
    ws = wb["DB"]
    for celula in ws[1]:
        celula.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(saida)
    return str(saida), {
        "linhas": len(resultado), "semCusto": int(resultado["Custo do produto"].isna().sum()),
        "canceladas": int(resultado["Status"].eq("Cancelado").sum()),
    }
