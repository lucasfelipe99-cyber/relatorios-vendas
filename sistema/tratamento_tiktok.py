"""Tratamento do relatório 'Todos pedido' exportado pelo TikTok Shop."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

from tratamento_mercado_livre import ErroDeEntrada, _carregar_custos

COLUNAS_SAIDA = [
    "Marketplace", "Marketplace venda", "Order ID", "Created Time", "Data",
    "Order Status", "Status", "Fulfillment Type", "SKU ID", "Product Name",
    "Seller SKU", "Preço Unitário", "Quantity", "Faturamento", "Desconto",
    "Rebate", "Comissão", "Frete", "Cancelamento", "Líquido", "Antecipa",
    "Imposto", "Custo do produto", "Gross margin", "Gross margin %",
]

OBRIGATORIAS = [
    "Order ID", "Order Status", "SKU ID", "Seller SKU", "Product Name", "Quantity",
    "SKU Subtotal Before Discount", "SKU Seller Discount", "Created Time", "Fulfillment Type",
]


def _coluna_indice(referencia: str) -> int:
    letras = re.match(r"[A-Z]+", referencia.upper())
    if not letras:
        return 0
    indice = 0
    for letra in letras.group(0):
        indice = indice * 26 + ord(letra) - 64
    return indice - 1


def _ler_xml_tiktok(caminho: Path) -> pd.DataFrame:
    """Lê o XLSX não canônico do TikTok, que repete um elemento row por célula."""
    try:
        with zipfile.ZipFile(caminho) as arquivo:
            nomes = [n for n in arquivo.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")]
            if not nomes:
                raise ErroDeEntrada("O arquivo TikTok não contém uma planilha de pedidos.")
            valores: dict[int, dict[int, Any]] = {}
            with arquivo.open(max(nomes, key=lambda n: arquivo.getinfo(n).file_size)) as fluxo:
                for evento, elemento in ET.iterparse(fluxo, events=("end",)):
                    if elemento.tag.rsplit("}", 1)[-1] != "c":
                        continue
                    ref = elemento.attrib.get("r", "")
                    numero = re.search(r"\d+", ref)
                    if not numero:
                        elemento.clear()
                        continue
                    linha = int(numero.group()) - 1
                    coluna = _coluna_indice(ref)
                    filho = next((f for f in elemento if f.tag.rsplit("}", 1)[-1] in {"v", "t"}), None)
                    valores.setdefault(linha, {})[coluna] = filho.text if filho is not None else None
                    elemento.clear()
    except zipfile.BadZipFile as exc:
        raise ErroDeEntrada("O arquivo selecionado não é um XLSX válido do TikTok.") from exc
    if not valores:
        raise ErroDeEntrada("Nenhuma linha foi encontrada no relatório TikTok.")
    max_coluna = max(max(linha) for linha in valores.values())
    matriz = [[valores.get(i, {}).get(j) for j in range(max_coluna + 1)] for i in range(max(valores) + 1)]
    return pd.DataFrame(matriz[1:], columns=matriz[0])


def _ler_pedidos(caminho: str | Path) -> pd.DataFrame:
    caminho = Path(caminho)
    try:
        df = pd.read_excel(caminho, sheet_name="OrderSKUList", dtype=str)
    except (ValueError, KeyError):
        df = pd.read_excel(caminho, dtype=str)
    if len(df.columns) <= 1:
        df = _ler_xml_tiktok(caminho)
    ausentes = [c for c in OBRIGATORIAS if c not in df.columns]
    if ausentes:
        raise ErroDeEntrada("Colunas ausentes no relatório TikTok: " + ", ".join(ausentes))
    return df


def _brl(valor: Any) -> float:
    if pd.isna(valor):
        return 0.0
    texto = str(valor).strip()
    if not texto or texto in {"-", "--"}:
        return 0.0
    texto = texto.replace("BRL", "").replace("R$", "").replace(" ", "")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError as exc:
        raise ErroDeEntrada(f"Valor monetário inválido no relatório TikTok: {valor!r}.") from exc


def transformar_tiktok(
    pedidos: pd.DataFrame, custos: pd.DataFrame, aliquota_imposto: float = .14,
    marketplace_venda: str = "Box fan",
) -> pd.DataFrame:
    ausentes = [c for c in OBRIGATORIAS if c not in pedidos.columns]
    if ausentes:
        raise ErroDeEntrada("Colunas ausentes no relatório TikTok: " + ", ".join(ausentes))
    df = pedidos.copy()
    df = df[df["Order ID"].notna() & df["Order ID"].astype(str).ne("Platform unique order ID.")].copy()
    df = df[df["Order ID"].astype(str).ne("x")].copy()
    df["Order ID"] = df["Order ID"].astype(str).str.strip()
    df["SKU Subtotal Before Discount"] = df["SKU Subtotal Before Discount"].map(_brl)
    df["SKU Seller Discount"] = df["SKU Seller Discount"].map(_brl)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").astype("Int64")
    if df["Quantity"].isna().any():
        raise ErroDeEntrada("A coluna 'Quantity' contém valores inválidos.")
    df["Faturamento"] = df["SKU Subtotal Before Discount"] - df["SKU Seller Discount"]
    df["Preço Unitário"] = df["Faturamento"] / df["Quantity"].replace(0, pd.NA)
    df["Desconto"] = 0.0
    df["Rebate"] = 0.0
    df["Data"] = pd.to_datetime(
        df["Created Time"].astype(str).str.split().str[0], format="%m/%d/%Y", errors="coerce"
    )
    if df["Data"].isna().any():
        raise ErroDeEntrada("A coluna 'Created Time' contém datas em formato inesperado.")
    regra_nova = df["Data"] > pd.Timestamp("2026-07-14")
    abaixo_50 = df["Faturamento"] < 50
    df["Comissão"] = df["Faturamento"] * -.06
    df.loc[regra_nova & abaixo_50, "Comissão"] = df.loc[regra_nova & abaixo_50, "Faturamento"] * -.10
    df.loc[df["Faturamento"].eq(0), "Comissão"] = 0
    df["Frete"] = df["Faturamento"] * -.06 - 4
    df.loc[regra_nova & ~abaixo_50, "Frete"] = df.loc[regra_nova & ~abaixo_50, "Faturamento"] * -.06 - 6
    df.loc[df["Faturamento"].eq(0), "Frete"] = 0
    venda = df["Order Status"].isin(["Enviado", "A ser enviado", "Para enviar"])
    df["Status"] = "Não pago"
    df.loc[df["Order Status"].eq("Cancelado"), "Status"] = "Cancelado"
    df.loc[venda, "Status"] = "Venda"
    df.loc[df["Faturamento"].eq(0), "Status"] = "Amostra Grátis"
    df["Líquido"] = df["Faturamento"] + df["Comissão"] + df["Frete"]
    df["Imposto"] = float(aliquota_imposto) * -df["Faturamento"]
    df["Seller SKU"] = df["Seller SKU"].astype(str).str.strip()
    # Sem Seller SKU: o SKU vira "Product Name Variation" (cada variação separada); variação "Padrão" fica só o título
    sem_sku = df["Seller SKU"].isna() | df["Seller SKU"].fillna("").str.lower().isin(["", "nan", "none", "-", "--"])
    if sem_sku.any():
        nome = df["Product Name"].map(lambda v: "" if pd.isna(v) else str(v))
        var = df["Variation"].map(lambda v: "" if pd.isna(v) else str(v).strip()) if "Variation" in df.columns \
            else pd.Series("", index=df.index)
        var = var.where(~var.str.lower().isin(["padrão", "padrao", "default"]), "")
        junto = (nome + " " + var).str.replace(r"\s+", " ", regex=True).str.strip()
        df.loc[sem_sku & junto.ne(""), "Seller SKU"] = junto[sem_sku & junto.ne("")]
    custo = custos[["SKU", "Valor total"]].copy()
    custo["SKU"] = custo["SKU"].astype(str).str.strip()
    df = df.merge(custo, left_on="Seller SKU", right_on="SKU", how="left", validate="many_to_one")
    df["Custo do produto"] = -df["Valor total"] * df["Quantity"]
    amostra = df["Status"].eq("Amostra Grátis")
    df.loc[amostra, "Custo do produto"] = 0
    df["Gross margin"] = df["Líquido"] + df["Imposto"] + df["Custo do produto"]
    df["Gross margin %"] = df["Gross margin"] / df["Faturamento"].replace(0, pd.NA)
    df.loc[amostra, "Gross margin %"] = 0
    df["Marketplace"] = "TikTok"
    df["Marketplace venda"] = marketplace_venda
    df["Cancelamento"] = 0.0
    df["Antecipa"] = 0.0
    return df[COLUNAS_SAIDA].reset_index(drop=True)


def processar_tiktok(
    caminho_pedidos: str | Path, caminho_custos: str | Path, caminho_saida: str | Path,
    aliquota_imposto: float = .14, marketplace_venda: str = "Box fan",
) -> tuple[str, dict[str, Any]]:
    resultado = transformar_tiktok(
        _ler_pedidos(caminho_pedidos), _carregar_custos(caminho_custos),
        aliquota_imposto, marketplace_venda,
    )
    saida = Path(caminho_saida).resolve()
    resultado.to_excel(saida, sheet_name="DB", index=False, engine="openpyxl")
    wb = load_workbook(saida); ws = wb["DB"]
    for celula in ws[1]: celula.font = Font(bold=True)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions; wb.save(saida)
    return str(saida), {
        "linhas": len(resultado), "semCusto": int(resultado["Custo do produto"].isna().sum()),
        "canceladas": int(resultado["Status"].eq("Cancelado").sum()),
        "amostrasGratis": int(resultado["Status"].eq("Amostra Grátis").sum()),
    }
