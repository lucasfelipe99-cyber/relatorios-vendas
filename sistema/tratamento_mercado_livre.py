"""Tratamento do relatório de vendas exportado pelo Mercado Livre."""

from __future__ import annotations

import argparse
import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font


COL_VENDA = "N.º de venda"
COL_DATA_VENDA = "Data da venda"
COL_ESTADO = "Estado"
COL_UNIDADES = "Unidades"
COL_RECEITA = "Receita por produtos (BRL)"
COL_TARIFA = "Tarifa de venda e impostos (BRL)"
COL_RECEITA_ENVIO = "Receita por envio (BRL)"
COL_TARIFA_ENVIO = "Tarifas de envio (BRL)"
COL_CUSTO_ENVIO = "Custo de envio com base nas medidas e peso declarados"
COL_CANCELAMENTO = "Cancelamentos e reembolsos (BRL)"
COL_TOTAL = "Total (BRL)"
COL_SKU = "SKU"
COL_ANUNCIO = "# de anúncio"
COL_TITULO = "Título do anúncio"
COL_PRECO = "Preço unitário de venda do anúncio (BRL)"
COL_TRANSPORTADOR = "Transportador"
COL_RASTREAMENTO = "Número de rastreamento"
COL_ENTREGA = "Forma de entrega"
COL_LINHA_BRANCA = "__linha_branca__"

# O Mercado Livre alterou este cabeçalho em setembro de 2026. Mantemos o
# nome canônico usado internamente para aceitar relatórios novos e antigos.
ALIASES_COLUNAS = {
    "Data de venda": COL_DATA_VENDA,
}

COLUNAS_OBRIGATORIAS = [
    COL_VENDA,
    COL_DATA_VENDA,
    COL_ESTADO,
    "Descrição do status",
    COL_UNIDADES,
    COL_RECEITA,
    "Receita por acréscimo no preço (pago pelo comprador)",
    COL_TARIFA,
    COL_RECEITA_ENVIO,
    COL_TARIFA_ENVIO,
    COL_CUSTO_ENVIO,
    COL_CANCELAMENTO,
    COL_TOTAL,
    COL_SKU,
    COL_ANUNCIO,
    COL_TITULO,
    COL_PRECO,
    COL_TRANSPORTADOR,
    COL_RASTREAMENTO,
    COL_ENTREGA,
]

COLUNAS_VALIDACAO_VERMELHA = [
    "Pertence a um kit",
    COL_UNIDADES,
    COL_RECEITA,
    "Receita por acréscimo no preço (pago pelo comprador)",
    "Taxa de parcelamento equivalente ao acréscimo",
    COL_TARIFA,
    COL_RECEITA_ENVIO,
    COL_TARIFA_ENVIO,
    COL_CUSTO_ENVIO,
    "Custo por diferenças nas medidas e no peso do pacote",
    "Descontos e bônus",
    COL_CANCELAMENTO,
    COL_TOTAL,
]

COLUNAS_SAIDA = [
    "Marketplace",
    "Marketplace venda",
    COL_VENDA,
    COL_DATA_VENDA,
    "Data",
    COL_ESTADO,
    "Type.1",
    COL_ENTREGA,
    COL_ANUNCIO,
    COL_TITULO,
    COL_SKU,
    COL_PRECO,
    COL_UNIDADES,
    "Faturamento",
    "Desconto",
    "rebate",
    "Comissão",
    "Frete",
    "Cancelamento",
    "Liquido",
    "Antecipa",
    "Imposto",
    "Custo do produto",
    "Gross margen",
    "Gross margen %",
    "Id mercado Pago",
]

MESES = {
    "janeiro": 1,
    "fevereiro": 2,
    "março": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12,
}


class ErroDeEntrada(ValueError):
    """Erro legível relacionado à estrutura dos arquivos de entrada."""


def _texto_vazio(valor: Any) -> bool:
    return pd.isna(valor) or (isinstance(valor, str) and valor.strip() == "")


def _vazios_para_na(serie: pd.Series) -> pd.Series:
    return serie.map(lambda valor: pd.NA if _texto_vazio(valor) else valor)


def _como_identificador(serie: pd.Series) -> pd.Series:
    def converter(valor: Any) -> Any:
        if _texto_vazio(valor):
            return pd.NA
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()

    return serie.map(converter).astype(object)


def _numerico(serie: pd.Series, nome: str) -> pd.Series:
    limpa = _vazios_para_na(serie)
    if limpa.dtype == object:
        limpa = limpa.map(
            lambda valor: (
                valor.replace(".", "").replace(",", ".")
                if isinstance(valor, str) and "," in valor
                else valor
            )
        )
    resultado = pd.to_numeric(limpa, errors="coerce")
    invalidos = limpa.notna() & resultado.isna()
    if invalidos.any():
        exemplos = limpa[invalidos].astype(str).head(3).tolist()
        raise ErroDeEntrada(
            f"A coluna '{nome}' contém valores que não são numéricos: {exemplos}."
        )
    return resultado.astype("float64")


def _validar_colunas(df: pd.DataFrame, obrigatorias: list[str], origem: str) -> None:
    ausentes = [coluna for coluna in obrigatorias if coluna not in df.columns]
    if ausentes:
        raise ErroDeEntrada(
            f"Colunas ausentes em {origem}: {', '.join(ausentes)}. "
            f"Colunas encontradas: {', '.join(map(str, df.columns))}."
        )


def _normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    renomear = {
        alias: canonica
        for alias, canonica in ALIASES_COLUNAS.items()
        if alias in df.columns and canonica not in df.columns
    }
    return df.rename(columns=renomear)


def _ler_planilha_com_colunas(
    caminho: str | Path, obrigatorias: list[str], aba: str | None = None
) -> pd.DataFrame:
    caminho = Path(caminho)
    if not caminho.exists():
        raise ErroDeEntrada(f"Arquivo não encontrado: {caminho}")
    if caminho.suffix.lower() == ".csv":
        df = pd.read_csv(caminho, sep=None, engine="python")
        _validar_colunas(df, obrigatorias, str(caminho))
        return df

    if aba:
        df = pd.read_excel(caminho, sheet_name=aba)
        _validar_colunas(df, obrigatorias, f"{caminho} (aba {aba})")
        return df

    abas = pd.ExcelFile(caminho).sheet_names
    for nome_aba in abas:
        candidato = pd.read_excel(caminho, sheet_name=nome_aba)
        if all(coluna in candidato.columns for coluna in obrigatorias):
            return candidato
    raise ErroDeEntrada(
        f"Nenhuma aba de {caminho} contém as colunas: {', '.join(obrigatorias)}."
    )


def _carregar_custos(caminho_custos: str | Path) -> pd.DataFrame:
    custos = _ler_planilha_com_colunas(
        caminho_custos, [COL_SKU, "Sub Categoria", "Valor"]
    ).copy()
    custos[COL_SKU] = _vazios_para_na(custos[COL_SKU])
    custos["Sub Categoria"] = custos["Sub Categoria"].astype("string").str.strip()
    custos["Valor"] = _numerico(custos["Valor"], "Valor")
    custos = custos[~custos["Sub Categoria"].isin(["Energia", "Mão De Obra"])]
    return (
        custos.dropna(subset=[COL_SKU])
        .groupby(COL_SKU, as_index=False, sort=False, dropna=False)["Valor"]
        .sum(min_count=1)
        .rename(columns={"Valor": "Valor total"})
    )


def _carregar_mapa_status(
    mapa_status: Mapping[str, str] | pd.DataFrame | str | Path,
) -> pd.DataFrame:
    if isinstance(mapa_status, Mapping):
        mapa = pd.DataFrame(list(mapa_status.items()), columns=["Status", "Type"])
    elif isinstance(mapa_status, pd.DataFrame):
        mapa = mapa_status.copy()
    else:
        caminho = Path(mapa_status)
        mapa = _ler_planilha_com_colunas(caminho, ["Status", "Type"])
        # O arquivo operacional pode ter uma aba Proc que ainda não contém
        # status recém-criados pelo Mercado Livre. Quando o próprio arquivo
        # tratado é fornecido, seu histórico DB completa esses casos sem
        # sobrescrever as classificações explícitas de Proc.
        if caminho.suffix.lower() != ".csv":
            excel = pd.ExcelFile(caminho)
            if "DB" in excel.sheet_names:
                historico = pd.read_excel(caminho, sheet_name="DB", usecols="A:Z")
                if {COL_ESTADO, "Type.1"}.issubset(historico.columns):
                    historico = historico[[COL_ESTADO, "Type.1"]].rename(
                        columns={COL_ESTADO: "Status", "Type.1": "Type"}
                    )
                    mapa = pd.concat([mapa, historico], ignore_index=True)
    _validar_colunas(mapa, ["Status", "Type"], "mapa_status")
    mapa = mapa[["Status", "Type"]].drop_duplicates("Status", keep="first")
    return mapa


def _classificar_status_automaticamente(estado: pd.Series) -> pd.Series:
    """Classifica o Estado sem depender de uma tabela externa.

    A prioridade é importante: estados de devolução podem também mencionar
    cancelamento ou reembolso, mas operacionalmente pertencem à Devolução.
    """
    texto = estado.astype("string").str.strip()
    normalizado = (
        texto.str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("ascii")
        .str.lower()
    )
    devolucao = normalizado.str.contains(
        r"devolu|devolv|reembolso|retorn(?:o|ando|ou)|produto devolvido",
        regex=True,
        na=False,
    )
    cancelada = normalizado.str.contains(
        r"cancel|compra recusada|venda recusada",
        regex=True,
        na=False,
    )
    resultado = pd.Series("Venda", index=estado.index, dtype="string")
    resultado.loc[cancelada] = "Cancelada"
    resultado.loc[devolucao] = "Devolução"
    # A devolução apenas "habilitada" ainda não ocorreu; no processo atual
    # esse estado continua representando uma venda.
    resultado.loc[normalizado.str.contains("mediacao com devolucao habilitada", na=False)] = "Venda"
    resultado.loc[estado.isna()] = pd.NA
    return resultado


def _converter_data_portugues(valor: Any) -> pd.Timestamp | pd.NaT:
    if _texto_vazio(valor):
        return pd.NaT
    texto = str(valor).strip().lower()
    correspondencia = re.match(r"^(\d{1,2})\s+de\s+([a-zç]+)\s+de\s+(\d{4})", texto)
    if not correspondencia:
        raise ErroDeEntrada(f"Data da venda em formato inesperado: {valor!r}")
    dia, mes_texto, ano = correspondencia.groups()
    mes = MESES.get(mes_texto)
    if mes is None:
        raise ErroDeEntrada(f"Mês em português não reconhecido: {mes_texto!r}")
    return pd.Timestamp(int(ano), mes, int(dia))


def _preenchimento_vermelho(celula: Any) -> bool:
    """Detecta os tons vermelhos usados pelo relatório para sinalizar erro."""
    preenchimento = celula.fill
    cor = preenchimento.fgColor
    if preenchimento.fill_type != "solid" or cor.type != "rgb":
        return False
    rgb = str(cor.rgb).upper()
    if len(rgb) == 8:
        rgb = rgb[2:]
    if len(rgb) != 6:
        return False
    try:
        vermelho, verde, azul = (int(rgb[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    # Inclui F4CCCC (tom observado no Mercado Livre) e tons equivalentes,
    # evitando confundir cinzas, amarelos e o cabeçalho institucional.
    return vermelho >= 220 and vermelho - verde >= 20 and vermelho - azul >= 20


def _celula_branca(celula: Any) -> bool:
    """Retorna True quando a célula não possui destaque visual de linha."""
    preenchimento = celula.fill
    if preenchimento.fill_type is None:
        return True
    cor = preenchimento.fgColor
    if preenchimento.fill_type == "solid" and cor.type == "rgb":
        rgb = str(cor.rgb).upper()
        return rgb in {"FFFFFFFF", "00FFFFFF"}
    return False


def _validar_linhas_vermelhas(caminho: Path) -> list[bool]:
    try:
        workbook = load_workbook(caminho, read_only=False, data_only=True)
    except Exception as exc:
        raise ErroDeEntrada(f"Não foi possível abrir o arquivo de vendas: {caminho.name}.") from exc
    try:
        if "Vendas BR" not in workbook.sheetnames:
            raise ErroDeEntrada(
                f"O arquivo {caminho} não contém a aba obrigatória 'Vendas BR'."
            )
        planilha = workbook["Vendas BR"]
        cabecalhos = {
            str(celula.value).strip(): celula.column
            for celula in planilha[6]
            if celula.value is not None
        }
        inicio = cabecalhos.get(COLUNAS_VALIDACAO_VERMELHA[0])
        fim = cabecalhos.get(COLUNAS_VALIDACAO_VERMELHA[-1])
        if inicio is None or fim is None or inicio > fim:
            raise ErroDeEntrada(
                "Não foi possível localizar o bloco de validação entre "
                f"'{COLUNAS_VALIDACAO_VERMELHA[0]}' e "
                f"'{COLUNAS_VALIDACAO_VERMELHA[-1]}'."
            )
        linhas = sorted(
            {
                celula.row
                for linha in planilha.iter_rows(
                    min_row=7, min_col=inicio, max_col=fim
                )
                for celula in linha
                if _preenchimento_vermelho(celula)
            }
        )
        # A primeira coluna acompanha a cor visual aplicada à linha pelo
        # Mercado Livre: filhos de pacote vêm destacados, enquanto uma nova
        # venda independente volta a ter fundo branco.
        linhas_brancas = [
            _celula_branca(planilha.cell(numero_linha, 1))
            for numero_linha in range(7, planilha.max_row + 1)
        ]
    finally:
        workbook.close()
    if linhas:
        exemplos = ", ".join(map(str, linhas[:8]))
        complemento = "…" if len(linhas) > 8 else ""
        raise ErroDeEntrada(
            "Arquivo não aceito: foram encontradas "
            f"{len(linhas)} linha(s) destacada(s) em vermelho "
            f"(linhas {exemplos}{complemento}). "
            "Baixe um novo relatório diretamente na plataforma do Mercado Livre e tente novamente."
        )
    return linhas_brancas


def _ler_vendas(caminho_vendas: str | Path) -> pd.DataFrame:
    caminho = Path(caminho_vendas)
    if not caminho.exists():
        raise ErroDeEntrada(f"Arquivo de vendas não encontrado: {caminho}")
    linhas_brancas = _validar_linhas_vermelhas(caminho)
    try:
        vendas = pd.read_excel(caminho, sheet_name="Vendas BR", skiprows=5)
    except ValueError as exc:
        raise ErroDeEntrada(
            f"O arquivo {caminho} não contém a aba obrigatória 'Vendas BR'."
        ) from exc
    vendas = _normalizar_colunas(vendas)
    _validar_colunas(vendas, COLUNAS_OBRIGATORIAS, f"{caminho} (aba Vendas BR)")
    vendas[COL_LINHA_BRANCA] = (
        linhas_brancas[: len(vendas)]
        + [False] * max(0, len(vendas) - len(linhas_brancas))
    )
    return vendas


def transformar_vendas(
    vendas: pd.DataFrame,
    custos: pd.DataFrame,
    aliquota_imposto: float,
    mapa_status: Mapping[str, str] | pd.DataFrame | None = None,
    marketplace_venda: str = "Box fan",
) -> pd.DataFrame:
    """Aplica a mesma transformação da consulta de vendas, sem gravar arquivo."""
    vendas = _normalizar_colunas(vendas)
    _validar_colunas(vendas, COLUNAS_OBRIGATORIAS, "relatório de vendas")
    _validar_colunas(custos, [COL_SKU, "Valor total"], "tabela de custos agregada")
    mapa = _carregar_mapa_status(mapa_status) if mapa_status is not None else None
    df = vendas.copy()

    # O relatório pode repetir nomes em blocos posteriores. O pandas preserva a
    # primeira ocorrência sem sufixo, que é o bloco usado pela consulta original.
    df[COL_VENDA] = _como_identificador(df[COL_VENDA])
    for coluna in [COL_ESTADO, COL_TRANSPORTADOR, COL_RASTREAMENTO, COL_ENTREGA]:
        df[coluna] = _vazios_para_na(df[coluna])

    estado_pacote = df[COL_ESTADO].astype("string").str.startswith("Pacote de", na=False)
    rastreio_vazio = df[COL_RASTREAMENTO].isna()
    transportador_preenchido = df[COL_TRANSPORTADOR].notna()
    linha_branca = (
        df[COL_LINHA_BRANCA].fillna(False).astype(bool)
        if COL_LINHA_BRANCA in df.columns
        else pd.Series(False, index=df.index)
    )
    df["Identificação"] = df[COL_RASTREAMENTO].astype(object)
    usar_numero_venda = estado_pacote | transportador_preenchido | linha_branca
    df.loc[rastreio_vazio & usar_numero_venda, "Identificação"] = df[COL_VENDA]
    df.loc[rastreio_vazio & ~usar_numero_venda, "Identificação"] = pd.NA
    df["Identificação"] = df["Identificação"].ffill()

    colunas_financeiras = [
        COL_TARIFA,
        COL_RECEITA_ENVIO,
        COL_TARIFA_ENVIO,
        COL_CUSTO_ENVIO,
        COL_CANCELAMENTO,
        COL_TOTAL,
        COL_RECEITA,
        COL_PRECO,
        COL_UNIDADES,
    ]
    for coluna in colunas_financeiras:
        df[coluna] = _numerico(df[coluna], coluna)

    # Tabela dos pais: contém os totais financeiros que serão rateados para
    # todos os itens filhos associados à mesma Identificação.
    contagens = df.groupby("Identificação", dropna=True)["Identificação"].transform("size")
    pais = df.loc[estado_pacote].copy()
    pais["Contagem VC"] = contagens.loc[estado_pacote] - 1
    pais["Frete Unitário"] = pais[COL_RECEITA_ENVIO].fillna(0) + pais[COL_TARIFA_ENVIO].fillna(0)
    pais["Rebate/Desconto"] = (
        pais[COL_TOTAL].fillna(0)
        - (
            pais[COL_RECEITA].fillna(0)
            + pais[COL_TARIFA].fillna(0)
            + pais[COL_RECEITA_ENVIO].fillna(0)
            + pais[COL_TARIFA_ENVIO].fillna(0)
            + pais[COL_CANCELAMENTO].fillna(0)
        )
    ).round(2)
    if (pais["Contagem VC"] <= 0).any():
        ids = pais.loc[pais["Contagem VC"] <= 0, "Identificação"].astype(str).tolist()
        raise ErroDeEntrada(f"Pacote sem item filho para rateio: {ids[:5]}")

    # Dados descritivos dos itens aparecem nas linhas filhas e, como no Power
    # Query, são preenchidos para cima antes da remoção das linhas pai.
    df[COL_ENTREGA] = df[COL_ENTREGA].ffill()
    for coluna in [COL_PRECO, COL_UNIDADES]:
        df[coluna] = df[coluna].bfill()
    for coluna in [COL_TITULO, COL_ANUNCIO, COL_SKU]:
        df[coluna] = _vazios_para_na(df[coluna]).bfill()

    trazer = [
        "Identificação",
        COL_VENDA,
        COL_RECEITA,
        COL_TARIFA,
        COL_CANCELAMENTO,
        COL_TOTAL,
        "Contagem VC",
        "Frete Unitário",
        "Rebate/Desconto",
    ]
    pais = pais[trazer].drop_duplicates("Identificação", keep="first").rename(
        columns={coluna: f"VC.{coluna}" for coluna in trazer if coluna != "Identificação"}
    )
    df = df.merge(pais, on="Identificação", how="left", validate="many_to_one")
    eh_kit = df[f"VC.{COL_RECEITA}"].notna()
    denominador = df[f"VC.{COL_RECEITA}"]
    if (eh_kit & (denominador == 0)).any():
        raise ErroDeEntrada("Há pacote com Receita por produtos igual a zero; o rateio é impossível.")

    df["Id mercado Pago"] = df[f"VC.{COL_VENDA}"].combine_first(df[COL_VENDA])
    df["Data"] = df[COL_DATA_VENDA].map(_converter_data_portugues)
    df["Faturamento"] = df[COL_RECEITA].fillna(0)
    df.loc[eh_kit, "Faturamento"] = (
        df.loc[eh_kit, COL_PRECO].fillna(0) * df.loc[eh_kit, COL_UNIDADES].fillna(0)
    )

    proporcao = df["Faturamento"] / denominador
    df["Comissão"] = df[COL_TARIFA].fillna(0)
    df.loc[eh_kit, "Comissão"] = (
        proporcao[eh_kit] * df.loc[eh_kit, f"VC.{COL_TARIFA}"].fillna(0)
    )
    df["Frete"] = df[COL_RECEITA_ENVIO].fillna(0) + df[COL_TARIFA_ENVIO].fillna(0)
    df.loc[eh_kit, "Frete"] = (
        proporcao[eh_kit] * df.loc[eh_kit, "VC.Frete Unitário"].fillna(0)
    )

    rebate_desconto = (
        df[COL_TOTAL].fillna(0)
        - (
            df["Faturamento"]
            + df[COL_TARIFA].fillna(0)
            + df[COL_RECEITA_ENVIO].fillna(0)
            + df[COL_TARIFA_ENVIO].fillna(0)
            + df[COL_CANCELAMENTO].fillna(0)
        )
    ).round(2)
    rateio_rebate = df["VC.Rebate/Desconto"] / df["VC.Contagem VC"]
    df["rebate"] = rebate_desconto.where(rebate_desconto > 0, 0)
    df["Desconto"] = rebate_desconto.where(rebate_desconto < 0, 0)
    df.loc[eh_kit, "rebate"] = rateio_rebate[eh_kit].where(rateio_rebate[eh_kit] > 0, 0)
    df.loc[eh_kit, "Desconto"] = rateio_rebate[eh_kit].where(rateio_rebate[eh_kit] < 0, 0)

    df["Cancelamento"] = df[COL_CANCELAMENTO].fillna(0)
    df.loc[eh_kit, "Cancelamento"] = (
        df.loc[eh_kit, f"VC.{COL_CANCELAMENTO}"].fillna(0)
        / df.loc[eh_kit, "VC.Contagem VC"]
    )
    df["Liquido"] = df[
        ["Faturamento", "Comissão", "Frete", "rebate", "Desconto", "Cancelamento"]
    ].sum(axis=1)

    df = df[df[COL_ESTADO].notna() & ~estado_pacote.reindex(df.index, fill_value=False)]
    df = df[df["Faturamento"] != 0].copy()
    df["Imposto"] = (-df["Faturamento"] - df["Desconto"]) * float(aliquota_imposto)
    df = df.merge(custos[[COL_SKU, "Valor total"]], on=COL_SKU, how="left", validate="many_to_one")
    df["Custo do produto"] = -df["Valor total"] * df[COL_UNIDADES]
    df["Gross margen"] = df["Liquido"] + df["Imposto"] + df["Custo do produto"]
    df["Gross margen %"] = df["Gross margen"] / df["Faturamento"]
    if mapa is None:
        df["Type.1"] = _classificar_status_automaticamente(df[COL_ESTADO])
    else:
        df = df.merge(
            mapa, left_on=COL_ESTADO, right_on="Status", how="left", validate="many_to_one"
        )
        df = df.rename(columns={"Type": "Type.1"})
    df["Marketplace"] = "Mercado Livre"
    df["Marketplace venda"] = marketplace_venda
    df["Antecipa"] = 0
    return df[COLUNAS_SAIDA].reset_index(drop=True)


def _formatar_saida(caminho: Path) -> None:
    wb = load_workbook(caminho)
    ws = wb["DB"]
    for celula in ws[1]:
        celula.font = Font(bold=True)
    for celula in ws["E"][1:]:
        celula.number_format = "yyyy-mm-dd"
    for coluna in range(12, 26):
        formato = "0.00%" if coluna == 25 else "#,##0.00"
        for celula in ws.iter_cols(min_col=coluna, max_col=coluna, min_row=2):
            for item in celula:
                item.number_format = formato
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho)


def processar_vendas(
    caminho_vendas: str | Path,
    caminho_custos: str | Path,
    aliquota_imposto: float,
    mapa_status: Mapping[str, str] | pd.DataFrame | str | Path | None = None,
    caminho_xlsx_saida: str | Path | None = None,
    marketplace_venda: str = "Box fan",
) -> str:
    """Processa as vendas e devolve o caminho absoluto do XLSX gerado."""
    vendas = _ler_vendas(caminho_vendas)
    custos = _carregar_custos(caminho_custos)
    mapa = _carregar_mapa_status(mapa_status) if mapa_status is not None else None
    resultado = transformar_vendas(
        vendas, custos, aliquota_imposto, mapa, marketplace_venda=marketplace_venda
    )
    origem = Path(caminho_vendas)
    saida = (
        Path(caminho_xlsx_saida)
        if caminho_xlsx_saida
        else origem.with_name(f"{origem.stem}_tratado.xlsx")
    )
    saida = saida.resolve()
    saida.parent.mkdir(parents=True, exist_ok=True)
    resultado.to_excel(saida, sheet_name="DB", index=False, engine="openpyxl")
    _formatar_saida(saida)
    return str(saida)


def _ler_mapa_json(valor: str) -> Mapping[str, str]:
    try:
        mapa = json.loads(valor)
    except json.JSONDecodeError as exc:
        raise argparse.ArgumentTypeError(f"JSON inválido: {exc}") from exc
    if not isinstance(mapa, dict):
        raise argparse.ArgumentTypeError("O mapa JSON deve ser um objeto Status -> Type.")
    return mapa


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("vendas", help="Relatório bruto XLSX com a aba 'Vendas BR'.")
    parser.add_argument("custos", help="Arquivo XLSX/XLSM/CSV da tabela de custos.")
    grupo = parser.add_mutually_exclusive_group(required=False)
    grupo.add_argument("--mapa-status", help="Arquivo XLSX/XLSM/CSV com Status e Type.")
    grupo.add_argument("--mapa-status-json", type=_ler_mapa_json, help="Objeto JSON Status -> Type.")
    parser.add_argument("--aliquota-imposto", type=float, default=0.14)
    parser.add_argument("--marketplace-venda", default="Box fan")
    parser.add_argument("--saida", help="Caminho do XLSX de saída.")
    args = parser.parse_args()
    mapa: Any = args.mapa_status if args.mapa_status else args.mapa_status_json
    try:
        caminho = processar_vendas(
            args.vendas,
            args.custos,
            args.aliquota_imposto,
            mapa,
            caminho_xlsx_saida=args.saida,
            marketplace_venda=args.marketplace_venda,
        )
    except ErroDeEntrada as exc:
        parser.error(str(exc))
    print(caminho)


if __name__ == "__main__":
    main()
