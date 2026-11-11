"""Tratamento Amazon a partir dos três relatórios operacionais."""
from __future__ import annotations
from pathlib import Path
from typing import Any
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from tratamento_mercado_livre import ErroDeEntrada, _carregar_custos

COLUNAS_SAIDA=["Marketplace","Marketplace venda","amazon-order-id","purchase-date","Data","order-status","Status","fulfillment-channel","asin","product-name","SKU.1","Preço Unitário","quantity","Faturamento","item-promotion-discount","Rebate","Comissão","Frete","Cancelamento","Liquido","Antecipa","Imposto","Custo do produto","Margin R$","Margin %"]
BASE_OBRIGATORIAS=["amazon-order-id","purchase-date","order-status","fulfillment-channel","product-name","sku","asin","quantity","item-price","item-tax","shipping-tax","item-promotion-discount","ship-promotion-discount"]
UNIFIED_OBRIGATORIAS=["tipo","id do pedido","sku","descrição","quantidade","vendas do produto","créditos de remessa","descontos promocionais","imposto de vendas coletados","tarifas de venda","taxas fba","taxas de outras transações"]
RECEBER_OBRIGATORIAS=["Tipo de transação","ID do pedido","Detalhes do produto","Tarifas da Amazon","Outros"]

def _validar(df,colunas,origem):
    ausentes=[c for c in colunas if c not in df.columns]
    if ausentes: raise ErroDeEntrada(f"Colunas ausentes em {origem}: {', '.join(ausentes)}.")

def _numero(valor:Any)->float:
    if pd.isna(valor) or str(valor).strip() in {"","-","--"}: return 0.0
    texto=str(valor).strip().replace("R$","").replace(" ","")
    if "," in texto and "." in texto: texto=texto.replace(".","").replace(",",".")
    elif "," in texto: texto=texto.replace(",",".")
    try:return float(texto)
    except ValueError as exc: raise ErroDeEntrada(f"Valor numérico inválido no relatório Amazon: {valor!r}.") from exc

def _ler_vendas(caminho):
    df=pd.read_csv(caminho,sep="\t",dtype=str,encoding="utf-8-sig",keep_default_na=False)
    _validar(df,BASE_OBRIGATORIAS,"TXT de pedidos")
    return df

def _ler_unificado(caminho):
    linhas=Path(caminho).read_text(encoding="utf-8-sig").splitlines()
    cabecalho=next((i for i,l in enumerate(linhas) if l.lower().startswith('"data/hora"')),None)
    if cabecalho is None: raise ErroDeEntrada("Não foi encontrado o cabeçalho no relatório unificado da Amazon.")
    df=pd.read_csv(caminho,skiprows=cabecalho,dtype=str,encoding="utf-8-sig",keep_default_na=False)
    _validar(df,UNIFIED_OBRIGATORIAS,"relatório unificado")
    return df

def _ler_receber(caminho):
    df=pd.read_csv(caminho,dtype=str,encoding="utf-8-sig",keep_default_na=False)
    _validar(df,RECEBER_OBRIGATORIAS,"relatório de transações a receber")
    return df

def _normalizar_sku(valor):
    sku="" if pd.isna(valor) else str(valor).strip()
    if sku.startswith("FBA"):
        sku=sku.split("-",1)[1] if "-" in sku else sku
    else: sku=sku.split("-",1)[0]
    sku=sku.split("-",1)[0]
    return {"4F":"GA1147","1222":"GA1222","8V":"GA1266","57":"GA1209","87":"GA1222"}.get(sku,sku)

def _preparar_unificado(df):
    u=df.copy()
    numericas=["quantidade","vendas do produto","créditos de remessa","descontos promocionais","imposto de vendas coletados","tarifas de venda","taxas fba","taxas de outras transações"]
    for c in numericas:u[c]=u[c].map(_numero)
    u["_frete_dba"]=u["taxas fba"]
    manuseio=u["descrição"].eq("Tarifa de manuseio com base no peso Delivery by Amazon")
    u.loc[manuseio,"_frete_dba"]=u.loc[manuseio,"taxas de outras transações"]
    frete=u.loc[u["_frete_dba"].ne(0),["id do pedido","_frete_dba"]].drop_duplicates("id do pedido",keep="first")
    pedidos=u[u["tipo"].eq("Pedido")].drop(columns=["_frete_dba"]).copy().merge(frete,on="id do pedido",how="left",validate="many_to_one")
    tarifas=pedidos.groupby(["id do pedido","sku"],as_index=False,sort=False).agg(
        quantidade=("quantidade","sum"),vendas_produto=("vendas do produto","sum"),
        descontos=("descontos promocionais","sum"),impostos=("imposto de vendas coletados","sum"),
        tarifas_venda=("tarifas de venda","sum"),frete=("_frete_dba","sum"),
        frete_cliente=("créditos de remessa","sum"))
    medias=tarifas[tarifas["quantidade"].eq(1)].drop_duplicates("sku",keep="first").copy()
    medias["pct_comissao"]=medias["tarifas_venda"]/medias["vendas_produto"].replace(0,pd.NA)
    medias["pct_frete"]=medias["frete"]/medias["vendas_produto"].replace(0,pd.NA)
    return tarifas,medias[["sku","pct_comissao","pct_frete"]]

def _preparar_receber(df):
    r=df[df["Tipo de transação"].ne("Reembolso")].copy()
    r["_tarifa"]=r["Tarifas da Amazon"].map(_numero);r["_outros"]=r["Outros"].map(_numero)
    r["_frete_total"]=r["_tarifa"].where(r["Detalhes do produto"].eq("Faturamento"),0)
    r["_tarifas_amazon"]=r["_tarifa"].where(r["Detalhes do produto"].ne("Faturamento"),0)
    g=r.groupby("ID do pedido",as_index=False,sort=False).agg(frete_total=("_frete_total","sum"),tarifas_amazon=("_tarifas_amazon","sum"),frete_cliente=("_outros","sum"),contagem=("ID do pedido","size"))
    g["frete_dba"]=(g["frete_total"]+g["frete_cliente"]).where(g["contagem"].ne(1),0)
    return g

def transformar_amazon(vendas,unificado,receber,custos,aliquota_imposto=.14,marketplace_venda="Box fan"):
    _validar(vendas,BASE_OBRIGATORIAS,"vendas");_validar(unificado,UNIFIED_OBRIGATORIAS,"unificado");_validar(receber,RECEBER_OBRIGATORIAS,"a receber")
    df=vendas.copy()
    for c in ["item-price","shipping-tax","item-tax","item-promotion-discount","ship-promotion-discount"]:df[c]=df[c].map(_numero)
    df["quantity"]=pd.to_numeric(df["quantity"].replace("",0),errors="coerce").fillna(0).astype(int)
    df["Data"]=pd.to_datetime(df["purchase-date"].astype(str).str.split("T").str[0],errors="coerce")
    vendas_status={"Pending - Waiting for Pick Up","Shipped - Picked Up","Shipped - Out for Delivery","Shipped - Delivered to Buyer","Shipped - Undeliverable","Shipped"}
    devolucoes={"Shipped - Returning to Seller","Shipped - Returned to Seller","Shipped - Lost in Transit"}
    df["Status"]=pd.Series(0,index=df.index,dtype=object)
    df.loc[df["order-status"].isin(vendas_status),"Status"]="Venda"
    df.loc[df["order-status"].eq("Pending"),"Status"]="Não pago"
    df.loc[df["order-status"].eq("Cancelled"),"Status"]="Cancelado"
    df.loc[df["order-status"].isin(devolucoes),"Status"]="Devolução"
    imposto_itens=df["item-tax"]+df["shipping-tax"]
    df["Faturamento"]=(df["item-price"]+imposto_itens).where(df["Status"].eq("Venda"),0)
    df["SKU.1"]=df["sku"].map(_normalizar_sku)
    df["Preço Unitário"]=(df["Faturamento"]/df["quantity"].replace(0,pd.NA)).where(df["Status"].eq("Venda"),0)
    tarifas,medias=_preparar_unificado(unificado); ar=_preparar_receber(receber)
    df=df.merge(tarifas[["id do pedido","sku","tarifas_venda","frete","frete_cliente"]],left_on=["amazon-order-id","sku"],right_on=["id do pedido","sku"],how="left",validate="many_to_one")
    df=df.merge(ar[["ID do pedido","tarifas_amazon","frete_dba"]],left_on="amazon-order-id",right_on="ID do pedido",how="left",validate="many_to_one")
    df=df.merge(medias,on="sku",how="left",validate="many_to_one")
    cont=df.groupby("amazon-order-id")["amazon-order-id"].transform("size")
    comissao=df["tarifas_venda"].copy()
    sem_unificado=comissao.isna();comissao=comissao.where(~sem_unificado,df["tarifas_amazon"])
    sem_receber=sem_unificado & df["tarifas_amazon"].isna();comissao=comissao.where(~sem_receber,df["pct_comissao"]*df["Faturamento"])
    df["Comissão"]=pd.to_numeric(comissao.where(df["Status"].eq("Venda"),0),errors="coerce")
    venda=df["Status"].eq("Venda")
    sem_comissao=df["Comissão"].isna() | df["Comissão"].eq(0)
    df.loc[venda & sem_comissao,"Comissão"]=-.13*df.loc[venda & sem_comissao,"Faturamento"]
    desconto_frete=-df["ship-promotion-discount"]
    frete=df["frete"].copy();sem_unificado=frete.isna()
    fallback=df["frete_dba"].copy();sem_receber=fallback.isna()
    fallback=fallback.where(~sem_receber,df["pct_frete"]*df["Faturamento"])
    nao_fba=~df["fulfillment-channel"].eq("Amazon")
    fallback=fallback.where(~(nao_fba & ~sem_receber),fallback/cont)
    frete=frete.where(~sem_unificado,fallback)
    frete_unificado_fba=df["frete"]+df["frete_cliente"]+desconto_frete
    frete=frete.where(~(~sem_unificado & df["fulfillment-channel"].eq("Amazon")),frete_unificado_fba)
    frete=frete.where(~(~sem_unificado & nao_fba),df["frete"]/cont)
    df["Frete"]=pd.to_numeric(frete.where(df["Status"].eq("Venda"),0),errors="coerce")
    sem_frete=df["Frete"].isna() | df["Frete"].eq(0)
    frete_premissa=df["Faturamento"].map(lambda valor:-6.75 if valor < 78 else -19.50)
    df.loc[venda & sem_frete,"Frete"]=frete_premissa[venda & sem_frete]
    df["Rebate"]=0.0;df["Cancelamento"]=0.0;df["item-promotion-discount"]*=-1
    df["Liquido"]=(df["Faturamento"]+df["Frete"]+df["Comissão"]+df["item-promotion-discount"]).where(df["Status"].eq("Venda"),0)
    df["Antecipa"]=0.0;df["Imposto"]=-float(aliquota_imposto)*(df["Faturamento"]+df["item-promotion-discount"])
    custo=custos[["SKU","Valor total"]].copy();custo["SKU"]=custo["SKU"].astype(str).str.strip()
    df=df.merge(custo,left_on="SKU.1",right_on="SKU",how="left",validate="many_to_one")
    df["Custo do produto"]=(-df["Valor total"]*df["quantity"]).where(df["Status"].eq("Venda"),0)
    df["Margin R$"]=df["Liquido"]+df["Imposto"]+df["Custo do produto"]
    df["Margin %"]=(df["Margin R$"]/df["Faturamento"].replace(0,pd.NA)).where(df["Status"].eq("Venda"),0)
    df["Marketplace"]="Amazon";df["Marketplace venda"]=marketplace_venda
    return df[COLUNAS_SAIDA].reset_index(drop=True)

def processar_amazon(caminho_vendas,caminho_unificado,caminho_receber,caminho_custos,caminho_saida,aliquota_imposto=.14,marketplace_venda="Box fan"):
    resultado=transformar_amazon(_ler_vendas(caminho_vendas),_ler_unificado(caminho_unificado),_ler_receber(caminho_receber),_carregar_custos(caminho_custos),aliquota_imposto,marketplace_venda)
    saida=Path(caminho_saida).resolve();resultado.to_excel(saida,sheet_name="DB",index=False,engine="openpyxl")
    wb=load_workbook(saida);ws=wb["DB"]
    for c in ws[1]:c.font=Font(bold=True)
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions;wb.save(saida)
    return str(saida),{"linhas":len(resultado),"semCusto":int(resultado["Custo do produto"].isna().sum()),"semTarifa":int(resultado["Comissão"].isna().sum()),"canceladas":int(resultado["Status"].eq("Cancelado").sum()),"devolucoes":int(resultado["Status"].eq("Devolução").sum())}
