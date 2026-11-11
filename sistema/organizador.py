# -*- coding: utf-8 -*-
"""Organizador: _entrada\\[Cliente] -> [raiz]\\[DD-MM]\\[Cliente]\\vendas_[mk]_[DD-MM].ext

Regras: nunca apaga nada; não reconhecido fica na _entrada e vai para o log;
pode rodar várias vezes sem duplicar (conteúdo idêntico já organizado é
movido para _entrada\\_ja_processados; versão diferente substitui e a anterior
vai para _versoes_anteriores).
"""
import csv
import datetime as dt
import hashlib
import io
import json
import os
import re
import shutil
import sys
import time
import zipfile

import comum
import compactados

TEMPORARIOS = (".crdownload", ".tmp", ".part", ".download")
IGNORAR = ("desktop.ini", "thumbs.db")
TABELAS_TEXTO = (".csv", ".txt", ".tsv")
SEGUNDOS_ESTAVEL = 20


# ---------------------------------------------------------------- utilidades
def sha256(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def destino_livre(caminho):
    if not os.path.exists(caminho):
        return caminho
    base, ext = os.path.splitext(caminho)
    n = 2
    while os.path.exists("%s (%d)%s" % (base, n, ext)):
        n += 1
    return "%s (%d)%s" % (base, n, ext)


def mover(origem, destino):
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    destino = destino_livre(destino)
    shutil.move(origem, destino)
    return destino


def texto_xlsx(caminho, limite=200000):
    """Textos de um .xlsx (sharedStrings + primeira planilha), sem acento e minúsculos."""
    partes = []
    try:
        with zipfile.ZipFile(caminho) as z:
            nomes = z.namelist()
            for n in ["xl/sharedStrings.xml"] + sorted(x for x in nomes if x.startswith("xl/worksheets/sheet"))[:1]:
                if n in nomes:
                    partes.append(z.read(n)[:limite].decode("utf-8", "ignore"))
    except (zipfile.BadZipFile, OSError):
        return ""
    txt = re.sub(r"<[^>]+>", " ", " ".join(partes))
    return comum.sem_acento(txt)


def ler_texto(caminho, limite=200000):
    with open(caminho, "rb") as f:
        bruto = f.read(limite)
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return bruto.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return bruto.decode("latin-1", "ignore"), "latin-1"


def conteudo_para_identificar(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".xlsx":
        return texto_xlsx(caminho)
    if ext in TABELAS_TEXTO:
        return comum.sem_acento(ler_texto(caminho, 50000)[0])
    if ext == ".zip":
        try:
            with zipfile.ZipFile(caminho) as z:
                return comum.sem_acento(" ".join(z.namelist()))
        except (zipfile.BadZipFile, OSError):
            return ""
    return ""


# ---------------------------------------------------------------- identificação
PASTA_TRATADOS_RECEBIDOS = "_tratados_recebidos"


def eh_tratado(caminho):
    """Planilha já tratada (saída do tratador: aba DB com Marketplace venda e Gross margin).
    Não é relatório bruto: se fosse organizada, substituiria o bruto do marketplace."""
    if os.path.splitext(caminho)[1].lower() not in (".xlsx", ".xlsm"):
        return False
    texto = texto_xlsx(caminho, 60000)
    return "marketplace venda" in texto and ("gross margen" in texto or "gross margin" in texto)


def identificar(cfg, cliente, caminho):
    """Retorna (mk_id, motivo) ou (None, motivo)."""
    nome = os.path.basename(caminho)
    nome_n = comum.sem_acento(nome)
    ext = os.path.splitext(nome)[1].lower()
    todos = cfg["marketplaces"]
    do_cliente = [m for m in cliente.get("marketplaces", []) if m in todos]
    # relatórios auxiliares (ex.: Amazon unificado / a receber) só valem se o cliente usa o marketplace "pai"
    mks = {k: v for k, v in todos.items() if not v.get("oculto") or v.get("pai") in do_cliente}

    # 0) já está no padrão vendas_<mk>_
    m = re.match(r"vendas_([a-z0-9-]+)_", nome_n)
    if m and m.group(1) in todos:
        return m.group(1), "nome no padrão"

    # 0b) relatórios auxiliares têm assinatura forte no conteúdo: confere antes do nome
    auxiliares = {k: v for k, v in mks.items() if v.get("oculto")}
    if auxiliares and ext in TABELAS_TEXTO:
        texto = conteudo_para_identificar(caminho)
        for k, v in auxiliares.items():
            termos = v.get("palavras_conteudo", [])
            acertos = sum(1 for p in termos if comum.sem_acento(p) in texto)
            if termos and acertos >= max(3, len(termos) - 1):
                return k, "conteúdo (%d termos)" % acertos
            if any(re.search(p, nome_n, re.I) for p in v.get("padroes_nome", [])):
                return k, "nome do arquivo"
    mks = {k: v for k, v in mks.items() if not v.get("oculto")}

    # 1) compactado (.zip, .rar do WinRAR, .7z): relatório de planilha dividido em partes (ex.: Shopee com muitas
    #    vendas: Order.all..._part_1_of_2.xlsx) pelo nome do arquivo ou das partes; senão é a Magalu
    if ext in compactados.EXTS_COMPACTADOS or (ext not in (".xlsx",) and compactados.tipo(caminho) in ("rar", "7z")):
        nomes = [nome_n]
        if compactados.tipo(caminho) == "zip":
            try:
                import zipfile
                with zipfile.ZipFile(caminho) as z:
                    nomes += [comum.sem_acento(os.path.basename(n)) for n in z.namelist() if n.lower().endswith((".xlsx", ".csv"))]
            except Exception:
                pass
        partes = [k for k, v in mks.items() if v["extensao"] == ".xlsx"
                  and any(re.search(pd_, n, re.I) for pd_ in v.get("padroes_nome", []) for n in nomes)]
        if len(partes) == 1:
            return partes[0], "compactado com o relatório em partes"
        zips = [k for k, v in mks.items() if v["extensao"] == ".zip"]
        if len(zips) == 1:
            return zips[0], "arquivo .zip"

    if ext not in (".xlsx",) + TABELAS_TEXTO:
        return None, "extensão não suportada (%s)" % (ext or "sem extensão")

    # 2) nome do arquivo
    por_nome = [k for k, v in mks.items() if v["extensao"] != ".zip"
                and any(re.search(p, nome_n, re.I) for p in v.get("padroes_nome", []))]
    if len(por_nome) == 1:
        return por_nome[0], "nome do arquivo"

    # 3) conteúdo
    texto = conteudo_para_identificar(caminho)
    if texto:
        pontos = {}
        for k, v in mks.items():
            if v["extensao"] == ".zip":
                continue
            pontos[k] = sum(1 for p in v.get("palavras_conteudo", []) if comum.sem_acento(p) in texto)
        candidatos = por_nome or list(pontos)
        melhor = max(candidatos, key=lambda k: (pontos.get(k, 0), k in do_cliente))
        empatados = [k for k in candidatos if pontos.get(k, 0) == pontos.get(melhor, 0)]
        if pontos.get(melhor, 0) >= 2 and len(empatados) == 1:
            return melhor, "conteúdo (%d termos)" % pontos[melhor]
        empatados_cliente = [k for k in empatados if k in do_cliente]
        if pontos.get(melhor, 0) >= 2 and len(empatados_cliente) == 1:
            return empatados_cliente[0], "conteúdo + cadastro do cliente"

    # 4) cliente com um único marketplace de planilha cadastrado (do_cliente nunca tem auxiliares)
    planilhas = [k for k in do_cliente if mks[k]["extensao"] != ".zip"]
    if len(planilhas) == 1:
        return planilhas[0], "único marketplace de planilha do cliente"

    return None, "não foi possível identificar o marketplace"


def periodo_do_nome(nome):
    """Tenta extrair duas datas AAAAMMDD/AAAA-MM-DD do nome do arquivo."""
    achados = re.findall(r"(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])", nome)
    datas = []
    for a, m, d in achados:
        try:
            datas.append(dt.date(int(a), int(m), int(d)))
        except ValueError:
            pass
    if len(datas) >= 2:
        return min(datas), max(datas)
    return None


# ---------------------------------------------------------------- conversão
def converter_para_xlsx(origem, destino):
    """CSV/TXT (Amazon) -> XLSX. Retorna mensagem de erro ou None."""
    try:
        from openpyxl import Workbook
    except ImportError:
        return "openpyxl não instalado"
    texto, _ = ler_texto(origem, 1 << 30)
    amostra = texto[:20000]
    primeira = amostra.splitlines()[0] if amostra else ""
    if "\t" in primeira:
        delim = "\t"
    else:
        try:
            delim = csv.Sniffer().sniff(amostra, delimiters=";,\t|").delimiter
        except csv.Error:
            delim = ","
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Relatorio")
    for linha in csv.reader(io.StringIO(texto), delimiter=delim):
        ws.append(linha)
    tmp = destino + ".parcial"
    wb.save(tmp)
    os.replace(tmp, destino)
    return None


# ---------------------------------------------------------------- processamento
def arquivo_marcador(final, cliente_nome, hoje):
    """Metadados (hash do original etc.) ficam no _sistema, fora da pasta do cliente."""
    return os.path.join(comum.PASTA_SISTEMA, "dados", "marcadores", comum.rotulo_dia(hoje) + "_" + str(hoje.year),
                        cliente_nome, os.path.basename(final) + ".json")


def processar_arquivo(cfg, cliente, caminho, mk_id, motivo, hoje=None):
    hoje = hoje or dt.date.today()
    mk = cfg["marketplaces"][mk_id]
    ext_orig = os.path.splitext(caminho)[1].lower()
    pasta = comum.pasta_dia(cfg, comum.pasta_de(cliente), hoje)
    os.makedirs(pasta, exist_ok=True)
    final = os.path.join(pasta, comum.nome_arquivo(mk_id, cfg, hoje))
    nome_orig = os.path.basename(caminho)
    per = periodo_do_nome(nome_orig)
    evento = {"tipo": "organizado", "cliente": cliente["id"], "mk": mk_id, "origem": nome_orig,
              "dia": comum.rotulo_dia(hoje), "data": hoje.isoformat(), "identificacao": motivo}
    if per:
        evento["inicio"], evento["fim"] = comum.fmt(per[0]), comum.fmt(per[1])

    precisa_converter = mk["extensao"] == ".xlsx" and ext_orig in TABELAS_TEXTO
    # relatório de planilha que veio compactado em partes (ex.: Shopee com muitas vendas): junta num .xlsx só
    precisa_juntar = mk["extensao"] == ".xlsx" and ext_orig in compactados.EXTS_COMPACTADOS and bool(compactados.tipo(caminho))
    if mk["extensao"] == ".xlsx" and ext_orig not in (".xlsx",) + TABELAS_TEXTO and not precisa_juntar:
        return False, "formato %s não serve para %s" % (ext_orig, mk["nome"])
    if mk["extensao"] == ".zip" and not compactados.tipo(caminho):
        return False, "%s deveria ser um arquivo compactado (.zip ou .rar)" % mk["nome"]
    # .rar, compactado dentro de compactado ou CSVs em subpasta: vira o .zip com os 2 CSVs na raiz
    precisa_normalizar = mk["extensao"] == ".zip" and not compactados.ja_plano(caminho)

    # duplicado exato do que já está organizado?
    hash_novo = sha256(caminho)
    marcador = arquivo_marcador(final, comum.pasta_de(cliente), hoje)
    hash_existente = None
    if os.path.exists(final):
        try:
            with open(marcador, "r", encoding="utf-8") as f:
                hash_existente = json.load(f).get("sha256_original")
        except (OSError, ValueError):
            hash_existente = None if (precisa_converter or precisa_normalizar or precisa_juntar) else sha256(final)
    if hash_existente == hash_novo:
        dup = mover(caminho, os.path.join(comum.pasta_entrada(cfg), "_ja_processados", comum.pasta_de(cliente),
                                          dt.datetime.now().strftime("%Y%m%d-%H%M%S_") + nome_orig))
        comum.log("[%s] %s: igual ao já organizado (%s) -> movido para %s" % (cliente["nome"], nome_orig, os.path.basename(final), dup))
        return True, "duplicado"

    # versão anterior diferente -> guarda
    if os.path.exists(final):
        ant = os.path.join(pasta, "_versoes_anteriores",
                           "%s_%s%s" % (os.path.splitext(os.path.basename(final))[0],
                                        dt.datetime.fromtimestamp(os.path.getmtime(final)).strftime("%H%M%S"), mk["extensao"]))
        mover(final, ant)
        if os.path.exists(marcador):
            mover(marcador, arquivo_marcador(ant, comum.pasta_de(cliente), hoje))
        comum.log("[%s] versão anterior de %s guardada em %s" % (cliente["nome"], os.path.basename(final), ant))

    if precisa_juntar:
        try:
            partes = compactados.juntar_planilhas(caminho, final)
        except ValueError as e:
            return False, str(e)
        orig_dest = mover(caminho, os.path.join(pasta, "_originais", os.path.splitext(os.path.basename(final))[0] + ext_orig))
        evento["original_guardado"] = orig_dest
        evento["partes"] = ["%s (%d linhas)" % p for p in partes]
        comum.log("[%s] %s: %d parte(s) juntadas: %s" % (cliente["nome"], nome_orig, len(partes), ", ".join(evento["partes"])))
    elif precisa_converter:
        erro = converter_para_xlsx(caminho, final)
        if erro:
            return False, "não consegui converter %s para .xlsx: %s" % (nome_orig, erro)
        orig_dest = mover(caminho, os.path.join(pasta, "_originais", os.path.splitext(os.path.basename(final))[0] + ext_orig))
        evento["original_guardado"] = orig_dest
    elif precisa_normalizar:
        try:
            dentro = compactados.para_zip_plano(caminho, final)
        except ValueError as e:
            return False, str(e)
        orig_dest = mover(caminho, os.path.join(pasta, "_originais", os.path.splitext(os.path.basename(final))[0] + ext_orig))
        evento["original_guardado"] = orig_dest
        evento["extraido"] = dentro
        comum.log("[%s] %s extraído: %s" % (cliente["nome"], nome_orig, ", ".join(dentro)))
    else:
        shutil.move(caminho, final)

    os.makedirs(os.path.dirname(marcador), exist_ok=True)
    with open(marcador, "w", encoding="utf-8") as f:
        json.dump({"sha256_original": hash_novo, "origem": nome_orig, "identificacao": motivo,
                   "organizado_em": dt.datetime.now().isoformat(timespec="seconds")}, f, ensure_ascii=False, indent=1)
    evento["destino"] = final
    comum.registrar_evento(evento)
    comum.log("[%s] %s -> %s (%s)" % (cliente["nome"], nome_orig, final, motivo))
    _tratar_depois(cfg, cliente, mk_id, hoje)
    return True, final


def _tratar_depois(cfg, cliente, mk_id, hoje):
    """Gera a versão tratada logo depois de organizar (falha no tratamento nunca derruba o organizador)."""
    if not cfg.get("tratamento", {}).get("automatico", True):
        return
    try:
        import tratador
        alvo = tratador.pai(cfg, mk_id)
        if alvo in cliente.get("marketplaces", []):
            tratador.tratar(cfg, cliente, alvo, hoje)
    except Exception as e:
        comum.log("[%s] tratamento não executado: %s" % (cliente["nome"], e))


def pastas_entrada(cfg, cliente):
    """_entrada\\Cliente\\Conta. Se o cliente tem uma conta só, arquivos soltos em _entrada\\Cliente
    (pasta de downloads antiga) também valem para ela."""
    pastas = [comum.pasta_entrada(cfg, comum.pasta_de(cliente))]
    if cliente.get("cliente_nome"):
        raiz_cliente = comum.pasta_entrada(cfg, cliente["cliente_nome"])
        ainda_baixa_la = False
        if not cliente.get("unica_conta") and cliente.get("perfil"):
            try:
                ainda_baixa_la = comum.mesma_pasta(comum.pasta_downloads_perfil(cfg, cliente["perfil"]).get("pasta"), raiz_cliente)
            except Exception:
                pass
        if cliente.get("unica_conta") or ainda_baixa_la:
            pastas.append(raiz_cliente)
    return pastas


def arquivos_entrada(cfg, cliente):
    saida = []
    for pasta in pastas_entrada(cfg, cliente):
        if not os.path.isdir(pasta):
            continue
        for nome in sorted(os.listdir(pasta)):
            caminho = os.path.join(pasta, nome)
            if not os.path.isfile(caminho) or nome.lower() in IGNORAR or nome.startswith("~$"):
                continue
            saida.append(caminho)
    return saida


def estado_pendencias():
    arq = os.path.join(comum.PASTA_SISTEMA, "dados", "nao_reconhecidos.json")
    try:
        with open(arq, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def salvar_pendencias(dados):
    arq = os.path.join(comum.PASTA_SISTEMA, "dados", "nao_reconhecidos.json")
    os.makedirs(os.path.dirname(arq), exist_ok=True)
    with open(arq, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)


class Trava:
    def __init__(self):
        self.arq = os.path.join(comum.PASTA_SISTEMA, "dados", "organizador.lock")

    def __enter__(self):
        os.makedirs(os.path.dirname(self.arq), exist_ok=True)
        for _ in range(60):
            try:
                fd = os.open(self.arq, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, str(os.getpid()).encode())
                os.close(fd)
                return self
            except FileExistsError:
                if time.time() - os.path.getmtime(self.arq) > 600:  # trava esquecida
                    os.replace(self.arq, self.arq + ".antiga")
                    continue
                time.sleep(1)
        raise RuntimeError("Outro organizador está rodando.")

    def __exit__(self, *a):
        try:
            os.replace(self.arq, self.arq + ".liberada")
        except OSError:
            pass


def executar(cfg=None, hoje=None, forcar_estavel=False):
    cfg = cfg or comum.carregar_config()
    hoje = hoje or dt.date.today()
    resumo = {"organizados": 0, "duplicados": 0, "nao_reconhecidos": 0, "aguardando": 0, "erros": 0, "detalhes": []}
    pendencias = {}
    anteriores = estado_pendencias()
    with Trava():
        comum.criar_pastas_base(cfg)
        for cliente in comum.contas(cfg):
            comum.criar_pastas_cliente(cfg, cliente)
            for caminho in arquivos_entrada(cfg, cliente):
                nome = os.path.basename(caminho)
                if nome.lower().endswith(TEMPORARIOS) or (
                        not forcar_estavel and time.time() - os.path.getmtime(caminho) < SEGUNDOS_ESTAVEL):
                    resumo["aguardando"] += 1
                    continue
                if eh_tratado(caminho):
                    # nunca entra no lugar de um relatório bruto; guardado à parte (nada é apagado)
                    dest = mover(caminho, os.path.join(comum.pasta_entrada(cfg), PASTA_TRATADOS_RECEBIDOS,
                                                       comum.pasta_de(cliente), nome))
                    resumo.setdefault("tratados_recebidos", 0)
                    resumo["tratados_recebidos"] += 1
                    comum.log("[%s] %s é uma planilha já tratada, não um relatório bruto: guardada em %s"
                              % (cliente["nome"], nome, os.path.dirname(dest)))
                    continue
                mk_id, motivo = identificar(cfg, cliente, caminho)
                if not mk_id:
                    resumo["nao_reconhecidos"] += 1
                    pendencias.setdefault(cliente["id"], []).append({"arquivo": nome, "motivo": motivo})
                    if not any(p["arquivo"] == nome for p in anteriores.get(cliente["id"], [])):
                        comum.log("[%s] NÃO RECONHECIDO: %s (%s) — mantido na _entrada" % (cliente["nome"], nome, motivo))
                    continue
                try:
                    ok, info = processar_arquivo(cfg, cliente, caminho, mk_id, motivo, hoje)
                except Exception as e:  # nunca derruba o lote inteiro
                    ok, info = False, "erro inesperado: %s" % e
                if ok and info == "duplicado":
                    resumo["duplicados"] += 1
                elif ok:
                    resumo["organizados"] += 1
                    resumo.setdefault("clientes_organizados", [])
                    if cliente["cliente_id"] not in resumo["clientes_organizados"]:
                        resumo["clientes_organizados"].append(cliente["cliente_id"])
                    resumo["detalhes"].append("%s: %s -> %s" % (cliente["nome"], nome, os.path.basename(info)))
                else:
                    resumo["erros"] += 1
                    pendencias.setdefault(cliente["id"], []).append({"arquivo": nome, "motivo": info})
                    comum.log("[%s] ERRO com %s: %s — mantido na _entrada" % (cliente["nome"], nome, info))
        salvar_pendencias(pendencias)
    if resumo.get("clientes_organizados"):
        # relatório consolidado do mês (pacote final) refeito com os tratados novos
        try:
            import consolidado
            consolidado.automatico(cfg, resumo["clientes_organizados"], "depois do organizador")
        except Exception as e:
            comum.log("CONSOLIDADO: falhou depois do organizador: %s" % e)
    comum.log("Execução: %(organizados)d organizado(s), %(duplicados)d duplicado(s), %(nao_reconhecidos)d não reconhecido(s), "
              "%(erros)d erro(s), %(aguardando)d aguardando download" % resumo)
    return resumo


def classificar_manual(cfg, cliente_id, arquivo, mk_id, hoje=None):
    cliente = comum.conta(cfg, cliente_id)
    if not cliente or mk_id not in cfg["marketplaces"]:
        return False, "cliente ou marketplace inválido"
    caminho = next((os.path.join(p, os.path.basename(arquivo)) for p in pastas_entrada(cfg, cliente)
                    if os.path.isfile(os.path.join(p, os.path.basename(arquivo)))), None)
    if not caminho:
        return False, "arquivo não está mais na _entrada"
    with Trava():
        ok, info = processar_arquivo(cfg, cliente, caminho, mk_id, "classificado manualmente no painel", hoje)
        pend = estado_pendencias()
        pend[cliente_id] = [p for p in pend.get(cliente_id, []) if p["arquivo"] != os.path.basename(arquivo)]
        salvar_pendencias(pend)
    return ok, info


if __name__ == "__main__":
    r = executar()
    print(json.dumps(r, ensure_ascii=False, indent=1))
