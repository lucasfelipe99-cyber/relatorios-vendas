# -*- coding: utf-8 -*-
"""Arquivos compactados da Magalu (.zip, .rar do WinRAR, .7z).

O tratador da Magalu lê um .zip com os dois CSVs (pedidos e pacotes) soltos na raiz.
Aqui qualquer compactado — inclusive .rar, compactado dentro de compactado e arquivos
dentro de subpastas — vira esse .zip "plano". O original nunca é apagado.

Para .rar/.7z usa o que tiver no Windows: WinRAR (UnRAR.exe/WinRAR.exe), 7-Zip ou o
tar.exe do próprio Windows.
"""
import glob
import os
import shutil
import subprocess
import tempfile
import zipfile

EXTS_COMPACTADOS = (".zip", ".rar", ".7z")
EXTS_DADOS = (".csv", ".txt", ".xlsx", ".xls")
_SEM_JANELA = 0x08000000 if os.name == "nt" else 0


def tipo(caminho):
    """Formato real pelo conteúdo (o nome pode enganar: .zip que é rar e vice-versa)."""
    try:
        with open(caminho, "rb") as f:
            cab = f.read(8)
    except OSError:
        return None
    if cab.startswith(b"PK\x03\x04") or cab.startswith(b"PK\x05\x06"):
        return "zip"
    if cab.startswith(b"Rar!\x1a\x07"):
        return "rar"
    if cab.startswith(b"7z\xbc\xaf\x27\x1c"):
        return "7z"
    return None


def _programas():
    pf = [os.environ.get(v) for v in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432")]
    pf = [p for p in pf if p] or [r"C:\Program Files", r"C:\Program Files (x86)"]
    cands = []
    for base in pf:
        cands += [("unrar", os.path.join(base, "WinRAR", "UnRAR.exe")),
                  ("winrar", os.path.join(base, "WinRAR", "WinRAR.exe")),
                  ("7z", os.path.join(base, "7-Zip", "7z.exe"))]
    for nome in ("unrar", "7z", "7za"):
        p = shutil.which(nome)
        if p:
            cands.append(("unrar" if nome == "unrar" else "7z", p))
    windir = os.environ.get("SystemRoot", r"C:\Windows")
    cands.append(("tar", os.path.join(windir, "System32", "tar.exe")))
    p = shutil.which("bsdtar") or shutil.which("tar")
    if p:
        cands.append(("tar", p))
    vistos, saida = set(), []
    for t, p in cands:
        if p and os.path.isfile(p) and p not in vistos:
            vistos.add(p)
            saida.append((t, p))
    return saida


def _extrair_externo(arquivo, destino):
    erros = []
    for t, exe in _programas():
        if t == "unrar":
            cmd = [exe, "x", "-y", "-o+", "-inul", arquivo, destino + os.sep]
        elif t == "winrar":
            cmd = [exe, "x", "-ibck", "-y", "-o+", arquivo, destino + os.sep]
        elif t == "7z":
            cmd = [exe, "x", "-y", "-o" + destino, arquivo]
        else:
            cmd = [exe, "-xf", arquivo, "-C", destino]
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300, creationflags=_SEM_JANELA)
            if r.returncode == 0 and any(os.scandir(destino)):
                return
            erros.append("%s: código %s" % (os.path.basename(exe), r.returncode))
        except (OSError, subprocess.SubprocessError) as e:
            erros.append("%s: %s" % (os.path.basename(exe), e))
    if not erros:
        raise ValueError("arquivo .rar/.7z e nenhum programa para abrir foi encontrado. Instale o WinRAR ou o 7-Zip.")
    raise ValueError("não consegui extrair o arquivo compactado (%s)." % "; ".join(erros))


def extrair_tudo(arquivo, destino, nivel=0):
    """Extrai o compactado em destino, abrindo também compactados de dentro (até 3 níveis)."""
    t = tipo(arquivo)
    if t == "zip":
        with zipfile.ZipFile(arquivo) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                alvo = os.path.normpath(os.path.join(destino, info.filename))
                if not alvo.startswith(os.path.normpath(destino) + os.sep):
                    continue   # caminho perigoso dentro do zip: ignora
                os.makedirs(os.path.dirname(alvo), exist_ok=True)
                with z.open(info) as o, open(alvo, "wb") as d:
                    shutil.copyfileobj(o, d)
    elif t in ("rar", "7z"):
        _extrair_externo(arquivo, destino)
    else:
        raise ValueError("%s não é um arquivo compactado (.zip/.rar/.7z)." % os.path.basename(arquivo))
    if nivel < 3:
        for raiz, _, nomes in os.walk(destino):
            for n in nomes:
                p = os.path.join(raiz, n)
                if p != arquivo and tipo(p):
                    sub = p + "_extraido"
                    os.makedirs(sub, exist_ok=True)
                    try:
                        extrair_tudo(p, sub, nivel + 1)
                    except ValueError:
                        pass


def arquivos_de_dados(pasta):
    return [p for p in glob.glob(os.path.join(pasta, "**", "*"), recursive=True)
            if os.path.isfile(p) and p.lower().endswith(EXTS_DADOS)]


def ja_plano(caminho):
    """True se já é .zip com um CSV de pedidos e um de pacotes na raiz (formato que o tratador lê)."""
    if tipo(caminho) != "zip":
        return False
    try:
        with zipfile.ZipFile(caminho) as z:
            nomes = [n for n in z.namelist() if "/" not in n.strip("/") and n.lower().endswith(".csv")]
    except (zipfile.BadZipFile, OSError):
        return False
    return (sum("pedidos" in n.lower() for n in nomes) == 1 and sum("pacotes" in n.lower() for n in nomes) == 1)


def para_zip_plano(arquivo, destino_zip):
    """Gera destino_zip com os CSVs de pedidos e pacotes na raiz. Devolve a lista de nomes colocados."""
    temp = tempfile.mkdtemp(prefix="magalu_")
    try:
        extrair_tudo(arquivo, temp)
        dados = arquivos_de_dados(temp)
        pedidos = [p for p in dados if "pedidos" in os.path.basename(p).lower() and p.lower().endswith(".csv")]
        pacotes = [p for p in dados if "pacotes" in os.path.basename(p).lower() and p.lower().endswith(".csv")]
        if len(pedidos) != 1 or len(pacotes) != 1:
            achados = ", ".join(sorted(os.path.basename(p) for p in dados)) or "nenhum arquivo de dados"
            raise ValueError("o compactado da Magalu precisa ter um CSV de pedidos e um CSV de pacotes. Encontrei: %s." % achados)
        parcial = destino_zip + ".parcial"
        with zipfile.ZipFile(parcial, "w", zipfile.ZIP_DEFLATED) as z:
            for p in (pedidos[0], pacotes[0]):
                z.write(p, os.path.basename(p))
        os.replace(parcial, destino_zip)
        return [os.path.basename(pedidos[0]), os.path.basename(pacotes[0])]
    finally:
        shutil.rmtree(temp, ignore_errors=True)


def _ordem_natural(p):
    import re
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", os.path.basename(p))]


def juntar_planilhas(arquivo, destino_xlsx):
    """Compactado com o relatório dividido em partes (ex.: Shopee com muitas vendas: Order.all..._part_1_of_2.xlsx,
    _part_2_of_2.xlsx): junta todas as partes num .xlsx só (cabeçalho uma vez, linhas de todas as partes, na ordem).
    As partes precisam ter as mesmas colunas. Devolve [(nome da parte, linhas)]."""
    import csv
    from openpyxl import Workbook, load_workbook
    temp = tempfile.mkdtemp(prefix="juntar_")
    try:
        extrair_tudo(arquivo, temp)
        partes = sorted([p for p in arquivos_de_dados(temp) if p.lower().endswith((".xlsx", ".csv", ".txt"))], key=_ordem_natural)
        if not partes:
            raise ValueError("o compactado não tem planilha (.xlsx ou .csv) dentro.")
        wb_out = Workbook()
        ws_out = wb_out.active
        cab, resumo, titulo = None, [], []

        def linhas_de(p):
            if p.lower().endswith(".xlsx"):
                wb = load_workbook(p, data_only=True)     # sem read_only: alguns relatórios vêm com a dimensão errada
                ws = wb.worksheets[0]
                titulo.append(ws.title)
                for r in ws.iter_rows(values_only=True):
                    yield list(r)
                wb.close()
            else:
                with open(p, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
                    amostra = f.read(4096)
                    f.seek(0)
                    try:
                        dialeto = csv.Sniffer().sniff(amostra, delimiters=",;\t")
                    except csv.Error:
                        dialeto = csv.excel
                    for r in csv.reader(f, dialeto):
                        yield r
        for p in partes:
            it = linhas_de(p)
            try:
                c = next(it)
            except StopIteration:
                continue
            while c and c[-1] in (None, ""):
                c = c[:-1]
            if cab is None:
                cab = c
                ws_out.append(cab)
            elif [str(x or "").strip() for x in c] != [str(x or "").strip() for x in cab]:
                raise ValueError("as partes do compactado têm colunas diferentes (%s): não dá para juntar." % os.path.basename(p))
            n = 0
            for r in it:
                if not any(v not in (None, "") for v in r):
                    continue
                ws_out.append(list(r[:len(cab)]))
                n += 1
            resumo.append((os.path.basename(p), n))
        if cab is None:
            raise ValueError("as planilhas do compactado estão vazias.")
        if titulo:
            ws_out.title = titulo[0][:31]        # mesma aba das partes (ex.: "orders" da Shopee)
        tmp = destino_xlsx + ".novo.xlsx"
        wb_out.save(tmp)
        os.replace(tmp, destino_xlsx)
        return resumo
    finally:
        shutil.rmtree(temp, ignore_errors=True)
