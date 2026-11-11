# -*- coding: utf-8 -*-
"""Leva os dados deste computador para o sistema online (Render).

Uso:  python exportar_nuvem.py [--so-zip] [--url https://xxx.onrender.com] [--usuario admin] [--senha ...]
        [--estacao NOME_DESTE_PC] [--zip caminho.zip]

1. Monta um .zip com o que o sistema online precisa:
   - sistema/config.json, sistema/dados/** (custos, de-para, traduções, consolidados, histórico…), sistema/magalu-ids.csv
   - raiz/** (pastas dos dias, _lancamentos, _entrada), sem cópias de segurança (_anteriores) nem o próprio sistema.
2. Envia o .zip em partes para /api/admin/importar (com a senha do admin). Sem --url, só gera o .zip.
"""
import argparse
import base64
import datetime as dt
import hashlib
import json
import os
import socket
import sys
import urllib.request
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
PULAR_RAIZ = {"_sistema", "_exportados", "_to_delete", "_enviados"}
PULAR_DIR = {"_anteriores", "__pycache__", "backups", "logs", ".venv", "app", "_importar"}
PARTE = 8 * 1024 * 1024


def montar_zip(sistema, raiz, destino, estacao, sistema_win, raiz_win):
    n = 0
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.writestr("manifesto.json", json.dumps({"sistema": sistema_win, "raiz": raiz_win, "estacao": estacao,
                                                 "gerado": dt.datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False))
        for nome in ("config.json", "magalu-ids.csv"):
            p = os.path.join(sistema, nome)
            if os.path.isfile(p):
                z.write(p, "sistema/" + nome)
                n += 1
        for base, prefixo, pular_topo in ((os.path.join(sistema, "dados"), "sistema/dados", set()), (raiz, "raiz", PULAR_RAIZ)):
            for pasta, subs, arqs in os.walk(base):
                rel = os.path.relpath(pasta, base)
                if rel == ".":
                    subs[:] = [s for s in subs if s not in pular_topo and s not in PULAR_DIR]
                else:
                    subs[:] = [s for s in subs if s not in PULAR_DIR]
                for a in arqs:
                    if a in ("rede.json", "estacoes.json") or a.startswith("~$") or a.endswith((".tmp", ".crdownload")):
                        continue
                    p = os.path.join(pasta, a)
                    nome_zip = prefixo + "/" + ("" if rel == "." else rel.replace(os.sep, "/") + "/") + a
                    try:
                        z.write(p, nome_zip)
                        n += 1
                    except OSError as e:
                        print("pulei %s (%s)" % (p, e))
    return n


def enviar(arq, url, usuario, senha, estacao):
    token = hashlib.sha256(("relatorios-vendas:" + senha).encode("utf-8")).hexdigest()
    tam = os.path.getsize(arq)
    total = max(1, (tam + PARTE - 1) // PARTE)
    ident = dt.datetime.now().strftime("imp%Y%m%d%H%M%S")
    with open(arq, "rb") as f:
        for i in range(total):
            corpo = json.dumps({"id": ident, "parte": i, "total": total, "estacao_padrao": estacao,
                                "conteudo": base64.b64encode(f.read(PARTE)).decode("ascii")}).encode("utf-8")
            req = urllib.request.Request(url.rstrip("/") + "/api/admin/importar", data=corpo, method="POST",
                                         headers={"X-Painel": "1", "Content-Type": "application/json",
                                                  "Authorization": "Bearer " + token, "X-Usuario": usuario})
            with urllib.request.urlopen(req, timeout=900) as r:
                resp = json.loads(r.read() or b"{}")
            if not resp.get("ok"):
                raise SystemExit("Erro na parte %d: %s" % (i + 1, resp.get("erro")))
            print("parte %d/%d enviada" % (i + 1, total))
    return resp


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sistema", default=AQUI)
    ap.add_argument("--raiz", default=None)
    ap.add_argument("--zip", default=None)
    ap.add_argument("--url", default="")
    ap.add_argument("--usuario", default="admin")
    ap.add_argument("--senha", default="")
    ap.add_argument("--estacao", default=(socket.gethostname() or "").split(".")[0])
    ap.add_argument("--ligar-estacao", action="store_true",
                    help="depois de enviar, este computador passa a ser estação do sistema online (baixa e envia para lá)")
    a = ap.parse_args()
    sistema = os.path.abspath(a.sistema)
    with open(os.path.join(sistema, "config.json"), "r", encoding="utf-8") as f:
        cfg = json.load(f)
    raiz = os.path.abspath(a.raiz) if a.raiz else cfg["pasta_raiz"]
    destino = a.zip or os.path.join(raiz, "_exportados", "nuvem_%s.zip" % dt.date.today().isoformat())
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    # o online troca estes caminhos (os gravados nos .json, sempre como o Windows os vê) pelas pastas de lá
    raiz_win = cfg.get("pasta_raiz") or raiz
    sistema_win = sistema if os.name == "nt" else raiz_win.rstrip("\\") + "\\_sistema"
    n = montar_zip(sistema, raiz, destino, a.estacao, sistema_win, raiz_win)
    print("zip pronto: %s (%d arquivos, %.1f MB)" % (destino, n, os.path.getsize(destino) / 1e6))
    if a.url:
        if not a.senha:
            raise SystemExit("Informe --senha (a senha do admin do sistema online).")
        print(enviar(destino, a.url, a.usuario, a.senha, a.estacao))
        if a.ligar_estacao:
            arq = os.path.join(sistema, "dados", "rede.json")
            try:
                with open(arq, "r", encoding="utf-8") as f:
                    rede = json.load(f)
            except (OSError, ValueError):
                rede = {}
            if os.path.isfile(arq):
                import shutil
                shutil.copy2(arq, arq + ".antes_online")
            rede.update(papel="estacao", servidor=a.url.rstrip("/"), senha=a.senha, nome=a.estacao, principal="1")
            rede.setdefault("usuario", "")
            with open(arq, "w", encoding="utf-8") as f:
                json.dump(rede, f, ensure_ascii=False, indent=1)
            print("Este computador agora é estação de %s (rede.json; o anterior ficou em rede.json.antes_online)." % a.url)


if __name__ == "__main__":
    main()
