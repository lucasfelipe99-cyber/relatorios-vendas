# -*- coding: utf-8 -*-
"""Abre o Chrome com cada perfil cadastrado (tarefa das 4h50)."""
import subprocess
import time

import comum


def abrir(cfg=None):
    if cfg is None:
        try:
            import rede
            cfg = rede.config_mesclada() if rede.papel() == "estacao" else None
        except Exception as e:
            comum.log("ABRIR CHROME: não consegui ler o computador principal: %s" % e)
            return {"ok": False, "erro": str(e)}
    cfg = cfg or comum.carregar_config()
    exe = comum.chrome_exe(cfg)
    if not exe:
        comum.log("ABRIR CHROME: chrome.exe não encontrado. Informe o caminho no painel (Ajustes).")
        return {"ok": False, "erro": "chrome.exe não encontrado"}
    abertos = []
    perfis = []
    for c in comum.contas(cfg, so_ativas=True):
        if c.get("perfil") and c.get("marketplaces") and c["perfil"] not in perfis and comum.deste_computador(c):
            perfis.append(c["perfil"])
    porta = cfg.get("porta_painel", 8765)
    for p in perfis:
        try:
            subprocess.Popen([exe, "--profile-directory=%s" % p, "--new-window", "http://localhost:%d/periodo" % porta])
            abertos.append(p)
            time.sleep(4)
        except OSError as e:
            comum.log("ABRIR CHROME: falha no perfil %s: %s" % (p, e))
    comum.log("ABRIR CHROME: %d perfil(is) aberto(s): %s" % (len(abertos), ", ".join(abertos) or "nenhum"))
    return {"ok": True, "perfis": abertos}


if __name__ == "__main__":
    abrir()
