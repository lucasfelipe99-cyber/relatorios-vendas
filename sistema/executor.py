# -*- coding: utf-8 -*-
"""Executor: o app aciona o Claude Code, que controla a extensão Claude in Chrome.

Para garantir que o download caia no perfil certo, roda UM cliente por vez:
fecha o Chrome, abre só o perfil daquele cliente, chama `claude -p --chrome`
com o texto do atalho, roda o organizador e fecha o Chrome antes do próximo.
"""
import datetime as dt
import json
import os
import shutil
import subprocess
import threading
import time

import comum

SEM_JANELA = getattr(subprocess, "CREATE_NO_WINDOW", 0)
ARQ_ESTADO = os.path.join(comum.PASTA_SISTEMA, "dados", "execucao.json")
ARQ_TRAVA = os.path.join(comum.PASTA_SISTEMA, "dados", "execucao.lock")
_parar = threading.Event()

PREAMBULO = (
    "Você está rodando de forma automática, sem ninguém olhando. Use SOMENTE as ferramentas do navegador (Claude in Chrome). "
    "Navegador: use o Chrome deste computador. Se houver mais de um navegador conectado, escolha o que está neste computador "
    "(onThisComputer/isLocal); se houver mais de um neste computador, NÃO escolha: pare e registre erro com motivo=outro e "
    "detalhe=varios-navegadores para cada marketplace. Abra abas novas para o trabalho e não feche abas que já estavam abertas."
)


# ---------------------------------------------------------------- Claude Code
def claude_exe(cfg=None):
    cfg = cfg or comum.carregar_config()
    if cfg.get("claude_exe") and os.path.exists(cfg["claude_exe"]):
        return cfg["claude_exe"]
    candidatos = [os.path.join(os.path.expanduser("~"), ".local", "bin", "claude.exe"),
                  os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "WinGet", "Links", "claude.exe")]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return shutil.which("claude") or ""


def status_claude(cfg=None):
    exe = claude_exe(cfg)
    if not exe:
        return {"instalado": False, "logado": False, "exe": "", "versao": ""}
    info = {"instalado": True, "exe": exe, "versao": "", "logado": False, "conta": ""}
    try:
        r = subprocess.run([exe, "--version"], capture_output=True, text=True, timeout=30, creationflags=SEM_JANELA)
        info["versao"] = (r.stdout or "").strip()
        r = subprocess.run([exe, "auth", "status"], capture_output=True, text=True, timeout=30, creationflags=SEM_JANELA)
        dados = json.loads(r.stdout or "{}")
        info["logado"] = dados.get("authMethod") == "claude.ai" or (r.returncode == 0 and dados.get("loggedIn", False))
        info["metodo"] = dados.get("authMethod")
        conta = dados.get("account") if isinstance(dados.get("account"), dict) else {}
        info["conta"] = dados.get("email") or conta.get("email", "")
    except Exception as e:
        info["erro"] = str(e)
    return info


def abrir_terminal(comando_ps):
    """Abre uma janela do PowerShell visível para você acompanhar (instalação/login)."""
    if os.name != "nt":
        return {"ok": False, "erro": "Disponível apenas no Windows."}
    subprocess.Popen(["powershell", "-NoExit", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", comando_ps],
                     creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0))
    return {"ok": True}


def instalar_claude():
    return abrir_terminal("Write-Host 'Instalando o Claude Code (oficial da Anthropic)...' -ForegroundColor Cyan; "
                          "irm https://claude.ai/install.ps1 | iex; "
                          "Write-Host ''; Write-Host 'Pronto. Volte ao app e clique em Fazer login.' -ForegroundColor Green")


def login_claude(cfg=None):
    exe = claude_exe(cfg)
    if not exe:
        return {"ok": False, "erro": "Instale o Claude Code primeiro."}
    return abrir_terminal("Write-Host 'Entre com a MESMA conta Claude usada na extensão do Chrome.' -ForegroundColor Cyan; "
                          "& '%s' auth login; Write-Host ''; Write-Host 'Pode fechar esta janela.' -ForegroundColor Green" % exe)


# ---------------------------------------------------------------- Chrome
def fechar_chrome(espera=25):
    if os.name != "nt" or not comum.chrome_aberto():
        return
    subprocess.run(["taskkill", "/IM", "chrome.exe"], capture_output=True, creationflags=SEM_JANELA)  # fechamento normal
    for _ in range(espera):
        time.sleep(1)
        if not comum.chrome_aberto():
            return
    subprocess.run(["taskkill", "/IM", "chrome.exe", "/F", "/T"], capture_output=True, creationflags=SEM_JANELA)
    time.sleep(2)


def abrir_perfil(cfg, perfil_dir):
    exe = comum.chrome_exe(cfg)
    if not exe:
        raise RuntimeError("chrome.exe não encontrado")
    porta = cfg.get("porta_painel", 8765)
    subprocess.Popen([exe, "--profile-directory=%s" % perfil_dir, "--new-window", "http://localhost:%d/periodo" % porta])


# ---------------------------------------------------------------- estado
def _gravar_estado(**kw):
    os.makedirs(os.path.dirname(ARQ_ESTADO), exist_ok=True)
    estado = ler_estado()
    estado.update(kw)
    estado["atualizado"] = dt.datetime.now().isoformat(timespec="seconds")
    tmp = ARQ_ESTADO + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=1)
    os.replace(tmp, ARQ_ESTADO)


def ler_estado():
    try:
        with open(ARQ_ESTADO, "r", encoding="utf-8") as f:
            e = json.load(f)
    except (OSError, ValueError):
        e = {}
    # trava esquecida (processo morreu)
    if e.get("rodando") and not _trava_viva():
        e["rodando"] = False
        e["etapa"] = "interrompido"
    return e


def _trava_viva():
    try:
        return time.time() - os.path.getmtime(ARQ_TRAVA) < 120
    except OSError:
        return False


def _tocar_trava():
    os.makedirs(os.path.dirname(ARQ_TRAVA), exist_ok=True)
    with open(ARQ_TRAVA, "w") as f:
        f.write(str(os.getpid()))


# ---------------------------------------------------------------- execução
def _rodar_claude(cfg, cliente, mks, log_arq, modelo=None):
    exe = claude_exe(cfg)
    prompt = PREAMBULO + "\n\n" + (comum.prompt_combinado(cfg, dict(cliente, marketplaces=mks)) if len(mks) > 1
                                   else comum.prompt_marketplace(cfg, cliente, mks[0]))
    trabalho = os.path.join(comum.PASTA_SISTEMA, "dados", "claude_trabalho")
    os.makedirs(trabalho, exist_ok=True)
    ex = cfg.get("execucao", {})
    cmd = [exe, "-p", prompt, "--chrome",
           "--allowedTools", "mcp__claude-in-chrome__*",
           "--disallowedTools", "Bash", "PowerShell", "Edit", "Write", "NotebookEdit", "WebFetch", "WebSearch",
           "--permission-prompts", "none",
           "--output-format", "json", "--no-session-persistence",
           "--max-turns", str(int(ex.get("max_turnos", 200)))]
    modelo = ex.get("modelo", "sonnet") if modelo is None else modelo
    if modelo:
        cmd += ["--model", modelo]
    limite = int(ex.get("limite_minutos_por_cliente", 40)) * 60
    inicio = time.time()
    with open(log_arq, "w", encoding="utf-8") as saida:
        saida.write("# %s — %s — %s\n# comando: claude -p <prompt> --chrome ...\n\n" % (
            dt.datetime.now().isoformat(timespec="seconds"), cliente["nome"], ", ".join(mks)))
        saida.flush()
        proc = subprocess.Popen(cmd, cwd=trabalho, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace", creationflags=SEM_JANELA,
                                start_new_session=(os.name != "nt"))
        _gravar_estado(pid_claude=proc.pid)
        saida_txt = []

        def ler():
            for linha in proc.stdout:
                saida_txt.append(linha)
                saida.write(linha)
                saida.flush()
        t = threading.Thread(target=ler, daemon=True)
        t.start()
        motivo_fim = None
        while proc.poll() is None:
            _tocar_trava()
            if _parar.is_set():
                motivo_fim = "parado por você"
            elif time.time() - inicio > limite:
                motivo_fim = "tempo limite (%d min)" % (limite // 60)
            if motivo_fim:
                parar_processo(proc.pid)
                break
            time.sleep(2)
        proc.wait(timeout=30)
        t.join(timeout=10)
    if _parar.is_set() and not motivo_fim:
        motivo_fim = "parado por você"
    texto = "".join(saida_txt).strip()
    resultado = {"codigo": proc.returncode, "duracao_min": round((time.time() - inicio) / 60, 1), "motivo_fim": motivo_fim,
                 "modelo": modelo or "padrão da conta", "mks": mks}
    try:
        j = json.loads(texto[texto.index("{"):]) if "{" in texto else {}
        if not j:
            raise ValueError("sem JSON")
        resultado["texto_completo"] = j.get("result") or ""
        resultado.update({"resposta": (j.get("result") or "")[-2000:], "erro": j.get("is_error"),
                          "turnos": j.get("num_turns"), "custo_estimado_usd": j.get("total_cost_usd")})
    except ValueError:
        resultado["resposta"] = texto[-2000:]
        resultado["erro"] = proc.returncode != 0
    if motivo_fim:
        resultado["erro"] = True
    return resultado


def aprender_roteiros(cfg_dummy, cliente, texto, inicio_iso):
    """Guarda os roteiros que o Claude descreveu, só dos marketplaces que registraram 'ok' nesta rodada."""
    import re
    blocos = re.findall(r'<roteiro\s+mk="([a-z0-9]+)"\s*>(.*?)</roteiro>', texto or "", re.S)
    if not blocos:
        return []
    oks = {e.get("mk") for e in comum.ler_historico(3000) if e.get("tipo") == "registro" and e.get("cliente") == cliente["id"]
           and e.get("status") == "ok" and e.get("ts", "") >= inicio_iso}
    salvos = []
    cfg = comum.carregar_config()
    for mk, rot in blocos:
        rot = "\n".join(l.rstrip() for l in rot.strip().splitlines() if l.strip())[:1500]
        if mk not in cfg["marketplaces"] or mk not in oks or len(rot) < 20:
            continue
        link = comum.link_cliente(cfg, cliente, mk)
        rots = cfg["marketplaces"][mk].setdefault("roteiros", {})
        anterior = rots.get(link, {})
        rots[link] = {"texto": rot, "atualizado": dt.datetime.now().isoformat(timespec="seconds"),
                      "rodadas": int(anterior.get("rodadas", 0)) + 1, "editado_por_voce": False}
        salvos.append(mk)
    if salvos:
        comum.salvar_config(cfg)
        comum.log("[%s] roteiro aprendido atualizado: %s" % (cliente["nome"], ", ".join(salvos)))
    return salvos


def faltando(cfg, cliente, mks):
    hoje = dt.date.today()
    return [m for m in mks if not os.path.exists(os.path.join(comum.pasta_dia(cfg, comum.pasta_de(cliente), hoje), comum.nome_arquivo(m, cfg, hoje)))]


def parar_processo(pid):
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=SEM_JANELA)
    else:
        try:
            os.killpg(os.getpgid(pid), 9)
        except OSError:
            pass


def _papel():
    try:
        import rede
        return rede.ler()
    except Exception:
        return {"papel": "sozinho", "nome": ""}


def contas_locais(cfg, contas):
    """Contas que este computador baixa. Sozinho: todas. Servidor: as sem computador ou deste. Estação: só as dela."""
    r = _papel()
    if r["papel"] == "sozinho":
        return contas
    if r["papel"] == "estacao":
        return [c for c in contas if (c.get("computador") or "") == r["nome"]]
    return [c for c in contas if (c.get("computador") or "") in ("", r["nome"])]


def rodar(cliente_ids=None, mk=None, origem="manual", cfg=None, pos_rodada=None, ver_faltando=None, aprender=None):
    """Roda os clientes indicados (ou todos os ativos), um por vez. Bloqueante.
    Na estação: cfg vem do servidor, pos_rodada envia os arquivos baixados, ver_faltando e aprender perguntam ao servidor."""
    cfg = cfg or comum.carregar_config()
    if ler_estado().get("rodando") and _trava_viva():
        return {"ok": False, "erro": "Já existe uma execução em andamento."}
    if not claude_exe(cfg):
        return {"ok": False, "erro": "Claude Code não instalado. Veja a aba Horários e ajustes > Claude Code."}
    clientes = [c for c in contas_locais(cfg, comum.filtrar_contas(cfg, cliente_ids)) if c.get("perfil") and c.get("marketplaces")]
    if not clientes:
        return {"ok": False, "erro": "Nenhuma conta ativa com perfil e marketplaces."}
    _parar.clear()
    _tocar_trava()
    _gravar_estado(rodando=True, origem=origem, inicio=dt.datetime.now().isoformat(timespec="seconds"),
                   fila=[c["nome"] for c in clientes], cliente=None, etapa="começando", resultados={})
    import organizador
    resultados = {}
    try:
        for c in clientes:
            if _parar.is_set():
                break
            mks = [mk] if mk else [m for m in c["marketplaces"] if comum.baixa_sozinho(cfg, m)]
            if not mks:
                continue
            pasta_logs = os.path.join(comum.PASTA_SISTEMA, "logs", "claude")
            os.makedirs(pasta_logs, exist_ok=True)
            log_arq = os.path.join(pasta_logs, "%s_%s.log" % (dt.datetime.now().strftime("%Y%m%d-%H%M%S"), c["id"]))
            _gravar_estado(cliente=c["nome"], cliente_id=c["id"], etapa="fechando o Chrome", log=log_arq)
            fechar_chrome()
            _gravar_estado(etapa="abrindo o Chrome no perfil do cliente")
            abrir_perfil(cfg, c["perfil"])
            for _ in range(int(cfg.get("execucao", {}).get("espera_chrome_seg", 15))):
                _tocar_trava()
                time.sleep(1)
            _gravar_estado(etapa="baixando: " + ", ".join(cfg["marketplaces"][m]["nome"] for m in mks))
            comum.log("[%s] EXECUÇÃO iniciada (%s) — %s" % (c["nome"], origem, ", ".join(mks)))
            ex = cfg.get("execucao", {})
            tentativas = [ex.get("modelo", "sonnet")]
            reforco = ex.get("modelo_reforco", "")
            if reforco != tentativas[0]:
                tentativas.append(reforco)
            pendentes, rodadas = list(mks), []
            for n, modelo in enumerate(tentativas):
                if _parar.is_set() or not pendentes:
                    break
                if n > 0:
                    _gravar_estado(etapa="reforço (modelo mais forte): " + ", ".join(cfg["marketplaces"][m]["nome"] for m in pendentes))
                    comum.log("[%s] tentando de novo com modelo mais forte: %s" % (c["nome"], ", ".join(pendentes)))
                    log_arq = log_arq.replace(".log", "_reforco.log")
                inicio_iso = dt.datetime.now().isoformat(timespec="seconds")
                try:
                    r = _rodar_claude(cfg, c, pendentes, log_arq, modelo)
                except Exception as e:
                    r = {"erro": True, "resposta": "Falha ao chamar o Claude Code: %s" % e, "mks": pendentes}
                r["log"] = log_arq
                r["aprendidos"] = (aprender or aprender_roteiros)(cfg, c, r.pop("texto_completo", ""), inicio_iso)
                _gravar_estado(etapa="aguardando downloads terminarem", resultados=dict(resultados, **{c["id"]: r}))
                for _ in range(int(ex.get("espera_download_seg", 30))):
                    _tocar_trava()
                    time.sleep(1)
                if pos_rodada:
                    pos_rodada(cfg, c)
                else:
                    organizador.executar(forcar_estavel=True)  # downloads em andamento (.crdownload) continuam ignorados
                antes = pendentes
                pendentes = (ver_faltando or faltando)(cfg, c, pendentes)
                r["recebidos"] = [m for m in antes if m not in pendentes]
                rodadas.append(r)
                comum.registrar_evento({"tipo": "execucao", "cliente": c["id"], "mks": antes, "origem": origem,
                                        "ok": not pendentes and not r.get("erro"), "motivo_fim": r.get("motivo_fim"),
                                        "duracao_min": r.get("duracao_min"), "turnos": r.get("turnos"),
                                        "custo_estimado_usd": r.get("custo_estimado_usd"), "modelo": r.get("modelo"),
                                        "recebidos": r["recebidos"], "aprendidos": r["aprendidos"], "log": log_arq,
                                        "reforco": n > 0})
                comum.log("[%s] rodada %s terminou em %.1f min, %s passos — recebidos: %s" % (
                    c["nome"], "de reforço" if n else "normal", r.get("duracao_min") or 0, r.get("turnos") or "?",
                    ", ".join(r["recebidos"]) or "nenhum"))
                if r.get("motivo_fim") == "parado por você":
                    break
            final = dict(rodadas[-1]) if rodadas else {"erro": True}
            final["rodadas"] = len(rodadas)
            final["turnos"] = sum((x.get("turnos") or 0) for x in rodadas) or None
            final["custo_estimado_usd"] = round(sum((x.get("custo_estimado_usd") or 0) for x in rodadas), 4) or None
            final["duracao_min"] = round(sum((x.get("duracao_min") or 0) for x in rodadas), 1)
            final["faltando"] = pendentes
            final["erro"] = bool(pendentes) or final.get("erro")
            final["aprendidos"] = sorted({m for x in rodadas for m in x.get("aprendidos", [])})
            resultados[c["id"]] = final
            fechar_chrome()
        if pos_rodada:
            for c in clientes:
                pos_rodada(cfg, c)
        else:
            organizador.executar(forcar_estavel=True)
    finally:
        _gravar_estado(rodando=False, etapa="parado" if _parar.is_set() else "concluído", cliente=None,
                       fim=dt.datetime.now().isoformat(timespec="seconds"), resultados=resultados, pid_claude=None)
        try:
            os.replace(ARQ_TRAVA, ARQ_TRAVA + ".fim")
        except OSError:
            pass
    return {"ok": True, "resultados": resultados}


def rodar_em_segundo_plano(cliente_ids=None, mk=None, origem="manual", so_locais=False):
    # contas de outros computadores (estações) vão para a fila delas (o agendamento de cada computador roda só as suas)
    r = _papel()
    enviados = []
    if r["papel"] == "servidor" and not so_locais:
        import rede
        cfg = comum.carregar_config()
        todas = [c for c in comum.filtrar_contas(cfg, cliente_ids) if c.get("perfil") and c.get("marketplaces")]
        remotas = {}
        for c in todas:
            comp = c.get("computador") or ""
            if comp and comp != r["nome"]:
                remotas.setdefault(comp, []).append(c["id"])
        for comp, ids in remotas.items():
            rede.enfileirar(comp, {"tipo": "executar", "conta_ids": ids, "mk": mk, "origem": origem})
            enviados.append("%s (%s)" % (comp, "online" if rede.online(comp) else "desligado: roda quando ligar"))
        if remotas and not [c for c in todas if (c.get("computador") or "") in ("", r["nome"])]:
            return {"ok": True, "enviado": enviados}
    if ler_estado().get("rodando") and _trava_viva():
        return {"ok": False, "erro": "Já existe uma execução em andamento."}
    if not claude_exe():
        return {"ok": False, "erro": "Claude Code não instalado. Veja Horários e ajustes > Claude Code."}
    threading.Thread(target=rodar, args=(cliente_ids, mk, origem), daemon=True).start()
    return {"ok": True, "enviado": enviados}


def parar():
    r = _papel()
    if r["papel"] == "servidor":
        import rede
        for nome in rede.ESTACOES:
            rede.enfileirar(nome, {"tipo": "parar"})
    _parar.set()
    pid = ler_estado().get("pid_claude")
    if pid:
        parar_processo(pid)
    _gravar_estado(etapa="parando…")
    return {"ok": True}


def testar_conexao(cliente_id, cfg=None):
    """Abre só o perfil do cliente e pergunta ao Claude Code quais navegadores ele enxerga."""
    cfg = cfg or comum.carregar_config()
    c = comum.conta(cfg, cliente_id)
    exe = claude_exe(cfg)
    if not c or not c.get("perfil"):
        return {"ok": False, "erro": "Cliente sem perfil."}
    if not exe:
        return {"ok": False, "erro": "Claude Code não instalado."}
    fechar_chrome()
    abrir_perfil(cfg, c["perfil"])
    time.sleep(int(cfg.get("execucao", {}).get("espera_chrome_seg", 15)))
    cmd = [exe, "-p", "Use a ferramenta list_connected_browsers do Claude in Chrome e responda em português, em uma linha por "
           "navegador: nome, se está neste computador e se está em uso. Não abra nem navegue em nenhuma página.",
           "--chrome", "--allowedTools", "mcp__claude-in-chrome__list_connected_browsers",
           "--disallowedTools", "Bash", "PowerShell", "Edit", "Write",
           "--permission-prompts", "none", "--output-format", "json", "--no-session-persistence", "--max-turns", "5"]
    trabalho = os.path.join(comum.PASTA_SISTEMA, "dados", "claude_trabalho")
    os.makedirs(trabalho, exist_ok=True)
    try:
        r = subprocess.run(cmd, cwd=trabalho, capture_output=True, text=True, encoding="utf-8", errors="replace",
                           timeout=240, creationflags=SEM_JANELA)
        txt = (r.stdout or "") + (r.stderr or "")
        try:
            j = json.loads(r.stdout[r.stdout.index("{"):])
            return {"ok": not j.get("is_error"), "resposta": j.get("result", txt)}
        except ValueError:
            return {"ok": r.returncode == 0, "resposta": txt[-3000:]}
    except subprocess.TimeoutExpired:
        return {"ok": False, "resposta": "O Claude Code não respondeu em 4 minutos."}
