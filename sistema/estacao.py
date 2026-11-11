# -*- coding: utf-8 -*-
"""Modo estação: este computador usa os dados do computador principal (servidor).

- O servidor local (127.0.0.1) repassa a tela e as ações para o servidor, com a senha da rede.
  A extensão do Chrome continua chamando localhost/periodo e /registrar normalmente.
- O que mexe NESTE computador (Chrome, perfis, pasta de downloads, abrir arquivo no Excel, Claude Code)
  é feito aqui mesmo.
- Um "agente" avisa o servidor a cada poucos segundos que está ligado, manda os perfis do Chrome daqui,
  recebe os pedidos de download (Executar) e envia para o servidor os relatórios baixados aqui.
"""
import base64
import datetime as dt
import json
import os
import queue
import shutil
import threading
import time
import traceback
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, quote

import comum
import rede

INTERVALO = 10          # segundos entre avisos ao servidor
IGNORAR = (".crdownload", ".part", ".tmp", ".download", ".parte")
_fila = queue.Queue()
_ultimo_erro = {"msg": None, "quando": None}
_ultimo_ok = {"quando": None}
_cfg_cache = {"cfg": None, "t": 0}


# ---------------------------------------------------------------- configuração
def cfg_rede(forcar=False):
    """Config do servidor + o que é deste computador (cache curto; se o servidor cair, usa a última)."""
    if not forcar and _cfg_cache["cfg"] and time.time() - _cfg_cache["t"] < 20:
        return _cfg_cache["cfg"]
    try:
        cfg = rede.config_mesclada()
        _cfg_cache.update(cfg=cfg, t=time.time())
        try:   # cópia local para uso sem rede (ex.: abrir o Chrome cedo)
            with open(os.path.join(comum.PASTA_SISTEMA, "dados", "config_servidor.json"), "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False)
        except OSError:
            pass
        return cfg
    except rede.ErroServidor:
        if _cfg_cache["cfg"]:
            return _cfg_cache["cfg"]
        try:
            with open(os.path.join(comum.PASTA_SISTEMA, "dados", "config_servidor.json"), "r", encoding="utf-8") as f:
                cfg = json.load(f)
            rede.mesclar(cfg)
            return cfg
        except (OSError, ValueError):
            raise


def minhas_contas(cfg=None):
    cfg = cfg or cfg_rede()
    return [c for c in comum.contas(cfg) if comum.deste_computador(c)]


# ---------------------------------------------------------------- enviar arquivos baixados aqui
def enviar_pendentes(cfg=None, contas=None):
    """Manda para o servidor os arquivos que chegaram na _entrada local das contas deste computador."""
    try:
        cfg = cfg or cfg_rede()
    except Exception as e:
        comum.log("Estação: sem computador principal para enviar arquivos (%s)." % e)
        return {"ok": False, "enviados": 0}
    contas = contas if contas is not None else minhas_contas(cfg)
    enviados = 0
    agora = time.time()
    for c in contas:
        pasta = comum.pasta_entrada(cfg, comum.pasta_de(c))
        try:
            os.makedirs(pasta, exist_ok=True)
            nomes = os.listdir(pasta)
        except OSError:
            continue
        for nome in nomes:
            arq = os.path.join(pasta, nome)
            if not os.path.isfile(arq) or nome.startswith(".") or nome.lower().endswith(IGNORAR) or nome.lower() == "desktop.ini":
                continue
            try:
                if agora - os.path.getmtime(arq) < 4:   # ainda sendo gravado
                    continue
                with open(arq, "rb") as f:
                    conteudo = base64.b64encode(f.read()).decode("ascii")
                r = rede.chamar("/api/agente/arquivo", {"conta_id": c["id"], "nome": nome, "conteudo": conteudo,
                                                        "organizar": False}, timeout=120)
                if not r.get("ok"):
                    comum.log("Estação: servidor recusou %s: %s" % (nome, r.get("erro")))
                    continue
                destino = os.path.join(comum.raiz(cfg), "_enviados", comum.pasta_de(c), dt.date.today().isoformat())
                os.makedirs(destino, exist_ok=True)
                alvo = os.path.join(destino, nome)
                if os.path.exists(alvo):
                    b, e = os.path.splitext(nome)
                    alvo = os.path.join(destino, "%s_%s%s" % (b, dt.datetime.now().strftime("%H%M%S"), e))
                shutil.move(arq, alvo)
                enviados += 1
                comum.log("[%s] enviado ao computador principal: %s" % (c["nome"], nome))
            except rede.ErroServidor as e:
                comum.log("Estação: não enviei %s (%s). Tento de novo depois." % (nome, e))
                return {"ok": False, "enviados": enviados}
            except OSError as e:
                comum.log("Estação: erro com %s: %s" % (nome, e))
    if enviados:
        try:
            rede.chamar("/api/agente/organizar", {}, timeout=300)
        except rede.ErroServidor as e:
            comum.log("Estação: %s" % e)
    return {"ok": True, "enviados": enviados}


# ---------------------------------------------------------------- execução local
def _ver_faltando(cfg, c, mks):
    try:
        return rede.chamar("/api/agente/faltando", {"conta_id": c["id"], "mks": mks}, timeout=60).get("faltando", mks)
    except rede.ErroServidor:
        return mks


def _aprender(cfg, c, texto, inicio_iso):
    try:
        return rede.chamar("/api/agente/roteiro", {"conta_id": c["id"], "texto": texto, "inicio_iso": inicio_iso},
                           timeout=60).get("aprendidos", [])
    except rede.ErroServidor:
        return []


def _pos_rodada(cfg, c):
    enviar_pendentes(cfg, [c])


def rodar_local(conta_ids=None, mk=None, origem="manual"):
    """Baixa (Claude + Chrome deste computador) as contas daqui. Bloqueante."""
    import executor
    cfg = cfg_rede(forcar=True)
    ids = [i for i in (conta_ids or [c["id"] for c in minhas_contas(cfg)])]
    return executor.rodar(ids or ["__nenhuma__"], mk, origem, cfg=cfg, pos_rodada=_pos_rodada,
                          ver_faltando=_ver_faltando, aprender=_aprender)


def _trabalhador():
    while True:
        t = _fila.get()
        try:
            if t.get("tipo") == "executar":
                comum.log("Estação: pedido de download recebido (%s)." % (t.get("origem") or "?"))
                rodar_local(t.get("conta_ids"), t.get("mk"), "%s (pela rede)" % (t.get("origem") or "servidor"))
        except Exception:
            traceback.print_exc()
            comum.log("Estação: erro ao executar: %s" % traceback.format_exc()[-400:])


def _agente():
    import executor
    import servidor
    while True:
        try:
            cfg = cfg_rede()
            perfis = servidor.info_perfis(cfg)
            r = rede.chamar("/api/agente/ping", {"nome": rede.ler()["nome"], "perfis": perfis,
                                                 "execucao": executor.ler_estado(),
                                                 "claude": bool(executor.claude_exe(cfg)), "chrome": bool(comum.chrome_exe(cfg))},
                            timeout=20)
            _ultimo_ok["quando"] = dt.datetime.now().isoformat(timespec="seconds")
            _ultimo_erro["msg"] = None
            for t in r.get("trabalhos") or []:
                if t.get("tipo") == "parar":
                    executor.parar()
                else:
                    _fila.put(t)
            if not executor.ler_estado().get("rodando"):
                enviar_pendentes(cfg)     # downloads feitos à mão nos perfis daqui
        except Exception as e:
            if str(e) != _ultimo_erro["msg"]:
                comum.log("Estação: %s" % e)
            _ultimo_erro.update(msg=str(e), quando=dt.datetime.now().isoformat(timespec="seconds"))
        time.sleep(INTERVALO)


def _registrar_evento_rede(original):
    def f(ev):
        original(ev)
        if isinstance(ev, dict) and ev.get("tipo") == "execucao":
            try:
                rede.chamar("/api/agente/evento", {"evento": dict(ev)}, timeout=20)
            except Exception:
                pass
    return f


# ---------------------------------------------------------------- servidor local (repasse)
def _estado_local(cfg):
    import executor
    import servidor
    d = rede.ler()
    return {"perfis": servidor.info_perfis(cfg), "chrome_aberto": comum.chrome_aberto(), "chrome_exe": comum.chrome_exe(cfg),
            "chrome_user_data": comum.chrome_user_data(cfg), "execucao": executor.ler_estado(),
            "claude_exe": executor.claude_exe(cfg), "pasta_sistema": comum.PASTA_SISTEMA,
            "app_desktop": servidor.GANCHOS.get("mostrar_janela") is not None,
            "rede_local": {"papel": "estacao", "nome": d["nome"], "usuario": d.get("usuario") or "", "servidor": d["servidor"], "tem_senha": bool(d["senha"]),
                           "ips": rede.ips_locais(), "ultimo_ok": _ultimo_ok["quando"], "erro": _ultimo_erro["msg"]}}


def _abrir_arquivo_aqui(dados):
    r = rede.chamar("/api/abrir-pasta", {"caminho": dados.get("caminho", "")}, timeout=30)
    if not r.get("ok") or not r.get("baixar"):
        return r
    corpo, _h = rede.chamar(r["baixar"], metodo="GET", timeout=120, bruto=True)
    pasta = os.path.join(comum.PASTA_SISTEMA, "dados", "abertos")
    os.makedirs(pasta, exist_ok=True)
    alvo = os.path.join(pasta, os.path.basename(r.get("nome") or "arquivo"))
    try:
        with open(alvo, "wb") as f:
            f.write(corpo)
    except PermissionError:   # já aberto no Excel
        b, e = os.path.splitext(alvo)
        alvo = "%s_%s%s" % (b, dt.datetime.now().strftime("%H%M%S"), e)
        with open(alvo, "wb") as f:
            f.write(corpo)
    if os.name == "nt":
        os.startfile(alvo)
    return {"ok": True, "aviso": "Cópia aberta deste computador (alterações nela não voltam para o principal)."}


def _conta_daqui(cfg, conta_id):
    c = comum.conta(cfg, conta_id)
    return c if c and comum.deste_computador(c) else None


def atualizar_do_principal():
    """Copia do computador principal os arquivos do sistema (.py, painel, scripts) para esta estação.
    Os arquivos substituídos vão para backups\antes_atualizar_<data>. config.json, dados e app não são tocados."""
    r = rede.chamar("/api/sistema/arquivos", {}, timeout=120)
    arqs = r.get("arquivos") or {}
    if not r.get("ok") or "servidor.py" not in arqs or "painel.html" not in arqs:
        raise rede.ErroServidor(r.get("erro") or "o computador principal não mandou os arquivos do sistema (ele precisa estar atualizado).")
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    bkp = os.path.join(comum.PASTA_SISTEMA, "backups", "antes_atualizar_" + carimbo)
    os.makedirs(bkp, exist_ok=True)
    trocados = []
    for nome, b64 in arqs.items():
        if os.path.basename(nome) != nome or nome.lower() in ("config.json", "python.txt"):
            continue
        novo = base64.b64decode(b64)
        alvo = os.path.join(comum.PASTA_SISTEMA, nome)
        if os.path.isfile(alvo):
            with open(alvo, "rb") as f:
                if f.read() == novo:
                    continue
            shutil.copy2(alvo, os.path.join(bkp, nome))
        with open(alvo + ".novo", "wb") as f:
            f.write(novo)
        os.replace(alvo + ".novo", alvo)
        trocados.append(nome)
    comum.log("Estação atualizada a partir do principal: %d arquivo(s) trocados%s" % (
        len(trocados), (": " + ", ".join(trocados[:15])) if trocados else ""))
    return trocados


def _local_post(caminho, dados):
    """Ações feitas neste computador. None = repassar ao servidor."""
    import servidor
    import executor
    if caminho == "/api/sistema/atualizar":
        # estação: busca o código do principal e remonta o app DESTE computador
        try:
            trocados = atualizar_do_principal()
        except rede.ErroServidor as e:
            return {"ok": False, "erro": "Não consegui buscar a versão do computador principal: %s" % e}
        r = servidor.atualizar_sistema(dados)
        if r.get("ok"):
            r["trocados"] = len(trocados)
        return r
    if caminho in ("/api/rede/ler", "/api/rede/salvar", "/api/rede/testar", "/api/rede/firewall", "/api/janela/mostrar",
                   "/api/agendar", "/api/claude/instalar", "/api/claude/login", "/api/claude/status"):
        if caminho == "/api/rede/ler":
            return {"ok": True, "rede": _estado_local(cfg_rede())["rede_local"]}
        return servidor.ACOES[caminho](dados)
    if caminho == "/api/abrir-pasta":
        return _abrir_arquivo_aqui(dados)
    if caminho == "/api/ajustes":
        # Chrome, Claude, horários e "baixar sozinho" são deste computador; o resto (links, instruções, modelos) é do principal
        locais = {k: dados.pop(k) for k in ("chrome_exe", "chrome_user_data", "claude_exe", "horarios") if k in dados}
        ex = dict(dados.get("execucao") or {})
        if "agendado" in ex:
            locais["execucao"] = {"agendado": ex.pop("agendado")}
            dados["execucao"] = ex
        if locais:
            cfg_local = comum.carregar_config()
            lex = cfg_local.get("execucao", {})
            locais.setdefault("execucao", {})
            for k in ("limite_minutos_por_cliente", "max_turnos", "modelo", "modelo_reforco"):
                locais["execucao"].setdefault(k, lex.get(k, {"limite_minutos_por_cliente": 40, "max_turnos": 200}.get(k, "")))
            r = servidor.salvar_ajustes(locais)
            if not r.get("ok"):
                return r
        _cfg_cache["t"] = 0
        try:
            return rede.chamar(caminho, dados, timeout=120)
        except rede.ErroServidor as e:
            return {"ok": True, "aviso": "Ajustes deste computador salvos. Os demais não chegaram ao principal: %s" % e}
    cfg = cfg_rede()
    if caminho == "/api/abrir-chrome":
        import abrir_chrome
        return abrir_chrome.abrir(cfg)
    if caminho == "/api/abrir-no-perfil":
        perfil = dados.get("perfil")
        if dados.get("cliente_id"):
            c = _conta_daqui(cfg, dados["cliente_id"])
            if not c:
                return None          # conta do principal: o servidor responde
            perfil = c.get("perfil")
        if not perfil:
            return {"ok": False, "erro": "Perfil não informado."}
        return comum.abrir_no_perfil(cfg, perfil, dados.get("url", ""))
    if caminho == "/api/downloads":
        c = _conta_daqui(cfg, dados.get("cliente_id"))
        if not c:
            return None
        if not c.get("perfil"):
            return {"ok": False, "erro": "Conta sem perfil do Chrome definido."}
        pasta = comum.pasta_entrada(cfg, comum.pasta_de(c))
        os.makedirs(pasta, exist_ok=True)
        return comum.alterar_pasta_downloads(cfg, c["perfil"], pasta)
    if caminho == "/api/claude/testar":
        c = _conta_daqui(cfg, dados.get("cliente_id"))
        if not c:
            return None
        return executor.testar_conexao(dados.get("cliente_id"), cfg=cfg)
    if caminho == "/api/executar":
        if dados.get("cliente_id"):
            if not _conta_daqui(cfg, dados["cliente_id"]):
                return None
            if executor.ler_estado().get("rodando"):
                return {"ok": False, "erro": "Já existe uma execução em andamento neste computador."}
            _fila.put({"tipo": "executar", "conta_ids": [dados["cliente_id"]], "mk": dados.get("mk"), "origem": "botão do app"})
            return {"ok": True}
        return None                  # todos: o servidor roda os dele e manda os daqui pela fila
    if caminho == "/api/executar/agendado":
        _fila.put({"tipo": "executar", "conta_ids": None, "mk": None, "origem": "agendamento"})
        return {"ok": True}
    if caminho == "/api/executar/parar":
        executor.parar()
        return None                  # e o servidor para os dele
    if caminho == "/api/organizar":
        enviar_pendentes(cfg)
        return None
    return None


class ProxyHandler(__import__("servidor").Handler):
    def _permitido(self):
        return self._local() and self._host_ok()

    def do_GET(self):
        if not self._permitido():
            return self._enviar(403, {"erro": "host não permitido"})
        url = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        try:
            if url.path in ("/", "/index.html"):
                with open(os.path.join(comum.PASTA_SISTEMA, "painel.html"), "rb") as f:
                    return self._enviar(200, f.read(), "text/html; charset=utf-8")
            if url.path == "/saude":
                return self._enviar(200, {"ok": True, "estacao": True})
            if url.path == "/periodo":
                return super().do_GET()          # mesma regra de datas, calculada aqui
            if url.path == "/api/estado":
                try:
                    cfg = cfg_rede()
                    est = rede.chamar(self.path, metodo="GET", timeout=30)       # /api/estado?mes=...
                except Exception as e:
                    d = rede.ler()
                    return self._enviar(200, {"erro_rede": str(e), "rede": {"papel": "estacao", "nome": d["nome"],
                                              "servidor": d["servidor"], "tem_senha": bool(d["senha"]), "ips": rede.ips_locais()}})
                loc = _estado_local(cfg)
                est["perfis_servidor"] = est.get("perfis") or []
                est["config"] = rede.mesclar(est.get("config") or {})
                est.update({k: v for k, v in loc.items() if k != "rede_local"})
                est["rede_servidor"] = est.get("rede")
                est["rede"] = loc["rede_local"]
                return self._enviar(200, est)
            # o resto (/registrar da extensão, /api/arquivo, ...) vai para o servidor
            corpo, h = rede.chamar(self.path, metodo="GET", timeout=120, bruto=True)
            return self._enviar(200, corpo, h.get("Content-Type", "application/octet-stream"))
        except rede.ErroServidor as e:
            if url.path == "/registrar":
                corpo = "<h1>Registro NÃO enviado</h1><p>%s</p>" % e
                return self._enviar(502, self._pagina("Erro", corpo), "text/html; charset=utf-8")
            return self._enviar(502, {"ok": False, "erro": str(e)})
        except Exception as e:
            traceback.print_exc()
            return self._enviar(500, {"erro": str(e)})

    def do_POST(self):
        if not self._permitido() or self.headers.get("X-Painel") != "1":
            return self._enviar(403, {"erro": "requisição não permitida"})
        caminho = urlparse(self.path).path
        try:
            tam = int(self.headers.get("Content-Length") or 0)
            bruto = self.rfile.read(tam) if tam else b""
            dados = json.loads(bruto or b"{}")
            r = _local_post(caminho, dados)
            if r is not None:
                return self._enviar(200, r)
            return self._enviar(200, rede.chamar(caminho, dados, timeout=600))
        except rede.ErroServidor as e:
            return self._enviar(200, {"ok": False, "erro": str(e)})
        except Exception as e:
            traceback.print_exc()
            return self._enviar(500, {"ok": False, "erro": "Erro interno: %s" % e})


_iniciado = {"v": False}


def iniciar_agente():
    if _iniciado["v"]:
        return
    _iniciado["v"] = True
    comum.registrar_evento = _registrar_evento_rede(comum.registrar_evento)
    threading.Thread(target=_agente, daemon=True).start()
    threading.Thread(target=_trabalhador, daemon=True).start()


def criar_servidor(cfg=None):
    cfg = cfg or comum.carregar_config()
    comum.criar_pastas_base(cfg)
    srv = ThreadingHTTPServer(("127.0.0.1", int(cfg.get("porta_painel", 8765))), ProxyHandler)
    iniciar_agente()
    return srv
