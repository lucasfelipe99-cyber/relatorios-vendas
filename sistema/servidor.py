# -*- coding: utf-8 -*-
"""Servidor local do painel (somente 127.0.0.1)."""
import datetime as dt
import html
import json
import os
import subprocess
import sys
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import comum
import organizador

TRAVA_CONFIG = comum.TRAVA_CFG   # a mesma trava usada pelos tratamentos ao completar a tabela de custos
GANCHOS = {"mostrar_janela": None, "sair": None}  # preenchido pelo app desktop


# ---------------------------------------------------------------- status do dia
def calcular_status(cfg, hoje=None):
    hoje = hoje or dt.date.today()
    dia = comum.rotulo_dia(hoje)
    ini, fim = comum.periodo(hoje)
    esperado = {"inicio": comum.fmt(ini), "fim": comum.fmt(fim)}
    hist = comum.ler_historico()
    pend = organizador.estado_pendencias()
    linhas, resumo = [], {"esperados": 0, "recebidos": 0, "faltando": 0, "erro": 0, "tratados": 0}
    for c in comum.contas(cfg, so_ativas=True):
        entradas = organizador.pastas_entrada(cfg, c)
        nao_rec = pend.get(c["id"], [])
        # arquivos que ainda estão na entrada (pendências podem ter sido resolvidas à mão)
        nao_rec = [p for p in nao_rec if any(os.path.exists(os.path.join(e, p["arquivo"])) for e in entradas)]
        for mk in c.get("marketplaces", []):
            if mk not in cfg["marketplaces"]:
                continue
            resumo["esperados"] += 1
            arq = os.path.join(comum.pasta_dia(cfg, comum.pasta_de(c), hoje), comum.nome_arquivo(mk, cfg, hoje))
            evs = [e for e in hist if e.get("cliente") == c["id"] and e.get("mk") == mk]
            hoje_evs = [e for e in evs if (e.get("data") or e.get("ts", ""))[:10] == hoje.isoformat()]   # dia de referência
            regs = [e for e in hoje_evs if e.get("tipo") == "registro"]
            ult_reg = regs[-1] if regs else None
            orgs = [e for e in evs if e.get("tipo") == "organizado"]
            ult_ok = orgs[-1]["ts"] if orgs else None
            linha = {"cliente_id": c["id"], "cliente": c["nome"], "grupo_id": c["cliente_id"], "grupo": c["cliente_nome"],
                     "conta": c["conta_nome"], "mk": mk, "mk_nome": cfg["marketplaces"][mk]["nome"],
                     "arquivo": arq, "ultima_extracao_ok": ult_ok, "periodo": None, "periodo_origem": None,
                     "pasta_dia": comum.pasta_dia(cfg, comum.pasta_de(c), hoje), "registro": ult_reg,
                     "manual": c.get("modo") == "manual" or not comum.baixa_sozinho(cfg, mk),
                     "sem_tratador": bool(cfg["marketplaces"][mk].get("sem_tratador")),
                     "tipos_upload": [{"mk": k, "nome": v["nome"]} for k, v in cfg["marketplaces"].items() if k == mk or v.get("pai") == mk]}
            if os.path.exists(arq):
                linha["estado"] = "recebido"
                resumo["recebidos"] += 1
                reg_ok = next((e for e in reversed(regs) if e.get("status") == "ok" and e.get("inicio")), None)
                org_hoje = next((e for e in reversed(hoje_evs) if e.get("tipo") == "organizado"), None)
                if reg_ok:
                    linha["periodo"], linha["periodo_origem"] = "%s a %s" % (reg_ok["inicio"], reg_ok["fim"]), "informado pela extensão"
                elif org_hoje and org_hoje.get("inicio"):
                    linha["periodo"], linha["periodo_origem"] = "%s a %s" % (org_hoje["inicio"], org_hoje["fim"]), "lido do nome do arquivo"
                else:
                    linha["periodo"], linha["periodo_origem"] = "%s a %s" % (esperado["inicio"], esperado["fim"]), "período esperado (não confirmado)"
            elif ult_reg and ult_reg.get("status") == "erro":
                linha["estado"] = "erro"
                linha["detalhe"] = "Extensão registrou erro: %s %s" % (ult_reg.get("motivo", ""), ult_reg.get("detalhe", ""))
                resumo["erro"] += 1
            elif nao_rec:
                linha["estado"] = "nao_reconhecido"
                linha["detalhe"] = "%d arquivo(s) parado(s) na _entrada sem reconhecimento" % len(nao_rec)
                resumo["erro"] += 1
            elif comum.ok_vendas(c["id"], mk, hoje):
                okm = comum.ok_vendas(c["id"], mk, hoje)
                linha["estado"] = "ok_manual"
                linha["detalhe"] = "OK manual: %s" % okm.get("motivo", "")
                resumo["recebidos"] += 1
            else:
                linha["estado"] = "faltando"
                resumo["faltando"] += 1
                if c.get("modo") == "manual":
                    linha["detalhe"] = "Conta manual: clique em Subir relatório."
                elif not comum.baixa_sozinho(cfg, mk):
                    linha["detalhe"] = "%s ainda não baixa sozinho: clique em Subir relatório." % cfg["marketplaces"][mk]["nome"]
            trats = [e for e in hoje_evs if e.get("tipo") == "tratamento"]
            linha["tratamento"] = trats[-1] if trats else None
            if trats and trats[-1].get("ok") and os.path.exists(trats[-1].get("saida", "")):
                resumo["tratados"] += 1
            elif linha["estado"] == "ok_manual" and not linha["sem_tratador"]:
                resumo["tratados"] += 1        # OK manual (sem relatório): não há o que tratar, conta como feito
            linhas.append(linha)
    return {"dia": dia, "data": hoje.isoformat(), "periodo_esperado": esperado, "linhas": linhas, "resumo": resumo,
            "pendencias": _pendencias_vivas(cfg, pend)}


def _pendencias_vivas(cfg, pend):
    saida = {}
    for k, v in pend.items():
        u = comum.conta(cfg, k)
        if not u:
            continue
        entradas = organizador.pastas_entrada(cfg, u)
        vivas = [p for p in v if any(os.path.exists(os.path.join(e, p["arquivo"])) for e in entradas)]
        if vivas:
            saida[k] = vivas
    return saida


def _rede():
    try:
        import rede
        return rede
    except Exception:
        return None


def _deste_computador(conta):
    return comum.deste_computador(conta)


def info_perfis(cfg):
    perfis = comum.listar_perfis(cfg)
    usados = {}
    for c in comum.contas(cfg):
        if not _deste_computador(c):
            continue
        if c.get("perfil"):
            usados.setdefault(c["perfil"], []).append(c)
    saida = []
    for p in perfis:
        d = comum.pasta_downloads_perfil(cfg, p["dir"])
        clientes = usados.get(p["dir"], [])
        alvo = comum.pasta_entrada(cfg, comum.pasta_de(clientes[0])) if len(clientes) == 1 else None
        alerta = None
        if len(clientes) > 1:
            alerta = "Este perfil está ligado a mais de uma conta (%s). Cada conta precisa de um perfil próprio." % ", ".join(c["nome"] for c in clientes)
        elif alvo and not comum.mesma_pasta(d["pasta"], alvo):
            alerta = "A pasta de downloads não é a _entrada de %s." % clientes[0]["nome"]
        elif alvo and d["pergunta"]:
            alerta = "O Chrome pergunta onde salvar cada arquivo; isso trava o download automático."
        saida.append({**p, "extensao": comum.extensao_instalada(cfg, p["dir"]), "downloads": d["pasta"],
                      "pergunta_onde_salvar": d["pergunta"], "erro_leitura": d["erro"],
                      "clientes": [c["nome"] for c in clientes], "conta_ids": [c["id"] for c in clientes],
                      "pasta_correta": alvo, "alerta": alerta})
    return saida


def _forecasts_definidos(cfg):
    try:
        import forecast
        h = dt.date.today()
        out = {}
        for c in cfg["clientes"]:
            d = forecast.ler_definido(c["id"], h.year, h.month)
            if d:
                out[c["id"]] = {"versao": d["versao"], "definido_em": d["definido_em"], "total": d["total"]}
        return out
    except Exception:
        return {}


def _situacao_lanc_mes_atual(cfg):
    try:
        import lancamentos, consolidado
        a, m = consolidado.mes_atual()
        return {u["id"]: lancamentos.situacao(cfg, lancamentos.dono(u), a, m) for u in comum.contas(cfg)}
    except Exception as e:
        comum.log("lançamentos: %s" % e)
        return {}


def _dia_ref(mes):
    """Dia de referência de um mês: o mês atual é hoje; um mês passado é o dia 1 do mês seguinte (a pasta desse dia
    guarda o relatório do mês inteiro, como no download automático do dia 1)."""
    hoje = dt.date.today()
    m = _mes_valido(mes) if mes else None
    if not m:
        return hoje
    a, mm = int(m[:4]), int(m[5:])
    if (a, mm) == _consolidado().mes_atual(hoje):
        return hoje
    d = dt.date(a + (mm == 12), 1 if mm == 12 else mm + 1, 1)
    return d if d <= hoje else hoje


def montar_estado(mes=None):
    cfg = comum.carregar_config()
    ref = _dia_ref(mes)
    horarios = comum.horarios_escalonados(cfg)
    atalhos = []
    for c in comum.contas(cfg):
        mks = [m for m in c.get("marketplaces", []) if comum.baixa_sozinho(cfg, m)]   # Temu/Via/Madeira/Site: sem atalho por enquanto
        atalhos.append({
            "cliente_id": c["id"], "cliente": c["nome"], "horario": horarios.get(c["id"]),
            "combinado": {"nome": comum.nome_atalho(c), "texto": comum.prompt_combinado(cfg, c)} if mks else None,
            "por_marketplace": [{"mk": m, "mk_nome": cfg["marketplaces"][m]["nome"], "nome": comum.nome_atalho(c, m),
                                 "link": comum.link_cliente(cfg, c, m), "link_status": cfg["marketplaces"][m]["link_status"]
                                 if not (c.get("links") or {}).get(m) else "personalizado para esta conta",
                                 "texto": comum.prompt_marketplace(cfg, c, m)} for m in mks],
        })
    ini, fim = comum.periodo()
    return {
        "config": cfg, "contas": comum.contas(cfg), "perfis": info_perfis(cfg), "status": calcular_status(cfg, ref), "atalhos": atalhos,
        "dia_ref": ref.isoformat(), "mes_ref": "%04d-%02d" % _consolidado().mes_atual(ref), "mes_passado": ref != dt.date.today(),
        "chrome_aberto": comum.chrome_aberto(), "chrome_exe": comum.chrome_exe(cfg),
        "chrome_user_data": comum.chrome_user_data(cfg), "link_extensao": comum.LINK_EXTENSAO,
        "pasta_sistema": comum.PASTA_SISTEMA, "nuvem": comum.NUVEM, "app_desktop": GANCHOS.get("mostrar_janela") is not None, "periodo": {"inicio": comum.fmt(ini), "fim": comum.fmt(fim)},
        "agora": dt.datetime.now().isoformat(timespec="seconds"),
        "execucao": _executor().ler_estado(), "claude_exe": _executor().claude_exe(cfg),
        "evolucao": [e for e in comum.ler_historico(5000) if e.get("tipo") == "execucao"][-20:][::-1],
        "consolidados": _consolidado().ler_estado(), "meses_consolidado": _consolidado().meses_disponiveis(cfg, quantos=12),
        "lancamentos": _situacao_lanc_mes_atual(cfg),
        "forecasts": _forecasts_definidos(cfg),
        "rede": _info_rede(),
        "atividade": (_atividade().recentes(30) if _atividade() else []),
        "equipe": (_atividade().quadro(cfg) if _atividade() else {}),
    }


def _atividade():
    try:
        import atividade
        return atividade
    except Exception:
        traceback.print_exc()
        return None


def _info_rede():
    r = _rede()
    if not r:
        return {"papel": "sozinho"}
    d = r.ler()
    out = {"papel": d["papel"], "nome": d["nome"], "usuario": d.get("usuario") or "", "servidor": d["servidor"], "tem_senha": bool(d["senha"]),
           "ips": r.ips_locais(), "estacoes": r.resumo_estacoes() if d["papel"] == "servidor" else {}}
    return out


def ler_rede(_dados):
    return {"ok": True, "rede": _info_rede()}


def salvar_rede(dados):
    r = _rede()
    try:
        d = r.salvar(dados)
    except ValueError as e:
        return {"ok": False, "erro": str(e)}
    return {"ok": True, "rede": _info_rede(), "aviso": "Feche e abra o app neste computador para aplicar (modo: %s)." % d["papel"]}


def testar_rede(dados):
    """Na estação: testa o endereço e a senha informados (antes de salvar)."""
    import urllib.request, urllib.error
    r = _rede()
    s = str(dados.get("servidor") or "").strip()
    if not s:
        return {"ok": False, "erro": "Informe o endereço do computador principal."}
    if not s.startswith("http"):
        s = "http://" + s
    if s.count(":") < 2 and not s.startswith("https://"):
        s += ":%d" % int(comum.carregar_config().get("porta_painel", 8765))
    req = urllib.request.Request(s.rstrip("/") + "/api/agente/config", data=b"{}", method="POST",
                                 headers={"X-Painel": "1", "Content-Type": "application/json",
                                          "Authorization": "Bearer " + r.token(str(dados.get("senha") or ""))})
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            cfg = json.loads(resp.read())["config"]
        return {"ok": True, "mensagem": "Conectado! O computador principal tem %d cliente(s)." % len(cfg.get("clientes", []))}
    except urllib.error.HTTPError as e:
        return {"ok": False, "erro": "Senha recusada." if e.code == 401 else "Erro %s no computador principal." % e.code}
    except Exception as e:
        return {"ok": False, "erro": "Não consegui conectar em %s (%s). Confira o IP, se o app está aberto lá e o firewall." % (s, getattr(e, "reason", e))}


def liberar_firewall(_dados):
    if os.name != "nt":
        return {"ok": False, "erro": "Disponível apenas no Windows."}
    porta = int(comum.carregar_config().get("porta_painel", 8765))
    cmd = ('netsh advfirewall firewall delete rule name="Relatorios de Vendas" >nul 2>&1 & '
           'netsh advfirewall firewall add rule name="Relatorios de Vendas" dir=in action=allow protocol=TCP localport=%d profile=private,domain' % porta)
    try:  # pede permissão de administrador (janela do Windows)
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Start-Process cmd -Verb RunAs -WindowStyle Hidden -ArgumentList '/c %s'" % cmd.replace("'", "''")],
                       timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return {"ok": True, "aviso": "Se o Windows pediu permissão e você aceitou, a porta %d foi liberada na rede privada." % porta}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


# ---------------------------------------------------------------- agente (estações chamam)
def agente_ping(dados, ip=None):
    return _rede().registrar_ping(dados, ip)


def agente_config(_dados):
    return {"ok": True, "config": comum.carregar_config()}


def agente_arquivo(dados):
    """A estação baixou um relatório: guarda na _entrada da conta aqui e organiza."""
    import base64
    cfg = comum.carregar_config()
    c = comum.conta(cfg, dados.get("conta_id"))
    if not c:
        return {"ok": False, "erro": "Conta desconhecida."}
    nome = os.path.basename(str(dados.get("nome") or ""))
    if not nome or nome.startswith("."):
        return {"ok": False, "erro": "Nome de arquivo inválido."}
    pasta = comum.pasta_entrada(cfg, comum.pasta_de(c))
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, nome)
    base, ext = os.path.splitext(destino)
    n = 1
    while os.path.exists(destino):
        destino = "%s (%d)%s" % (base, n, ext)
        n += 1
    with open(destino + ".parte", "wb") as f:
        f.write(base64.b64decode(dados.get("conteudo") or ""))
    os.replace(destino + ".parte", destino)
    comum.log("[%s] arquivo recebido da estação %s: %s" % (c["nome"], dados.get("estacao") or "?", nome))
    if dados.get("organizar", True):
        organizador.executar(forcar_estavel=True)
    return {"ok": True}


def agente_organizar(_dados):
    return {"ok": True, "resumo": organizador.executar(forcar_estavel=True)}


def agente_faltando(dados):
    cfg = comum.carregar_config()
    c = comum.conta(cfg, dados.get("conta_id"))
    if not c:
        return {"ok": False, "erro": "Conta desconhecida."}
    return {"ok": True, "faltando": _executor().faltando(cfg, c, dados.get("mks") or [])}


def agente_roteiro(dados):
    cfg = comum.carregar_config()
    c = comum.conta(cfg, dados.get("conta_id"))
    if not c:
        return {"ok": False, "erro": "Conta desconhecida."}
    return {"ok": True, "aprendidos": _executor().aprender_roteiros(cfg, c, dados.get("texto") or "", dados.get("inicio_iso"))}


def agente_evento(dados):
    """Eventos de execução da estação (histórico do painel)."""
    ev = dados.get("evento") or {}
    if isinstance(ev, dict) and ev.get("tipo"):
        ev["estacao"] = dados.get("estacao")
        comum.registrar_evento(ev)
    if dados.get("log"):
        comum.log("[estação %s] %s" % (dados.get("estacao") or "?", str(dados["log"])[:500]))
    return {"ok": True}


# ---------------------------------------------------------------- ações
def salvar_cliente(dados):
    """Cliente = grupo com uma ou mais contas (uma por navegador/perfil do Chrome)."""
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        erro = comum.validar_nome_cliente(dados.get("nome"))
        if erro:
            return {"ok": False, "erro": erro}
        nome = dados["nome"].strip()
        cid = dados.get("id")
        outros = [c for c in cfg["clientes"] if c["id"] != cid]
        if any(comum.sem_acento(c["nome"]) == comum.sem_acento(nome) for c in outros):
            return {"ok": False, "erro": "Já existe um cliente com esse nome."}
        contas_in = dados.get("contas") or []
        if not contas_in:
            return {"ok": False, "erro": "Cadastre pelo menos uma conta (navegador) para o cliente."}
        atual = next((c for c in cfg["clientes"] if c["id"] == cid), None) if cid else None
        if cid and not atual:
            return {"ok": False, "erro": "Cliente não encontrado."}
        # o mesmo nome de perfil pode existir em computadores diferentes (cada Chrome tem os seus)
        perfis_outros = {(ct.get("computador") or "", ct.get("perfil")): "%s / %s" % (c["nome"], ct["nome"])
                         for c in outros for ct in c.get("contas", []) if ct.get("perfil")}
        antigas = {ct["id"]: ct for ct in (atual or {}).get("contas", [])}
        novas, nomes, perfis = [], set(), set()
        for i, ci in enumerate(contas_in, 1):
            cn = (ci.get("nome") or "").strip() or nome
            erro = comum.validar_nome_cliente(cn)
            if erro:
                return {"ok": False, "erro": "Conta %d: %s" % (i, erro)}
            if comum.sem_acento(cn) in nomes:
                return {"ok": False, "erro": "Duas contas com o mesmo nome (%s)." % cn}
            nomes.add(comum.sem_acento(cn))
            modo = "manual" if ci.get("modo") == "manual" else "automatico"
            perfil = "" if modo == "manual" else (ci.get("perfil") or "")   # conta manual: relatórios sobem pelo app
            comp = "" if modo == "manual" else str(ci.get("computador") or "").strip()
            chave = (comp, perfil)
            if perfil and (chave in perfis or chave in perfis_outros):
                dono = perfis_outros.get(chave, "outra conta deste cliente")
                return {"ok": False, "erro": "O perfil da conta \"%s\" já está ligado a %s. Use um perfil por conta." % (cn, dono)}
            if perfil:
                perfis.add(chave)
            mks = [m for m in ci.get("marketplaces", []) if m in cfg["marketplaces"] and not cfg["marketplaces"][m].get("oculto")]
            links = {m: (ci.get("links") or {}).get(m, "").strip() for m in mks if (ci.get("links") or {}).get(m, "").strip()}
            for m, l in links.items():
                if not l.startswith("https://"):
                    return {"ok": False, "erro": "Conta \"%s\": o link de %s precisa começar com https://" % (cn, cfg["marketplaces"][m]["nome"])}
            ct_id = ci.get("id") if ci.get("id") in antigas else comum.gerar_id(cn, {x["id"] for x in novas} | set(antigas))
            ct = dict(antigas.get(ct_id, {}))
            ct.update({"id": ct_id, "nome": cn, "modo": modo, "perfil": perfil, "computador": comp, "marketplaces": mks, "links": links,
                       "ativo": bool(ci.get("ativo", True))})
            novas.append(ct)
        avisos = []
        if atual:
            if atual["nome"] != nome:
                avisos.append("Nome do cliente alterado: as pastas antigas de \"%s\" foram mantidas (nada é apagado)." % atual["nome"])
            for ct in novas:
                ant = antigas.get(ct["id"])
                if ant and ant["nome"] != ct["nome"]:
                    avisos.append("Conta renomeada de \"%s\" para \"%s\": pastas antigas mantidas." % (ant["nome"], ct["nome"]))
            removidas = [a["nome"] for k, a in antigas.items() if k not in {x["id"] for x in novas}]
            if removidas:
                avisos.append("Conta(s) removida(s) do cadastro: %s. As pastas foram mantidas." % ", ".join(removidas))
            c = atual
        else:
            c = {"id": comum.gerar_id(nome, {x["id"] for x in cfg["clientes"]})}
            cfg["clientes"].append(c)
        if atual and (atual["nome"] != nome or any(antigas.get(ct["id"], ct)["nome"] != ct["nome"] for ct in novas)):
            avisos.append("Confira em Perfis e downloads: a pasta de downloads das contas precisa apontar para a _entrada nova.")
        c.update({"nome": nome, "ativo": bool(dados.get("ativo", True)), "contas": novas})
        c.setdefault("imposto", float(cfg.get("tratamento", {}).get("imposto_padrao", 0.14)))
        comum.criar_pastas_base(cfg)
        comum.criar_pastas_cliente(cfg, c)
        comum.salvar_config(cfg)
        pastas = [comum.pasta_entrada(cfg, comum.pasta_de(u)) for u in comum.contas_do_cliente(cfg, c)]
        return {"ok": True, "cliente": c, "aviso": " ".join(avisos) or None, "pastas_entrada": pastas}


def remover_cliente(dados):
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        antes = len(cfg["clientes"])
        cfg["clientes"] = [c for c in cfg["clientes"] if c["id"] != dados.get("id")]
        if len(cfg["clientes"]) == antes:
            return {"ok": False, "erro": "Cliente não encontrado."}
        comum.salvar_config(cfg)
        return {"ok": True, "aviso": "Cliente removido do cadastro. As pastas e arquivos dele foram mantidos."}


def salvar_ajustes(dados):
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        h = dados.get("horarios")
        if h:
            for k in ("abrir_chrome", "inicio_atalhos", "organizador_inicio", "organizador_fim"):
                v = str(h.get(k, cfg["horarios"][k]))
                try:
                    dt.datetime.strptime(v, "%H:%M")
                except ValueError:
                    return {"ok": False, "erro": "Horário inválido em %s: %s (use HH:MM)" % (k, v)}
                cfg["horarios"][k] = v
            for k in ("intervalo_entre_perfis_min", "organizador_intervalo_min"):
                try:
                    n = int(h.get(k, cfg["horarios"][k]))
                    assert 1 <= n <= 120
                except (ValueError, AssertionError):
                    return {"ok": False, "erro": "Intervalo inválido em %s (1 a 120 minutos)" % k}
                cfg["horarios"][k] = n
        for mk, txt in (dados.get("instrucoes") or {}).items():
            txt = (txt or "").strip()
            if mk in cfg["marketplaces"] and txt:
                if len(txt) > 4000:
                    return {"ok": False, "erro": "Instruções de %s muito longas (máx. 4000 caracteres)." % cfg["marketplaces"][mk]["nome"]}
                cfg["marketplaces"][mk]["instrucoes"] = txt
        for mk, link in (dados.get("links") or {}).items():
            if mk in cfg["marketplaces"] and link:
                if not link.startswith("https://"):
                    return {"ok": False, "erro": "Links precisam começar com https://"}
                if link != cfg["marketplaces"][mk]["link"]:
                    cfg["marketplaces"][mk]["link"] = link
                    cfg["marketplaces"][mk]["link_status"] = "ajustado por você em " + dt.date.today().strftime("%d/%m/%Y")
        ex = dados.get("execucao")
        if ex:
            atual = cfg.setdefault("execucao", {})
            try:
                atual["limite_minutos_por_cliente"] = max(5, min(180, int(ex.get("limite_minutos_por_cliente", atual.get("limite_minutos_por_cliente", 40)))))
                atual["max_turnos"] = max(20, min(1000, int(ex.get("max_turnos", atual.get("max_turnos", 200)))))
            except ValueError:
                return {"ok": False, "erro": "Limites de execução inválidos."}
            atual["modelo"] = str(ex.get("modelo", atual.get("modelo", ""))).strip()
            atual["modelo_reforco"] = str(ex.get("modelo_reforco", atual.get("modelo_reforco", ""))).strip()
            if "modo_paralelo" in ex:
                atual["modo_paralelo"] = bool(ex["modo_paralelo"])
            if "agendado" in ex:
                atual["agendado"] = bool(ex["agendado"])
        for mk, por_link in (dados.get("roteiros") or {}).items():
            if mk not in cfg["marketplaces"]:
                continue
            rots = cfg["marketplaces"][mk].setdefault("roteiros", {})
            for link, txt in (por_link or {}).items():
                txt = (txt or "").strip()[:1500]
                if not txt:
                    rots.pop(link, None)
                elif link in rots and rots[link].get("texto") != txt:
                    rots[link].update({"texto": txt, "editado_por_voce": True,
                                       "atualizado": dt.datetime.now().isoformat(timespec="seconds")})
        for chave in ("chrome_exe", "chrome_user_data", "claude_exe"):
            if chave in dados:
                cfg[chave] = dados[chave].strip()
        comum.salvar_config(cfg)
        return {"ok": True}


def alterar_downloads(dados):
    cfg = comum.carregar_config()
    c = comum.conta(cfg, dados.get("cliente_id"))
    if not c or not c.get("perfil"):
        return {"ok": False, "erro": "Conta sem perfil do Chrome definido."}
    return comum.alterar_pasta_downloads(cfg, c["perfil"], comum.pasta_entrada(cfg, comum.pasta_de(c)))


def caminho_permitido(alvo):
    cfg = comum.carregar_config()
    raiz = os.path.normcase(os.path.abspath(comum.raiz(cfg)))
    sistema = os.path.normcase(os.path.abspath(comum.PASTA_SISTEMA))
    a = os.path.normcase(os.path.abspath(alvo))
    if os.path.basename(a) in ("rede.json", "estacoes.json"):
        return False
    return a.startswith(raiz) or a.startswith(sistema)


def abrir_pasta(dados):
    cfg = comum.carregar_config()
    alvo = dados.get("caminho", "")
    if not caminho_permitido(alvo):
        return {"ok": False, "erro": "Só abro pastas dentro de " + comum.raiz(cfg)}
    if dados.get("_remoto"):   # pedido vindo de outro computador: o arquivo é baixado lá, não aberto aqui
        if os.path.isfile(alvo):
            from urllib.parse import quote
            return {"ok": True, "baixar": "/api/arquivo?caminho=" + quote(alvo), "nome": os.path.basename(alvo)}
        return {"ok": False, "erro": "Pastas abrem só no computador principal (%s). Os arquivos podem ser abertos daqui." % alvo}
    if os.path.isfile(alvo):
        if os.name == "nt":
            os.startfile(alvo)
        return {"ok": True}
    if not os.path.isdir(alvo):
        return {"ok": False, "erro": "A pasta ainda não existe (nenhum relatório organizado nela hoje)."}
    if os.name == "nt":
        os.startfile(alvo)
    return {"ok": True}


def reaplicar_agendamento(_dados):
    if os.name != "nt":
        return {"ok": False, "erro": "Disponível apenas no Windows."}
    script = os.path.join(comum.PASTA_SISTEMA, "agendar.ps1")
    env = {k: v for k, v in os.environ.items() if k.upper() not in ("PYTHONPATH", "PYTHONHOME")}
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script], env=env,
                       capture_output=True, text=True, timeout=120,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return {"ok": r.returncode == 0, "saida": (r.stdout or "")[-3000:], "erro": (r.stderr or "")[-3000:] or None}


def abrir_chrome_agora(_dados):
    import abrir_chrome
    return abrir_chrome.abrir()


def abrir_no_perfil(dados):
    cfg = comum.carregar_config()
    perfil = dados.get("perfil")
    if dados.get("cliente_id"):
        c = comum.conta(cfg, dados["cliente_id"])
        if not c or not c.get("perfil"):
            return {"ok": False, "erro": "Cliente sem perfil do Chrome definido."}
        perfil = c["perfil"]
    if not perfil:
        return {"ok": False, "erro": "Perfil não informado."}
    return comum.abrir_no_perfil(cfg, perfil, dados.get("url", ""))


def mostrar_janela(_dados):
    f = GANCHOS.get("mostrar_janela")
    if not f:
        return {"ok": False, "erro": "sem janela"}
    f()
    return {"ok": True}


def atualizar_sistema(_dados):
    """Botão "Atualizar sistema": abre o atualizar.ps1 (janela do PowerShell), que espera o app fechar,
    remonta o PainelRelatorios.exe com os arquivos atuais de _sistema e abre o app de novo."""
    script = os.path.join(comum.PASTA_SISTEMA, "atualizar.ps1")
    if os.name != "nt" or not os.path.isfile(script):
        return {"ok": False, "erro": "Atualizador não encontrado (%s)." % script}
    try:
        if _executor().ler_estado().get("rodando"):
            return {"ok": False, "erro": "Há um download de relatórios em andamento. Espere terminar (ou pare) e atualize depois."}
    except Exception:
        pass
    env = {k: v for k, v in os.environ.items()
           if not k.upper().startswith(("_PYI", "_MEI")) and k.upper() not in ("PYTHONPATH", "PYTHONHOME")}
    flags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-AppPid", str(os.getpid())],
                     cwd=comum.PASTA_SISTEMA, env=env, creationflags=flags, close_fds=True)
    comum.log("Atualização do app pedida pelo painel: fechando para remontar.")

    def fechar():
        import time
        time.sleep(1.5)
        f = GANCHOS.get("sair")
        if f:
            f()
        else:
            os._exit(0)
    threading.Thread(target=fechar, daemon=True).start()
    return {"ok": True}


# arquivos do sistema que a estação copia do principal no "Atualizar sistema" (nunca config.json, dados, .venv, app)
ARQUIVOS_SISTEMA_EXTRA = ("painel.html", "atualizar.ps1", "ATUALIZAR.bat", "agendar.ps1", "remover_agendamentos.ps1", "icone.ico")


def arquivos_sistema(_dados=None):
    """Código atual do principal (todos os .py + painel e scripts), em base64, para a estação se atualizar."""
    import base64
    import glob
    arqs = {}
    nomes = [os.path.basename(p) for p in glob.glob(os.path.join(comum.PASTA_CODIGO, "*.py"))] + list(ARQUIVOS_SISTEMA_EXTRA)
    for n in sorted(set(nomes)):
        p = os.path.join(comum.PASTA_CODIGO, n)
        if os.path.isfile(p):
            with open(p, "rb") as f:
                arqs[n] = base64.b64encode(f.read()).decode("ascii")
    return {"ok": True, "arquivos": arqs, "gerado": dt.datetime.now().isoformat(timespec="seconds")}


def _executor():
    import executor
    return executor


def _consolidado():
    import consolidado
    return consolidado


def _tratador():
    import tratador
    return tratador


def enviar_custos(dados):
    import base64
    tipo = dados.get("tipo", "custos")
    if tipo not in ("custos", "ids_shopee"):
        return {"ok": False, "erro": "Tipo inválido."}
    cid = dados.get("cliente_id")
    try:
        conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
    except ValueError:
        return {"ok": False, "erro": "Arquivo corrompido no envio."}
    if not conteudo:
        return {"ok": False, "erro": "Selecione um arquivo."}
    if len(conteudo) > 30 * 1024 * 1024:
        return {"ok": False, "erro": "Arquivo muito grande (máx. 30 MB)."}
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        c = next((c for c in cfg["clientes"] if c["id"] == cid), None)
        if not c:
            return {"ok": False, "erro": "Cliente não encontrado."}
        mes = None
        if tipo == "custos":
            mes = _mes_valido(dados.get("mes")) or _tratador().mes_chave(*_tratador().mes_do_periodo())
        try:
            info = _tratador().guardar_arquivo(cid, tipo, dados.get("nome", "arquivo"), conteudo, mes)
        except ValueError as e:
            return {"ok": False, "erro": str(e)}
        if tipo == "custos":
            _guardar_custos_mes(c, mes, info)
        else:
            c[tipo] = info
        comum.salvar_config(cfg)
    retratados = (_retratar_mes(cid, mes) if tipo == "custos" else _retratar_hoje(cid)) if dados.get("retratar", True) else []
    return {"ok": True, "info": info, "retratados": retratados}


def tabela_custos(dados):
    """Tabela de custos que vale para o mês, para editar no app."""
    cfg = comum.carregar_config()
    c = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
    if not c:
        return {"ok": False, "erro": "Cliente não encontrado."}
    mes = _mes_valido(dados.get("mes")) or _tratador().mes_chave(*_tratador().mes_do_periodo())
    a, m = (int(x) for x in mes.split("-"))
    try:
        linhas, info, mes_usado = _tratador().ler_tabela_mes(c, a, m)
    except Exception as e:
        return {"ok": False, "erro": "Não consegui ler a tabela de custos: %s" % e}
    auto = (info or {}).get("adicionados_auto") or {}
    # quantas vezes cada SKU aparece nos tratados de hoje (para priorizar o que mais vende)
    return {"ok": True, "mes": mes, "mes_usado": mes_usado, "linhas": linhas, "auto": auto,
            "info": {k: v for k, v in (info or {}).items() if k not in ("adicionados_auto", "historico")},
            "versao": (info or {}).get("versao")}


def titulos_custos(dados):
    """Para o editor de custos: títulos dos SKUs vendidos (dos tratados do mês e dos 3 meses anteriores),
    para preencher a Descrição de quem está sem. Não grava nada: o editor preenche e você confere e salva."""
    import collections
    cfg = comum.carregar_config()
    c = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
    if not c:
        return {"ok": False, "erro": "Cliente não encontrado."}
    mes = _mes_valido(dados.get("mes")) or _tratador().mes_chave(*_tratador().mes_do_periodo())
    a, m = (int(x) for x in mes.split("-"))
    cons, trat = _consolidado(), _tratador()
    contagem = {}
    meses, aa, mm = [], a, m
    for _ in range(4):
        meses.append((aa, mm))
        aa, mm = (aa - 1, 12) if mm == 1 else (aa, mm - 1)
    from openpyxl import load_workbook
    lidos = set()
    for (aa, mm) in meses:
        dia, _ = cons.dia_fonte(cfg, aa, mm)
        if not dia:
            continue
        for u in comum.contas_do_cliente(cfg, c):
            for mk in u.get("marketplaces", []):
                if mk not in trat.SKU_TRATADO:
                    continue
                p = cons._tratado_da_conta(cfg, u, mk, dia)
                if not os.path.isfile(p):
                    p, _d = cons._tratado_anterior(cfg, u, mk, aa, mm, dia)
                if not p or p in lidos or not os.path.isfile(p):
                    continue
                lidos.add(p)
                col_sku, col_nome = trat.SKU_TRATADO[mk]
                try:
                    wb = load_workbook(p, read_only=True, data_only=True)
                    ws = wb["DB"] if "DB" in wb.sheetnames else wb.worksheets[0]
                    it = ws.iter_rows(values_only=True)
                    cab = {str(x).strip(): j for j, x in enumerate(next(it)) if x is not None}
                    js, jn = cab.get(col_sku), cab.get(col_nome)
                    if js is not None and jn is not None:
                        for r in it:
                            k = trat._chave_sku(r[js] if js < len(r) else "")
                            t = " ".join(str(r[jn] if jn < len(r) else "" or "").split())[:300]
                            if k and t and t.lower() not in ("nan", "none"):
                                contagem.setdefault(k, collections.Counter())[t] += 1
                    wb.close()
                except Exception as e:
                    comum.log("títulos para custos: não li %s: %s" % (p, e))
    titulos = {k: v.most_common(1)[0][0] for k, v in contagem.items()}
    sem_zero = {}
    for k, t in titulos.items():
        sem_zero.setdefault(k.lstrip("0") or "0", t)
    pedidos = dados.get("skus") or []
    achados = {}
    for s_ in pedidos[:20000]:
        k = trat._chave_sku(s_)
        t = titulos.get(k) or sem_zero.get(k.lstrip("0") or "0")
        if t:
            achados[s_] = t
    return {"ok": True, "titulos": achados, "arquivos": len(lidos)}


def salvar_tabela_custos(dados):
    cid = dados.get("cliente_id")
    mes = _mes_valido(dados.get("mes"))
    if not mes:
        return {"ok": False, "erro": "Mês inválido."}
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        c = next((c for c in cfg["clientes"] if c["id"] == cid), None)
        if not c:
            return {"ok": False, "erro": "Cliente não encontrado."}
        a, m = (int(x) for x in mes.split("-"))
        info_ant, _ = _tratador().escolher_custos(c, a, m)
        if dados.get("versao_base") and info_ant and info_ant.get("versao") and int(dados["versao_base"]) != int(info_ant["versao"]) \
                and (info_ant.get("mes") == mes):
            return {"ok": False, "erro": "A tabela deste mês foi alterada por outra pessoa enquanto você editava (versão %s). "
                                         "Abra o editor de novo para ver a versão atual." % info_ant["versao"], "conflito": True}
        try:
            linhas, info = _tratador().salvar_tabela_mes(cid, mes, dados.get("linhas"))
        except ValueError as e:
            return {"ok": False, "erro": str(e)}
        info["nome_original"] = "editada no app" + ((" por " + dados["estacao"]) if dados.get("estacao") else "")
        auto = (info_ant or {}).get("adicionados_auto") or {}
        if auto:   # continua marcando os que o sistema acrescentou e ainda estão sem custo
            info["adicionados_auto"] = {k: v for k, v in auto.items() if any(l["SKU"] == k and l["Custo"] is None for l in linhas)}
        _tratador()._info_para_salvar(info, linhas)
        _guardar_custos_mes(c, mes, info)
        comum.salvar_config(cfg)
    retratados = _retratar_mes(cid, mes) if dados.get("retratar", True) else []
    return {"ok": True, "info": {k: v for k, v in info.items() if k != "adicionados_auto"}, "retratados": retratados}


def exportar_custos_editor(dados):
    cfg = comum.carregar_config()
    c = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
    if not c:
        return {"ok": False, "erro": "Cliente não encontrado."}
    mes = _mes_valido(dados.get("mes")) or "mes"
    destino, n = _tratador().exportar_custos(c, mes, dados.get("linhas"), dados.get("filtro") or "")
    if not n:
        return {"ok": False, "erro": "Nada para extrair com esse filtro."}
    return {"ok": True, "caminho": destino, "linhas": n}


def importar_custos_editor(dados):
    """Lê um Excel/CSV (SKU e Custo; Descrição e Categoria opcionais) e devolve as linhas para o editor juntar."""
    import base64
    try:
        conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
    except ValueError:
        return {"ok": False, "erro": "Arquivo corrompido no envio."}
    if not conteudo:
        return {"ok": False, "erro": "Escolha um arquivo."}
    nome = os.path.basename(dados.get("nome") or "custos.xlsx")
    ext = os.path.splitext(nome)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv", ".xls"):
        return {"ok": False, "erro": "Use .xlsx, .xlsm ou .csv."}
    import tempfile as _tf
    pasta = _tf.mkdtemp(prefix="impcustos_")
    arq = os.path.join(pasta, "custos" + ext)
    with open(arq, "wb") as f:
        f.write(conteudo)
    try:
        linhas, info = _tratador().ler_planilha_custos_parcial(arq)
    except Exception as e:
        return {"ok": False, "erro": "Arquivo não aceito: %s" % e}
    finally:
        import shutil
        shutil.rmtree(pasta, ignore_errors=True)
    return {"ok": True, "linhas": linhas, "invalidos": info["invalidos"], "aba": info["aba"]}


def _mes_valido(m):
    try:
        a, mm = (int(x) for x in str(m).split("-"))
        assert 2000 < a < 2100 and 1 <= mm <= 12
        return "%04d-%02d" % (a, mm)
    except (ValueError, AssertionError, TypeError):
        return None


def _guardar_custos_mes(c, mes, info):
    """Uma tabela por mês; vale sempre a última versão enviada para o mês."""
    meses = c.setdefault("custos_meses", {})
    ant = meses.get(mes) or {}
    info["versao"] = int(ant.get("versao", 0) or 0) + 1
    if ant:
        hist = ant.get("historico", [])
        hist.append({k: ant.get(k) for k in ("versao", "nome_original", "atualizado", "skus")})
        info["historico"] = hist[-20:]
    meses[mes] = info
    c["custos"] = info          # última enviada (compatibilidade)


def _retratar_mes(cliente_id, mes):
    """Depois de trocar a tabela de um mês: refaz os tratados que usam aquele mês e o consolidado do mês."""
    try:
        ano, mm = (int(x) for x in mes.split("-"))
    except (ValueError, AttributeError):
        return []
    if (ano, mm) == _tratador().mes_do_periodo():
        return _retratar_hoje(cliente_id)
    import consolidado
    cfg = comum.carregar_config()
    dia, _ = consolidado.dia_fonte(cfg, ano, mm)
    if not dia:
        return []
    res = []
    for u in comum.filtrar_contas(cfg, [cliente_id], so_ativas=False):
        mks = [m for m in u.get("marketplaces", []) if m in cfg["marketplaces"]
               and os.path.exists(os.path.join(comum.pasta_dia(cfg, comum.pasta_de(u), dia), comum.nome_arquivo(m, cfg, dia)))]
        if mks:
            res += [{"conta": u["nome"], "mk": e["mk"], "ok": e.get("ok"), "erro": e.get("erro")}
                    for e in _tratador().tratar_cliente(cfg, u, dia, mks, "custos de %s alterados" % mes)]
    cli = next((c for c in cfg["clientes"] if c["id"] == cliente_id), None)
    if cli:
        consolidado.gerar(cfg, cli, ano, mm, origem="custos de %s alterados" % mes)
    return res


def planilhas_lista(dados):
    import planilhas
    return dict(planilhas.listar(comum.carregar_config(), dados.get("dia")), ok=True)


def planilhas_ler(dados):
    import planilhas
    try:
        return planilhas.ler(comum.carregar_config(), dados.get("caminho"), dados.get("aba"), dados.get("pagina", 1),
                             dados.get("por_pagina", 100), dados.get("busca", ""), dados.get("ordem"),
                             dados.get("desc"), dados.get("filtros"))
    except ValueError as e:
        return {"ok": False, "erro": str(e)}
    except Exception as e:
        return {"ok": False, "erro": "Não consegui ler este arquivo: %s" % e}


def modelo_custos(dados):
    p = _tratador().caminho_modelo_custos()
    if os.name == "nt":
        os.startfile(p)
    return {"ok": True, "caminho": p}


def copiar_custos(dados):
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        de = next((c for c in cfg["clientes"] if c["id"] == dados.get("de")), None)
        para = next((c for c in cfg["clientes"] if c["id"] == dados.get("para")), None)
        if not de or not para or de is para:
            return {"ok": False, "erro": "Escolha clientes diferentes."}
        mes = _mes_valido(dados.get("mes")) or _tratador().mes_chave(*_tratador().mes_do_periodo())
        a, m = (int(x) for x in mes.split("-"))
        info_de, _ = _tratador().escolher_custos(de, a, m)
        origem = _tratador().caminho_custos(de, a, m)
        if not origem:
            return {"ok": False, "erro": "%s ainda não tem tabela de custos." % de["nome"]}
        with open(origem, "rb") as f:
            conteudo = f.read()
        try:
            info = _tratador().guardar_arquivo(para["id"], "custos", (info_de or {}).get("nome_original") or os.path.basename(origem), conteudo, mes)
        except ValueError as e:
            return {"ok": False, "erro": str(e)}
        info["copiado_de"] = de["nome"]
        _guardar_custos_mes(para, mes, info)
        comum.salvar_config(cfg)
    return {"ok": True, "info": info, "retratados": _retratar_mes(para["id"], mes)}


def salvar_imposto(dados):
    try:
        pct = float(str(dados.get("imposto_pct", "")).replace(",", "."))
        assert 0 <= pct < 100
    except (ValueError, AssertionError):
        return {"ok": False, "erro": "Informe o imposto em %, entre 0 e 99,99."}
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        c = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
        if not c:
            return {"ok": False, "erro": "Cliente não encontrado."}
        c["imposto"] = round(pct / 100, 6)
        comum.salvar_config(cfg)
    return {"ok": True, "retratados": _retratar_hoje(c["id"])}


def salvar_antecipa(dados):
    """Antecipação Shopee por conta: ativa (com %) ou desativada (0). Retrata o dia do cliente."""
    try:
        pct = float(str(dados.get("pct", "0")).replace(",", ".")) if dados.get("ativa") else 0.0
        assert 0 <= pct < 50
    except (ValueError, AssertionError):
        return {"ok": False, "erro": "Informe a antecipação em %, entre 0 e 49,99."}
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        c = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
        if not c:
            return {"ok": False, "erro": "Cliente não encontrado."}
        ct = next((x for x in c.get("contas", []) if x["id"] == dados.get("conta_id")), None)
        if not ct:
            return {"ok": False, "erro": "Conta não encontrada."}
        c.setdefault("shopee_antecipa", {})[ct["id"]] = round(pct / 100, 6)
        comum.salvar_config(cfg)
    comum.log("[%s / %s] antecipação Shopee: %s" % (c["nome"], ct["nome"], ("%.2f%%" % pct) if pct else "desativada"))
    return {"ok": True, "retratados": _retratar_hoje(c["id"])}


def _retratar_hoje(cliente_id):
    """Depois de mudar custos/imposto, refaz os tratados de hoje que já existem (o bruto continua igual)."""
    cfg = comum.carregar_config()
    hoje = dt.date.today()
    res = []
    for u in comum.filtrar_contas(cfg, [cliente_id], so_ativas=False):
        mks = [m for m in u.get("marketplaces", []) if m in cfg["marketplaces"]
               and os.path.exists(os.path.join(comum.pasta_dia(cfg, comum.pasta_de(u), hoje), comum.nome_arquivo(m, cfg, hoje)))]
        if mks:
            res += [{"conta": u["nome"], "mk": e["mk"], "ok": e.get("ok"), "erro": e.get("erro")}
                    for e in _tratador().tratar_cliente(cfg, u, hoje, mks, "custos/imposto alterados")]
    if res:
        _consolidar_depois([cliente_id], "custos/imposto alterados")
    return res


def _consolidar_depois(ids_contas_ou_clientes, origem, dia=None):
    """Refaz o relatório consolidado do mês dos clientes afetados (não derruba a ação se falhar).
    dia: dia de referência de um mês passado (refaz o pacote daquele mês)."""
    try:
        import consolidado
        cfg = comum.carregar_config()
        grupos = {comum.conta(cfg, i)["cliente_id"] if comum.conta(cfg, i) else i for i in ids_contas_ou_clientes if i}
        grupos = [g for g in grupos if any(c["id"] == g for c in cfg["clientes"])]
        if dia and consolidado.mes_atual(dia) != consolidado.mes_atual():
            a, m = consolidado.mes_atual(dia)
            if not cfg.get("consolidado", {}).get("automatico", True):
                return []
            return consolidado.gerar_clientes(cfg, sorted(grupos), a, m, origem=origem)
        return consolidado.automatico(cfg, grupos, origem)
    except Exception as e:
        comum.log("CONSOLIDADO: falhou (%s): %s" % (origem, e))
        return []


def classificar(d):
    ok, info = organizador.classificar_manual(comum.carregar_config(), d.get("cliente_id"), d.get("arquivo", ""), d.get("mk"))
    if ok:
        _consolidar_depois([d.get("cliente_id")], "arquivo classificado")
    return {"ok": ok, "info": info}


def subir_relatorio(dados):
    """Upload manual de um relatório de vendas (conta manual ou quando o download automático falhou).

    O arquivo vai para a _entrada da conta e passa pelo mesmo caminho do download: renomeia, guarda versão
    anterior, converte/extrai se precisar, trata e refaz o consolidado. Nada é apagado."""
    import base64
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    if not u or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    nome = os.path.basename(str(dados.get("nome") or "")).strip()
    if not nome or nome.startswith("."):
        return {"ok": False, "erro": "Nome de arquivo inválido."}
    try:
        conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
    except ValueError:
        return {"ok": False, "erro": "Arquivo corrompido no envio."}
    if not conteudo:
        return {"ok": False, "erro": "Arquivo vazio."}
    if len(conteudo) > 80 * 1024 * 1024:
        return {"ok": False, "erro": "Arquivo muito grande (máx. 80 MB)."}
    ref = _dia_ref(dados.get("mes"))       # mês passado: vai para a pasta do dia 1 do mês seguinte (o mês inteiro)
    entrada = comum.pasta_entrada(cfg, comum.pasta_de(u))
    os.makedirs(entrada, exist_ok=True)
    destino = organizador.destino_livre(os.path.join(entrada, nome))
    with open(destino, "wb") as f:
        f.write(conteudo)
    # Amazon tem 3 relatórios: se o arquivo for um auxiliar (unificado / a receber), usa o tipo certo
    familia = [k for k, v in cfg["marketplaces"].items() if k == mk or v.get("pai") == mk]
    escolhido = dados.get("tipo") if dados.get("tipo") in familia else None
    if not escolhido:
        achado, _ = organizador.identificar(cfg, u, destino)
        escolhido = achado if achado in familia else mk
    if organizador.eh_tratado(destino):
        os.replace(destino, destino + ".tratado-recusado")
        return {"ok": False, "erro": "Este arquivo já é uma planilha tratada (aba DB). Suba o relatório bruto do marketplace."}
    ok, info = organizador.classificar_manual(cfg, u["id"], os.path.basename(destino), escolhido, hoje=ref)
    if not ok:
        return {"ok": False, "erro": "Não consegui organizar: %s (o arquivo ficou na _entrada)." % info}
    comum.registrar_evento({"tipo": "upload_manual", "cliente": u["id"], "mk": escolhido, "origem": nome, "data": ref.isoformat()})
    cons = _consolidar_depois([u["id"]], "relatório subido pelo app", ref) or []
    alvo = cfg["marketplaces"][escolhido].get("pai") or escolhido
    aviso = None
    if ref != dt.date.today():
        mes_ref = "%04d-%02d" % _consolidado().mes_atual(ref)
        ruim = [r for r in cons if r.get("mes") == mes_ref and not r.get("ok")]
        if ruim:
            aviso = "Pacote de %s não gerado: %s Confira se o relatório é mesmo do mês %s/%s (mês inteiro)." % (
                mes_ref[5:] + "/" + mes_ref[:4], ruim[0].get("erro") or "", mes_ref[5:], mes_ref[:4])
    trat = None if info == "duplicado" else next(
        (e for e in reversed(comum.ler_historico(300)) if e.get("tipo") == "tratamento"
         and e.get("cliente") == u["id"] and e.get("mk") == alvo), None)
    return {"ok": True, "mk": escolhido, "mk_nome": cfg["marketplaces"][escolhido]["nome"],
            "arquivo": info if info != "duplicado" else None, "duplicado": info == "duplicado", "aviso": aviso,
            "tratamento": {k: trat.get(k) for k in ("ok", "pendente", "erro", "mk")} if trat else None}


def _lanc():
    import lancamentos
    return lancamentos


def subir_lancamento(dados):
    """ADS / afiliados subidos à mão: guarda por CONTA, marketplace e mês e refaz o consolidado do(s) mês(es)."""
    import base64
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    cli = _cliente_por_id(cfg, u["cliente_id"]) if u else None
    tipo, mk = dados.get("tipo"), dados.get("mk")
    if not u or not cli or tipo not in _lanc().TIPOS or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta, tipo ou marketplace inválido."}
    try:
        conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
    except ValueError:
        return {"ok": False, "erro": "Arquivo corrompido no envio."}
    if not conteudo:
        return {"ok": False, "erro": "Arquivo vazio."}
    nome = os.path.basename(str(dados.get("nome") or "arquivo"))
    if os.path.splitext(nome)[1].lower() not in (".xlsx", ".xlsm", ".csv", ".txt"):
        return {"ok": False, "erro": "Formato não aceito. Use .xlsx ou .csv."}
    if tipo == "visitas":
        # Mercado Livre > Métricas > Negócio: visitas TOTAIS de cada dia (vários dias num arquivo só) -> balizam a conversão
        import tempfile
        fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(nome)[1].lower())
        try:
            with os.fdopen(fd, "wb") as f:
                f.write(conteudo)
            eh_total = _lanc().ler_visitas_total(tmp) is not None
        except ValueError:
            eh_total = True             # é o relatório, mas com problema: guardar_visitas_total dá a mensagem
        except Exception:
            eh_total = False
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        if eh_total:
            if mk != "mercadolivre":
                return {"ok": False, "erro": "Este é o relatório de evolução do negócio do Mercado Livre: suba na linha do Mercado Livre."}
            try:
                meses, info = _lanc().guardar_visitas_total(cfg, _lanc().dono(u), mk, nome, conteudo,
                                                            mes_alvo=_mes_valido(dados.get("mes_alvo")))
            except ValueError as e:
                return {"ok": False, "erro": "Arquivo não aceito: %s" % e}
            import consolidado
            cons = []
            for m in (meses if dados.get("refazer_consolidado", True) else []):
                a, mm = (int(x) for x in m["mes"].split("-"))
                r = consolidado.gerar(cfg, cli, a, mm, origem="visitas totais subidas")
                cons.append({k: r.get(k) for k in ("mes", "ok", "conferido", "erro")})
            return {"ok": True, "meses": meses, "info": info, "consolidado": cons, "dia_nome": None}
    dia = dados.get("dia") or None
    dia_nome = None
    if dados.get("mes_nome"):      # vários dias de uma vez: o dia vem do nome do arquivo
        m = _mes_valido(dados["mes_nome"])
        if not m:
            return {"ok": False, "erro": "Mês inválido."}
        try:
            dia_nome = _lanc()._dia_forcado(tipo, mk, nome) or _lanc().dia_do_nome(nome, int(m[:4]), int(m[5:]))
        except ValueError as e:
            return {"ok": False, "erro": "Arquivo não aceito: %s" % e}
        if not dia_nome:
            return {"ok": False, "erro": "Não achei o dia no nome do arquivo (ex.: 05.xlsx ou 2026-10-05.xlsx)."}
        dia = dia_nome.isoformat()
    try:
        camp = dados.get("campanha") or None
        if dados.get("detalhe"):   # relatório de dentro de uma campanha fechada (por anúncio)
            if tipo != "ads":
                return {"ok": False, "erro": "Detalhe de campanha só vale para ADS."}
            camp = _lanc().DETALHE + str(dados["detalhe"]).strip()[:80]
        elif camp and (tipo != "ads" or not any(c["id"] == camp for c in _lanc().campanhas(cfg, _lanc().dono(u), mk))):
            return {"ok": False, "erro": "Campanha não encontrada."}
        meses, info = _lanc().guardar(cfg, _lanc().dono(u), tipo, mk, nome, conteudo, dia, campanha=camp,
                                      mes_alvo=_mes_valido(dados.get("mes_alvo")))   # só o mês escolhido na tela
    except ValueError as e:
        return {"ok": False, "erro": "Arquivo não aceito: %s" % e}
    except Exception as e:
        return {"ok": False, "erro": "Não consegui ler o arquivo: %s" % e}
    import consolidado
    cons = []
    for m in (meses if dados.get("refazer_consolidado", True) else []):   # vários arquivos: refaz só no último
        a, mm = (int(x) for x in m["mes"].split("-"))
        r = consolidado.gerar(cfg, cli, a, mm, origem="%s subido" % _lanc().TIPOS[tipo]["nome"])
        cons.append({k: r.get(k) for k in ("mes", "ok", "conferido", "erro")})
    if dia_nome and not info.get("sem_data"):
        dias_arq = {d for m in meses for d in m["dias_enviados"]}
        if dia_nome.isoformat() not in dias_arq:
            info["aviso"] = ("O arquivo tem coluna de data (%s) diferente do dia do nome (%s): valeram as datas do arquivo. " % (
                ", ".join(sorted(dias_arq)), dia_nome.strftime("%d/%m/%Y")) + info.get("aviso", "")).strip()
    return {"ok": True, "meses": meses, "info": info, "consolidado": cons, "dia_nome": dia_nome.isoformat() if dia_nome else None}


def subir_depara(dados):
    """De-para da Shopee (planilha mass_update): ID do produto -> SKU, por conta. Refaz o consolidado do mês atual."""
    import base64
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    cli = _cliente_por_id(cfg, u["cliente_id"]) if u else None
    if not u or not cli:
        return {"ok": False, "erro": "Conta inválida."}
    try:
        conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
    except ValueError:
        return {"ok": False, "erro": "Arquivo corrompido no envio."}
    nome = os.path.basename(str(dados.get("nome") or "arquivo"))
    if os.path.splitext(nome)[1].lower() not in (".xlsx", ".xlsm", ".xls", ".csv"):
        return {"ok": False, "erro": "Formato não aceito. Use o .xlsx exportado da Shopee."}
    try:
        mk = dados.get("mk") or "shopee"
        if not _lanc().tem_depara(mk):
            return {"ok": False, "erro": "Este marketplace não usa de-para."}
        info = _lanc().guardar_depara(cfg, _lanc().dono(u), nome, conteudo, mk)
    except ValueError as e:
        return {"ok": False, "erro": "Arquivo não aceito: %s" % e}
    except Exception as e:
        return {"ok": False, "erro": "Não consegui ler o arquivo: %s" % e}
    import consolidado
    a, m = consolidado.mes_atual()
    r = consolidado.gerar(cfg, cli, a, m, origem="de-para %s subido" % cfg["marketplaces"][mk]["nome"])
    return {"ok": True, "info": info, "consolidado": [{k: r.get(k) for k in ("mes", "ok", "conferido", "erro")}]}


def editar_campanha(dados):
    """Campanhas de ADS da conta/marketplace (Magalu e outros, menos Mercado Livre): criar, renomear, desativar, ativar."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    if not u or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    try:
        c = _lanc().editar_campanha(cfg, _lanc().dono(u), mk, dados.get("acao"), dados.get("nome"), dados.get("id"))
    except ValueError as e:
        return {"ok": False, "erro": str(e)}
    return {"ok": True, "campanha": c}


def ok_lancamento(dados):
    """OK manual em dias de ADS/afiliados (ex.: dia sem investimento). dados: conta_id, tipo, mk, dias[], motivo, campanha, remover."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk, tipo = dados.get("mk"), dados.get("tipo")
    if not u or mk not in cfg["marketplaces"] or tipo not in _lanc().TIPOS:
        return {"ok": False, "erro": "Conta, tipo ou marketplace inválido."}
    dias = [d for d in (dados.get("dias") or []) if _dia_valido(d)]
    if not dias:
        return {"ok": False, "erro": "Escolha pelo menos um dia."}
    feitos = _lanc().marcar_ok(cfg, _lanc().dono(u), tipo, mk, dias, dados.get("motivo"), dados.get("campanha") or None,
                               bool(dados.get("remover")))
    return {"ok": True, "dias": feitos}


def depara_pendentes(dados):
    """Para preencher o de-para no app: anúncios do ADS que ficaram sem SKU, os já preenchidos e os SKUs conhecidos."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    if not u or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    m = _mes_valido(dados.get("mes")) or "%04d-%02d" % _consolidado().mes_atual()
    est = (_consolidado().ler_estado().get(u["cliente_id"]) or {}).get(m) or {}
    if not est.get("ok") and est.get("ultimo_ok"):
        est = est["ultimo_ok"]
    if dados.get("tipo") == "visitas":
        sv = ((est.get("lancamentos") or {}).get("visitas") or {}).get("%s|%s" % (u["id"], mk)) or {}
        pend = [{"id": k, "nome": (sv.get("sem_sku_nomes") or {}).get(k, ""), "gasto": v} for k, v in (sv.get("sem_sku_ids") or {}).items()]
    else:
        sd = ((est.get("lancamentos") or {}).get("ads_sem_depara") or {}).get("%s|%s" % (u["id"], mk)) or {}
        pend = [{"id": k, "nome": (sd.get("nomes") or {}).get(k, ""), "gasto": v} for k, v in (sd.get("ids") or {}).items()]
    man = _lanc().depara_manual(cfg, _lanc().dono(u), mk)
    # sugestões: a tabela de custos do mês (tem todos os SKUs vendidos, com descrição)
    sug = []
    try:
        cli = next(c for c in cfg["clientes"] if c["id"] == u["cliente_id"])
        linhas, _, _ = _tratador().ler_tabela_mes(cli, int(m[:4]), int(m[5:]))
        sug = [{"sku": l["SKU"], "desc": l.get("Descrição") or ""} for l in linhas][:5000]
    except Exception:
        pass
    return {"ok": True, "mes": m, "pendentes": pend, "manuais": man, "sugestoes": sug,
            "consolidado_em": est.get("gerado"), "por_nome": mk == "tiktokshop",
            "tipo": dados.get("tipo") or "ads"}


def salvar_depara_app(dados):
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    if not u or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    itens = dados.get("itens") or {}
    if not isinstance(itens, dict) or not itens:
        return {"ok": False, "erro": "Nada para salvar."}
    r = _lanc().salvar_depara_manual(cfg, _lanc().dono(u), mk, itens, por=dados.get("estacao") or "")
    _consolidar_depois([u["id"]], "de-para preenchido no app")
    return dict(r, ok=True)


def limpar_lancamento(dados):
    """Limpa ADS/afiliados de uma conta×marketplace: dias escolhidos ou o mês inteiro (tudo=True)."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk, tipo = dados.get("mk"), dados.get("tipo")
    if not u or mk not in cfg["marketplaces"] or tipo not in _lanc().TIPOS:
        return {"ok": False, "erro": "Conta, tipo ou marketplace inválido."}
    m = _mes_valido(dados.get("mes"))
    if not m:
        return {"ok": False, "erro": "Mês inválido."}
    dias = None if dados.get("tudo") else [d for d in (dados.get("dias") or []) if _dia_valido(d)]
    if dias is not None and not dias:
        return {"ok": False, "erro": "Escolha pelo menos um dia (ou o mês inteiro)."}
    try:
        camp = dados.get("campanha")
        r = _lanc().limpar(cfg, _lanc().dono(u), tipo, mk, int(m[:4]), int(m[5:]), dias,
                           campanha=(str(camp) if camp not in (None, "", False) else None))
    except ValueError as e:
        return {"ok": False, "erro": str(e)}
    a_, m_ = int(m[:4]), int(m[5:])
    if (a_, m_) != _consolidado().mes_atual():       # mês passado: refaz o pacote daquele mês
        cli = _cliente_por_id(cfg, u["cliente_id"])
        if cli:
            _consolidado().gerar(cfg, cli, a_, m_, origem="lançamentos limpos")
    else:
        _consolidar_depois([u["id"]], "lançamentos limpos")
    return dict(r, ok=True)


def limpar_vendas(dados):
    """Tira o relatório de vendas do dia (de um marketplace, ou de todos da conta) para enviar de novo.
    Nada é apagado: bruto, original e tratado vão para [pasta do dia]\\_limpos\\[hora]."""
    import shutil
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    if not u:
        return {"ok": False, "erro": "Conta inválida."}
    dia = dt.date.fromisoformat(dados["dia"]) if _dia_valido(dados.get("dia") or "") else dt.date.today()
    mks = [dados["mk"]] if dados.get("mk") else list(u.get("marketplaces", []))
    if any(m not in cfg["marketplaces"] for m in mks):
        return {"ok": False, "erro": "Marketplace inválido."}
    pasta = comum.pasta_dia(cfg, comum.pasta_de(u), dia)
    rot = comum.rotulo_dia(dia)
    destino = os.path.join(pasta, "_limpos", dt.datetime.now().strftime("%H%M%S"))
    movidos = []
    for mk in mks:
        # o marketplace e os relatórios auxiliares dele (ex.: Amazon unificado / a receber)
        fam = [mk] + [k for k, v in cfg["marketplaces"].items() if v.get("pai") == mk]
        cands = []
        for k in fam:
            cands.append(os.path.join(pasta, comum.nome_arquivo(k, cfg, dia)))
            cands += [os.path.join(pasta, "_originais", f) for f in (os.listdir(os.path.join(pasta, "_originais"))
                      if os.path.isdir(os.path.join(pasta, "_originais")) else []) if f.startswith("vendas_%s_%s" % (k, rot))]
        cands.append(_tratador().caminho_saida(cfg, u, mk, dia))
        for c in cands:
            if os.path.isfile(c):
                os.makedirs(destino, exist_ok=True)
                alvo = os.path.join(destino, os.path.basename(c))
                try:
                    shutil.move(c, alvo)
                except PermissionError:
                    return {"ok": False, "erro": "Feche o arquivo %s no Excel e tente de novo." % os.path.basename(c)}
                movidos.append(os.path.basename(c))
        if comum.ok_vendas(u["id"], mk, dia):
            comum.marcar_ok_vendas(u["id"], mk, dia, remover=True)
            movidos.append("OK manual (%s)" % mk)
    if not movidos:
        return {"ok": False, "erro": "Não há relatório de %s para limpar em %s." % (
            ", ".join(cfg["marketplaces"][m]["nome"] for m in mks), dia.strftime("%d/%m/%Y"))}
    comum.log("[%s] vendas limpas em %s (%s): %s" % (u["nome"], rot, ", ".join(mks), ", ".join(movidos)))
    comum.registrar_evento({"tipo": "vendas_limpas", "cliente": u["id"], "mks": mks, "dia": dia.isoformat()})
    _consolidar_depois([u["id"]], "vendas limpas", dia)
    return {"ok": True, "movidos": movidos, "pasta": destino if os.path.isdir(destino) else None}


def ok_vendas(dados):
    """OK manual no relatório de vendas do dia (ex.: marketplace sem venda / sem relatório)."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    if not u or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    dia = dados.get("dia") or dt.date.today().isoformat()
    if not _dia_valido(dia):
        return {"ok": False, "erro": "Dia inválido."}
    comum.marcar_ok_vendas(u["id"], mk, dia, dados.get("motivo"), bool(dados.get("remover")))
    comum.log("[%s] OK manual de vendas %s em %s %s" % (u["nome"], mk, dia, "removido" if dados.get("remover") else "dado"))
    _consolidar_depois([u["id"]], "OK manual de vendas", dt.date.fromisoformat(dia[:10]))
    return {"ok": True}


def _dia_valido(d):
    try:
        dt.date.fromisoformat(str(d)[:10])
        return True
    except ValueError:
        return False


def visao_meses(dados):
    """Mês a mês: o que está alimentado (vendas tratadas, ADS, afiliados, custos, forecast, consolidado)."""
    import consolidado, tratador, forecast
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    if not cli:
        return {"ok": False, "erro": "Cliente não encontrado."}
    quantos = max(1, min(24, int(dados.get("quantos") or 12)))
    hoje = dt.date.today()
    est = consolidado.ler_estado().get(cli["id"], {})
    contas = [u for u in comum.contas_do_cliente(cfg, cli) if u.get("ativo", True)]
    meses = []
    for m in consolidado.meses_disponiveis(cfg, hoje, quantos):
        a, mm = int(m["chave"][:4]), int(m["chave"][5:])
        item = {"chave": m["chave"], "rotulo": m["rotulo"], "atual": m["chave"] == "%04d-%02d" % consolidado.mes_atual(hoje)}
        item["vendas"] = consolidado.cobertura_vendas(cfg, cli, a, mm, hoje)
        lanc = {"ads": [], "afiliados": [], "visitas": []}
        for u in contas:
            sit = _lanc().situacao(cfg, _lanc().dono(u), a, mm)
            esp = sit.get("dias_esperados", 0)
            for tipo in ("ads", "afiliados", "visitas"):
                for mk in u.get("marketplaces", []):
                    x = sit.get(tipo, {}).get(mk)
                    falt = len(x.get("faltam", [])) if x else esp
                    lanc[tipo].append({"conta": u["conta_nome"], "conta_id": u["id"], "mk": mk,
                                       "nome": cfg["marketplaces"].get(mk, {}).get("nome", mk),
                                       "enviado": bool(x), "unico": bool(x and x.get("unico")),
                                       "esperados": esp, "faltam": 0 if (x and x.get("unico")) else falt,
                                       "ok_manual": len((x or {}).get("ok_manual", []))})
        item.update(lanc)
        info, mes_usado = tratador.escolher_custos(cli, a, mm)
        item["custos"] = {"tem": bool(info), "mes_usado": mes_usado, "proprio": mes_usado == m["chave"]}
        d = forecast.ler_definido(cli["id"], a, mm)
        item["forecast"] = {"definido": bool(d), "versao": d.get("versao") if d else None}
        r = est.get(m["chave"])
        item["consolidado"] = {k: r.get(k) for k in ("ok", "conferido", "gerado", "erro", "parcial", "pronto")} if r else None
        meses.append(item)
    return {"ok": True, "meses": meses}


def painel_controle(dados):
    """Números do mês para o painel (lidos do consolidado conferido). Sem cliente: uma linha por cliente."""
    import painel_dados, consolidado
    cfg = comum.carregar_config()
    m = _mes_valido(dados.get("mes") or "") or "%04d-%02d" % consolidado.mes_atual()
    a, mm = int(m[:4]), int(m[5:])
    try:
        if dados.get("cliente_id"):
            cli = _cliente_por_id(cfg, dados["cliente_id"])
            if not cli:
                return {"ok": False, "erro": "Cliente não encontrado."}
            return painel_dados.resumo(cfg, cli, a, mm, filtro=dados.get("filtro"))
        return {"ok": True, "mes": m, "clientes": painel_dados.resumo_todos(cfg, a, mm, filtro=dados.get("filtro"))}
    except Exception as e:
        comum.log("painel: %s" % e)
        return {"ok": False, "erro": "Não consegui montar o painel: %s" % str(e)[:300]}


def painel_conferir(dados):
    """Marca/desmarca um marketplace (e conta) como conferido pelos números macro no período mostrado."""
    import painel_dados
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    m = _mes_valido(dados.get("mes") or "")
    if not cli or not m or not dados.get("marketplace"):
        return {"ok": False, "erro": "Cliente, mês ou marketplace inválido."}
    painel_dados.salvar_conferencia(cli["id"], m, dados["marketplace"], dados.get("conta") or "", dados.get("de"), dados.get("ate"),
                                    dados.get("valores"), dados.get("informados"), dados.get("nota"), dados.get("por"),
                                    remover=bool(dados.get("remover")))
    comum.log("[%s] painel: %s %s %s (%s) %s" % (cli["nome"], m, dados["marketplace"], dados.get("conta") or "",
                                                 "%s a %s" % (dados.get("de") or "início", dados.get("ate") or "fim"),
                                                 "conferência retirada" if dados.get("remover") else "conferido"))
    return {"ok": True}


def traducao_listar(dados):
    """Todos os anúncios das vendas tratadas do mês (por marketplace e conta), com o SKU que veio da plataforma e a
    tradução feita no app, para corrigir SKU errado cadastrado pelo cliente."""
    import pandas as pd
    import consolidado
    trat = _tratador()
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    m = _mes_valido(dados.get("mes") or "") or trat.mes_chave(*trat.mes_do_periodo())
    if not cli:
        return {"ok": False, "erro": "Escolha o cliente no topo da tela."}
    dia = _dia_ref(m)
    regras = trat.ler_traducoes(cli["id"])
    linhas, faltam = [], []
    for u in comum.filtrar_contas(cfg, [cli["id"]], so_ativas=False):
        for mk in u.get("marketplaces", []):
            if mk not in trat.SKU_TRATADO or mk not in consolidado.MAPA:
                continue
            arq = trat.caminho_saida(cfg, u, mk, dia)
            nome_mk = cfg["marketplaces"].get(mk, {}).get("nome", mk)
            if not os.path.isfile(arq):
                faltam.append("%s · %s" % (u["nome"], nome_mk))
                continue
            M = consolidado.MAPA[mk]
            col_sku, col_tit = trat.SKU_TRATADO[mk]
            try:
                df = consolidado.ler_excel(arq, sheet_name="DB")
            except Exception as e:
                faltam.append("%s · %s (não abriu: %s)" % (u["nome"], nome_mk, str(e)[:80]))
                continue
            txt = lambda c: df[c].map(trat._chave_an) if c in df.columns else pd.Series([""] * len(df), index=df.index)
            num = lambda c: pd.to_numeric(df[c], errors="coerce").fillna(0.0) if c in df.columns else pd.Series([0.0] * len(df), index=df.index)
            t = pd.DataFrame({"an": txt(M.get("anuncio")), "sku": txt(col_sku),
                              "tit": df[col_tit].astype(str) if col_tit in df.columns else "",
                              "venda": (df[M["status"]].astype(str).str.strip().str.lower() == "venda") if M.get("status") in df.columns else True,
                              "qtd": num(M.get("qtd")), "fat": num(M.get("fat")), "custo": num("Custo do produto")})
            t.loc[t["sku"].str.lower().isin(["nan", "none"]), "sku"] = ""
            for (an, sku), g in t.groupby(["an", "sku"], sort=False):
                v = g[g["venda"]]
                qtd = float(v["qtd"].sum())
                regra, chave = None, None
                for k in (trat.chave_traducao(mk, an, "x"), trat.chave_traducao(mk, "*", "x")):
                    pre = k[:-1]
                    for kk, rr in regras.items():
                        if kk.startswith(pre) and trat._chave_sku(rr.get("sku")) == trat._chave_sku(sku):
                            regra, chave = rr, kk
                            break
                    if regra:
                        break
                linhas.append({"mk": mk, "mk_nome": nome_mk, "conta": u["nome"], "anuncio": an, "titulo": str(g["tit"].iloc[0])[:200],
                               "sku": sku, "sku_original": regra["sku_original"] if regra else sku,
                               "sku_correto": regra["sku"] if regra else "", "escopo": ("sku" if regra and regra.get("anuncio") == "*" else "anuncio"),
                               "regra": chave, "regra_quando": (regra or {}).get("quando"),
                               "qtd": qtd, "fat": round(float(v["fat"].sum()), 2), "linhas": int(len(g)),
                               "custo_unit": round(-float(v["custo"].sum()) / qtd, 2) if qtd else 0.0,
                               "sem_custo": bool(qtd and not float(v["custo"].sum()))})
    linhas.sort(key=lambda x: (-x["fat"], x["mk"], x["anuncio"]))
    sug = []
    try:
        tab, _, _ = trat.ler_tabela_mes(cli, int(m[:4]), int(m[5:]))
        sug = [{"sku": l["SKU"], "desc": l.get("Descrição") or "", "valor": l.get("Custo")} for l in tab][:8000]
    except Exception:
        pass
    return {"ok": True, "mes": m, "dia": dia.isoformat(), "linhas": linhas, "faltam": faltam, "sugestoes": sug,
            "regras": len(regras)}


def traducao_salvar(dados):
    """Grava as traduções e refaz os tratados do mês e o pacote final."""
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    m = _mes_valido(dados.get("mes") or "")
    if not cli or not m:
        return {"ok": False, "erro": "Cliente ou mês inválido."}
    n = _tratador().salvar_traducoes(cli["id"], dados.get("itens") or [], dados.get("por") or "")
    comum.log("[%s] tradução de SKU: %d regra(s) alterada(s) (%s)" % (cli["nome"], n, m))
    ret = _retratar_mes(cli["id"], m) if n else []
    return {"ok": True, "alteradas": n, "retratados": ret,
            "erros": [r for r in ret if not r.get("ok")]}


def historico_cargas(dados):
    import painel_dados
    cfg = comum.carregar_config()
    return {"ok": True, "cargas": painel_dados.cargas(cfg, dados.get("cliente_id") or None, int(dados.get("limite") or 60))}


def resumo_tratado(dados):
    """Raio-x de um tratado (o que tem no arquivo + checagens)."""
    import painel_dados
    cfg = comum.carregar_config()
    p = os.path.abspath(str(dados.get("arquivo") or ""))
    raiz = os.path.abspath(comum.raiz(cfg))
    if not p.startswith(raiz + os.sep) or not p.lower().endswith(".xlsx") or not os.path.isfile(p):
        return {"ok": False, "erro": "Arquivo não encontrado."}
    try:
        return painel_dados.resumo_tratado(p, dados.get("mk"))
    except Exception as e:
        return {"ok": False, "erro": "Não consegui ler o tratado: %s" % str(e)[:300]}


def ajuste_ads(dados):
    """Ajuste manual de ADS (campanha apagada): acao salvar | remover. Refaz o consolidado dos meses tocados."""
    cfg = comum.carregar_config()
    u = comum.conta(cfg, dados.get("conta_id"))
    mk = dados.get("mk")
    cli = _cliente_por_id(cfg, u["cliente_id"]) if u else None
    if not u or not cli or mk not in cfg["marketplaces"]:
        return {"ok": False, "erro": "Conta ou marketplace inválido."}
    try:
        if dados.get("acao") == "remover":
            item = _lanc().remover_ajuste(cfg, _lanc().dono(u), dados.get("id"))
        else:
            item = _lanc().salvar_ajuste(cfg, _lanc().dono(u), mk, dados.get("nome"), dados.get("gasto"), dados.get("de"), dados.get("ate"),
                                         dados.get("receita"), dados.get("cliques"), dados.get("motivo"), dados.get("id") or None,
                                         tipo=dados.get("tipo") or "ads")
    except (ValueError, TypeError) as e:
        return {"ok": False, "erro": str(e) or "Valores inválidos."}
    import consolidado
    d1, d2 = dt.date.fromisoformat(item["de"]), dt.date.fromisoformat(item["ate"])
    meses, d = set(), d1
    while d <= d2:
        meses.add((d.year, d.month))
        d += dt.timedelta(days=1)
    cons = []
    for a, m in sorted(meses):
        r = consolidado.gerar(cfg, cli, a, m, origem="ajuste manual de %s" % ("afiliados" if item.get("tipo") == "afiliados" else "ADS"))
        cons.append({k: r.get(k) for k in ("mes", "ok", "conferido", "erro")})
    return {"ok": True, "ajuste": item, "consolidado": cons}


def reprocessar_lancamentos(dados):
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    if not cli:
        return {"ok": False, "erro": "Selecione um cliente."}
    _lanc().migrar_legado(cfg, cli)
    meses, erros = [], []
    for u in comum.contas_do_cliente(cfg, cli):
        r = _lanc().reprocessar(cfg, _lanc().dono(u))
        meses += r["meses"]
        erros += r["erros"]
    import consolidado
    for mes in sorted({m["mes"] for m in meses}):
        consolidado.gerar(cfg, cli, int(mes[:4]), int(mes[5:]), origem="ADS/afiliados reprocessados")
    return {"ok": True, "meses": meses, "erros": erros}


def _reprocessar_antigos():
    """Na abertura do app: se a leitura de ADS/afiliados melhorou, relê os arquivos já enviados (uma vez)."""
    try:
        cfg = comum.carregar_config()
        import consolidado
        for cli in cfg.get("clientes", []):
            mudou = _lanc().migrar_legado(cfg, cli)
            meses = set()
            for u in comum.contas_do_cliente(cfg, cli):
                d = _lanc().dono(u)
                if os.path.isdir(os.path.join(_lanc().pasta(cfg, d), "_originais")) and (mudou or _lanc().versao_dados(cfg, d) < _lanc().LEITOR_VERSAO):
                    meses |= {m["mes"] for m in _lanc().reprocessar(cfg, d)["meses"]}
            for mes in sorted(meses):
                consolidado.gerar(cfg, cli, int(mes[:4]), int(mes[5:]), origem="ADS/afiliados por conta")
    except Exception as e:
        comum.log("Reprocessar ADS/afiliados falhou: %s" % e)


REGRA_NAO_VENDAS = 8   # sobe quando muda a regra de cancelada/devolução/não pago no tratador: refaz os tratados de hoje (uma vez)


def _retratar_regra_nao_vendas():
    """Na abertura do app (computador principal): se a regra das linhas que não são venda mudou, refaz os tratados
    de hoje que já existem (o bruto continua igual) e o pacote final dos clientes afetados."""
    arq = os.path.join(comum.PASTA_SISTEMA, "dados", "regras_tratamento.json")
    try:
        with open(arq, "r", encoding="utf-8") as f:
            feito = json.load(f).get("nao_vendas", 0)
    except (OSError, ValueError):
        feito = 0
    if feito >= REGRA_NAO_VENDAS:
        return
    try:
        cfg = comum.carregar_config()
        hoje = dt.date.today()
        afetados = set()
        for u in comum.filtrar_contas(cfg, None, so_ativas=False):
            mks = [m for m in u.get("marketplaces", []) if m in cfg["marketplaces"]
                   and os.path.isfile(_tratador().caminho_saida(cfg, u, m, hoje))]
            if mks:
                _tratador().tratar_cliente(cfg, u, hoje, mks, "regra nova: cancelada/devolução/não pago")
                afetados.add(u["cliente_id"])
        if afetados:
            _consolidar_depois(sorted(afetados), "regra nova: cancelada/devolução/não pago")
        os.makedirs(os.path.dirname(arq), exist_ok=True)
        with open(arq, "w", encoding="utf-8") as f:
            json.dump({"nao_vendas": REGRA_NAO_VENDAS, "em": dt.datetime.now().isoformat(timespec="seconds")}, f)
        comum.log("Tratados de hoje refeitos com a regra nova de cancelada/devolução/não pago (%d cliente(s))" % len(afetados))
    except Exception as e:
        comum.log("Refazer tratados com a regra nova falhou: %s" % e)


def situacao_lancamentos(dados):
    cfg = comum.carregar_config()
    cli = next((c for c in cfg["clientes"] if c["id"] == dados.get("cliente_id")), None)
    if not cli:
        return {"ok": False, "erro": "Cliente não encontrado."}
    try:
        a, m = (int(x) for x in str(dados.get("mes")).split("-"))
    except ValueError:
        return {"ok": False, "erro": "Mês inválido."}
    return {"ok": True, "situacao": {u["id"]: _lanc().situacao(cfg, _lanc().dono(u), a, m) for u in comum.contas_do_cliente(cfg, cli)}}


def modelo_lancamento(dados):
    if dados.get("tipo") not in _lanc().TIPOS:
        return {"ok": False, "erro": "Tipo inválido."}
    p = _lanc().modelo(comum.carregar_config(), dados["tipo"])
    if os.name == "nt":
        os.startfile(p)
    return {"ok": True, "caminho": p}


def _cliente_por_id(cfg, cid):
    return next((c for c in cfg["clientes"] if c["id"] == cid), None)


def forecast_simular(dados):
    import forecast
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    if not cli:
        return {"ok": False, "erro": "Selecione um cliente."}
    hoje = dt.date.today()
    if dados.get("ajustes") is not None:
        forecast.salvar_ajustes(cli["id"], hoje.year, hoje.month, dados["ajustes"])
    alvo = None
    try:
        alvo = float(str(dados.get("alvo")).replace(",", ".")) / 100 if dados.get("alvo") not in (None, "") else None
    except ValueError:
        alvo = None
    sim = forecast.simular(cfg, cli, hoje, alvo=alvo)
    if sim.get("erro"):
        return {"ok": False, "erro": sim["erro"]}
    return {"ok": True, "simulacao": sim, "definido": forecast.comparar(cfg, cli, hoje)}


def forecast_definir(dados):
    import forecast
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    if not cli:
        return {"ok": False, "erro": "Selecione um cliente."}
    hoje = dt.date.today()
    ja = forecast.ler_definido(cli["id"], hoje.year, hoje.month)
    motivo = (dados.get("motivo") or "").strip()
    if ja and not motivo:
        return {"ok": False, "erro": "O forecast deste mês já está definido. Para alterar, informe o motivo."}
    return forecast.definir(cfg, cli, hoje, motivo)


def gerar_consolidado(dados):
    import consolidado
    cfg = comum.carregar_config()
    ano = mes = None
    if dados.get("mes"):
        try:
            ano, mes = (int(x) for x in str(dados["mes"]).split("-"))
            assert 1 <= mes <= 12
        except (ValueError, AssertionError):
            return {"ok": False, "erro": "Mês inválido."}
    ids = [dados["cliente_id"]] if dados.get("cliente_id") else None
    res = consolidado.gerar_clientes(cfg, ids, ano, mes, origem="botão do app")
    return {"ok": True, "resultados": [{k: r.get(k) for k in ("cliente", "mes", "ok", "conferido", "erro", "linhas")} for r in res]}


def pacote_pronto(dados):
    """Pacote final pronto do cliente no mês (subir ou remover). Com ele, o pacote do sistema sai do arquivo subido
    e só as visitas alimentadas no sistema entram por cima."""
    import base64
    import consolidado
    cfg = comum.carregar_config()
    cli = _cliente_por_id(cfg, dados.get("cliente_id"))
    m = _mes_valido(dados.get("mes"))
    if not cli or not m:
        return {"ok": False, "erro": "Cliente ou mês inválido."}
    a, mm = int(m[:4]), int(m[5:])
    try:
        if dados.get("remover"):
            r = consolidado.remover_pronto(cfg, cli, a, mm)
        else:
            try:
                conteudo = base64.b64decode(dados.get("conteudo") or "", validate=True)
            except ValueError:
                return {"ok": False, "erro": "Arquivo corrompido no envio."}
            nome = os.path.basename(str(dados.get("nome") or "pacote.xlsx"))
            if not nome.lower().endswith((".xlsx", ".xlsm", ".xls")):
                return {"ok": False, "erro": "Envie o pacote em Excel (.xlsx)."}
            r = consolidado.guardar_pronto(cfg, cli, a, mm, nome, conteudo)
    except ValueError as e:
        return {"ok": False, "erro": "Pacote não aceito: %s" % e}
    return {"ok": bool(r.get("ok")), "erro": r.get("erro"), "linhas": r.get("linhas"), "conferido": r.get("conferido"),
            "aviso": r.get("aviso_pronto"), "visitas_do_sistema": r.get("visitas_do_sistema", [])}


def tratar_agora(dados):
    cfg = comum.carregar_config()
    clientes = comum.filtrar_contas(cfg, [dados["cliente_id"]] if dados.get("cliente_id") else None)
    ref = _dia_ref(dados.get("mes"))
    res = []
    for c in clientes:
        mks = [dados["mk"]] if dados.get("mk") else None
        for e in _tratador().tratar_cliente(cfg, c, ref, mks=mks, origem="botão do app"):
            res.append({"cliente": c["nome"], "mk": e["mk"], "ok": e.get("ok"), "pendente": e.get("pendente"), "erro": e.get("erro")})
    _consolidar_depois([c["id"] for c in clientes], "tratamento pelo app", ref)
    return {"ok": True, "resultados": res}


ACOES = {
    "/api/custos/enviar": enviar_custos,
    "/api/custos/copiar": copiar_custos,
    "/api/custos/modelo": modelo_custos,
    "/api/custos/tabela": tabela_custos,
    "/api/custos/titulos": titulos_custos,
    "/api/custos/salvar": salvar_tabela_custos,
    "/api/custos/exportar": exportar_custos_editor,
    "/api/custos/importar": importar_custos_editor,
    "/api/planilhas/lista": planilhas_lista,
    "/api/planilhas/ler": planilhas_ler,
    "/api/cliente/imposto": salvar_imposto,
    "/api/cliente/antecipa": salvar_antecipa,
    "/api/tratar": tratar_agora,
    "/api/consolidado/gerar": gerar_consolidado,
    "/api/consolidado/pronto": pacote_pronto,
    "/api/vendas/subir": subir_relatorio,
    "/api/lancamento/subir": subir_lancamento,
    "/api/forecast/simular": forecast_simular,
    "/api/forecast/definir": forecast_definir,
    "/api/lancamento/situacao": situacao_lancamentos,
    "/api/lancamento/reprocessar": reprocessar_lancamentos,
    "/api/lancamento/depara": subir_depara,
    "/api/lancamento/campanha": editar_campanha,
    "/api/lancamento/ok": ok_lancamento,
    "/api/lancamento/ajuste": ajuste_ads,
    "/api/vendas/ok": ok_vendas,
    "/api/vendas/limpar": limpar_vendas,
    "/api/lancamento/limpar": limpar_lancamento,
    "/api/lancamento/depara/pendentes": depara_pendentes,
    "/api/lancamento/depara/salvar": salvar_depara_app,
    "/api/meses": visao_meses,
    "/api/painel": painel_controle,
    "/api/painel/conferir": painel_conferir,
    "/api/traducao/listar": traducao_listar,
    "/api/traducao/salvar": traducao_salvar,
    "/api/cargas": historico_cargas,
    "/api/tratado/resumo": resumo_tratado,
    "/api/lancamento/modelo": modelo_lancamento,
    "/api/executar": lambda d: _executor().rodar_em_segundo_plano([d["cliente_id"]] if d.get("cliente_id") else None,
                                                                  d.get("mk"), "botão do app"),
    "/api/executar/parar": lambda d: _executor().parar(),
    "/api/executar/agendado": lambda d: _executor().rodar_em_segundo_plano(None, None, "agendamento", so_locais=True),
    "/api/claude/instalar": lambda d: _executor().instalar_claude(),
    "/api/claude/login": lambda d: _executor().login_claude(),
    "/api/claude/status": lambda d: {"ok": True, **_executor().status_claude()},
    "/api/claude/testar": lambda d: _executor().testar_conexao(d.get("cliente_id")),
    "/api/abrir-no-perfil": abrir_no_perfil,
    "/api/janela/mostrar": mostrar_janela,
    "/api/sistema/atualizar": atualizar_sistema,
    "/api/sistema/arquivos": arquivos_sistema,
    "/api/cliente": salvar_cliente,
    "/api/cliente/remover": remover_cliente,
    "/api/ajustes": salvar_ajustes,
    "/api/downloads": alterar_downloads,
    "/api/organizar": lambda d: {"ok": True, "resumo": organizador.executar(forcar_estavel=True)},
    "/api/classificar": classificar,
    "/api/abrir-pasta": abrir_pasta,
    "/api/agendar": reaplicar_agendamento,
    "/api/abrir-chrome": abrir_chrome_agora,
    "/api/rede/ler": ler_rede,
    "/api/rede/salvar": salvar_rede,
    "/api/rede/testar": testar_rede,
    "/api/rede/firewall": liberar_firewall,
    "/api/agente/config": agente_config,
    "/api/agente/arquivo": agente_arquivo,
    "/api/agente/organizar": agente_organizar,
    "/api/agente/faltando": agente_faltando,
    "/api/agente/roteiro": agente_roteiro,
    "/api/agente/evento": agente_evento,
}
# ações que só fazem sentido no próprio computador (nunca pela rede)
SO_LOCAL = {"/api/rede/salvar", "/api/rede/firewall", "/api/janela/mostrar", "/api/sistema/atualizar", "/api/agendar", "/api/claude/instalar", "/api/claude/login"}
# ações que mexem no Chrome deste computador: a estação faz as dela sozinha; as que chegam aqui pela rede são de contas daqui
NO_CHROME_DAQUI = {"/api/abrir-no-perfil", "/api/abrir-chrome", "/api/downloads"}



PAGINA_LOGIN = """<!doctype html><html lang="pt-br"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Entrar · Relatórios de Vendas</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#f3f5f9;--card:#fff;--tx:#141a26;--mut:#5d6779;--li:#dfe4ec;--pri:#2c5bd9;--pri2:#1f46b3;--err:#b42318;--errbg:#fdecea}
@media (prefers-color-scheme:dark){:root{--bg:#0d1320;--card:#151d2c;--tx:#e8edf6;--mut:#94a0b6;--li:#26324a;--pri:#5b8cff;--pri2:#4a78e8;--err:#ff8a80;--errbg:#3a1d1f}}
*{box-sizing:border-box}html,body{height:100%}
body{margin:0;font-family:Inter,Segoe UI,system-ui,sans-serif;background:var(--bg);color:var(--tx);display:grid;place-items:center;padding:16px;
  background-image:radial-gradient(circle at 15% 20%,color-mix(in srgb,var(--pri) 14%,transparent),transparent 40%),radial-gradient(circle at 85% 80%,color-mix(in srgb,var(--pri) 10%,transparent),transparent 45%)}
.card{width:100%;max-width:380px;background:var(--card);border:1px solid var(--li);border-radius:16px;padding:32px 28px;box-shadow:0 20px 50px -20px rgba(15,25,50,.35)}
.logo{display:flex;align-items:center;gap:12px;margin-bottom:22px}
.logo i{width:42px;height:42px;border-radius:11px;background:var(--pri);display:grid;place-items:center}
.logo svg{width:24px;height:24px;stroke:#fff;fill:none;stroke-width:2;stroke-linecap:round}
.logo b{font-size:18px;display:block}.logo span{font-size:13px;color:var(--mut)}
h1{font-size:22px;margin:0 0 4px}p.sub{margin:0 0 22px;color:var(--mut);font-size:14px}
label{display:block;font-size:13px;font-weight:600;margin:0 0 6px}
input{width:100%;font:inherit;font-size:15px;padding:11px 12px;border:1px solid var(--li);border-radius:10px;background:transparent;color:var(--tx);margin-bottom:16px;outline:none}
input:focus{border-color:var(--pri);box-shadow:0 0 0 3px color-mix(in srgb,var(--pri) 22%,transparent)}
.senha{position:relative}.senha button{position:absolute;right:8px;top:7px;border:0;background:none;color:var(--mut);font:inherit;font-size:12px;cursor:pointer;padding:6px}
button.entrar{width:100%;font:inherit;font-weight:600;font-size:15px;padding:12px;border:0;border-radius:10px;background:var(--pri);color:#fff;cursor:pointer}
button.entrar:hover{background:var(--pri2)}
.erro{background:var(--errbg);color:var(--err);border-radius:10px;padding:10px 12px;font-size:13px;margin-bottom:16px}
.erro:empty{display:none}.rodape{margin-top:18px;text-align:center;font-size:12px;color:var(--mut)}
</style></head><body><main class="card">
<div class="logo"><i><svg viewBox="0 0 24 24"><path d="M5 19V11M10 19V6M15 19v-6M20 19V9"/></svg></i><div><b>Relatórios de Vendas</b><span>Marketplaces</span></div></div>
<h1>Entrar</h1><p class="sub">Use o seu usuário e a sua senha.</p>
<div class="erro">{ERRO}</div>
<form method="post" action="/login">
<label for="u">Usuário</label><input id="u" name="usuario" autocomplete="username" autocapitalize="none" spellcheck="false" required autofocus>
<label for="s">Senha</label><div class="senha"><input id="s" type="password" name="senha" autocomplete="current-password" required>
<button type="button" onclick="var i=document.getElementById('s');i.type=i.type==='password'?'text':'password';this.textContent=i.type==='password'?'mostrar':'ocultar'">mostrar</button></div>
<button class="entrar">Entrar</button></form>
<div class="rodape">Esqueceu a senha? Peça ao admin para criar uma nova.</div>
</main></body></html>"""


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    server_version = "PainelRelatorios/1.0"

    def log_message(self, *a):
        pass

    def _local(self):
        if comum.NUVEM:
            return False         # online: todo acesso vem de fora e precisa do login
        return self.client_address[0] in ("127.0.0.1", "::1", "::ffff:127.0.0.1")

    def _host_ok(self):
        porta = self.server.server_address[1]
        return self.headers.get("Host", "") in ("localhost:%d" % porta, "127.0.0.1:%d" % porta)

    def _sessao_ok(self):
        """Online: cookie de sessão de um usuário (login e senha próprios)."""
        self._usuario = None
        if not comum.NUVEM:
            return False
        import usuarios
        for parte in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = parte.strip().partition("=")
            if k == "rv_sessao":
                self._usuario = usuarios.sessao(v)
                return self._usuario is not None
        return False

    def _token_ok(self):
        if self._sessao_ok():
            return True
        r = _rede()
        if not r or r.papel() != "servidor":
            return False
        t = r.token()
        if not t:
            return False
        import hmac
        auth = self.headers.get("Authorization", "")
        if auth.startswith("Bearer ") and hmac.compare_digest(auth[7:].strip(), t):
            return True
        for parte in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = parte.strip().partition("=")
            if k == "rv_token" and hmac.compare_digest(v, t):
                return True
        return False

    def _permitido(self):
        """local: só com Host localhost (proteção contra sites maliciosos). Rede: precisa da senha."""
        if self._local() and self._host_ok():
            return True
        if comum.NUVEM:
            ip = self.client_ip()
            if _limitar_login(ip):
                return False
            ok = self._token_ok()
            if not ok and ("Bearer " in (self.headers.get("Authorization") or "") or "rv_token=" in (self.headers.get("Cookie") or "")
                           or "rv_sessao=" in (self.headers.get("Cookie") or "")):
                _limitar_login(ip, falhou=True)      # chave errada também conta como tentativa
            return ok
        return not self._local() and self._token_ok()

    def _remoto(self):
        return not self._local()

    def client_ip(self):
        return (self.headers.get("X-Forwarded-For") or self.client_address[0]).split(",")[0].strip()

    def _identidade(self):
        """(quem, nome da pessoa, computador) de quem fez o pedido."""
        from urllib.parse import unquote
        u = getattr(self, "_usuario", None)
        if u:
            return "web|%s" % u["login"], u.get("nome") or u["login"], "online"
        if self._remoto():
            pc = unquote(self.headers.get("X-Estacao") or "") or ("navegador " + self.client_address[0])
            usuario = unquote(self.headers.get("X-Usuario") or "") or pc
        else:
            r = _rede()
            d = r.ler() if r else {"nome": "este computador", "usuario": ""}
            pc, usuario = d["nome"], d.get("usuario") or d["nome"]
        return "%s|%s" % (pc, usuario), usuario, pc

    def _login(self, erro=""):
        if comum.NUVEM:
            if not os.environ.get("ADMIN_SENHA"):
                erro = erro or "Senha do admin ainda não definida (variável ADMIN_SENHA no Render)."
            return self._enviar(200 if not erro else 401, PAGINA_LOGIN.replace("{ERRO}", html.escape(erro)),
                                "text/html; charset=utf-8")
        corpo = ("<h1>Relatórios de Vendas</h1><p>Este é o computador principal. Digite a senha da rede para entrar.</p>"
                 "<form method='post' action='/login'><input type='password' name='senha' autofocus "
                 "style='font-size:1.1em;padding:8px;width:240px'> <button style='font-size:1.1em;padding:8px 16px'>Entrar</button></form>"
                 "<p style='color:#b42318'>%s</p>" % html.escape(erro))
        return self._enviar(200 if not erro else 401, self._pagina("Entrar", corpo), "text/html; charset=utf-8")

    def _usuarios(self, caminho):
        """Cadastro de usuários do online (admin) e troca da própria senha (todos)."""
        import usuarios
        u = getattr(self, "_usuario", None)
        tam = int(self.headers.get("Content-Length") or 0)
        dados = json.loads(self.rfile.read(tam) or b"{}") if tam else {}
        if not comum.NUVEM or not u:
            return {"ok": False, "erro": "Disponível só no sistema online, com login."}
        try:
            if caminho == "/api/usuarios/minha-senha":
                usuarios.trocar_minha_senha(u["login"], dados.get("atual"), dados.get("nova"))
                comum.log("Usuário %s trocou a própria senha." % u["login"])
                return {"ok": True}
            if u.get("perfil") != "admin":
                return {"ok": False, "erro": "Só um admin pode cadastrar usuários."}
            if caminho == "/api/usuarios/listar":
                return {"ok": True, "usuarios": usuarios.listar()}
            if caminho == "/api/usuarios/salvar":
                r = usuarios.salvar(dados, u["login"])
                comum.log("Usuário %s salvo por %s." % (r["login"], u["login"]))
                return {"ok": True, "usuario": r}
            if caminho == "/api/usuarios/remover":
                if (dados.get("login") or "").lower() == u["login"]:
                    return {"ok": False, "erro": "Você não pode remover o seu próprio usuário."}
                ok = usuarios.remover(dados.get("login"))
                comum.log("Usuário %s removido por %s." % (dados.get("login"), u["login"]))
                return {"ok": ok}
        except ValueError as e:
            return {"ok": False, "erro": str(e)}
        return {"ok": False, "erro": "ação desconhecida"}

    def _enviar(self, codigo, corpo, tipo="application/json; charset=utf-8"):
        if isinstance(corpo, (dict, list)):
            corpo = json.dumps(corpo, ensure_ascii=False)
        dados = corpo.encode("utf-8") if isinstance(corpo, str) else corpo
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(dados)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(dados)

    def _pagina(self, titulo, corpo):
        return ("<!doctype html><html lang='pt-br'><head><meta charset='utf-8'><title>%s</title>"
                "<style>body{font-family:Segoe UI,system-ui,sans-serif;max-width:640px;margin:48px auto;padding:0 16px;color:#1d2433}"
                "b{font-size:1.6em}code{background:#eef1f6;padding:2px 6px;border-radius:4px}</style></head><body>%s</body></html>"
                % (html.escape(titulo), corpo))

    def do_GET(self):
        url = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(url.query).items()}
        if url.path == "/saude":
            return self._enviar(200, {"ok": True})
        if url.path == "/sair":
            self.send_response(303)
            self.send_header("Set-Cookie", "rv_token=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0" + ("; Secure" if comum.NUVEM else ""))
            if comum.NUVEM:
                self.send_header("Set-Cookie", "rv_sessao=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0; Secure")
            self.send_header("Location", "/")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if not self._permitido():
            if self._remoto() and url.path in ("/", "/index.html", "/login") and _rede() and _rede().papel() == "servidor":
                return self._login()
            return self._enviar(403 if self._local() else 401, {"erro": "host não permitido" if self._local() else "senha necessária"})
        try:
            if url.path == "/api/arquivo":
                alvo = q.get("caminho", "")
                if not caminho_permitido(alvo) or not os.path.isfile(alvo):
                    return self._enviar(404, {"erro": "arquivo não encontrado"})
                with open(alvo, "rb") as f:
                    dados = f.read()
                from urllib.parse import quote
                self.send_response(200)
                self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Length", str(len(dados)))
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(os.path.basename(alvo)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(dados)
                return
            if url.path == "/api/consolidado/zip":
                # todos os pacotes finais gerados do mês (ou só do cliente escolhido), num .zip
                import io
                import zipfile
                from urllib.parse import quote
                mes = _mes_valido(q.get("mes")) or "%04d-%02d" % _consolidado().mes_atual()
                cid = q.get("cliente") or ""
                est = _consolidado().ler_estado()
                buf = io.BytesIO()
                n = 0
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for c in comum.carregar_config().get("clientes", []):
                        if (cid and c["id"] != cid) or (not cid and not c.get("ativo", True)):
                            continue
                        r = (est.get(c["id"]) or {}).get(mes) or {}
                        if not r.get("ok") and r.get("ultimo_ok"):
                            r = r["ultimo_ok"]
                        arq = r.get("arquivo") if r.get("ok") else None
                        if arq and caminho_permitido(arq) and os.path.isfile(arq):
                            z.write(arq, os.path.basename(arq))
                            n += 1
                if not n:
                    return self._enviar(404, {"erro": "nenhum pacote final gerado para %s" % mes})
                dados = buf.getvalue()
                nome = "pacotes_finais_%s.zip" % mes
                self.send_response(200)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Length", str(len(dados)))
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(nome))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(dados)
                return
            if url.path in ("/", "/index.html"):
                with open(os.path.join(comum.PASTA_CODIGO, "painel.html"), "rb") as f:
                    return self._enviar(200, f.read(), "text/html; charset=utf-8")
            if url.path == "/api/estado":
                est = montar_estado(q.get("mes"))
                quem, usuario, pc = self._identidade()
                est["eu"] = {"quem": quem, "usuario": usuario, "pc": pc}
                u = getattr(self, "_usuario", None)
                if u:
                    est["eu"].update(login=u["login"], admin=u.get("perfil") == "admin", principal=bool(u.get("principal")))
                return self._enviar(200, est)
            if url.path == "/api/equipe":
                at = _atividade()
                quem, usuario, pc = self._identidade()
                eu = {"quem": quem, "usuario": usuario, "pc": pc}
                u = getattr(self, "_usuario", None)
                if u:
                    eu.update(login=u["login"], admin=u.get("perfil") == "admin", principal=bool(u.get("principal")))
                return self._enviar(200, {"quadro": at.quadro(comum.carregar_config()), "recentes": at.recentes(30), "eu": eu})
            if url.path == "/periodo":
                ini, fim = comum.periodo()
                if q.get("formato") == "json":
                    return self._enviar(200, {"inicio": comum.fmt(ini), "fim": comum.fmt(fim)})
                corpo = ("<h1>Período do relatório</h1><p>INICIO: <b id='inicio'>%s</b></p><p>FIM: <b id='fim'>%s</b></p>"
                         "<p>Hoje: %s. Regra: dia 1 = mês anterior inteiro; demais dias = dia 1 do mês até ontem.</p>"
                         % (comum.fmt(ini), comum.fmt(fim), dt.date.today().strftime("%d/%m/%Y")))
                return self._enviar(200, self._pagina("Período", corpo), "text/html; charset=utf-8")
            if url.path == "/registrar":
                return self._registrar(q)
            if url.path == "/saude":
                return self._enviar(200, {"ok": True})
            return self._enviar(404, {"erro": "não encontrado"})
        except Exception as e:
            traceback.print_exc()
            return self._enviar(500, {"erro": str(e)})

    def _registrar(self, q):
        cfg = comum.carregar_config()
        c = comum.conta(cfg, q.get("cliente"))
        mk = q.get("mk")
        status = q.get("status")
        if not c or mk not in cfg["marketplaces"] or status not in ("ok", "erro"):
            corpo = "<h1>Registro NÃO aceito</h1><p>Cliente, marketplace ou status inválido.</p>"
            return self._enviar(400, self._pagina("Erro", corpo), "text/html; charset=utf-8")
        ev = {"tipo": "registro", "cliente": c["id"], "mk": mk, "status": status}
        for k in ("inicio", "fim", "motivo", "detalhe"):
            if q.get(k):
                ev[k] = q[k][:300]
        comum.registrar_evento(ev)
        comum.log("[%s] extensão registrou %s em %s %s" % (c["nome"], status.upper(), cfg["marketplaces"][mk]["nome"],
                                                          ("— " + ev.get("motivo", "") + " " + ev.get("detalhe", "")) if status == "erro" else ""))
        corpo = "<h1>Registrado</h1><p>%s · %s · <b>%s</b></p><p>Pode fechar esta aba.</p>" % (
            html.escape(c["nome"]), html.escape(cfg["marketplaces"][mk]["nome"]), "OK" if status == "ok" else "ERRO")
        return self._enviar(200, self._pagina("Registrado", corpo), "text/html; charset=utf-8")

    def do_POST(self):
        caminho = urlparse(self.path).path
        if caminho == "/login" and self._remoto():
            tam = int(self.headers.get("Content-Length") or 0)
            form = parse_qs((self.rfile.read(tam) or b"").decode("utf-8"))
            senha = form.get("senha", [""])[0]
            r = _rede()
            import hmac
            if comum.NUVEM:
                import usuarios
                if _limitar_login(self.client_ip()):
                    return self._login("Muitas tentativas erradas. Espere 15 minutos.")
                u = usuarios.entrar(form.get("usuario", [""])[0], senha)
                _limitar_login(self.client_ip(), falhou=u is None)
                if not u:
                    return self._login("Usuário ou senha incorretos.")
                comum.log("Login no sistema online: %s" % u["login"])
                self.send_response(303)
                self.send_header("Set-Cookie", "rv_sessao=%s; Path=/; HttpOnly; SameSite=Lax; Max-Age=%d; Secure"
                                 % (usuarios.sessao_nova(u), usuarios.DURACAO))
                self.send_header("Location", "/")
                self.send_header("Content-Length", "0")
                self.end_headers()
                return
            usuario_ok = True
            certo = usuario_ok and r and r.papel() == "servidor" and r.token() and hmac.compare_digest(r.token(senha), r.token())
            if comum.NUVEM:
                _limitar_login(self.client_ip(), falhou=not certo)
            if certo:
                self.send_response(303)
                self.send_header("Set-Cookie", "rv_token=%s; Path=/; HttpOnly; SameSite=Strict; Max-Age=2592000%s"
                                 % (r.token(), "; Secure" if comum.NUVEM else ""))
                self.send_header("Location", "/")
                self.send_header("Content-Length", "0")
                self.end_headers()
                return
            return self._login("Usuário ou senha incorretos.")
        if not self._permitido() or self.headers.get("X-Painel") != "1":
            return self._enviar(403 if self._local() else 401, {"erro": "requisição não permitida"})
        remoto = self._remoto()
        if remoto and caminho in SO_LOCAL:
            return self._enviar(200, {"ok": False, "erro": "Essa ação só pode ser feita no próprio computador principal."})
        if remoto and caminho in NO_CHROME_DAQUI:
            return self._enviar(200, {"ok": False, "erro": "Essa conta baixa no computador principal: faça isso lá (o Chrome dela está naquele computador)."})
        quem, usuario, pc = self._identidade()
        if caminho.startswith("/api/usuarios/"):
            return self._enviar(200, self._usuarios(caminho))
        if caminho == "/api/agente/ping":
            acao = lambda d: agente_ping(d, self.client_address[0])
        elif caminho == "/api/presenca":
            acao = lambda d: _atividade().presenca(quem, usuario, pc, d)
        elif caminho == "/api/presenca/sair":
            acao = lambda d: _atividade().sair(quem)
        else:
            acao = ACOES.get(caminho)
        if not acao:
            return self._enviar(404, {"erro": "ação desconhecida"})
        try:
            tam = int(self.headers.get("Content-Length") or 0)
            dados = json.loads(self.rfile.read(tam) or b"{}") if tam else {}
            if remoto:
                dados["_remoto"] = True
                dados["estacao"] = pc
            at = _atividade()
            desc = at.descrever(caminho, dados, comum.carregar_config()) if at else None
            if desc:
                conf = at.conflito(desc, quem)
                if conf:
                    return self._enviar(200, {"ok": False, "erro": conf, "conflito": True})
                return self._enviar(200, at.executar(desc, quem, usuario, pc, lambda: acao(dados)))
            return self._enviar(200, acao(dados))
        except Exception as e:
            traceback.print_exc()
            return self._enviar(500, {"ok": False, "erro": "Erro interno: %s" % e})


ML_ANTIGO = ("No menu do vendedor, entre em Vendas. Use a opção de baixar o relatório de vendas em Excel (botão de download/exportar "
             "da lista de vendas). Escolha período personalizado de INICIO até FIM, gere e baixe o arquivo .xlsx.")
ML_NOVO = ("Na página de Vendas, PRIMEIRO filtre o período: abra o filtro de data, escolha período personalizado de INICIO até FIM e aplique o filtro. "
           "Confira que a lista mostra o período certo. Só então clique no botão de download (baixar/exportar) da lista de vendas. "
           "No canto inferior direito da tela vai aparecer um aviso com o relatório sendo gerado (carregando). Não saia da página: aguarde até o aviso "
           "indicar que está pronto (confira a cada 30 segundos, no máximo 10 minutos) e clique em Baixar nesse aviso para baixar o arquivo .xlsx.")


MG_ANTIGO = "Abra a lista de pedidos do portal do seller. Filtre pelo período de INICIO até FIM, use a opção de exportar o relatório de pedidos/vendas e baixe."
MG_NOVO = 'Abra a lista de pedidos do portal do seller. Filtre pelo período de INICIO até FIM e use a opção de exportar o relatório de pedidos/vendas. A Magalu gera o arquivo em segundo plano: depois de pedir a exportação, abra https://seller.magalu.com/central-arquivos (Central de Arquivos), aguarde o relatório que você acabou de pedir ficar pronto (atualize a página a cada 30 segundos, no máximo 15 minutos) e clique para baixar o .zip. Baixe só o arquivo gerado agora (o mais recente), não arquivos antigos. Não extraia nem abra o .zip.'


AMZ_ANTIGO = "Abra o menu Relatórios > Fulfillment (ou Pedidos > Relatórios de pedidos se a loja só tiver envio próprio)"
AMZ_NOVO = "Baixe 3 relatórios do período de INICIO até FIM, um de cada vez (aguarde cada um ficar pronto: confira a cada 30 segundos, no máximo 15 minutos):\n1) PEDIDOS: menu Relatórios > Fulfillment (ou Pedidos > Relatórios de pedidos se a loja só tiver envio próprio), relatório 'Todos os pedidos' (All Orders), 'Data exata' de INICIO até FIM, Solicitar download e baixe o arquivo .txt original (NÃO escolha Excel).\n2) RELATÓRIO DE TRANSAÇÕES UNIFICADO: menu Pagamentos > Biblioteca de relatórios > Relatórios por período (Transação / Custom Unified Transaction), período de INICIO até FIM, gere e baixe o .csv.\n3) TRANSAÇÕES A RECEBER: menu Pagamentos > Visualizar transações, filtre o período de INICIO até FIM e baixe o .csv."
AMZ_FONTE = 'Seller Central Brasil. São 3 relatórios: Todos os pedidos (.txt), relatório de transações unificado (.csv) e transações a receber (.csv) — o tratador precisa dos três. Caminhos dos menus a confirmar no 1º teste.'
MG_EXTRA = ' O .zip completo vem com dois arquivos dentro (pedidos e pacotes): é esse que deve ser baixado.'
PADRAO_AUX = {
    "amazon-unificado": {
        "nome": "Amazon — relatório unificado",
        "oculto": True,
        "pai": "amazon",
        "extensao": ".csv",
        "link": "",
        "link_status": "",
        "link_fonte": "",
        "padroes_nome": [
            "customunifiedtransaction",
            "unified.?transaction",
            "transacao.?unificad"
        ],
        "palavras_conteudo": [
            "data/hora",
            "id do pedido",
            "vendas do produto",
            "creditos de remessa",
            "tarifas de venda",
            "taxas fba"
        ],
        "instrucoes": ""
    },
    "amazon-receber": {
        "nome": "Amazon — transações a receber",
        "oculto": True,
        "pai": "amazon",
        "extensao": ".csv",
        "link": "",
        "link_status": "",
        "link_fonte": "",
        "padroes_nome": [
            "transacoes.?a.?receber",
            "a.?receber"
        ],
        "palavras_conteudo": [
            "tipo de transacao",
            "id do pedido",
            "detalhes do produto",
            "tarifas da amazon"
        ],
        "instrucoes": ""
    }
}


def _migrar(cfg):
    """Atualiza instruções padrão antigas (só se você ainda não tiver editado)."""
    with open(comum.ARQ_CONFIG, "r", encoding="utf-8") as f:
        mudou = any("contas" not in c for c in json.load(f).get("clientes", []))  # cadastro antigo -> grupos com contas
    # tabela de custos única -> tabela do mês em que foi enviada (vale para os meses seguintes até ter outra)
    for c in cfg.get("clientes", []):
        k = c.get("custos") or {}
        if k.get("arquivo") and not c.get("custos_meses"):
            try:
                d = dt.datetime.fromisoformat(k.get("atualizado", "")).date()
                mes = "%04d-%02d" % (d.year, d.month)
            except ValueError:
                mes = "%04d-%02d" % _tratador().mes_do_periodo()
            k = dict(k, mes=mes, versao=1)
            c["custos_meses"] = {mes: k}
            c["custos"] = k
            mudou = True
    for mk, antigo, novo in (("mercadolivre", ML_ANTIGO, ML_NOVO), ("magalu", MG_ANTIGO, MG_NOVO)):
        m = cfg.get("marketplaces", {}).get(mk)
        if m and m.get("instrucoes", "").startswith(antigo):
            m["instrucoes"] = novo
            mudou = True
    mks = cfg.setdefault("marketplaces", {})
    for k, v in comum.MARKETPLACES_NOVOS.items():   # Temu, Via Varejo, Madeira Madeira, Site
        if k not in mks:
            mks[k] = json.loads(json.dumps(v))
            mudou = True
    if (mks.get("viavarejo") or {}).pop("sem_tratador", None) is not None:   # Casas Bahia ganhou tratador (26/10/2026)
        mudou = True
    tm = mks.get("temu")
    if tm is not None and tm.pop("sem_tratador", None) is not None:   # Temu ganhou tratador (26/10/2026)
        novo = comum.MARKETPLACES_NOVOS["temu"]
        tm["padroes_nome"] = list(dict.fromkeys(tm.get("padroes_nome", []) + novo["padroes_nome"]))
        tm["palavras_conteudo"] = list(dict.fromkeys(tm.get("palavras_conteudo", []) + novo["palavras_conteudo"]))
        mudou = True
    mm = mks.get("madeiramadeira")
    if mm is not None and mm.pop("sem_tratador", None) is not None:   # Madeira Madeira ganhou tratador (26/10/2026)
        novo = comum.MARKETPLACES_NOVOS["madeiramadeira"]
        mm["padroes_nome"] = list(dict.fromkeys(mm.get("padroes_nome", []) + novo["padroes_nome"]))
        mm["palavras_conteudo"] = list(dict.fromkeys(mm.get("palavras_conteudo", []) + novo["palavras_conteudo"]))
        mudou = True
    # tratador de vendas: relatórios auxiliares da Amazon e instruções novas
    for k in ("amazon-unificado", "amazon-receber"):
        if k not in mks:
            mks[k] = PADRAO_AUX[k]
            mudou = True
    amz = mks.get("amazon")
    if amz and amz.get("instrucoes", "").startswith(AMZ_ANTIGO):
        amz["instrucoes"] = AMZ_NOVO
        amz["link_fonte"] = AMZ_FONTE
        mudou = True
    mg = mks.get("magalu")
    if mg and mg.get("instrucoes", "") == MG_NOVO:
        mg["instrucoes"] = MG_NOVO + MG_EXTRA
        mudou = True
    if "tratamento" not in cfg:
        cfg["tratamento"] = {"automatico": True, "imposto_padrao": 0.14}
        mudou = True
    ex = cfg.setdefault("execucao", {})
    if "modelo_reforco" not in ex:  # configuração anterior às otimizações
        if not ex.get("modelo"):
            ex["modelo"] = "sonnet"
        ex["modelo_reforco"] = ""
        mudou = True
    if mudou:
        comum.salvar_config(cfg)


_TENTATIVAS = {}


def _limitar_login(ip, falhou=None):
    """Online: 8 senhas erradas em 15 minutos bloqueiam aquele endereço por 15 minutos (e há um total geral de 60/15 min,
    para quem tenta de vários endereços). falhou=None só consulta; True/False registra o resultado. Devolve True se bloqueado."""
    import time as _t
    agora = _t.time()
    for k in list(_TENTATIVAS):
        _TENTATIVAS[k] = [x for x in _TENTATIVAS[k] if agora - x < 900]
    if falhou:
        _TENTATIVAS.setdefault(ip, []).append(agora)
        _TENTATIVAS.setdefault("*", []).append(agora)
    elif falhou is False:
        _TENTATIVAS.pop(ip, None)
    return len(_TENTATIVAS.get(ip, [])) >= 8 or len(_TENTATIVAS.get("*", [])) >= 60


# ---------------------------------------------------------------- online: importar os dados do computador
def admin_importar(dados):
    """Recebe em partes o .zip gerado no computador (exportar_nuvem.py) e, na última parte, coloca os dados no disco:
    sistema/... -> pasta do sistema (config, dados) e raiz/... -> pasta dos relatórios. Caminhos do Windows guardados
    nos .json passam a apontar para as pastas daqui; contas sem computador ficam com a estação indicada."""
    import base64
    import re as _re
    import shutil
    import zipfile
    if not comum.NUVEM:
        return {"ok": False, "erro": "Importação só existe no sistema online."}
    ident = _re.sub(r"[^\w-]", "", str(dados.get("id") or ""))[:40]
    parte, total = int(dados.get("parte") or 0), int(dados.get("total") or 0)
    if not ident or total < 1 or not (0 <= parte < total):
        return {"ok": False, "erro": "Parte inválida."}
    pasta = os.path.join(comum.PASTA_SISTEMA, "_importar")
    os.makedirs(pasta, exist_ok=True)
    arq = os.path.join(pasta, ident + ".zip")
    # cada parte num arquivo próprio: reenviar uma parte (depois de uma queda) não estraga o pacote
    p_arq = os.path.join(pasta, "%s.%05d.parte" % (ident, parte))
    with open(p_arq + ".tmp", "wb") as f:
        f.write(base64.b64decode(dados.get("conteudo") or ""))
    os.replace(p_arq + ".tmp", p_arq)
    if parte < total - 1:
        return {"ok": True, "parte": parte}
    faltam = [i for i in range(total) if not os.path.isfile(os.path.join(pasta, "%s.%05d.parte" % (ident, i)))]
    if faltam:
        return {"ok": False, "erro": "Faltam partes: %s. Rode de novo." % faltam[:10]}
    with open(arq, "wb") as d:
        for i in range(total):
            pi = os.path.join(pasta, "%s.%05d.parte" % (ident, i))
            with open(pi, "rb") as o:
                shutil.copyfileobj(o, d)
    for i in range(total):
        os.remove(os.path.join(pasta, "%s.%05d.parte" % (ident, i)))
    with zipfile.ZipFile(arq) as z:
        man = json.loads(z.read("manifesto.json").decode("utf-8"))
        cfg = comum.carregar_config()
        raiz = comum.raiz(cfg)
        n = 0
        for info in z.infolist():
            nome = info.filename
            if info.is_dir() or nome == "manifesto.json" or ".." in nome.split("/"):
                continue
            if nome.startswith("sistema/"):
                rel, base = nome[len("sistema/"):], comum.PASTA_SISTEMA
                if os.path.basename(rel) in ("rede.json", "estacoes.json"):
                    continue
            elif nome.startswith("raiz/"):
                rel, base = nome[len("raiz/"):], raiz
            else:
                continue
            alvo = os.path.join(base, *rel.split("/"))
            os.makedirs(os.path.dirname(alvo), exist_ok=True)
            with z.open(info) as o, open(alvo + ".importando", "wb") as d:
                shutil.copyfileobj(o, d)
            os.replace(alvo + ".importando", alvo)        # nunca deixa um arquivo pela metade
            n += 1
    # caminhos do Windows -> pastas daqui
    velhos = [(str(man.get("sistema") or ""), comum.PASTA_SISTEMA), (str(man.get("raiz") or ""), raiz)]
    velhos = [(v.rstrip("\\/"), novo) for v, novo in velhos if v]
    velhos.sort(key=lambda x: -len(x[0]))

    def trocar(v):
        if isinstance(v, str):
            for velho, novo in velhos:
                if v.lower().startswith(velho.lower()) and (len(v) == len(velho) or v[len(velho)] in "\\/"):
                    resto = v[len(velho):].replace("\\", "/").lstrip("/")
                    return os.path.join(novo, resto) if resto else novo
            return v
        if isinstance(v, list):
            return [trocar(x) for x in v]
        if isinstance(v, dict):
            return {k: trocar(x) for k, x in v.items()}
        return v
    trocados = 0
    for pasta_j, _subs, arqs in os.walk(os.path.join(comum.PASTA_SISTEMA, "dados")):
        for a in arqs:
            p = os.path.join(pasta_j, a)
            try:
                if a.endswith(".json"):
                    with open(p, "r", encoding="utf-8") as f:
                        d = json.load(f)
                    with open(p, "w", encoding="utf-8") as f:
                        json.dump(trocar(d), f, ensure_ascii=False)
                    trocados += 1
                elif a.endswith(".jsonl"):
                    with open(p, "r", encoding="utf-8") as f:
                        linhas = [json.dumps(trocar(json.loads(x)), ensure_ascii=False) for x in f if x.strip()]
                    with open(p, "w", encoding="utf-8") as f:
                        f.write("\n".join(linhas) + ("\n" if linhas else ""))
                    trocados += 1
            except (OSError, ValueError):
                pass
    estacao = str(dados.get("estacao_padrao") or man.get("estacao") or "").strip()
    with TRAVA_CONFIG:
        cfg = comum.carregar_config()
        cfg = trocar(cfg)
        sem = 0
        for c in cfg.get("clientes", []):
            for u in c.get("contas", []):
                if estacao and not u.get("computador"):
                    u["computador"] = estacao
                    sem += 1
        comum.salvar_config(cfg)
    os.remove(arq)
    comum.log("Importação do computador: %d arquivo(s); %d .json ajustados; %d conta(s) passaram para '%s'."
              % (n, trocados, sem, estacao))
    threading.Thread(target=_aquecer_painel, args=(5,), daemon=True).start()
    return {"ok": True, "fim": True, "arquivos": n, "json_ajustados": trocados, "contas_para_estacao": sem, "estacao": estacao}


ACOES["/api/admin/importar"] = admin_importar


def _aquecer_painel(espera=0):
    """Prepara em segundo plano os números do painel (lê cada consolidado uma vez e guarda o resumo em disco)."""
    time.sleep(espera)
    import painel_dados
    painel_dados.aquecer()


def criar_servidor(cfg=None):
    cfg = cfg or comum.carregar_config()
    with TRAVA_CONFIG:
        _migrar(cfg)
    comum.criar_pastas_base(cfg)
    threading.Thread(target=_reprocessar_antigos, daemon=True).start()
    r = _rede()
    papel = r.papel() if r else "sozinho"
    if papel == "estacao":
        import estacao
        return estacao.criar_servidor(cfg)
    threading.Thread(target=_retratar_regra_nao_vendas, daemon=True).start()
    if comum.NUVEM:
        threading.Thread(target=_aquecer_painel, args=(30,), daemon=True).start()
    host = "0.0.0.0" if papel == "servidor" else "127.0.0.1"   # servidor: aceita a rede (com senha)
    porta = int(os.environ.get("PORT") or 0) if comum.NUVEM else 0
    return ThreadingHTTPServer((host, porta or int(cfg.get("porta_painel", 8765))), Handler)


def main():
    if sys.stdout is None or sys.stderr is None:  # pythonw: sem console
        pasta = os.path.join(comum.PASTA_SISTEMA, "logs")
        os.makedirs(pasta, exist_ok=True)
        saida = open(os.path.join(pasta, "painel.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or saida
        sys.stderr = sys.stderr or saida
    srv = criar_servidor()
    print("Painel em http://localhost:%d" % srv.server_address[1])
    srv.serve_forever()


if __name__ == "__main__":
    main()
