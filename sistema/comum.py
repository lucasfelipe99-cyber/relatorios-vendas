# -*- coding: utf-8 -*-
"""Funções compartilhadas: configuração, pastas, período, perfis do Chrome, prompts e status."""
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata

def _achar_pasta_sistema():
    """Pasta com config.json. No .exe (PyInstaller) sobe a partir do executável até achar."""
    if os.environ.get("RELATORIOS_SISTEMA"):
        return os.environ["RELATORIOS_SISTEMA"]
    if getattr(sys, "frozen", False):
        p = os.path.dirname(os.path.abspath(sys.executable))
        for _ in range(5):
            if os.path.exists(os.path.join(p, "config.json")):
                return p
            p = os.path.dirname(p)
    return os.path.dirname(os.path.abspath(__file__))


PASTA_SISTEMA = _achar_pasta_sistema()
# Online (Render): RV_NUVEM=1. Os dados ficam no disco (RELATORIOS_SISTEMA / RELATORIOS_RAIZ) e o código na pasta do app.
NUVEM = os.environ.get("RV_NUVEM") == "1"
PASTA_CODIGO = PASTA_SISTEMA if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
RAIZ_NUVEM = os.environ.get("RELATORIOS_RAIZ") or os.path.join(os.path.dirname(PASTA_SISTEMA), "Relatorios")
TRAVA_CFG = threading.RLock()   # quem lê-altera-grava o config.json (app e tratamentos)
ARQ_CONFIG = os.environ.get("RELATORIOS_CONFIG") or os.path.join(PASTA_SISTEMA, "config.json")
EXTENSAO_ID = "fcoeoabgfenejglbffodgkkbkcdhcgfn"
LINK_EXTENSAO = "https://chromewebstore.google.com/detail/claude/" + EXTENSAO_ID
CARACTERES_INVALIDOS = '\\/:*?"<>|'


# ---------------------------------------------------------------- config
def _config_inicial():
    """Online, na primeira vez (disco vazio): começa do modelo que vem com o código."""
    if NUVEM and not os.path.exists(ARQ_CONFIG):
        os.makedirs(os.path.dirname(ARQ_CONFIG), exist_ok=True)
        modelo = os.path.join(PASTA_CODIGO, "config.json")
        shutil.copy2(modelo, ARQ_CONFIG)


def carregar_config():
    _config_inicial()
    with open(ARQ_CONFIG, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    migrar_contas(cfg)  # em memória; o servidor grava a migração ao iniciar
    if NUVEM:
        cfg["pasta_raiz"] = RAIZ_NUVEM          # o config pode ter vindo do Windows (C:\\Relatórios)
        cfg["pasta_entrada"] = "_entrada"
        cfg["nuvem"] = True
    return cfg


def salvar_config(cfg):
    pasta = os.path.dirname(ARQ_CONFIG)
    fd, tmp = tempfile.mkstemp(prefix="config_", suffix=".json", dir=pasta)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    if os.path.exists(ARQ_CONFIG):
        os.makedirs(pasta_dados(cfg, "backups"), exist_ok=True)
        shutil.copy2(ARQ_CONFIG, os.path.join(pasta_dados(cfg, "backups"), "config_ultimo.json"))
    os.replace(tmp, ARQ_CONFIG)


def pasta_dados(cfg, nome):
    """Subpastas do sistema: logs, dados, backups."""
    p = os.path.join(PASTA_SISTEMA, nome)
    os.makedirs(p, exist_ok=True)
    return p


# ---------------------------------------------------------------- pastas
def raiz(cfg):
    return cfg["pasta_raiz"]


def pasta_entrada(cfg, cliente_nome=None):
    p = os.path.join(raiz(cfg), cfg.get("pasta_entrada", "_entrada"))
    return os.path.join(p, cliente_nome) if cliente_nome else p


def rotulo_dia(data=None):
    data = data or dt.date.today()
    return data.strftime("%d-%m")


def pasta_dia(cfg, cliente_nome, data=None):
    return os.path.join(raiz(cfg), rotulo_dia(data), cliente_nome)


def nome_arquivo(mk_id, cfg, data=None):
    return "vendas_%s_%s%s" % (mk_id, rotulo_dia(data), cfg["marketplaces"][mk_id]["extensao"])


def criar_pastas_base(cfg):
    os.makedirs(pasta_entrada(cfg), exist_ok=True)


def pasta_de(unidade):
    """Subcaminho da conta: Cliente\\Conta (cai para o nome se for um cadastro antigo)."""
    return unidade.get("pasta") or unidade["nome"]


def criar_pastas_cliente(cfg, cliente):
    """Aceita um cliente (grupo) ou uma conta já 'achatada'."""
    if "contas" in cliente:
        for u in contas_do_cliente(cfg, cliente):
            os.makedirs(pasta_entrada(cfg, pasta_de(u)), exist_ok=True)
    else:
        os.makedirs(pasta_entrada(cfg, pasta_de(cliente)), exist_ok=True)


# ---------------------------------------------------------------- grupos e contas
SEP_CONTA = "--"  # gerar_id nunca produz "--", então o id da conta é único e reversível


SHOPEE_ANTECIPA_PADRAO = 0.025


def _unidade(cliente, conta, total_contas):
    return {
        "id": cliente["id"] + SEP_CONTA + conta["id"],
        "cliente_id": cliente["id"], "cliente_nome": cliente["nome"],
        "conta_id": conta["id"], "conta_nome": conta["nome"],
        "nome": cliente["nome"] if conta["nome"] == cliente["nome"] else "%s / %s" % (cliente["nome"], conta["nome"]),
        "loja": conta["nome"],
        "pasta": os.path.join(cliente["nome"], conta["nome"]),
        "unica_conta": total_contas == 1,
        "modo": "manual" if conta.get("modo") == "manual" else "automatico",
        "perfil": "" if conta.get("modo") == "manual" else conta.get("perfil", ""),
        "marketplaces": conta.get("marketplaces", []), "links": conta.get("links", {}),
        "computador": "" if conta.get("modo") == "manual" else (conta.get("computador") or ""),
        "ativo": bool(cliente.get("ativo", True)) and bool(conta.get("ativo", True)),
        "imposto": cliente.get("imposto"), "custos": cliente.get("custos"), "custos_meses": cliente.get("custos_meses") or {},
        "ids_shopee": cliente.get("ids_shopee"),
        # antecipação Shopee da conta, em fração (padrão 2,5%; 0 = a loja não usa antecipação)
        "shopee_antecipa": (cliente.get("shopee_antecipa") or {}).get(conta["id"], SHOPEE_ANTECIPA_PADRAO),
    }


# Marketplaces só cadastrados (sem download automático e sem tratador por enquanto): o relatório é subido pelo app
# e fica guardado na pasta do dia, sem tratamento. Entram nas configurações antigas pelo _migrar do servidor.
MARKETPLACES_NOVOS = {
    "temu": {
        "nome": "Temu",
        "padroes_nome": [
            "temu",
            "order_export"
        ],
        "palavras_conteudo": [
            "temu",
            "sku de contribuicao",
            "preco-base dos produtos",
            "desconto da temu",
            "approved courier list"
        ],
        "extensao": ".xlsx",
        "link": "",
        "link_status": "a definir",
        "link_fonte": "Ainda não definido: por enquanto o relatório é subido pelo app (Vendas > Subir relatório).",
        "instrucoes": "",
        "sem_download": True
    },
    "viavarejo": {
        "nome": "Via Varejo",
        "padroes_nome": [
            "via.?varejo",
            "casas.?bahia",
            "grupo.?casas"
        ],
        "palavras_conteudo": [
            "via varejo",
            "casas bahia"
        ],
        "extensao": ".xlsx",
        "link": "",
        "link_status": "a definir",
        "link_fonte": "Ainda não definido: por enquanto o relatório é subido pelo app (Vendas > Subir relatório).",
        "instrucoes": "",
        "sem_download": True
    },
    "madeiramadeira": {
        "nome": "Madeira Madeira",
        "padroes_nome": [
            "madeira.?madeira",
            "report_pedido"
        ],
        "palavras_conteudo": [
            "madeiramadeira",
            "madeira madeira",
            "pedido site mm",
            "nome do seller",
            "status cliente entrega"
        ],
        "extensao": ".xlsx",
        "link": "",
        "link_status": "a definir",
        "link_fonte": "Ainda não definido: por enquanto o relatório é subido pelo app (Vendas > Subir relatório).",
        "instrucoes": "",
        "sem_download": True
    },
    "site": {
        "nome": "Site",
        "padroes_nome": [
            "^site[ _.-]",
            "loja.?propria",
            "vendas.?site"
        ],
        "palavras_conteudo": [],
        "extensao": ".xlsx",
        "link": "",
        "link_status": "a definir",
        "link_fonte": "Ainda não definido: por enquanto o relatório é subido pelo app (Vendas > Subir relatório).",
        "instrucoes": "",
        "sem_download": True,
        "sem_tratador": True
    }
}


def _arq_ok_vendas():
    return os.path.join(PASTA_SISTEMA, "dados", "ok_vendas.json")


def oks_vendas():
    """{conta_id: {mk: {"AAAA-MM-DD": {"motivo", "quando"}}}} — dia sem relatório de vendas aceito à mão."""
    try:
        with open(_arq_ok_vendas(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def ok_vendas(conta_id, mk, dia):
    return (oks_vendas().get(conta_id, {}).get(mk, {}) or {}).get(dia.isoformat() if hasattr(dia, "isoformat") else str(dia))


def marcar_ok_vendas(conta_id, mk, dia, motivo="", remover=False):
    todos = oks_vendas()
    d = (dia.isoformat() if hasattr(dia, "isoformat") else str(dia))[:10]
    m = todos.setdefault(conta_id, {}).setdefault(mk, {})
    if remover:
        m.pop(d, None)
    else:
        m[d] = {"motivo": (motivo or "OK manual").strip()[:200], "quando": dt.datetime.now().isoformat(timespec="seconds")}
    os.makedirs(os.path.dirname(_arq_ok_vendas()), exist_ok=True)
    with open(_arq_ok_vendas() + ".novo", "w", encoding="utf-8") as f:
        json.dump(todos, f, ensure_ascii=False, indent=1)
    os.replace(_arq_ok_vendas() + ".novo", _arq_ok_vendas())


def baixa_sozinho(cfg, mk):
    return mk in cfg.get("marketplaces", {}) and not cfg["marketplaces"][mk].get("sem_download")


def contas_do_cliente(cfg, cliente):
    lista = cliente.get("contas") or []
    return [_unidade(cliente, ct, len(lista)) for ct in lista]


def contas(cfg, so_ativas=False):
    """Todas as contas (uma por navegador) de todos os clientes, no formato usado pela rotina."""
    migrar_contas(cfg)  # aceita cadastro antigo mesmo se vier de fora do carregar_config
    saida = []
    for c in cfg.get("clientes", []):
        for u in contas_do_cliente(cfg, c):
            if not so_ativas or u["ativo"]:
                saida.append(u)
    return saida


def conta(cfg, uid):
    return next((u for u in contas(cfg) if u["id"] == uid), None)


def filtrar_contas(cfg, ids, so_ativas=True):
    """ids pode ter ids de conta ou de cliente (grupo inteiro)."""
    return [u for u in contas(cfg, so_ativas) if not ids or u["id"] in ids or u["cliente_id"] in ids]


def migrar_contas(cfg):
    """Cadastro antigo (perfil/marketplaces no cliente) -> cliente com uma conta de mesmo nome."""
    mudou = False
    for c in cfg.get("clientes", []):
        if "contas" not in c:
            c["contas"] = [{"id": gerar_id(c["nome"], set()), "nome": c["nome"], "perfil": c.pop("perfil", ""),
                            "marketplaces": c.pop("marketplaces", []), "links": c.pop("links", {}), "ativo": True}]
            mudou = True
    return mudou


def validar_nome_cliente(nome):
    nome = (nome or "").strip()
    if not nome:
        return "Informe o nome do cliente."
    if any(c in CARACTERES_INVALIDOS for c in nome):
        return "O nome não pode conter: " + " ".join(CARACTERES_INVALIDOS)
    if nome.startswith("_") or nome.endswith(".") or nome.endswith(" "):
        return "O nome não pode começar com _ nem terminar com ponto ou espaço."
    return None


def sem_acento(txt):
    return "".join(c for c in unicodedata.normalize("NFKD", txt or "") if not unicodedata.combining(c)).lower()


def gerar_id(nome, existentes):
    base = re.sub(r"[^a-z0-9]+", "-", sem_acento(nome)).strip("-") or "cliente"
    base = base[:30]
    novo, n = base, 2
    while novo in existentes:
        novo = "%s-%d" % (base, n)
        n += 1
    return novo


# ---------------------------------------------------------------- período
def periodo(hoje=None):
    """(inicio, fim) do relatório. Dia 1 -> mês anterior inteiro; senão dia 1 do mês até ontem."""
    hoje = hoje or dt.date.today()
    if hoje.day == 1:
        fim = hoje - dt.timedelta(days=1)
        inicio = fim.replace(day=1)
    else:
        inicio = hoje.replace(day=1)
        fim = hoje - dt.timedelta(days=1)
    return inicio, fim


def fmt(d):
    return d.strftime("%d/%m/%Y")


# ---------------------------------------------------------------- Chrome
def chrome_user_data(cfg):
    if cfg.get("chrome_user_data"):
        return cfg["chrome_user_data"]
    local = os.environ.get("LOCALAPPDATA", "")
    return os.path.join(local, "Google", "Chrome", "User Data")


def chrome_exe(cfg):
    if cfg.get("chrome_exe") and os.path.exists(cfg["chrome_exe"]):
        return cfg["chrome_exe"]
    candidatos = [
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return ""


def listar_perfis(cfg):
    """Perfis do Chrome deste computador: [{dir, nome, email}]."""
    ud = chrome_user_data(cfg)
    perfis = []
    try:
        with open(os.path.join(ud, "Local State"), "r", encoding="utf-8") as f:
            info = json.load(f).get("profile", {}).get("info_cache", {})
        for d, v in info.items():
            perfis.append({"dir": d, "nome": v.get("name") or d, "email": v.get("user_name", "")})
    except (OSError, ValueError):
        pass
    if not perfis and os.path.isdir(ud):  # sem Local State: varre as pastas
        for d in sorted(os.listdir(ud)):
            if (d == "Default" or d.startswith("Profile ")) and os.path.exists(os.path.join(ud, d, "Preferences")):
                perfis.append({"dir": d, "nome": d, "email": ""})
    perfis.sort(key=lambda p: (p["dir"] != "Default", sem_acento(p["nome"])))
    return perfis


def arquivo_preferences(cfg, perfil_dir):
    return os.path.join(chrome_user_data(cfg), perfil_dir, "Preferences")


def extensao_instalada(cfg, perfil_dir):
    pasta = os.path.join(chrome_user_data(cfg), perfil_dir, "Extensions", EXTENSAO_ID)
    return os.path.isdir(pasta) and bool(os.listdir(pasta))


def pasta_downloads_perfil(cfg, perfil_dir):
    """Pasta de downloads configurada no perfil (None = padrão do Windows)."""
    try:
        with open(arquivo_preferences(cfg, perfil_dir), "r", encoding="utf-8") as f:
            prefs = json.load(f)
    except (OSError, ValueError):
        return {"pasta": None, "pergunta": None, "erro": "Não foi possível ler as preferências do perfil."}
    dl = prefs.get("download", {})
    pasta = dl.get("default_directory") or prefs.get("savefile", {}).get("default_directory")
    return {"pasta": pasta, "pergunta": bool(dl.get("prompt_for_download", False)), "erro": None}


def chrome_aberto():
    if os.name != "nt":
        return False
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
                             capture_output=True, text=True, timeout=15,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
        return "chrome.exe" in out.lower()
    except Exception:
        return False


def abrir_no_perfil(cfg, perfil_dir, url):
    """Abre uma URL no Chrome, no perfil indicado."""
    exe = chrome_exe(cfg)
    if not exe:
        return {"ok": False, "erro": "chrome.exe não encontrado. Informe o caminho em Horários e ajustes."}
    if not url.startswith("https://"):
        return {"ok": False, "erro": "Só abro links https://"}
    if perfil_dir not in [p["dir"] for p in listar_perfis(cfg)]:
        return {"ok": False, "erro": "Perfil do Chrome não encontrado: " + perfil_dir}
    subprocess.Popen([exe, "--profile-directory=%s" % perfil_dir, url])
    return {"ok": True}


def mesma_pasta(a, b):
    if not a or not b:
        return False
    return os.path.normcase(os.path.normpath(a)) == os.path.normcase(os.path.normpath(b))


def alterar_pasta_downloads(cfg, perfil_dir, nova_pasta):
    """Backup + edição do Preferences + verificação. Exige Chrome fechado."""
    if chrome_aberto():
        return {"ok": False, "erro": "O Chrome está aberto. Feche TODAS as janelas do Chrome (inclusive na bandeja do relógio) e tente de novo."}
    arq = arquivo_preferences(cfg, perfil_dir)
    if not os.path.exists(arq):
        return {"ok": False, "erro": "Arquivo de preferências do perfil não encontrado: " + arq}
    os.makedirs(nova_pasta, exist_ok=True)
    try:
        with open(arq, "r", encoding="utf-8") as f:
            prefs = json.load(f)
    except ValueError:
        return {"ok": False, "erro": "Arquivo de preferências ilegível; nada foi alterado."}
    pasta_bkp = pasta_dados(cfg, "backups")
    carimbo = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    seguro = re.sub(r"[^A-Za-z0-9]+", "_", perfil_dir)
    backup = os.path.join(pasta_bkp, "Preferences_%s_%s.json" % (seguro, carimbo))
    shutil.copy2(arq, backup)
    dl = prefs.setdefault("download", {})
    dl["default_directory"] = nova_pasta
    dl["prompt_for_download"] = False
    dl["directory_upgrade"] = True
    prefs.setdefault("savefile", {})["default_directory"] = nova_pasta
    tmp = arq + ".novo"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(prefs, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, arq)
    conf = pasta_downloads_perfil(cfg, perfil_dir)
    ok = mesma_pasta(conf["pasta"], nova_pasta) and conf["pergunta"] is False
    return {"ok": ok, "backup": backup, "pasta_lida": conf["pasta"],
            "erro": None if ok else "A alteração não foi confirmada na releitura. O backup está em " + backup}


# ---------------------------------------------------------------- prompts
BLOCO_PERIODO = (
    "PERÍODO: abra http://localhost:{porta}/periodo e use as datas INICIO e FIM mostradas. "
    "Se a página não abrir, calcule: se hoje é dia 1, INICIO = dia 1 do mês anterior e FIM = último dia do mês anterior; "
    "nos outros dias, INICIO = dia 1 do mês atual e FIM = ontem. Use o formato DD/MM/AAAA."
)

BLOCO_ECONOMIA = (
    "ECONOMIA: para entender a página, prefira ler o texto e os elementos (read_page, find, get_page_text). "
    "Use captura de tela só quando precisar ver algo visual, como um aviso de download ou um botão que não aparece no texto. "
    "Não repita ações que já deram certo e não navegue por páginas fora do roteiro."
)

BLOCO_SEGURANCA = (
    "REGRAS (obrigatórias):\n"
    "- Se aparecer tela de login, pedido de senha, código de verificação (SMS, e-mail ou app), captcha ou qualquer checagem de segurança: PARE na hora. "
    "Não digite nada, não tente entrar e não tente contornar. Registre como erro (veja abaixo) e siga para o próximo item.\n"
    "- Não altere configurações da conta, não edite nem cancele pedidos, não aceite termos e não clique em nada além do necessário para exportar.\n"
    "- Deixe o download ir para a pasta padrão do Chrome; não renomeie, não mova e não abra o arquivo baixado.\n"
    "- Se a exportação falhar por outro motivo (botão não encontrado, erro na página, relatório não ficou pronto no tempo indicado), pare esse item e registre como erro."
)


def link_cliente(cfg, cliente, mk_id):
    return (cliente.get("links") or {}).get(mk_id) or cfg["marketplaces"][mk_id]["link"]


def roteiro(cfg, cliente, mk_id):
    """Roteiro aprendido para este marketplace + link (o mesmo painel vale para todos os clientes)."""
    return ((cfg["marketplaces"][mk_id].get("roteiros") or {}).get(link_cliente(cfg, cliente, mk_id)) or {}).get("texto", "")


def _bloco_mk(cfg, cliente, mk_id, numero=None, aprender=True):
    mk = cfg["marketplaces"][mk_id]
    porta = cfg.get("porta_painel", 8765)
    titulo = ("%d) " % numero if numero else "") + mk["nome"]
    ext = ".zip (não extraia)" if mk["extensao"] == ".zip" else "Excel (.xlsx)"
    base = "http://localhost:%d/registrar?cliente=%s&mk=%s" % (porta, cliente["id"], mk_id)
    rot = roteiro(cfg, cliente, mk_id)
    partes = ["### " + titulo, "Acesse: " + link_cliente(cfg, cliente, mk_id), mk["instrucoes"]]
    if rot:
        partes.append("ROTEIRO QUE FUNCIONOU DA ÚLTIMA VEZ (siga direto, sem explorar; se a tela estiver diferente, adapte):\n" + rot)
    partes += ["Formato esperado: %s." % ext,
               "Ao terminar, registre o resultado abrindo UMA destas URLs (troque o que está entre < >):",
               "- deu certo: %s&status=ok&inicio=<INICIO>&fim=<FIM>" % base,
               "- deu problema: %s&status=erro&motivo=<login|codigo|captcha|outro>&detalhe=<texto curto sem acentos>" % base]
    if aprender:
        partes.append('Se o download deu certo, inclua na sua resposta final um bloco <roteiro mk="%s"> ... </roteiro> com os passos '
                      "exatos que funcionaram, em no máximo 10 linhas curtas (nomes de botões, campos, menus e URLs que usou). "
                      "Não coloque dados do cliente (nomes, e-mails, valores, números de pedido)." % mk_id)
    return "\n".join(partes)


def prompt_marketplace(cfg, cliente, mk_id):
    mk = cfg["marketplaces"][mk_id]
    porta = cfg.get("porta_painel", 8765)
    return "\n\n".join([
        "Tarefa: baixar o relatório de vendas do %s do cliente %s. Este perfil do Chrome já está logado nas contas do cliente."
        % (mk["nome"], cliente["nome"]),
        BLOCO_PERIODO.format(porta=porta),
        BLOCO_SEGURANCA,
        BLOCO_ECONOMIA,
        _bloco_mk(cfg, cliente, mk_id),
    ])


def prompt_combinado(cfg, cliente):
    porta = cfg.get("porta_painel", 8765)
    mks = [m for m in cliente.get("marketplaces", []) if baixa_sozinho(cfg, m)]
    nomes = ", ".join(cfg["marketplaces"][m]["nome"] for m in mks)
    paralelo = cfg.get("execucao", {}).get("modo_paralelo", True) and len(mks) > 1
    if paralelo:
        abertura = (
            "Tarefa: baixar os relatórios de vendas do cliente %s: %s. Este perfil do Chrome já está logado nas contas do cliente. "
            "Um problema em um marketplace NÃO impede os outros.\n\n"
            "COMO TRABALHAR (em paralelo, para não ficar parado esperando):\n"
            "FASE 1 — PEDIR: para cada marketplace, em uma ABA PRÓPRIA (deixe-a aberta), faça os passos até PEDIR o relatório "
            "(filtrar o período e clicar em exportar/download). Assim que o marketplace começar a gerar o arquivo (aparecer carregando, "
            "processando ou o pedido entrar no histórico/central de arquivos), NÃO espere: passe para o próximo marketplace.\n"
            "FASE 2 — BUSCAR: depois de pedir todos, percorra as abas em rodízio. Em cada aba, confira se o relatório já está pronto "
            "(leia o texto da página; use captura de tela só se precisar). Se estiver pronto, baixe, registre o resultado e não volte mais a ela. "
            "Se não estiver, vá para a próxima aba. Quando tiver passado por todas sem nada novo, espere cerca de 60 segundos antes da próxima volta "
            "(não fique checando sem parar). Limite total de espera: 15 minutos; o que não ficar pronto até lá, registre como erro com "
            "detalhe=nao-ficou-pronto.\n"
            "No Mercado Livre o aviso de geração fica no canto inferior direito da própria página de Vendas: mantenha essa aba aberta e volte a ela. "
            "Na Magalu, acompanhe na Central de Arquivos." % (cliente["nome"], nomes))
    else:
        abertura = ("Tarefa: baixar os relatórios de vendas do cliente %s, nesta ordem: %s. Este perfil do Chrome já está logado nas contas do cliente. "
                    "Faça um marketplace por vez, cada um em uma aba nova; um problema em um deles NÃO impede os seguintes." % (cliente["nome"], nomes))
    partes = [abertura, BLOCO_PERIODO.format(porta=porta), BLOCO_SEGURANCA, BLOCO_ECONOMIA]
    if paralelo:
        partes.append("Abaixo estão os passos de cada marketplace. Na FASE 1 faça só até pedir o relatório; na FASE 2, a parte de esperar/baixar.")
    for i, m in enumerate(mks, 1):
        partes.append(_bloco_mk(cfg, cliente, m, i))
    partes.append("No final, responda com uma linha por marketplace: nome, ok ou erro, e o motivo do erro se houver.")
    return "\n\n".join(partes)


def nome_atalho(cliente, mk_id=None):
    return "vendas-" + cliente["id"] + ("-" + mk_id if mk_id else "")


def horarios_escalonados(cfg):
    h = cfg["horarios"]
    hh, mm = [int(x) for x in h["inicio_atalhos"].split(":")]
    base = dt.datetime(2000, 1, 1, hh, mm)
    passo = int(h.get("intervalo_entre_perfis_min", 5))
    saida, i = {}, 0
    for c in contas(cfg, so_ativas=True):
        if c.get("marketplaces") and c.get("modo") != "manual":
            saida[c["id"]] = (base + dt.timedelta(minutes=passo * i)).strftime("%H:%M")
            i += 1
    return saida


# ---------------------------------------------------------------- histórico
def arq_historico():
    return os.path.join(PASTA_SISTEMA, "dados", "historico.jsonl")


MUDOU = [0]   # conta alterações (log/eventos) para as telas se atualizarem sozinhas


def registrar_evento(ev):
    MUDOU[0] += 1
    os.makedirs(os.path.dirname(arq_historico()), exist_ok=True)
    ev = dict(ev)
    ev.setdefault("ts", dt.datetime.now().isoformat(timespec="seconds"))
    with open(arq_historico(), "a", encoding="utf-8") as f:
        f.write(json.dumps(ev, ensure_ascii=False) + "\n")


def ler_historico(limite=20000):
    try:
        with open(arq_historico(), "r", encoding="utf-8") as f:
            linhas = f.readlines()[-limite:]
    except OSError:
        return []
    saida = []
    for l in linhas:
        try:
            saida.append(json.loads(l))
        except ValueError:
            pass
    return saida


def log(msg, cfg=None):
    MUDOU[0] += 1
    pasta = os.path.join(PASTA_SISTEMA, "logs")
    os.makedirs(pasta, exist_ok=True)
    agora = dt.datetime.now()
    with open(os.path.join(pasta, "organizador_%s.log" % agora.strftime("%Y-%m")), "a", encoding="utf-8") as f:
        f.write("%s  %s\n" % (agora.strftime("%Y-%m-%d %H:%M:%S"), msg))


def deste_computador(conta):
    """Uso em rede: a conta baixa neste computador? (sozinho: sempre)."""
    try:
        import rede
        d = rede.ler()
    except Exception:
        return True
    comp = conta.get("computador") or ""
    if d["papel"] == "estacao":
        # a estação "principal" (ligada ao sistema online) também baixa as contas sem computador definido
        return comp == d["nome"] or (comp == "" and d.get("principal") == "1")
    if d["papel"] == "servidor":
        if NUVEM:
            return False          # online não tem Chrome: quem baixa são os computadores (estações)
        return comp in ("", d["nome"])
    return True
