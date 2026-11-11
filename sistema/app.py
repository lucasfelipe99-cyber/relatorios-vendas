# -*- coding: utf-8 -*-
"""App desktop "Relatórios de Vendas": janela própria + ícone na bandeja.

Uso:
  PainelRelatorios.exe              abre a janela (ou traz a já aberta para frente)
  PainelRelatorios.exe --bandeja    inicia só no ícone da bandeja (usado ao entrar no Windows)
  PainelRelatorios.exe --organizar  roda o organizador e sai (tarefa agendada)
  PainelRelatorios.exe --abrir-chrome  abre o Chrome com os perfis e sai
  PainelRelatorios.exe --rodar-todos   baixa os relatórios de todos os clientes (tarefa das 5h)
O servidor local continua rodando enquanto o app estiver na bandeja: a extensão
usa /periodo e /registrar. Fechar a janela só esconde; para encerrar use "Sair".
"""
import datetime as dt
import json
import os
import sys
import threading
import time
import urllib.request
import webbrowser

if getattr(sys, "frozen", False) is False:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import comum  # noqa: E402

TITULO = "Relatórios de Vendas"


def _log_sem_console():
    if sys.stdout is None or sys.stderr is None:
        pasta = os.path.join(comum.PASTA_SISTEMA, "logs")
        os.makedirs(pasta, exist_ok=True)
        f = open(os.path.join(pasta, "app.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or f
        sys.stderr = sys.stderr or f


def _chamar(url, metodo="GET", timeout=3):
    req = urllib.request.Request(url, method=metodo, data=b"{}" if metodo == "POST" else None,
                                 headers={"X-Painel": "1", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


class App:
    def __init__(self, iniciar_oculto):
        self.cfg = comum.carregar_config()
        self.url = "http://localhost:%d" % int(self.cfg.get("porta_painel", 8765))
        self.iniciar_oculto = iniciar_oculto
        self.janela = None
        self.icone = None
        self.saindo = False
        self.servidor = None
        self.webview_ok = False

    # ------------------------------------------------ servidor
    def iniciar_servidor(self):
        import servidor
        try:
            self.servidor = servidor.criar_servidor(self.cfg)
        except OSError:
            # porta ocupada: outra cópia do app já está aberta? pede para ela aparecer
            try:
                if not self.iniciar_oculto and _chamar(self.url + "/api/janela/mostrar", "POST").get("ok"):
                    return "outra_instancia"
            except Exception:
                pass
            try:
                if _chamar(self.url + "/saude").get("ok"):
                    print("Servidor já está no ar por outro processo; abrindo só a janela.")
                    return "sem_servidor_proprio"
            except Exception:
                pass
            raise
        servidor.GANCHOS["mostrar_janela"] = self.mostrar
        servidor.GANCHOS["sair"] = self.sair          # botão "Atualizar sistema"
        threading.Thread(target=self.servidor.serve_forever, daemon=True).start()
        return "ok"

    # ------------------------------------------------ janela
    def mostrar(self):
        if self.janela is not None and self.webview_ok:
            try:
                self.janela.show()
                self.janela.restore()
                return
            except Exception:
                pass
        webbrowser.open(self.url)

    def _ao_fechar(self):
        if self.saindo or self.icone is None:
            return True
        self.janela.hide()  # fechar = esconder; o app continua na bandeja
        return False

    # ------------------------------------------------ bandeja
    def _imagem_icone(self):
        from PIL import Image, ImageDraw
        arq = os.path.join(comum.PASTA_SISTEMA, "icone.ico")
        if os.path.exists(arq):
            return Image.open(arq)
        img = Image.new("RGBA", (64, 64), (36, 87, 214, 255))
        d = ImageDraw.Draw(img)
        for i, h in enumerate((18, 30, 42)):
            d.rectangle((12 + i * 15, 52 - h, 22 + i * 15, 52), fill="white")
        return img

    def iniciar_bandeja(self):
        try:
            import pystray
        except Exception as e:  # sem bandeja, o app ainda funciona com a janela
            print("Bandeja indisponível:", e)
            return
        M = pystray.MenuItem
        menu = pystray.Menu(
            M("Abrir painel", lambda *_: self.mostrar(), default=True),
            M("Rodar organizador agora", lambda *_: self._organizar()),
            M("Abrir pasta de hoje", lambda *_: self._abrir_pasta_hoje()),
            M("Baixar relatórios de todos agora", lambda *_: self._baixar_todos()),
            M("Abrir Chrome com os perfis", lambda *_: self._abrir_chrome()),
            pystray.Menu.SEPARATOR,
            M("Sair (para a rotina de amanhã)", lambda *_: self.sair()),
        )
        self.icone = pystray.Icon("relatorios", self._imagem_icone(), TITULO, menu)
        self.icone.run_detached()
        threading.Thread(target=self._vigiar_status, daemon=True).start()

    def _notificar(self, texto):
        try:
            if self.icone:
                self.icone.notify(texto, TITULO)
        except Exception:
            pass

    def _organizar(self):
        import organizador
        try:
            r = organizador.executar(forcar_estavel=True)
            self._notificar("%d organizado(s), %d não reconhecido(s), %d erro(s)." % (r["organizados"], r["nao_reconhecidos"], r["erros"]))
        except Exception as e:
            self._notificar("Organizador: %s" % e)

    def _abrir_pasta_hoje(self):
        cfg = comum.carregar_config()
        p = os.path.join(comum.raiz(cfg), comum.rotulo_dia())
        alvo = p if os.path.isdir(p) else comum.raiz(cfg)
        if os.name == "nt":
            os.startfile(alvo)

    def _baixar_todos(self):
        import executor
        r = executor.rodar_em_segundo_plano(origem="menu da bandeja")
        self._notificar("Baixando relatórios… acompanhe no painel." if r.get("ok") else r.get("erro", "erro"))

    def _abrir_chrome(self):
        import abrir_chrome
        r = abrir_chrome.abrir()
        self._notificar("Chrome aberto em %d perfil(is)." % len(r.get("perfis", [])) if r.get("ok") else r.get("erro", "erro"))

    def _vigiar_status(self):
        """Atualiza a dica do ícone e avisa uma vez por dia, depois do horário final do organizador."""
        import servidor
        marca = os.path.join(comum.PASTA_SISTEMA, "dados", "ultimo_aviso.txt")
        while not self.saindo:
            try:
                cfg = comum.carregar_config()
                r = servidor.calcular_status(cfg)["resumo"]
                if self.icone:
                    self.icone.title = "%s — hoje: %d de %d recebidos" % (TITULO, r["recebidos"], r["esperados"])
                fim = dt.datetime.strptime(cfg["horarios"]["organizador_fim"], "%H:%M").time()
                hoje = dt.date.today().isoformat()
                avisado = open(marca).read().strip() if os.path.exists(marca) else ""
                if r["esperados"] and dt.datetime.now().time() >= fim and avisado != hoje:
                    pend = r["faltando"] + r["erro"]
                    self._notificar("Todos os %d relatórios de hoje chegaram." % r["esperados"] if not pend
                                    else "%d de %d relatórios faltando ou com erro. Abra o painel." % (pend, r["esperados"]))
                    os.makedirs(os.path.dirname(marca), exist_ok=True)
                    open(marca, "w").write(hoje)
            except Exception as e:
                print("vigia:", e)
            time.sleep(60)

    def sair(self):
        self.saindo = True
        if self.icone:
            try:
                self.icone.stop()
            except Exception:
                pass
        if self.servidor:
            threading.Thread(target=self.servidor.shutdown, daemon=True).start()
        if self.janela is not None and self.webview_ok:
            try:
                self.janela.destroy()
                return
            except Exception:
                pass
        os._exit(0)

    # ------------------------------------------------ principal
    def rodar(self):
        estado = self.iniciar_servidor()
        if estado == "outra_instancia":
            return
        self.iniciar_bandeja()
        try:
            import webview
            webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
            webview.settings["ALLOW_DOWNLOADS"] = True          # botão Baixar (pacote final)
            self.janela = webview.create_window(TITULO, self.url, width=1240, height=860, min_size=(820, 560),
                                                hidden=self.iniciar_oculto, text_select=True)
            self.janela.events.closing += self._ao_fechar
            self.webview_ok = True
            armazenamento = os.path.join(comum.PASTA_SISTEMA, "dados", "janela")
            webview.start(gui="edgechromium" if os.name == "nt" else None, private_mode=False, storage_path=armazenamento)
        except Exception as e:
            # sem WebView2: usa o navegador padrão e segue na bandeja
            print("Janela própria indisponível (%s); usando o navegador." % e)
            self.webview_ok = False
            if not self.iniciar_oculto:
                webbrowser.open(self.url)
            while not self.saindo:  # mantém o servidor no ar mesmo sem janela
                time.sleep(1)
        if not self.saindo:  # janela encerrada sem bandeja
            self.sair()


def main(argv):
    _log_sem_console()
    if "--organizar" in argv:
        try:
            import rede
            if rede.papel() == "estacao":   # estação: só envia os arquivos baixados aqui para o computador principal
                import estacao
                estacao.enviar_pendentes()
                return
        except Exception as e:
            comum.log("Organizador (estação): %s" % e)
            return
        import organizador
        organizador.executar()
        return
    if "--rodar-todos" in argv:
        # Tarefa das 5h. Se o app já estiver aberto na bandeja, pede a ele (mostra o progresso no painel).
        cfg = comum.carregar_config()
        estacao_rede = False
        try:
            import rede
            if rede.papel() == "estacao":
                estacao_rede = True
                cfg = dict(rede.config_mesclada(), porta_painel=cfg.get("porta_painel", 8765))
        except Exception as e:
            comum.log("Tarefa agendada: computador principal inacessível (%s)." % e)
            return
        if not cfg.get("execucao", {}).get("agendado", False):
            comum.log("Download automático desligado: a tarefa agendada não baixou nada (use \"Baixar todos agora\").")
            return
        url = "http://localhost:%d" % int(cfg.get("porta_painel", 8765))
        try:
            if _chamar(url + "/api/executar/agendado", "POST", timeout=10).get("ok"):
                return
        except Exception:
            pass
        if estacao_rede:
            import estacao
            estacao.rodar_local(None, None, "agendamento")
            return
        import executor
        executor.rodar(origem="agendamento")
        return
    if "--abrir-chrome" in argv:
        import abrir_chrome
        abrir_chrome.abrir()
        return
    App(iniciar_oculto="--bandeja" in argv).rodar()


if __name__ == "__main__":
    main(sys.argv[1:])
