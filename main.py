# =============================================================================
# CIEBP - Descanso de Tela Institucional
# Versão: 1.0
# Descrição: Tela de recepção profissional para eventos do CIEBP.
#            Exibe informações do evento, relógio em tempo real e
#            toca rádio online enquanto aguarda o início da atividade.
# =============================================================================

import tkinter as tk
from tkinter import ttk, font as tkfont, colorchooser
import json
import os
import sys
import time
import math
import random
import threading
import queue
import subprocess
import shutil
from datetime import datetime
from pathlib import Path

# Pillow: manipulação de imagens
try:
    from PIL import Image, ImageTk, ImageFilter, ImageDraw
    PIL_DISPONIVEL = True
except ImportError:
    PIL_DISPONIVEL = False
    print("[AVISO] Pillow não instalado. As imagens de fundo não serão exibidas.")
    print("        Execute: pip install Pillow")

# pywin32 (opcional): Windows Media Player via COM — controle completo de volume.
# Sem ele, o programa usa PowerShell como fallback (built-in no Windows 10/11).
try:
    import pythoncom
    import win32com.client
    WMP_DISPONIVEL = True
except ImportError:
    WMP_DISPONIVEL = False

try:
    import pygame
    PYGAME_DISPONIVEL = True
except ImportError:
    PYGAME_DISPONIVEL = False

# Verifica se o PowerShell está disponível (fallback de áudio sem pip)
PS_DISPONIVEL = shutil.which("powershell") is not None

# Diretório base: funciona tanto rodando pelo Python quanto como .exe PyInstaller
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def carregar_config():
    """Carrega as configurações do arquivo config.json."""
    caminho = BASE_DIR / "config.json"
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"[ERRO] Arquivo config.json não encontrado em: {caminho}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"[ERRO] config.json com formato inválido: {e}")
        sys.exit(1)


def salvar_config(config):
    """Persiste as configurações atuais no config.json."""
    caminho = BASE_DIR / "config.json"
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[AVISO] Não foi possível salvar config.json: {e}")


def criar_imagem_fallback(largura, altura, cor_destaque="#00BCD4"):
    """
    Cria uma imagem de fundo padrão caso a imagem do espaço não seja encontrada.
    Retorna um objeto PIL.Image com gradiente escuro institucional.
    """
    img = Image.new("RGB", (largura, altura), "#0D1B2A")
    draw = ImageDraw.Draw(img)

    # Gradiente manual: faixas horizontais variando a luminosidade
    for y in range(altura):
        fator = y / altura
        r = int(13 + fator * 20)
        g = int(27 + fator * 30)
        b = int(42 + fator * 50)
        draw.line([(0, y), (largura, y)], fill=(r, g, b))

    # Acento de cor do espaço no topo e rodapé
    try:
        r_hex = int(cor_destaque[1:3], 16)
        g_hex = int(cor_destaque[3:5], 16)
        b_hex = int(cor_destaque[5:7], 16)
    except Exception:
        r_hex, g_hex, b_hex = 0, 188, 212

    for espessura in range(6):
        draw.line([(0, espessura), (largura, espessura)],
                  fill=(r_hex, g_hex, b_hex))
        draw.line([(0, altura - 1 - espessura), (largura, altura - 1 - espessura)],
                  fill=(r_hex, g_hex, b_hex))

    return img


def redimensionar_imagem(img_pil, largura, altura):
    """Redimensiona a imagem PIL para cobrir toda a tela (cover)."""
    prop_tela = largura / altura
    prop_img = img_pil.width / img_pil.height

    if prop_img > prop_tela:
        # Imagem mais larga que a tela: ajusta pela altura
        novo_h = altura
        novo_w = int(img_pil.width * altura / img_pil.height)
    else:
        # Imagem mais alta que a tela: ajusta pela largura
        novo_w = largura
        novo_h = int(img_pil.height * largura / img_pil.width)

    img_redim = img_pil.resize((novo_w, novo_h), Image.LANCZOS)

    # Recorta o centro
    x0 = (novo_w - largura) // 2
    y0 = (novo_h - altura) // 2
    return img_redim.crop((x0, y0, x0 + largura, y0 + altura))


def escurecer_imagem(img_pil, fator=0.45):
    """Aplica uma camada escura semitransparente sobre a imagem para facilitar a leitura."""
    overlay = Image.new("RGBA", img_pil.size, (0, 0, 0, int(255 * fator)))
    img_rgba = img_pil.convert("RGBA")
    combinada = Image.alpha_composite(img_rgba, overlay)
    return combinada.convert("RGB")


# =============================================================================
# JANELA DE CONFIGURAÇÃO (LAUNCHER)
# =============================================================================

class JanelaConfig(tk.Tk):
    """
    Janela inicial onde o usuário escolhe o espaço e a rádio
    antes de iniciar o descanso de tela em tela cheia.
    """

    def __init__(self, config):
        super().__init__()
        self.config_dados = config
        self.data_atual = datetime.now().strftime("%d/%m/%Y")
        self.espaco_escolhido = tk.StringVar(value=config.get("espaco_padrao", "Hub de Inovação"))
        self.radio_escolhida = tk.StringVar(value=config.get("radio_padrao", "Música local (MP3)"))

        # Variáveis dos campos editáveis do evento
        self.var_evento = tk.StringVar(value=config.get("evento", "Evento CIEBP"))
        self.var_data = tk.StringVar(value=self.data_atual)
        self.var_professores = tk.StringVar(value=", ".join(config.get("professores", [])))
        self.var_mensagem = tk.StringVar(value=config.get("mensagem_boas_vindas", ""))
        cores_texto = config.get("cores_texto", {})
        self.var_cores_texto = {
            "evento": tk.StringVar(value=cores_texto.get("evento", "#FB1874")),
            "data": tk.StringVar(value=cores_texto.get("data", "#FB1874")),
            "professores": tk.StringVar(value=cores_texto.get("professores", "#FB1874")),
            "mensagem": tk.StringVar(value=cores_texto.get("mensagem", "#E6EEF5")),
        }
        self._botoes_cor_texto = {}
        intermitencias = config.get("intermitencia_textos", {})
        self.var_intermitencia_textos = {
            "hora": tk.StringVar(value=intermitencias.get("hora", "Desligado")),
            "evento": tk.StringVar(value=intermitencias.get("evento", "Desligado")),
            "legenda": tk.StringVar(value=intermitencias.get("legenda", "Desligado")),
            "data": tk.StringVar(value=intermitencias.get("data", "Desligado")),
            "professores": tk.StringVar(value=intermitencias.get("professores", "Desligado")),
            "mensagem": tk.StringVar(value=intermitencias.get("mensagem", "Desligado")),
        }
        opacidades = config.get("opacidade_textos", {})
        self.var_opacidade_textos = {
            chave: tk.StringVar(value=opacidades.get(chave, "100%"))
            for chave in ("hora", "evento", "legenda", "data", "professores", "mensagem")
        }
        self.var_posicao_relogio = tk.StringVar(value=config.get("posicao_relogio", "Central (grande)"))
        self.var_audio_local = tk.StringVar(value=config.get("audio_local", ""))
        self.var_fundo = tk.StringVar(value=config.get("fundo_padrao", "Padrão do espaço"))
        self.var_alerta_teste = tk.StringVar(value="Aviso de almoço")
        self.var_intervalo_tubarao = tk.StringVar(
            value=str(config.get("intervalo_tubarao_minutos", 10))
        )
        agendamentos = config.get("agendamentos_banner", [])
        self.var_agendamentos = []
        for indice in range(8):
            item = agendamentos[indice] if indice < len(agendamentos) else {}
            horario = "".join(caractere for caractere in str(item.get("horario", "")) if caractere.isdigit())[:4]
            if len(horario) == 4:
                horario = f"{horario[:2]}:{horario[2:]}"
            self.var_agendamentos.append((
                tk.StringVar(value=horario),
                tk.StringVar(value=str(item.get("mensagem", ""))),
            ))

        self.title("CIEBP - Descanso de Tela")
        self.resizable(False, True)
        self._construir_interface()
        self._ajustar_altura()

    def _ajustar_altura(self):
        """
        Calcula a altura ideal da janela levando em conta o conteúdo rolável
        e o botão fixo no fundo, limitando pela área útil real da tela.
        """
        self.update_idletasks()
        larg = 580

        # Altura do conteúdo rolável
        alt_conteudo = self._frame_interno.winfo_reqheight() + 4

        # Altura do botão fixo
        alt_btn = self._btn_iniciar.winfo_reqheight()

        # Desconta: barra de tarefas (~40px) + barra de título (~30px) + margens (~50px)
        alt_tela_util = self.winfo_screenheight() - 180

        # Espaço para o scroll = total útil - botão fixo
        alt_scroll_max = alt_tela_util - alt_btn
        alt_scroll = min(alt_conteudo, alt_scroll_max)
        alt_janela = alt_scroll + alt_btn

        x = (self.winfo_screenwidth() - larg) // 2
        y = max(5, (self.winfo_screenheight() - alt_janela) // 2)
        self.geometry(f"{larg}x{alt_janela}+{x}+{y}")

        # Habilita scrollbar apenas se o conteúdo não coube na área rolável
        if alt_conteudo > alt_scroll:
            self._scrollbar.pack(side="right", fill="y")
            self._canvas_scroll.pack(side="left", fill="both", expand=True)
        else:
            self._scrollbar.pack_forget()

        self._canvas_scroll.configure(scrollregion=self._canvas_scroll.bbox("all"))

    def _construir_interface(self):
        """Monta os widgets dentro de um frame rolável."""
        self.configure(bg="#0D1B2A")

        # --- Container rolável ---
        frame_outer = tk.Frame(self, bg="#0D1B2A")
        frame_outer.pack(fill="both", expand=True)

        self._scrollbar = tk.Scrollbar(frame_outer, orient="vertical",
                                        bg="#1E3A5F", troughcolor="#0D1B2A",
                                        activebackground="#00BCD4")
        self._canvas_scroll = tk.Canvas(frame_outer, bg="#0D1B2A",
                                         highlightthickness=0,
                                         yscrollcommand=self._scrollbar.set)
        self._scrollbar.config(command=self._canvas_scroll.yview)
        self._canvas_scroll.pack(side="left", fill="both", expand=True)
        # scrollbar só aparece se necessário (ver _ajustar_altura)

        self._frame_interno = tk.Frame(self._canvas_scroll, bg="#0D1B2A")
        _win_id = self._canvas_scroll.create_window(
            (0, 0), window=self._frame_interno, anchor="nw"
        )

        # Ajusta largura do frame interno ao canvas
        def _on_resize(e):
            self._canvas_scroll.itemconfig(_win_id, width=e.width)
        self._canvas_scroll.bind("<Configure>", _on_resize)

        # Rola com a roda do mouse
        def _on_wheel(e):
            self._canvas_scroll.yview_scroll(-1 * (e.delta // 120), "units")
        self._canvas_scroll.bind("<MouseWheel>", _on_wheel)
        self._frame_interno.bind("<MouseWheel>", _on_wheel)

        F = self._frame_interno   # atalho para adicionar widgets

        # --- Cabeçalho ---
        frame_cabecalho = tk.Frame(F, bg="#0D1B2A")
        frame_cabecalho.pack(fill="x", padx=30, pady=(12, 4))

        tk.Label(
            frame_cabecalho,
            text="CIEBP",
            font=("Segoe UI", 22, "bold"),
            fg="#00BCD4",
            bg="#0D1B2A"
        ).pack()

        tk.Label(
            frame_cabecalho,
            text="Tela de espera para eventos do CIEBP",
            font=("Segoe UI", 10),
            fg="#89A0B7",
            bg="#0D1B2A",
            wraplength=500,
            justify="center"
        ).pack(pady=(1, 0))

        # Separador
        tk.Frame(F, bg="#1E3A5F", height=1).pack(fill="x", padx=30, pady=8)

        # --- Configuração do evento ---
        frame_ev = tk.Frame(F, bg="#0D1B2A")
        frame_ev.pack(fill="x", padx=30, pady=(0, 4))

        tk.Label(frame_ev, text="Dados do evento",
                 font=("Segoe UI", 11, "bold"), fg="#E3EDF5",
                 bg="#0D1B2A", anchor="w").pack(fill="x", pady=(0, 2))

        def _campo(rotulo, var, dica="", parent=frame_ev):
            """Cria um par rótulo + campo de texto estilizado."""
            tk.Label(parent, text=rotulo, font=("Segoe UI", 9, "bold"),
                     fg="#B8C7D6", bg="#0D1B2A", anchor="w").pack(fill="x")
            if dica:
                tk.Label(
                    parent,
                    text=dica,
                    font=("Segoe UI", 8),
                    fg="#7E94A8",
                    bg="#0D1B2A",
                    anchor="w",
                    justify="left",
                    wraplength=500,
                ).pack(fill="x", pady=(0, 2))
            e = tk.Entry(parent, textvariable=var, font=("Segoe UI", 11),
                         bg="#132336", fg="#ECEFF1", insertbackground="#ECEFF1",
                         relief="flat", bd=0, highlightthickness=1,
                         highlightbackground="#1E3A5F",
                         highlightcolor="#00BCD4")
            e.pack(fill="x", ipady=4, pady=(0, 6))
            e.bind("<MouseWheel>", _on_wheel)

        _campo("Nome do evento", self.var_evento)
        _campo("Data do evento", self.var_data, "O campo já vem com a data de hoje, mas você pode alterar.")
        _campo("Professores", self.var_professores, "Separe os nomes por vírgula.")
        _campo("Mensagem da tela", self.var_mensagem, "Esse texto aparece na tela principal.")

        frame_cores = tk.Frame(F, bg="#10243A")
        frame_cores.pack(fill="x", padx=30, pady=(2, 8))
        tk.Label(frame_cores, text="Cores dos textos", font=("Segoe UI", 10, "bold"), fg="#E3EDF5", bg="#10243A", anchor="w").pack(fill="x", padx=10, pady=(8, 1))
        tk.Label(frame_cores, text="Escolha uma cor para cada texto do evento.", font=("Segoe UI", 8), fg="#9CB2C6", bg="#10243A", anchor="w").pack(fill="x", padx=10, pady=(0, 5))
        for chave, rotulo in (("evento", "Nome do evento"), ("data", "Data"), ("professores", "Professores"), ("mensagem", "Mensagem")):
            linha_cor = tk.Frame(frame_cores, bg="#10243A")
            linha_cor.pack(fill="x", padx=10, pady=2)
            tk.Label(linha_cor, text=rotulo, font=("Segoe UI", 9, "bold"), fg="#B8C7D6", bg="#10243A", width=19, anchor="w").pack(side="left")
            botao_cor = tk.Button(linha_cor, textvariable=self.var_cores_texto[chave], font=("Consolas", 9, "bold"), fg="#0D1B2A", bg=self.var_cores_texto[chave].get(), activeforeground="#0D1B2A", activebackground=self.var_cores_texto[chave].get(), relief="flat", bd=0, cursor="hand2", command=lambda item=chave: self._selecionar_cor_texto(item))
            botao_cor.pack(side="right", ipadx=8, ipady=3)
            self._botoes_cor_texto[chave] = botao_cor

        frame_intermitencia = tk.Frame(F, bg="#10243A")
        frame_intermitencia.pack(fill="x", padx=30, pady=(0, 8))
        tk.Label(frame_intermitencia, text="Intermitencia dos textos", font=("Segoe UI", 10, "bold"), fg="#E3EDF5", bg="#10243A", anchor="w").pack(fill="x", padx=10, pady=(8, 1))
        tk.Label(frame_intermitencia, text="Desligado mantem visivel; Ocultar remove o elemento da tela; as demais opcoes definem o intervalo de piscar.", font=("Segoe UI", 8), fg="#9CB2C6", bg="#10243A", anchor="w", wraplength=500, justify="left").pack(fill="x", padx=10, pady=(0, 5))
        opcoes_intermitencia = ("Desligado", "Ocultar", "1 segundo", "2 segundos", "5 segundos", "10 segundos", "30 segundos", "1 minuto", "2 minutos", "5 minutos")
        for chave, rotulo in (("hora", "Hora"), ("evento", "Nome do evento"), ("legenda", "Legenda do espaco"), ("data", "Data"), ("professores", "Professores"), ("mensagem", "Mensagem")):
            linha_intermitencia = tk.Frame(frame_intermitencia, bg="#10243A")
            linha_intermitencia.pack(fill="x", padx=10, pady=2)
            tk.Label(linha_intermitencia, text=rotulo, font=("Segoe UI", 9, "bold"), fg="#B8C7D6", bg="#10243A", width=19, anchor="w").pack(side="left")
            combo_intermitencia = ttk.Combobox(linha_intermitencia, textvariable=self.var_intermitencia_textos[chave], values=opcoes_intermitencia, state="readonly", width=15, font=("Segoe UI", 9), style="Institucional.TCombobox")
            combo_intermitencia.pack(side="right", ipady=2)

        frame_aparencia = tk.Frame(F, bg="#10243A")
        frame_aparencia.pack(fill="x", padx=30, pady=(0, 8))
        tk.Label(frame_aparencia, text="Opacidade e posicao do relogio", font=("Segoe UI", 10, "bold"), fg="#E3EDF5", bg="#10243A", anchor="w").pack(fill="x", padx=10, pady=(8, 1))
        tk.Label(frame_aparencia, text="A opacidade reduz a intensidade visual sem ocultar o texto. O relogio compacto fica no canto inferior direito.", font=("Segoe UI", 8), fg="#9CB2C6", bg="#10243A", anchor="w", wraplength=500, justify="left").pack(fill="x", padx=10, pady=(0, 5))
        opcoes_opacidade = ("100%", "75%", "50%", "25%")
        for chave, rotulo in (("hora", "Hora"), ("evento", "Nome do evento"), ("legenda", "Legenda do espaco"), ("data", "Data"), ("professores", "Professores"), ("mensagem", "Mensagem")):
            linha_opacidade = tk.Frame(frame_aparencia, bg="#10243A")
            linha_opacidade.pack(fill="x", padx=10, pady=2)
            tk.Label(linha_opacidade, text=rotulo, font=("Segoe UI", 9, "bold"), fg="#B8C7D6", bg="#10243A", width=19, anchor="w").pack(side="left")
            combo_opacidade = ttk.Combobox(linha_opacidade, textvariable=self.var_opacidade_textos[chave], values=opcoes_opacidade, state="readonly", width=15, font=("Segoe UI", 9), style="Institucional.TCombobox")
            combo_opacidade.pack(side="right", ipady=2)

        linha_relogio = tk.Frame(frame_aparencia, bg="#10243A")
        linha_relogio.pack(fill="x", padx=10, pady=(7, 8))
        tk.Label(linha_relogio, text="Posicao do relogio", font=("Segoe UI", 9, "bold"), fg="#B8C7D6", bg="#10243A", width=19, anchor="w").pack(side="left")
        combo_relogio = ttk.Combobox(linha_relogio, textvariable=self.var_posicao_relogio, values=("Central (grande)", "Inferior direito (compacto)"), state="readonly", width=24, font=("Segoe UI", 9), style="Institucional.TCombobox")
        combo_relogio.pack(side="right", ipady=2)

        # Separador
        tk.Frame(F, bg="#1E3A5F", height=1).pack(fill="x", padx=30, pady=(2, 6))

        # --- Seleção de Espaço ---
        frame_espaco = tk.Frame(F, bg="#0D1B2A")
        frame_espaco.pack(fill="x", padx=30, pady=(0, 6))

        tk.Label(
            frame_espaco,
            text="Espaço do evento",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))

        combo_estilo = ttk.Style()
        combo_estilo.theme_use("clam")
        combo_estilo.configure(
            "Institucional.TCombobox",
            fieldbackground="#1E3A5F",
            background="#1E3A5F",
            foreground="#FFFFFF",
            selectforeground="#FFFFFF",
            selectbackground="#00BCD4",
            bordercolor="#00BCD4",
            arrowcolor="#00BCD4",
            font=("Segoe UI", 11)
        )
        combo_estilo.map("Institucional.TCombobox",
            fieldbackground=[("readonly", "#1E3A5F"), ("disabled", "#0D1B2A")],
            foreground=[("readonly", "#FFFFFF"), ("disabled", "#546E7A")],
            selectbackground=[("readonly", "#00BCD4")],
            selectforeground=[("readonly", "#FFFFFF")],
            background=[("readonly", "#1E3A5F"), ("active", "#00BCD4")],
        )
        # Estilo do menu suspenso (listbox)
        self.option_add("*TCombobox*Listbox.background", "#1E3A5F")
        self.option_add("*TCombobox*Listbox.foreground", "#FFFFFF")
        self.option_add("*TCombobox*Listbox.selectBackground", "#00BCD4")
        self.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")

        espacos = list(self.config_dados.get("espacos", {}).keys())
        self.combo_espaco = ttk.Combobox(
            frame_espaco,
            textvariable=self.espaco_escolhido,
            values=espacos,
            state="readonly",
            style="Institucional.TCombobox",
            font=("Segoe UI", 11),
            height=8
        )
        self.combo_espaco.pack(fill="x", ipady=4)

        # --- Seleção de Imagem de Fundo ---
        frame_fundo = tk.Frame(F, bg="#0D1B2A")
        frame_fundo.pack(fill="x", padx=30, pady=(0, 6))

        tk.Label(
            frame_fundo,
            text="Imagem de fundo",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))

        base_dir = BASE_DIR
        opcoes_fundo = ["Padrão do espaço"] + [
            f"Fundo {i}" for i in range(1, 7)
            if (base_dir / "assets" / f"fundo{i}.png").exists()
        ]
        opcoes_fundo.extend(("Ilha do Pescador", "Mundo de Blocos", "Voo no Rio"))
        self.combo_fundo = ttk.Combobox(
            frame_fundo,
            textvariable=self.var_fundo,
            values=opcoes_fundo,
            state="readonly",
            style="Institucional.TCombobox",
            font=("Segoe UI", 11),
            height=8
        )
        self.combo_fundo.pack(fill="x", ipady=4)

        frame_tubarao = tk.Frame(F, bg="#0D1B2A")
        frame_tubarao.pack(fill="x", padx=30, pady=(0, 6))

        tk.Label(
            frame_tubarao,
            text="Intervalo do tubarão",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))

        tk.Label(
            frame_tubarao,
            text="Informe em minutos. Esse ajuste vale para o fundo Ilha do Pescador. Se colocar 0, o tubarão fica desativado.",
            font=("Segoe UI", 8),
            fg="#7E94A8",
            bg="#0D1B2A",
            anchor="w",
            justify="left",
            wraplength=500,
        ).pack(fill="x", pady=(0, 2))

        self._entry_tubarao = tk.Entry(
            frame_tubarao,
            textvariable=self.var_intervalo_tubarao,
            font=("Segoe UI", 11),
            bg="#132336",
            fg="#ECEFF1",
            insertbackground="#ECEFF1",
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground="#1E3A5F",
            highlightcolor="#00BCD4",
        )
        self._entry_tubarao.pack(fill="x", ipady=4)
        self._entry_tubarao.bind("<MouseWheel>", _on_wheel)

        # --- Agendamentos de mensagens em banner aereo ---
        frame_agendamentos = tk.Frame(F, bg="#0D1B2A")
        frame_agendamentos.pack(fill="x", padx=30, pady=(8, 6))
        tk.Label(
            frame_agendamentos,
            text="Mensagens agendadas no avião",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w",
        ).pack(fill="x", pady=(0, 2))
        tk.Label(
            frame_agendamentos,
            text="Preencha até 8 horários no formato HH:MM. Em cada horário, um avião leva a mensagem em um banner no alto da tela.",
            font=("Segoe UI", 8),
            fg="#7E94A8",
            bg="#0D1B2A",
            anchor="w",
            justify="left",
            wraplength=500,
        ).pack(fill="x", pady=(0, 5))

        cabecalho = tk.Frame(frame_agendamentos, bg="#0D1B2A")
        cabecalho.pack(fill="x", pady=(0, 2))
        tk.Label(cabecalho, text="Horário", font=("Segoe UI", 8, "bold"), fg="#B8C7D6", bg="#0D1B2A", width=10, anchor="w").pack(side="left")
        tk.Label(cabecalho, text="Mensagem", font=("Segoe UI", 8, "bold"), fg="#B8C7D6", bg="#0D1B2A", anchor="w").pack(side="left", fill="x", expand=True)
        for indice, (var_horario, var_mensagem) in enumerate(self.var_agendamentos, start=1):
            linha = tk.Frame(frame_agendamentos, bg="#0D1B2A")
            linha.pack(fill="x", pady=2)
            tk.Label(linha, text=f"{indice}.", font=("Segoe UI", 9), fg="#7E94A8", bg="#0D1B2A", width=3, anchor="w").pack(side="left")
            campo_horario = tk.Entry(
                linha, textvariable=var_horario, width=8, font=("Segoe UI", 10),
                bg="#132336", fg="#ECEFF1", insertbackground="#ECEFF1", relief="flat", bd=0,
                highlightthickness=1, highlightbackground="#1E3A5F", highlightcolor="#00BCD4",
            )
            campo_horario.pack(side="left", ipady=3, padx=(0, 5))
            campo_mensagem = tk.Entry(
                linha, textvariable=var_mensagem, font=("Segoe UI", 10),
                bg="#132336", fg="#ECEFF1", insertbackground="#ECEFF1", relief="flat", bd=0,
                highlightthickness=1, highlightbackground="#1E3A5F", highlightcolor="#00BCD4",
            )
            campo_mensagem.pack(side="left", fill="x", expand=True, ipady=3)
            campo_horario.bind("<MouseWheel>", _on_wheel)
            campo_horario.bind("<KeyRelease>", self._formatar_horario_agendamento)
            campo_horario.bind("<FocusOut>", self._formatar_horario_agendamento)
            campo_mensagem.bind("<MouseWheel>", _on_wheel)

        # --- Seleção de Rádio ---
        frame_radio = tk.Frame(F, bg="#0D1B2A")
        frame_radio.pack(fill="x", padx=30, pady=(0, 6))

        tk.Label(
            frame_radio,
            text="Áudio de fundo",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))

        radios = list(self.config_dados.get("radios", {}).keys())
        self.combo_radio = ttk.Combobox(
            frame_radio,
            textvariable=self.radio_escolhida,
            values=radios,
            state="readonly",
            style="Institucional.TCombobox",
            font=("Segoe UI", 11),
            height=8
        )
        self.combo_radio.pack(fill="x", ipady=4)
        self.combo_radio.bind("<<ComboboxSelected>>", self._atualizar_campo_mp3)

        # --- Arquivo de áudio local (MP3) ---
        self._frame_mp3 = tk.Frame(frame_radio, bg="#0D1B2A")
        if self.radio_escolhida.get() == "Música local (MP3)":
            self._frame_mp3.pack(fill="x", pady=(4, 0))

        tk.Label(self._frame_mp3, text="Arquivo de áudio (MP3, WAV, WMA...)",
                 font=("Segoe UI", 9, "bold"), fg="#B8C7D6",
                 bg="#0D1B2A", anchor="w").pack(fill="x")
        frame_mp3_row = tk.Frame(self._frame_mp3, bg="#0D1B2A")
        frame_mp3_row.pack(fill="x")
        self._entry_mp3 = tk.Entry(
            frame_mp3_row, textvariable=self.var_audio_local,
            font=("Segoe UI", 9), bg="#132336", fg="#ECEFF1",
            insertbackground="#ECEFF1", relief="flat", bd=0,
            highlightthickness=1, highlightbackground="#1E3A5F",
            highlightcolor="#00BCD4")
        self._entry_mp3.pack(side="left", fill="x", expand=True, ipady=4)
        self._entry_mp3.bind("<MouseWheel>", _on_wheel)
        tk.Label(
            self._frame_mp3,
            text="Escolha um arquivo local para tocar durante a espera.",
            font=("Segoe UI", 8),
            fg="#7E94A8",
            bg="#0D1B2A",
            anchor="w",
            justify="left",
            wraplength=500,
        ).pack(fill="x", pady=(0, 2))
        tk.Button(
            frame_mp3_row, text="Escolher arquivo",
            font=("Segoe UI", 9), fg="#FFFFFF", bg="#1E3A5F",
            activebackground="#00BCD4", activeforeground="#FFFFFF",
            relief="flat", cursor="hand2", bd=0, padx=10,
            command=self._selecionar_mp3
        ).pack(side="right", padx=(4, 0))

        # --- Status do áudio ---
        frame_ast = tk.Frame(F, bg="#0D1B2A")
        frame_ast.pack(fill="x", padx=30, pady=(4, 2))
        if PYGAME_DISPONIVEL:
            _cor_a = "#4CAF50"
            _txt_a = "Áudio pronto para uso."
        elif WMP_DISPONIVEL:
            _cor_a = "#4CAF50"
            _txt_a = "Áudio pronto para uso."
        elif PS_DISPONIVEL:
            _cor_a = "#FF9800"
            _txt_a = "Áudio disponível em modo compatível."
        else:
            _cor_a = "#F44336"
            _txt_a = "Áudio indisponível. Instale pywin32 para habilitar."
        tk.Label(frame_ast, text=_txt_a, font=("Segoe UI", 8),
                 fg=_cor_a, bg="#0D1B2A", anchor="w",
                 wraplength=462, justify="left").pack(fill="x")

        # --- Dica de atalhos ---
        frame_dica = tk.Frame(F, bg="#0F2336", bd=0)
        frame_dica.pack(fill="x", padx=30, pady=(4, 6))
        tk.Label(
            frame_dica,
            text="Atalhos: ESC fecha a tela, setas ajustam o volume e M liga ou desliga o som.",
            font=("Segoe UI", 8),
            fg="#607D8B",
            bg="#0F2336",
            justify="left",
            padx=12,
            pady=4
        ).pack(fill="x")

        # --- Teste de aviso ---
        frame_teste = tk.Frame(F, bg="#0D1B2A")
        frame_teste.pack(fill="x", padx=30, pady=(0, 6))
        tk.Label(
            frame_teste,
            text="Teste de aviso",
            font=("Segoe UI", 11, "bold"),
            fg="#E3EDF5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))
        tk.Label(
            frame_teste,
            text="Use este botão para visualizar o alerta sem esperar o horário real.",
            font=("Segoe UI", 8),
            fg="#7E94A8",
            bg="#0D1B2A",
            anchor="w",
            justify="left",
            wraplength=500,
        ).pack(fill="x", pady=(0, 4))
        self.combo_alerta_teste = ttk.Combobox(
            frame_teste,
            textvariable=self.var_alerta_teste,
            values=[
                "Aviso de almoço",
                "Saída do 1º professor",
                "Saída do 2º professor",
            ],
            state="readonly",
            style="Institucional.TCombobox",
            font=("Segoe UI", 11),
            height=6
        )
        self.combo_alerta_teste.pack(fill="x", ipady=4, pady=(0, 6))
        tk.Button(
            frame_teste,
            text="Testar aviso agora",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF",
            bg="#FF7043",
            activebackground="#F4511E",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=14,
            pady=8,
            command=self._testar_alerta
        ).pack(anchor="center")

        # --- Botão Sobre ---
        frame_sobre = tk.Frame(F, bg="#0D1B2A")
        frame_sobre.pack(fill="x", padx=30, pady=(0, 8))
        tk.Button(
            frame_sobre,
            text="Sobre o projeto",
            font=("Segoe UI", 9),
            fg="#90A4AE",
            bg="#0D1B2A",
            activebackground="#132336",
            activeforeground="#00BCD4",
            relief="flat",
            cursor="hand2",
            bd=0,
            command=self._mostrar_sobre
        ).pack(anchor="center")

        # --- Botão Iniciar (fixo no fundo da janela, fora do scroll) ---
        self._btn_iniciar = tk.Button(
            self,             # pai = janela principal, não o frame rolável
            text="Iniciar tela de espera",
            font=("Segoe UI", 13, "bold"),
            fg="#FFFFFF",
            bg="#00BCD4",
            activebackground="#0097A7",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=20,
            pady=10,
            command=self._iniciar
        )
        # pack ANTES de _ajustar_altura para que a altura seja considerada
        self._btn_iniciar.pack(side="bottom", fill="x", padx=0)

        self._btn_iniciar.bind("<Enter>", lambda e: self._btn_iniciar.configure(bg="#0097A7"))
        self._btn_iniciar.bind("<Leave>", lambda e: self._btn_iniciar.configure(bg="#00BCD4"))
        self.bind("<Return>", self._iniciar_por_tecla)
        self.bind("<KP_Enter>", self._iniciar_por_tecla)

    def _formatar_horario_agendamento(self, event=None):
        """Aplica a mascara HH:MM enquanto o horario e digitado."""
        campo = event.widget
        digitos = "".join(caractere for caractere in campo.get() if caractere.isdigit())[:4]
        if len(digitos) <= 2:
            valor_formatado = digitos
        else:
            valor_formatado = f"{digitos[:2]}:{digitos[2:]}"
        if campo.get() != valor_formatado:
            campo.delete(0, "end")
            campo.insert(0, valor_formatado)
        campo.icursor("end")

    def _selecionar_cor_texto(self, chave):
        """Abre o seletor de cores para um texto do evento."""
        cor_inicial = self.var_cores_texto[chave].get()
        _, cor_escolhida = colorchooser.askcolor(color=cor_inicial, parent=self, title="Escolha a cor do texto")
        if not cor_escolhida:
            return
        self.var_cores_texto[chave].set(cor_escolhida.upper())
        botao = self._botoes_cor_texto[chave]
        botao.configure(bg=cor_escolhida, activebackground=cor_escolhida)

    def _atualizar_campo_mp3(self, event=None):
        """Mostra/oculta o campo de arquivo MP3 conforme a seleção da rádio."""
        if self.radio_escolhida.get() == "Música local (MP3)":
            self._frame_mp3.pack(fill="x", pady=(4, 0))
        else:
            self._frame_mp3.pack_forget()
        self.after(50, self._ajustar_altura)

    def _selecionar_mp3(self):
        """Abre o diálogo de arquivo para escolher um arquivo de áudio."""
        from tkinter import filedialog
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo de áudio",
            filetypes=[
                ("Áudio", "*.mp3 *.wav *.ogg *.aac *.m4a *.wma *.flac"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if caminho:
            self.var_audio_local.set(caminho)

    def _iniciar_por_tecla(self, event=None):
        """Inicia a tela de espera ao pressionar Enter em qualquer ponto do formulario."""
        self._iniciar()
        return "break"

    def _iniciar(self):
        """Aplica as configurações do formulário e abre a tela cheia."""
        dados_tela = self._preparar_dados_tela(salvar=True)
        self.destroy()
        app = DescansoTela(*dados_tela)
        app.mainloop()

    def _testar_alerta(self):
        """Abre a tela cheia em modo de teste do aviso."""
        mapa_alerta = {
            "Aviso de almoço": "almoco",
            "Saída do 1º professor": "saida_1",
            "Saída do 2º professor": "saida_2",
        }
        alerta_teste = mapa_alerta.get(self.var_alerta_teste.get(), "almoco")
        dados_tela = self._preparar_dados_tela(salvar=False, alerta_teste=alerta_teste)
        self.destroy()
        app = DescansoTela(*dados_tela)
        app.mainloop()

    def _preparar_dados_tela(self, salvar=False, alerta_teste=None):
        """Atualiza os dados do formulário e retorna os argumentos da tela principal."""
        # Atualiza config em memória com os valores digitados
        self.config_dados["evento"] = self.var_evento.get().strip() or "Evento CIEBP"
        self.config_dados["data"] = self.var_data.get().strip() or self.data_atual
        profs_raw = self.var_professores.get()
        self.config_dados["professores"] = [p.strip() for p in profs_raw.split(",") if p.strip()]
        self.config_dados["mensagem_boas_vindas"] = self.var_mensagem.get().strip()
        self.config_dados["cores_texto"] = {
            chave: variavel.get().strip() for chave, variavel in self.var_cores_texto.items()
        }
        self.config_dados["intermitencia_textos"] = {
            chave: variavel.get().strip() for chave, variavel in self.var_intermitencia_textos.items()
        }
        self.config_dados["opacidade_textos"] = {
            chave: variavel.get().strip() for chave, variavel in self.var_opacidade_textos.items()
        }
        self.config_dados["posicao_relogio"] = self.var_posicao_relogio.get().strip()
        try:
            intervalo_tubarao = max(0, int(float(self.var_intervalo_tubarao.get().strip() or "10")))
        except ValueError:
            intervalo_tubarao = 10
        self.config_dados["intervalo_tubarao_minutos"] = intervalo_tubarao
        agendamentos = []
        for var_horario, var_mensagem in self.var_agendamentos:
            horario = var_horario.get().strip()
            mensagem = var_mensagem.get().strip()
            try:
                hora, minuto = (int(valor) for valor in horario.split(":", 1))
                horario_valido = 0 <= hora <= 23 and 0 <= minuto <= 59
            except (TypeError, ValueError):
                horario_valido = False
            if horario_valido and mensagem:
                agendamentos.append({"horario": f"{hora:02d}:{minuto:02d}", "mensagem": mensagem})
        self.config_dados["agendamentos_banner"] = agendamentos

        espaco = self.espaco_escolhido.get()
        radio = self.radio_escolhida.get()
        fundo = self.var_fundo.get()
        if fundo == "Padrão do espaço":
            fundo_path = ""
        elif fundo == "Ilha do Pescador":
            fundo_path = "__pescador__"
        elif fundo == "Mundo de Blocos":
            fundo_path = "__blocos__"
        elif fundo == "Voo no Rio":
            fundo_path = "__voo__"
        else:
            num = fundo.split()[-1]  # "Fundo 3" -> "3"
            fundo_path = f"assets/fundo{num}.png"
        if radio == "Música local (MP3)":
            self.config_dados["audio_local"] = self.var_audio_local.get().strip()

        if salvar:
            # Persiste todas as escolhas para próxima abertura
            self.config_dados["espaco_padrao"] = espaco
            self.config_dados["radio_padrao"] = radio
            self.config_dados["fundo_padrao"] = fundo
            salvar_config(self.config_dados)

        return (self.config_dados, espaco, radio, fundo_path, alerta_teste)

    def _mostrar_sobre(self):
        """Exibe a janela de créditos do projeto."""
        from tkinter import messagebox
        import tkinter.simpledialog

        janela = tk.Toplevel(self)
        janela.title("Sobre o Projeto")
        janela.configure(bg="#0D1B2A")
        janela.resizable(False, False)
        janela.grab_set()

        larg, alt = 520, 420
        x = self.winfo_x() + (self.winfo_width() - larg) // 2
        y = self.winfo_y() + (self.winfo_height() - alt) // 2
        janela.geometry(f"{larg}x{alt}+{x}+{y}")

        # Cabeçalho
        tk.Label(janela, text="CIEBP - Tela de espera",
                 font=("Segoe UI", 15, "bold"), fg="#00BCD4",
                 bg="#0D1B2A").pack(pady=(28, 4))

        tk.Label(janela, text="Centro de Inovação da Escola Básica Paulista",
                 font=("Segoe UI", 9), fg="#546E7A",
                 bg="#0D1B2A").pack()

        # Separador
        tk.Frame(janela, bg="#1E3A5F", height=1).pack(fill="x", padx=40, pady=16)

        # Texto do projeto
        texto_projeto = (
            "Este programa ajuda na recepção dos eventos do CIEBP.\n\n"
            "Ele mostra, em tela cheia, o nome do evento, a data, os professores,\n"
            "a mensagem principal e um áudio de fundo opcional.\n\n"
            "A ideia é deixar o ambiente organizado e acolhedor antes do início das atividades."
        )
        tk.Label(janela, text=texto_projeto,
                 font=("Segoe UI", 10), fg="#CFD8DC", bg="#0D1B2A",
                 justify="left", wraplength=440, padx=20).pack()

        # Separador
        tk.Frame(janela, bg="#1E3A5F", height=1).pack(fill="x", padx=40, pady=16)

        # Créditos
        tk.Label(janela, text="Desenvolvido por",
                 font=("Segoe UI", 8), fg="#546E7A", bg="#0D1B2A").pack()
        tk.Label(janela, text="Prof. Júlio César Valera",
                 font=("Segoe UI", 13, "bold"), fg="#FFFFFF", bg="#0D1B2A").pack(pady=(2, 0))
        tk.Label(janela, text="CIEBP  ·  Ribeirão Preto / SP",
                 font=("Segoe UI", 9), fg="#90A4AE", bg="#0D1B2A").pack(pady=(2, 0))
        tk.Label(janela, text="juliovalera@professor.educacao.sp.gov.br",
                 font=("Segoe UI", 9, "italic"), fg="#00BCD4", bg="#0D1B2A").pack(pady=(2, 0))

        # Botão fechar
        tk.Button(janela, text="Fechar",
                  font=("Segoe UI", 10, "bold"), fg="#FFFFFF", bg="#00BCD4",
                  activebackground="#0097A7", activeforeground="#FFFFFF",
                  relief="flat", cursor="hand2", bd=0, padx=30, pady=8,
                  command=janela.destroy).pack(pady=20)


# =============================================================================
# TELA PRINCIPAL - DESCANSO DE TELA
# =============================================================================

class DescansoTela(tk.Tk):
    """
    Janela principal em tela cheia do descanso de tela institucional.
    Exibe imagem de fundo, informações do evento e relógio em tempo real.
    """

    def __init__(self, config, espaco, radio_nome, fundo_path="", alerta_teste=None):
        super().__init__()

        self.config_dados = config
        self.espaco_atual = espaco
        self.radio_nome = radio_nome
        self.fundo_path = fundo_path  # caminho relativo ou "" para usar o do espaço
        self.alerta_teste = alerta_teste

        # Configurações de volume (0–100)
        self.volume = config.get("volume_padrao", 60)
        self.mudo = False

        # Controle de áudio
        self._player = None           # thread — backend MCI ou WMP COM
        self._proc_radio = None       # processo — backend PowerShell
        self._backend_audio = None    # "mci" | "ps" | None
        self._url_radio_atual = ""    # URL ativa
        self._parar_mci = False       # sinal de parada para o thread MCI
        self._mci_alias = "ciebp_audio"  # alias usado no MCI
        self._fila_mci = queue.Queue()   # fila de comandos para o thread MCI

        # Referências às imagens (evitar coleta de lixo pelo GC)
        self._img_fundo = None
        self._img_logo = None
        self._alerta_tags = ("alerta_horario",)
        self._efeitos_tags = "alerta_efeitos"
        self._confetes = []
        self._foguetes = []
        self._estouros = []
        self._animacao_alerta_ativa = False
        self._animacao_alerta_after = None
        self._ultimo_alerta_animado = None
        self._ultimo_confete = 0.0
        self._ultimo_foguete = 0.0
        self._contador_foguetes = 0
        self._banner_agendado_ativo = None
        self._banner_agendado_after = None
        self._fila_banners_agendados = []
        self._agendamentos_disparados = set()
        self._tags_fundo_pescador = "fundo_pescador"
        self._animacao_pescador_ativa = False
        self._animacao_pescador_after = None
        self._pescador_cena = {}
        self._tags_fundo_blocos = "fundo_blocos"
        self._animacao_blocos_ativa = False
        self._animacao_blocos_after = None
        self._blocos_cena = {}
        self._tags_fundo_voo = "fundo_voo"
        self._animacao_voo_ativa = False
        self._animacao_voo_after = None
        self._voo_cena = {}

        self._configurar_janela()
        self._construir_canvas()
        self._configurar_atalhos()
        self._atualizar_tela()
        self._loop_relogio()
        # Inicia a rádio depois que a janela estiver completamente visível
        self.after(400, self._iniciar_radio)

    # ------------------------------------------------------------------
    # CONFIGURAÇÃO DA JANELA
    # ------------------------------------------------------------------

    def _configurar_janela(self):
        """Configura a janela principal: tela cheia, cursor escondido, fundo preto."""
        self.title("CIEBP - Descanso de Tela")
        self.configure(bg="#000000")

        # Tela cheia nativa do Tkinter (Windows)
        self.state("zoomed")          # Maximiza primeiro
        self.attributes("-fullscreen", True)  # Depois força tela cheia real

        # Esconde o cursor do mouse (modo apresentação)
        self.config(cursor="none")

        # Impede que a janela seja redimensionada saindo da tela cheia
        self.resizable(False, False)

    def _configurar_atalhos(self):
        """Registra os atalhos de teclado."""
        self.bind("<Escape>", self._sair_tela_cheia)
        self.bind("<Up>", self._volume_mais)
        self.bind("<Down>", self._volume_menos)
        self.bind("<m>", self._alternar_mudo)
        self.bind("<M>", self._alternar_mudo)

    # ------------------------------------------------------------------
    # CANVAS E DESENHO DA INTERFACE
    # ------------------------------------------------------------------

    def _construir_canvas(self):
        """Cria o canvas que ocupa toda a tela."""
        self.canvas = tk.Canvas(
            self,
            bg="#000000",
            highlightthickness=0,
            bd=0
        )
        self.canvas.pack(fill="both", expand=True)

    def _atualizar_tela(self):
        """
        Redesenha toda a interface no canvas.
        Chamado uma vez na inicialização e novamente se a janela for redimensionada.
        """
        self.update_idletasks()
        larg = self.winfo_width()
        alt = self.winfo_height()

        if larg < 100 or alt < 100:
            # Dimensões ainda não estabilizaram, aguarda um frame
            self.after(50, self._atualizar_tela)
            return

        self._parar_animacao_pescador()
        self._parar_animacao_blocos()
        self._parar_animacao_voo()
        self.canvas.delete("all")
        self._desenhar_fundo(larg, alt)
        self._desenhar_painel_central(larg, alt)  # inclui relógio e logo
        self._desenhar_rodape(larg, alt)
        self._desenhar_info_radio(larg, alt)
        self._atualizar_alerta_horario(datetime.now())

    def _desenhar_fundo(self, larg, alt):
        """Carrega e desenha a imagem de fundo com overlay escuro."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")

        if self.fundo_path == "__pescador__":
            self._desenhar_fundo_pescador(larg, alt, cor_destaque)
            self.canvas.create_rectangle(0, 0, larg, 5, fill=cor_destaque, outline="")
            return
        if self.fundo_path == "__blocos__":
            self._desenhar_fundo_blocos(larg, alt, cor_destaque)
            self.canvas.create_rectangle(0, 0, larg, 5, fill=cor_destaque, outline="")
            return
        if self.fundo_path == "__voo__":
            self._desenhar_fundo_voo(larg, alt, cor_destaque)
            self.canvas.create_rectangle(0, 0, larg, 5, fill=cor_destaque, outline="")
            return

        # Usa fundo escolhido manualmente ou o fundo padrão do espaço
        caminho_img = self.fundo_path if self.fundo_path else espaco_info.get("imagem", "")

        # Caminho absoluto em relação ao diretório do script
        base_dir = BASE_DIR
        caminho_completo = base_dir / caminho_img

        img_pil = None

        if PIL_DISPONIVEL:
            if caminho_completo.exists():
                try:
                    img_pil = Image.open(caminho_completo)
                except Exception as e:
                    print(f"[AVISO] Não foi possível abrir a imagem '{caminho_completo}': {e}")
            else:
                print(f"[AVISO] Imagem não encontrada: {caminho_completo}")
                print("         Usando imagem de fundo padrão.")

            if img_pil is None:
                img_pil = criar_imagem_fallback(larg, alt, cor_destaque)

            img_pil = redimensionar_imagem(img_pil, larg, alt)
            img_pil = escurecer_imagem(img_pil, fator=0.30)

            self._img_fundo = ImageTk.PhotoImage(img_pil)
            self.canvas.create_image(0, 0, image=self._img_fundo, anchor="nw")
        else:
            # Sem Pillow: fundo preto sólido
            self.canvas.create_rectangle(0, 0, larg, alt, fill="#0D1B2A", outline="")

        # Acento colorido no topo
        self.canvas.create_rectangle(0, 0, larg, 5, fill=cor_destaque, outline="")

    def _mix_cor(self, cor1, cor2, fator):
        """Interpola duas cores hexadecimais."""
        fator = max(0.0, min(1.0, fator))
        c1 = [int(cor1[i:i+2], 16) for i in (1, 3, 5)]
        c2 = [int(cor2[i:i+2], 16) for i in (1, 3, 5)]
        comp = [int(a + (b - a) * fator) for a, b in zip(c1, c2)]
        return "#{:02X}{:02X}{:02X}".format(*comp)

    def _obter_intervalo_tubarao_segundos(self):
        """Retorna o intervalo configurado do tubarão em segundos."""
        try:
            minutos = float(self.config_dados.get("intervalo_tubarao_minutos", 10))
        except (TypeError, ValueError):
            minutos = 10
        minutos = max(0.0, minutos)
        if minutos == 0:
            return None
        return minutos * 60

    def _aproximar_valor(self, atual, alvo, passo):
        """Aproxima um valor do alvo sem ultrapassar."""
        if atual < alvo:
            return min(alvo, atual + passo)
        if atual > alvo:
            return max(alvo, atual - passo)
        return atual

    def _obter_estado_solar(self, agora, larg, alt):
        """Calcula céu, mar e posição do sol conforme o horário."""
        minutos = agora.hour * 60 + agora.minute + agora.second / 60
        nascer = 5 * 60 + 30
        por = 18 * 60 + 30
        horizonte = int(alt * 0.6)

        if minutos <= nascer:
            return {
                "ceu_topo": "#051126",
                "ceu_base": "#102545",
                "mar": "#0C2745",
                "sol_visivel": False,
                "sol_cor": "#FFD86B",
                "sol_x": int(larg * 0.1),
                "sol_y": horizonte + 20,
            }

        if minutos >= por:
            return {
                "ceu_topo": "#061225",
                "ceu_base": "#13233E",
                "mar": "#0B223B",
                "sol_visivel": False,
                "sol_cor": "#FFD86B",
                "sol_x": int(larg * 0.9),
                "sol_y": horizonte + 20,
            }

        progresso = (minutos - nascer) / (por - nascer)
        x = int(larg * (0.12 + progresso * 0.76))
        curva = 1 - (2 * progresso - 1) ** 2
        y = int(horizonte - 35 - curva * (alt * 0.34))

        if progresso < 0.18:
            fator = progresso / 0.18
            topo = self._mix_cor("#18345A", "#64A3FF", fator)
            base = self._mix_cor("#F48C6A", "#BFE8FF", fator)
            mar = self._mix_cor("#17405C", "#2474B5", fator)
            sol = self._mix_cor("#FFB35A", "#FFE08A", fator)
        elif progresso < 0.75:
            fator = (progresso - 0.18) / 0.57
            topo = self._mix_cor("#64A3FF", "#8AC6FF", fator)
            base = self._mix_cor("#BFE8FF", "#DFF6FF", fator)
            mar = self._mix_cor("#2474B5", "#1E87C8", fator)
            sol = self._mix_cor("#FFE08A", "#FFF0A6", fator)
        else:
            fator = (progresso - 0.75) / 0.25
            topo = self._mix_cor("#8AC6FF", "#20345F", fator)
            base = self._mix_cor("#DFF6FF", "#F8906E", fator)
            mar = self._mix_cor("#1E87C8", "#1B496D", fator)
            sol = self._mix_cor("#FFF0A6", "#FF9B58", fator)

        return {
            "ceu_topo": topo,
            "ceu_base": base,
            "mar": mar,
            "sol_visivel": True,
            "sol_cor": sol,
            "sol_x": x,
            "sol_y": y,
        }

    def _desenhar_fundo_pescador(self, larg, alt, cor_destaque):
        """Desenha o fundo animado da ilha do pescador."""
        horizonte = int(alt * 0.6)
        estado = self._obter_estado_solar(datetime.now(), larg, alt)
        tags = self._tags_fundo_pescador

        bandas = []
        for indice in range(8):
            y1 = int((horizonte / 8) * indice)
            y2 = int((horizonte / 8) * (indice + 1))
            ret = self.canvas.create_rectangle(0, y1, larg, y2, fill=estado["ceu_topo"], outline="", tags=tags)
            bandas.append(ret)

        brilho = self.canvas.create_oval(
            int(larg * 0.32), int(horizonte * 0.42),
            int(larg * 0.7), int(horizonte + 120),
            fill="#FFFFFF", outline="", stipple="gray50", tags=tags
        )
        mar = self.canvas.create_rectangle(0, horizonte, larg, alt, fill=estado["mar"], outline="", tags=tags)
        reflexo = self.canvas.create_polygon(
            int(larg * 0.48), horizonte,
            int(larg * 0.54), horizonte,
            int(larg * 0.66), alt,
            int(larg * 0.4), alt,
            fill="#BFE8FF", outline="", stipple="gray50", tags=tags
        )

        sol = self.canvas.create_oval(0, 0, 0, 0, fill=estado["sol_cor"], outline="", tags=tags)
        sol_aura = self.canvas.create_oval(0, 0, 0, 0, fill="#FFF5B8", outline="", stipple="gray50", tags=tags)

        nuvens = []
        for base_x, base_y, escala in ((0.18, 0.18, 1.0), (0.74, 0.14, 0.8)):
            grupo = []
            for dx, dy, raio in ((0, 0, 34), (28, -10, 28), (56, 0, 32)):
                grupo.append(self.canvas.create_oval(0, 0, 0, 0, fill="#F5FBFF", outline="", stipple="gray50", tags=tags))
            nuvens.append({"itens": grupo, "base_x": base_x, "base_y": base_y, "escala": escala})

        ilha = self.canvas.create_oval(
            int(larg * 0.08), int(alt * 0.72),
            int(larg * 0.36), int(alt * 0.86),
            fill="#6E5B37", outline="#A48B57", width=2, tags=tags
        )
        areia = self.canvas.create_oval(
            int(larg * 0.12), int(alt * 0.76),
            int(larg * 0.3), int(alt * 0.84),
            fill="#CFAE72", outline="", tags=tags
        )
        coqueiro_offset_x = -20
        coqueiro_topo_x = int(larg * 0.24) + coqueiro_offset_x
        coqueiro_topo_y = int(alt * 0.56)
        coqueiro_base_x = int(larg * 0.21) + coqueiro_offset_x
        coqueiro_base_y = int(alt * 0.73)
        coqueiro_tronco = self.canvas.create_line(
            coqueiro_base_x, coqueiro_base_y,
            coqueiro_topo_x, coqueiro_topo_y,
            fill="#70411E", width=11, smooth=True, tags=tags
        )
        marcas_tronco = []
        for proporcao in (0.18, 0.34, 0.50, 0.66):
            marca_x = coqueiro_base_x + (coqueiro_topo_x - coqueiro_base_x) * proporcao
            marca_y = coqueiro_base_y + (coqueiro_topo_y - coqueiro_base_y) * proporcao
            marcas_tronco.append(self.canvas.create_line(
                marca_x - 5,
                marca_y + 1,
                marca_x + 6,
                marca_y - 2,
                fill="#C08A52", width=2, tags=tags
            ))
        folhas = []
        for pontos, cor in [
            ((0, 0, -24, -8, -57, -29, -67, -43, -34, -41, -8, -22), "#238A50"),
            ((0, 0, -12, -22, -33, -54, -41, -73, -12, -57, 4, -27), "#2FA85F"),
            ((0, 0, 5, -23, 12, -57, 25, -75, 30, -47, 18, -17), "#3ABA69"),
            ((0, 0, 23, -13, 53, -30, 75, -30, 51, -12, 15, -3), "#269452"),
            ((0, 0, 20, 2, 52, 7, 68, 20, 34, 18, 8, 8), "#1E7D46"),
            ((0, 0, -15, 3, -47, 16, -61, 31, -28, 24, -4, 9), "#207B45"),
            ((0, 0, -2, -14, 1, -43, 10, -59, 17, -31, 10, -8), "#52C975"),
        ]:
            coords = []
            for indice in range(0, len(pontos), 2):
                coords.extend((
                    coqueiro_topo_x + pontos[indice],
                    coqueiro_topo_y + pontos[indice + 1],
                ))
            folhas.append(self.canvas.create_polygon(
                *coords,
                fill=cor, outline="#17643A", width=1, smooth=True, splinesteps=24, tags=tags
            ))
        cocos = [
            self.canvas.create_oval(0, 0, 0, 0, fill="#9C612C", outline="#E5AD61", width=1, tags=tags)
            for _ in range(3)
        ]
        for coco, dx, dy in zip(cocos, (-11, 3, 13), (7, 9, 5)):
            self.canvas.coords(coco, coqueiro_topo_x + dx - 6, coqueiro_topo_y + dy - 6, coqueiro_topo_x + dx + 6, coqueiro_topo_y + dy + 6)

        trapiche = self.canvas.create_polygon(
            int(larg * 0.28), int(alt * 0.76),
            int(larg * 0.43), int(alt * 0.68),
            int(larg * 0.45), int(alt * 0.7),
            int(larg * 0.3), int(alt * 0.79),
            fill="#7C5632", outline="#A67C52", width=2, tags=tags
        )

        agua_ripples = [
            self.canvas.create_arc(0, 0, 0, 0, start=10, extent=160, style="arc", outline="#D7F6FF", width=2, tags=tags),
            self.canvas.create_arc(0, 0, 0, 0, start=10, extent=160, style="arc", outline="#7FDBFF", width=2, tags=tags),
            self.canvas.create_arc(0, 0, 0, 0, start=10, extent=160, style="arc", outline="#7FDBFF", width=2, tags=tags),
        ]

        pescador = {
            "cabeca": self.canvas.create_oval(0, 0, 0, 0, fill="#F2C79B", outline="", tags=tags),
            "chapeu": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#B5651D", outline="", tags=tags),
            "chapeu_faixa": self.canvas.create_line(0, 0, 0, 0, fill="#F2D38B", width=3, tags=tags),
            "orelha": self.canvas.create_oval(0, 0, 0, 0, fill="#E9B987", outline="", tags=tags),
            "nariz": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#E9B987", outline="#C88E62", width=1, tags=tags),
            "olho": self.canvas.create_oval(0, 0, 0, 0, fill="#1B2838", outline="", tags=tags),
            "barba": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#4D372A", outline="", tags=tags),
            "tronco": self.canvas.create_line(0, 0, 0, 0, fill="#F6F1E9", width=8, tags=tags),
            "camisa": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#E7F4F0", outline="#77AFAA", width=1, tags=tags),
            "colete": self.canvas.create_line(0, 0, 0, 0, fill="#2E7E79", width=3, tags=tags),
            "bermuda": self.canvas.create_line(0, 0, 0, 0, fill="#395B72", width=7, tags=tags),
            "perna1": self.canvas.create_line(0, 0, 0, 0, fill="#213547", width=6, tags=tags),
            "perna2": self.canvas.create_line(0, 0, 0, 0, fill="#213547", width=6, tags=tags),
            "sapato1": self.canvas.create_line(0, 0, 0, 0, fill="#352E29", width=4, tags=tags),
            "sapato2": self.canvas.create_line(0, 0, 0, 0, fill="#352E29", width=4, tags=tags),
            "braco": self.canvas.create_line(0, 0, 0, 0, fill="#F2C79B", width=5, tags=tags),
            "vara": self.canvas.create_line(0, 0, 0, 0, fill="#3A2A1A", width=3, smooth=True, tags=tags),
            "linha": self.canvas.create_line(0, 0, 0, 0, fill="#E7F7FF", width=2, tags=tags),
            "banco": self.canvas.create_rectangle(0, 0, 0, 0, fill="#8B5A2B", outline="", tags=tags),
            "captura": self.canvas.create_text(0, 0, text="", font=("Segoe UI", 18, "bold"), fill="#FFE08A", tags=tags),
            "captura_peixe_corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#7DFF9B", outline="#DFFFEA", width=2, state="hidden", tags=tags),
            "captura_peixe_cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#55D68C", outline="#DFFFEA", width=2, state="hidden", tags=tags),
            "captura_peixe_olho": self.canvas.create_oval(0, 0, 0, 0, fill="#163040", outline="", state="hidden", tags=tags),
            "captura_bota_cano": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#FFD28A", outline="#FFF1CC", width=2, state="hidden", tags=tags),
            "captura_bota_sola": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#C48A53", outline="#FFF1CC", width=2, state="hidden", tags=tags),
            "captura_bigorna_base": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#8D9AA6", outline="#D8E0E6", width=2, state="hidden", tags=tags),
            "captura_bigorna_topo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#A7B3BD", outline="#D8E0E6", width=2, state="hidden", tags=tags),
            "captura_pneu_externo": self.canvas.create_oval(0, 0, 0, 0, fill="#2E3640", outline="#717B86", width=2, state="hidden", tags=tags),
            "captura_pneu_interno": self.canvas.create_oval(0, 0, 0, 0, fill="#0D1B2A", outline="", state="hidden", tags=tags),
            "captura_pneu_aro": self.canvas.create_oval(0, 0, 0, 0, fill="#72818A", outline="#C7D2D8", width=1, state="hidden", tags=tags),
            "captura_pneu_brilho": self.canvas.create_arc(0, 0, 0, 0, start=115, extent=125, style="arc", outline="#D5E1E7", width=2, state="hidden", tags=tags),
        }
        trofeu_ilha = {
            "peixe_corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#7DFF9B", outline="#DFFFEA", width=2, state="hidden", tags=tags),
            "peixe_cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#55D68C", outline="#DFFFEA", width=2, state="hidden", tags=tags),
            "peixe_olho": self.canvas.create_oval(0, 0, 0, 0, fill="#163040", outline="", state="hidden", tags=tags),
            "bota_cano": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#FFD28A", outline="#FFF1CC", width=2, state="hidden", tags=tags),
            "bota_sola": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#C48A53", outline="#FFF1CC", width=2, state="hidden", tags=tags),
            "bigorna_base": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#8D9AA6", outline="#D8E0E6", width=2, state="hidden", tags=tags),
            "bigorna_topo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#A7B3BD", outline="#D8E0E6", width=2, state="hidden", tags=tags),
            "pneu_externo": self.canvas.create_oval(0, 0, 0, 0, fill="#2E3640", outline="#717B86", width=2, state="hidden", tags=tags),
            "pneu_interno": self.canvas.create_oval(0, 0, 0, 0, fill="#0D1B2A", outline="", state="hidden", tags=tags),
            "pneu_aro": self.canvas.create_oval(0, 0, 0, 0, fill="#72818A", outline="#C7D2D8", width=1, state="hidden", tags=tags),
            "pneu_brilho": self.canvas.create_arc(0, 0, 0, 0, start=115, extent=125, style="arc", outline="#D5E1E7", width=2, state="hidden", tags=tags),
        }

        ondas = []
        for frac, amp in ((0.52, 0.16), (0.67, 0.13), (0.8, 0.11)):
            ondas.append(self.canvas.create_line(0, 0, 0, 0, fill="#89D4FF", width=2, smooth=True, tags=tags))

        peixinhos_salto = []
        for _ in range(3):
            peixinhos_salto.append({
                "corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#7DFF9B", outline="#E5FFEF", width=1, state="hidden", tags=tags),
                "cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#45C985", outline="#E5FFEF", width=1, state="hidden", tags=tags),
                "olho": self.canvas.create_oval(0, 0, 0, 0, fill="#163040", outline="", state="hidden", tags=tags),
            })
        golfinhos_salto = []
        for _ in range(2):
            golfinhos_salto.append({
                "corpo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, fill="#5F93A8", outline="#B7E1EC", width=1, smooth=True, splinesteps=16, state="hidden", tags=tags),
                "focinho": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#6EA1B4", outline="#B7E1EC", width=1, state="hidden", tags=tags),
                "barbatana": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#477B92", outline="#B7E1EC", width=1, state="hidden", tags=tags),
                "nadadeira_lateral": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#4D849A", outline="#B7E1EC", width=1, state="hidden", tags=tags),
                "cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#4D849A", outline="#B7E1EC", width=1, state="hidden", tags=tags),
                "barriga": self.canvas.create_arc(0, 0, 0, 0, start=190, extent=160, style="arc", outline="#D9F1F5", width=2, state="hidden", tags=tags),
                "olho": self.canvas.create_oval(0, 0, 0, 0, fill="#102934", outline="", state="hidden", tags=tags),
            })

        tubarao = {
            "corpo": self.canvas.create_polygon(
                0, 0, 0, 0, 0, 0,
                fill="#163A50", outline="#2A5970", width=2,
                smooth=True, splinesteps=24, state="hidden", tags=tags
            ),
            "barbatana": self.canvas.create_polygon(
                0, 0, 0, 0, 0, 0,
                fill="#1E4C65", outline="#2A5970", width=2,
                smooth=True, state="hidden", tags=tags
            ),
            "cauda": self.canvas.create_polygon(
                0, 0, 0, 0, 0, 0, 0, 0,
                fill="#173E56", outline="#2A5970", width=2,
                smooth=True, state="hidden", tags=tags
            ),
            "barriga": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#B5D6DE", outline="", state="hidden", tags=tags),
            "olho": self.canvas.create_oval(0, 0, 0, 0, fill="#F3F7E9", outline="#163040", width=1, state="hidden", tags=tags),
            "pupila": self.canvas.create_oval(0, 0, 0, 0, fill="#101820", outline="", state="hidden", tags=tags),
            "boca": self.canvas.create_line(0, 0, 0, 0, fill="#15232A", width=2, smooth=True, state="hidden", tags=tags),
            "dentes": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, fill="#FFFFFF", outline="#C9E8EE", width=1, state="hidden", tags=tags),
        }
        gaivota = {
            "sombra": self.canvas.create_oval(0, 0, 0, 0, fill="#93A0AF", outline="", state="hidden", tags=tags),
            "asa_esq": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#E8EDF2", outline="#9AA8B6", width=1, smooth=True, splinesteps=16, state="hidden", tags=tags),
            "asa_dir": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#C8D1DB", outline="#8B9AA9", width=1, smooth=True, splinesteps=16, state="hidden", tags=tags),
            "cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, 0, 0, fill="#F4F6F8", outline="#AAB5BF", width=1, state="hidden", tags=tags),
            "corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#F9FBFC", outline="#AAB5BF", width=1, state="hidden", tags=tags),
            "peito": self.canvas.create_oval(0, 0, 0, 0, fill="#FFFFFF", outline="", state="hidden", tags=tags),
            "cabeca": self.canvas.create_oval(0, 0, 0, 0, fill="#FFFFFF", outline="#AAB5BF", width=1, state="hidden", tags=tags),
            "bico": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F2B544", outline="#D98024", width=1, state="hidden", tags=tags),
            "olho": self.canvas.create_oval(0, 0, 0, 0, fill="#163040", outline="", state="hidden", tags=tags),
            "carga_peixe_corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#7DFF9B", outline="#DFFFEA", width=1, state="hidden", tags=tags),
            "carga_peixe_cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#55D68C", outline="#DFFFEA", width=1, state="hidden", tags=tags),
            "carga_peixe_olho": self.canvas.create_oval(0, 0, 0, 0, fill="#163040", outline="", state="hidden", tags=tags),
        }
        caranguejo = {
            "corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#E85A47", outline="#FFC2A6", width=2, state="hidden", tags=tags),
            "garra_esq": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F26B56", outline="#FFC2A6", width=2, state="hidden", tags=tags),
            "garra_dir": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F26B56", outline="#FFC2A6", width=2, state="hidden", tags=tags),
            "perna1": self.canvas.create_line(0, 0, 0, 0, 0, 0, fill="#C94736", width=2, smooth=True, state="hidden", tags=tags),
            "perna2": self.canvas.create_line(0, 0, 0, 0, 0, 0, fill="#C94736", width=2, smooth=True, state="hidden", tags=tags),
            "perna3": self.canvas.create_line(0, 0, 0, 0, 0, 0, fill="#C94736", width=2, smooth=True, state="hidden", tags=tags),
            "perna4": self.canvas.create_line(0, 0, 0, 0, 0, 0, fill="#C94736", width=2, smooth=True, state="hidden", tags=tags),
            "olho_esq": self.canvas.create_line(0, 0, 0, 0, fill="#FFD7C7", width=2, state="hidden", tags=tags),
            "olho_dir": self.canvas.create_line(0, 0, 0, 0, fill="#FFD7C7", width=2, state="hidden", tags=tags),
            "pupila_esq": self.canvas.create_oval(0, 0, 0, 0, fill="#1E232A", outline="", state="hidden", tags=tags),
            "pupila_dir": self.canvas.create_oval(0, 0, 0, 0, fill="#1E232A", outline="", state="hidden", tags=tags),
        }
        acoes = {
            "coco_caindo": self.canvas.create_oval(0, 0, 0, 0, fill="#9C612C", outline="#E5AD61", width=2, state="hidden", tags=tags),
            "coco_aberto": self.canvas.create_oval(0, 0, 0, 0, fill="#8B5428", outline="#F5DFB4", width=2, state="hidden", tags=tags),
            "coco_polpa": self.canvas.create_arc(0, 0, 0, 0, start=180, extent=180, style="pieslice", fill="#FFF1D2", outline="#F5DFB4", state="hidden", tags=tags),
            "aviao_corpo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#EAF5FA", outline="#8197A5", width=1, state="hidden", tags=tags),
            "aviao_asa": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#C5D7E2", outline="#8197A5", width=1, state="hidden", tags=tags),
            "aviao_cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#C5D7E2", outline="#8197A5", width=1, state="hidden", tags=tags),
            "binoculo_esq": self.canvas.create_oval(0, 0, 0, 0, fill="#263D4A", outline="#9EC2D0", width=2, state="hidden", tags=tags),
            "binoculo_dir": self.canvas.create_oval(0, 0, 0, 0, fill="#263D4A", outline="#9EC2D0", width=2, state="hidden", tags=tags),
            "binoculo_ponte": self.canvas.create_line(0, 0, 0, 0, fill="#9EC2D0", width=3, state="hidden", tags=tags),
        }

        intervalo_tubarao = self._obter_intervalo_tubarao_segundos()

        self._pescador_cena = {
            "larg": larg,
            "alt": alt,
            "horizonte": horizonte,
            "bands": bandas,
            "brilho": brilho,
            "mar": mar,
            "reflexo": reflexo,
            "sol": sol,
            "sol_aura": sol_aura,
            "nuvens": nuvens,
            "ilha": ilha,
            "areia": areia,
            "coqueiro_tronco": coqueiro_tronco,
            "marcas_tronco": marcas_tronco,
            "folhas": folhas,
            "cocos": cocos,
            "coqueiro_deslocamento": 0.0,
            "trapiche": trapiche,
            "ondas": ondas,
            "peixinhos_salto": peixinhos_salto,
            "salto_peixes_ativo": False,
            "salto_peixes_inicio": 0.0,
            "salto_peixes_duracao": 0.0,
            "salto_peixes_x": float(larg * 0.58),
            "salto_peixes_sentido": 1,
            "proximo_salto_peixes": time.time() + random.uniform(20, 40),
            "golfinhos_salto": golfinhos_salto,
            "salto_golfinhos_ativo": False,
            "salto_golfinhos_inicio": 0.0,
            "salto_golfinhos_duracao": 0.0,
            "salto_golfinhos_x": float(larg * 0.68),
            "salto_golfinhos_sentido": 1,
            "proximo_salto_golfinhos": time.time() + random.uniform(55, 100),
            "ripples": agua_ripples,
            "pescador": pescador,
            "trofeu_ilha": trofeu_ilha,
            "tubarao": tubarao,
            "gaivota": gaivota,
            "caranguejo": caranguejo,
            "acoes": acoes,
            "base_x": int(larg * 0.4),
            "base_y": int(alt * 0.69),
            # O pescador para ao lado da base do tronco antes de balancar o coqueiro.
            "coqueiro_x": int(larg * 0.21) - 20,
            "coqueiro_y": int(alt * 0.73),
            "pescador_x": float(int(larg * 0.4)),
            "pescador_y": float(int(alt * 0.69)),
            "refugio_x": float(int(larg * 0.24)),
            "refugio_y": float(int(alt * 0.73)),
            "pescador_modo": "pescando",
            "estado": "idle",
            "evento_inicio": time.time(),
            "evento_ate": time.time() + random.uniform(14, 22),
            "proximo_evento": time.time() + random.uniform(18, 30),
            "acao_tipo": "",
            "captura_tipo": "",
            "trofeu_tipo": "",
            "trofeu_instante": 0.0,
            "trofeu_x": int(larg * 0.205),
            "trofeu_y": int(alt * 0.79),
            "gaivota_ativa": False,
            "gaivota_inicio": 0.0,
            "gaivota_duracao": 0.0,
            "gaivota_pouso": 0.0,
            "tubarao_ativo": False,
            "tubarao_inicio": 0.0,
            "tubarao_duracao": 0.0,
            "tubarao_sentido": 1,
            "tubarao_base_y": alt * 0.8,
            "tubarao_escala": 1.0,
            "tubarao_fase": 0.0,
            "tubarao_x": None,
            "tubarao_y": None,
            "proximo_tubarao": (time.time() + intervalo_tubarao) if intervalo_tubarao else float("inf"),
            "caranguejo_ativo": False,
            "caranguejo_inicio": 0.0,
            "caranguejo_duracao": 0.0,
            "caranguejo_x": None,
            "caranguejo_y": None,
            "proximo_caranguejo": (time.time() + (intervalo_tubarao * 0.5)) if intervalo_tubarao else float("inf"),
        }

        self._atualizar_cena_pescador(datetime.now(), time.time())
        self._animacao_pescador_ativa = True
        self._animacao_pescador_after = self.after(120, self._loop_animacao_pescador)

    def _ocultar_forma_captura(self, grupo):
        """Esconde as formas visuais de peixe e bota."""
        for chave in (
            "peixe_corpo", "peixe_cauda", "peixe_olho",
            "bota_cano", "bota_sola",
            "bigorna_base", "bigorna_topo",
            "pneu_externo", "pneu_interno", "pneu_aro", "pneu_brilho",
            "carga_peixe_corpo", "carga_peixe_cauda", "carga_peixe_olho",
        ):
            if chave in grupo:
                self.canvas.itemconfig(grupo[chave], state="hidden")

    def _desenhar_forma_captura(self, grupo, tipo, x, y, escala=1.0):
        """Desenha peixe ou objetos engraçados em uma posição específica."""
        self._ocultar_forma_captura(grupo)

        if tipo == "bota":
            # Bota de pescador com cano inclinado, bico arredondado e salto.
            self.canvas.coords(
                grupo["bota_cano"],
                x - 15 * escala, y - 32 * escala,
                x + 7 * escala, y - 29 * escala,
                x + 5 * escala, y - 5 * escala,
                x + 18 * escala, y + 1 * escala,
                x + 12 * escala, y + 9 * escala,
                x - 18 * escala, y + 7 * escala,
                x - 15 * escala, y - 2 * escala,
            )
            self.canvas.coords(
                grupo["bota_sola"],
                x - 20 * escala, y + 6 * escala,
                x + 15 * escala, y + 8 * escala,
                x + 21 * escala, y + 13 * escala,
                x + 13 * escala, y + 17 * escala,
                x - 21 * escala, y + 14 * escala,
            )
            self.canvas.itemconfig(grupo["bota_cano"], fill="#D99A4B", outline="#FFF0C9", width=2)
            self.canvas.itemconfig(grupo["bota_sola"], fill="#704321", outline="#D8A36B", width=2)
            self.canvas.itemconfig(grupo["bota_cano"], state="normal")
            self.canvas.itemconfig(grupo["bota_sola"], state="normal")
            return

        if tipo == "bigorna":
            # Silhueta classica: chifre longo, mesa, cintura estreita e base pesada.
            self.canvas.coords(
                grupo["bigorna_topo"],
                x - 28 * escala, y - 18 * escala,
                x + 5 * escala, y - 18 * escala,
                x + 15 * escala, y - 13 * escala,
                x + 31 * escala, y - 7 * escala,
                x + 12 * escala, y - 5 * escala,
                x + 7 * escala, y + 1 * escala,
                x - 18 * escala, y + 1 * escala,
                x - 12 * escala, y - 8 * escala,
            )
            self.canvas.coords(
                grupo["bigorna_base"],
                x - 17 * escala, y,
                x + 9 * escala, y,
                x + 13 * escala, y + 10 * escala,
                x + 24 * escala, y + 15 * escala,
                x + 19 * escala, y + 21 * escala,
                x - 28 * escala, y + 21 * escala,
                x - 31 * escala, y + 15 * escala,
                x - 20 * escala, y + 10 * escala,
            )
            self.canvas.itemconfig(grupo["bigorna_topo"], fill="#B7C2CB", outline="#E7EEF2", width=2)
            self.canvas.itemconfig(grupo["bigorna_base"], fill="#61717C", outline="#B7C2CB", width=2)
            self.canvas.itemconfig(grupo["bigorna_topo"], state="normal")
            self.canvas.itemconfig(grupo["bigorna_base"], state="normal")
            return

        if tipo == "pneu":
            self.canvas.coords(
                grupo["pneu_externo"],
                x - 22 * escala, y - 21 * escala,
                x + 22 * escala, y + 21 * escala,
            )
            self.canvas.coords(
                grupo["pneu_interno"],
                x - 13 * escala, y - 13 * escala,
                x + 13 * escala, y + 13 * escala,
            )
            self.canvas.coords(
                grupo["pneu_aro"],
                x - 7 * escala, y - 7 * escala,
                x + 7 * escala, y + 7 * escala,
            )
            self.canvas.coords(
                grupo["pneu_brilho"],
                x - 19 * escala, y - 18 * escala,
                x + 19 * escala, y + 18 * escala,
            )
            self.canvas.itemconfig(grupo["pneu_externo"], fill="#252C32", outline="#A9B6BD", width=2)
            self.canvas.itemconfig(grupo["pneu_interno"], fill="#10171C")
            self.canvas.itemconfig(grupo["pneu_aro"], fill="#6F818A", outline="#D5E1E7", width=1)
            self.canvas.itemconfig(grupo["pneu_externo"], state="normal")
            self.canvas.itemconfig(grupo["pneu_interno"], state="normal")
            self.canvas.itemconfig(grupo["pneu_aro"], state="normal")
            self.canvas.itemconfig(grupo["pneu_brilho"], state="normal")
            return

        if tipo == "peixinho":
            meio = 12 * escala
            altura = 8 * escala
            corpo_cor = "#9BE7FF"
            cauda_cor = "#69D8F6"
            outline = "#EAFBFF"
            olho = 4 * escala
            olho_x = 6 * escala
            cauda_dx = 14 * escala
            cauda_dy = 10 * escala
        else:
            meio = 18 * escala
            altura = 12 * escala
            corpo_cor = "#7DFF9B"
            cauda_cor = "#55D68C"
            outline = "#E5FFEF"
            olho = 5 * escala
            olho_x = 8 * escala
            cauda_dx = 18 * escala
            cauda_dy = 14 * escala

        self.canvas.coords(
            grupo["peixe_corpo"],
            x - meio, y - altura,
            x + meio, y + altura,
        )
        self.canvas.coords(
            grupo["peixe_cauda"],
            x + meio - 2 * escala, y,
            x + meio + cauda_dx, y - cauda_dy,
            x + meio + cauda_dx, y + cauda_dy,
        )
        self.canvas.coords(
            grupo["peixe_olho"],
            x - olho_x, y - olho / 2,
            x - olho_x + olho, y + olho / 2,
        )
        self.canvas.itemconfig(grupo["peixe_corpo"], fill=corpo_cor, outline=outline, state="normal")
        self.canvas.itemconfig(grupo["peixe_cauda"], fill=cauda_cor, outline=outline, state="normal")
        self.canvas.itemconfig(grupo["peixe_olho"], state="normal")

    def _atualizar_cena_pescador(self, agora, agora_ts):
        """Atualiza o fundo animado da ilha do pescador."""
        cena = self._pescador_cena
        if not cena:
            return

        larg = cena["larg"]
        alt = cena["alt"]
        horizonte = cena["horizonte"]
        estado_solar = self._obter_estado_solar(agora, larg, alt)
        cena["estado_solar"] = estado_solar

        for indice, item in enumerate(cena["bands"]):
            fator = indice / max(1, len(cena["bands"]) - 1)
            cor = self._mix_cor(estado_solar["ceu_topo"], estado_solar["ceu_base"], fator)
            self.canvas.itemconfig(item, fill=cor)

        self.canvas.itemconfig(cena["mar"], fill=estado_solar["mar"])
        brilho_cor = self._mix_cor(estado_solar["ceu_base"], "#FFFFFF", 0.38)
        reflexo_cor = self._mix_cor(estado_solar["sol_cor"], "#DFF6FF", 0.55)
        self.canvas.itemconfig(cena["brilho"], fill=brilho_cor)
        self.canvas.itemconfig(cena["reflexo"], fill=reflexo_cor)

        if estado_solar["sol_visivel"]:
            sx = estado_solar["sol_x"]
            sy = estado_solar["sol_y"]
            brilho_larg = 170
            brilho_alt = 130
            reflexo_topo_y = horizonte
            reflexo_base_y = alt
            reflexo_topo_larg = 28
            reflexo_base_larg = 130

            self.canvas.coords(
                cena["brilho"],
                sx - brilho_larg, sy - brilho_alt,
                sx + brilho_larg, sy + brilho_alt,
            )
            self.canvas.coords(
                cena["reflexo"],
                sx - reflexo_topo_larg, reflexo_topo_y,
                sx + reflexo_topo_larg, reflexo_topo_y,
                sx + reflexo_base_larg, reflexo_base_y,
                sx - reflexo_base_larg, reflexo_base_y,
            )
            self.canvas.coords(cena["sol"], sx - 38, sy - 38, sx + 38, sy + 38)
            self.canvas.coords(cena["sol_aura"], sx - 82, sy - 82, sx + 82, sy + 82)
            self.canvas.itemconfig(cena["brilho"], state="normal")
            self.canvas.itemconfig(cena["reflexo"], state="normal")
            self.canvas.itemconfig(cena["sol"], fill=estado_solar["sol_cor"], state="normal")
            self.canvas.itemconfig(
                cena["sol_aura"],
                fill=self._mix_cor(estado_solar["sol_cor"], "#FFFDE7", 0.58),
                state="normal",
            )
        else:
            self.canvas.itemconfig(cena["brilho"], state="hidden")
            self.canvas.itemconfig(cena["reflexo"], state="hidden")
            self.canvas.itemconfig(cena["sol"], state="hidden")
            self.canvas.itemconfig(cena["sol_aura"], state="hidden")

        # Deriva lenta, com ritmos ligeiramente diferentes para nao parecer repetitiva.
        fase_nuvem = agora_ts * 0.045
        for indice, nuvem in enumerate(cena["nuvens"]):
            desloc_x = (
                math.sin(fase_nuvem + indice * 1.7) * 38
                + math.sin(fase_nuvem * 0.37 + indice * 2.4) * 10
            )
            desloc_y = math.cos(fase_nuvem * 0.72 + indice * 1.3) * 7
            base_x = larg * nuvem["base_x"] + desloc_x
            base_y = horizonte * nuvem["base_y"] + desloc_y
            escala = nuvem["escala"]
            specs = ((0, 0, 34), (28, -10, 28), (56, 0, 32))
            for item, (dx, dy, raio) in zip(nuvem["itens"], specs):
                r = raio * escala
                x = base_x + dx * escala
                y = base_y + dy * escala
                self.canvas.coords(item, x - r, y - r, x + r, y + r)

        fase_onda = agora_ts * 1.2
        for indice, item in enumerate(cena["ondas"]):
            y_base = alt * (0.78 + indice * 0.06)
            amplitude = 7 + indice * 2
            pontos = []
            for passo in range(11):
                x = (larg / 10) * passo
                y = y_base + math.sin(fase_onda + passo * 0.85 + indice) * amplitude
                pontos.extend((x, y))
            self.canvas.coords(item, *pontos)

        if (not cena["salto_peixes_ativo"]) and agora_ts >= cena["proximo_salto_peixes"]:
            cena["salto_peixes_ativo"] = True
            cena["salto_peixes_inicio"] = agora_ts
            cena["salto_peixes_duracao"] = random.uniform(3.0, 4.2)
            cena["salto_peixes_x"] = random.uniform(larg * 0.46, larg * 0.72)
            cena["salto_peixes_sentido"] = random.choice((-1, 1))

        if cena["salto_peixes_ativo"]:
            progresso_salto = (agora_ts - cena["salto_peixes_inicio"]) / max(0.1, cena["salto_peixes_duracao"])
            if progresso_salto >= 1.0:
                cena["salto_peixes_ativo"] = False
                cena["proximo_salto_peixes"] = agora_ts + random.uniform(24, 52)
                for peixinho in cena["peixinhos_salto"]:
                    for item in peixinho.values():
                        self.canvas.itemconfig(item, state="hidden")
            else:
                for indice, peixinho in enumerate(cena["peixinhos_salto"]):
                    progresso_individual = progresso_salto * 1.35 - indice * 0.20
                    if not 0.0 <= progresso_individual <= 1.0:
                        for item in peixinho.values():
                            self.canvas.itemconfig(item, state="hidden")
                        continue
                    sentido = cena["salto_peixes_sentido"]
                    peixe_x = cena["salto_peixes_x"] + sentido * ((progresso_individual - 0.5) * 150 + indice * 9)
                    peixe_y = alt * 0.81 - math.sin(progresso_individual * math.pi) * (52 + indice * 7)
                    tamanho = 9 + indice * 1.5
                    self.canvas.coords(
                        peixinho["corpo"],
                        peixe_x - tamanho, peixe_y - tamanho * 0.55,
                        peixe_x + tamanho, peixe_y + tamanho * 0.55,
                    )
                    self.canvas.coords(
                        peixinho["cauda"],
                        peixe_x - sentido * tamanho * 0.7, peixe_y,
                        peixe_x - sentido * tamanho * 1.8, peixe_y - tamanho,
                        peixe_x - sentido * tamanho * 1.8, peixe_y + tamanho,
                    )
                    olho_x = peixe_x + sentido * tamanho * 0.35
                    self.canvas.coords(peixinho["olho"], olho_x - 2, peixe_y - 2, olho_x + 2, peixe_y + 2)
                    for item in peixinho.values():
                        self.canvas.itemconfig(item, state="normal")
        else:
            for peixinho in cena["peixinhos_salto"]:
                for item in peixinho.values():
                    self.canvas.itemconfig(item, state="hidden")

        if (not cena["salto_golfinhos_ativo"]) and agora_ts >= cena["proximo_salto_golfinhos"]:
            cena["salto_golfinhos_ativo"] = True
            cena["salto_golfinhos_inicio"] = agora_ts
            cena["salto_golfinhos_duracao"] = random.uniform(5.5, 7.0)
            cena["salto_golfinhos_x"] = random.uniform(larg * 0.57, larg * 0.78)
            cena["salto_golfinhos_sentido"] = random.choice((-1, 1))

        if cena["salto_golfinhos_ativo"]:
            progresso_golfinhos = (agora_ts - cena["salto_golfinhos_inicio"]) / max(0.1, cena["salto_golfinhos_duracao"])
            if progresso_golfinhos >= 1.0:
                cena["salto_golfinhos_ativo"] = False
                cena["proximo_salto_golfinhos"] = agora_ts + random.uniform(70, 135)
                for golfinho in cena["golfinhos_salto"]:
                    for item in golfinho.values():
                        self.canvas.itemconfig(item, state="hidden")
            else:
                for indice, golfinho in enumerate(cena["golfinhos_salto"]):
                    progresso_individual = progresso_golfinhos * 1.22 - indice * 0.27
                    if not 0.0 <= progresso_individual <= 1.0:
                        for item in golfinho.values():
                            self.canvas.itemconfig(item, state="hidden")
                        continue
                    sentido = cena["salto_golfinhos_sentido"]
                    golfinho_x = cena["salto_golfinhos_x"] + sentido * ((progresso_individual - 0.5) * 175 + indice * 16)
                    golfinho_y = alt * 0.745 - math.sin(progresso_individual * math.pi) * (58 + indice * 8)
                    tamanho = 20 - indice * 2
                    self.canvas.coords(
                        golfinho["corpo"],
                        golfinho_x + sentido * tamanho * 1.3, golfinho_y,
                        golfinho_x + sentido * tamanho * 0.25, golfinho_y - tamanho * 0.7,
                        golfinho_x - sentido * tamanho * 0.85, golfinho_y - tamanho * 0.48,
                        golfinho_x - sentido * tamanho * 1.15, golfinho_y + tamanho * 0.08,
                        golfinho_x - sentido * tamanho * 0.45, golfinho_y + tamanho * 0.42,
                    )
                    self.canvas.coords(
                        golfinho["focinho"],
                        golfinho_x + sentido * tamanho * 1.18, golfinho_y - tamanho * 0.12,
                        golfinho_x + sentido * tamanho * 1.90, golfinho_y - tamanho * 0.02,
                        golfinho_x + sentido * tamanho * 1.18, golfinho_y + tamanho * 0.16,
                    )
                    self.canvas.coords(
                        golfinho["barbatana"],
                        golfinho_x - sentido * tamanho * 0.05, golfinho_y - tamanho * 0.32,
                        golfinho_x - sentido * tamanho * 0.42, golfinho_y - tamanho * 1.05,
                        golfinho_x + sentido * tamanho * 0.25, golfinho_y - tamanho * 0.24,
                    )
                    self.canvas.coords(
                        golfinho["nadadeira_lateral"],
                        golfinho_x + sentido * tamanho * 0.05, golfinho_y + tamanho * 0.12,
                        golfinho_x - sentido * tamanho * 0.35, golfinho_y + tamanho * 0.78,
                        golfinho_x + sentido * tamanho * 0.42, golfinho_y + tamanho * 0.30,
                    )
                    self.canvas.coords(
                        golfinho["cauda"],
                        golfinho_x - sentido * tamanho * 1.05, golfinho_y,
                        golfinho_x - sentido * tamanho * 1.72, golfinho_y - tamanho * 0.62,
                        golfinho_x - sentido * tamanho * 1.50, golfinho_y,
                        golfinho_x - sentido * tamanho * 1.72, golfinho_y + tamanho * 0.62,
                    )
                    self.canvas.coords(
                        golfinho["barriga"],
                        golfinho_x - tamanho, golfinho_y - tamanho * 0.15,
                        golfinho_x + tamanho, golfinho_y + tamanho * 0.72,
                    )
                    olho_x = golfinho_x + sentido * tamanho * 0.80
                    self.canvas.coords(golfinho["olho"], olho_x - 2, golfinho_y - tamanho * 0.35 - 2, olho_x + 2, golfinho_y - tamanho * 0.35 + 2)
                    for item in golfinho.values():
                        self.canvas.itemconfig(item, state="normal")
        else:
            for golfinho in cena["golfinhos_salto"]:
                for item in golfinho.values():
                    self.canvas.itemconfig(item, state="hidden")

        tubarao = cena["tubarao"]
        intervalo_tubarao = self._obter_intervalo_tubarao_segundos()
        if intervalo_tubarao is None:
            cena["proximo_tubarao"] = float("inf")

        if (not cena["tubarao_ativo"]) and agora_ts >= cena["proximo_tubarao"]:
            cena["tubarao_ativo"] = True
            cena["tubarao_inicio"] = agora_ts
            cena["tubarao_duracao"] = random.uniform(34.0, 42.0)
            cena["tubarao_sentido"] = random.choice([-1, 1])
            cena["tubarao_base_y"] = random.uniform(alt * 0.75, alt * 0.84)
            cena["tubarao_escala"] = random.uniform(0.92, 1.1)
            cena["tubarao_fase"] = random.uniform(0.0, math.tau)

        if cena["tubarao_ativo"]:
            progresso_tubarao = (agora_ts - cena["tubarao_inicio"]) / max(1.0, cena["tubarao_duracao"])
            if progresso_tubarao >= 1.0:
                cena["tubarao_ativo"] = False
                cena["proximo_tubarao"] = (agora_ts + intervalo_tubarao) if intervalo_tubarao else float("inf")
                cena["tubarao_x"] = None
                cena["tubarao_y"] = None
                for chave in tubarao:
                    self.canvas.itemconfig(tubarao[chave], state="hidden")
            else:
                escala = cena["tubarao_escala"]
                # Mantém o tubarão apenas na área livre do mar, sem alcançar
                # a região da ilha, do trapiche e do pescador.
                faixa_inicio = max(larg * 0.60, cena["base_x"] + 240 * escala)
                faixa_fim = larg * 0.98
                x_esquerda = faixa_inicio + 26 * escala
                x_direita = faixa_fim - 92 * escala
                sentido_inicial = cena["tubarao_sentido"]

                if sentido_inicial > 0:
                    inicio_x = x_esquerda
                    retorno_x = x_direita
                else:
                    inicio_x = x_direita
                    retorno_x = x_esquerda

                if progresso_tubarao < 0.5:
                    trecho = progresso_tubarao * 2
                    tub_x = inicio_x + (retorno_x - inicio_x) * trecho
                    sentido = 1 if retorno_x > inicio_x else -1
                else:
                    trecho = (progresso_tubarao - 0.5) * 2
                    tub_x = retorno_x + (inicio_x - retorno_x) * trecho
                    sentido = 1 if inicio_x > retorno_x else -1

                fase_tub = cena["tubarao_fase"] + progresso_tubarao * math.tau * 1.8
                tub_y = cena["tubarao_base_y"] + math.sin(fase_tub) * (18 * escala)
                corpo_dx = sentido * 70 * escala
                corpo_dy = math.cos(fase_tub) * 2.5
                cena["tubarao_x"] = tub_x
                cena["tubarao_y"] = tub_y

                self.canvas.coords(
                    tubarao["corpo"],
                    tub_x + corpo_dx, tub_y + corpo_dy,
                    tub_x + sentido * 18 * escala, tub_y - 22 * escala,
                    tub_x - sentido * 28 * escala, tub_y - 18 * escala,
                    tub_x - sentido * 52 * escala, tub_y - 7 * escala,
                    tub_x - sentido * 52 * escala, tub_y + 7 * escala,
                    tub_x - sentido * 28 * escala, tub_y + 18 * escala,
                    tub_x + sentido * 18 * escala, tub_y + 22 * escala,
                )
                self.canvas.coords(
                    tubarao["barbatana"],
                    tub_x - sentido * 10 * escala, tub_y - 8 * escala,
                    tub_x - sentido * 24 * escala, tub_y - 46 * escala,
                    tub_x + sentido * 8 * escala, tub_y - 12 * escala,
                )
                self.canvas.coords(
                    tubarao["cauda"],
                    tub_x - sentido * 50 * escala, tub_y,
                    tub_x - sentido * 82 * escala, tub_y - 22 * escala,
                    tub_x - sentido * 68 * escala, tub_y - 2 * escala,
                    tub_x - sentido * 82 * escala, tub_y + 22 * escala,
                )
                frente_x = tub_x + sentido * 48 * escala
                self.canvas.coords(
                    tubarao["barriga"],
                    tub_x + sentido * 38 * escala, tub_y + 8 * escala,
                    tub_x + sentido * 8 * escala, tub_y + 20 * escala,
                    tub_x - sentido * 28 * escala, tub_y + 12 * escala,
                )
                self.canvas.coords(
                    tubarao["olho"],
                    frente_x - 5 * escala, tub_y - 12 * escala,
                    frente_x + 5 * escala, tub_y - 2 * escala,
                )
                self.canvas.coords(
                    tubarao["pupila"],
                    frente_x - sentido * 1 * escala - 2 * escala, tub_y - 10 * escala,
                    frente_x - sentido * 1 * escala + 2 * escala, tub_y - 6 * escala,
                )
                self.canvas.coords(
                    tubarao["boca"],
                    frente_x + sentido * 12 * escala, tub_y + 4 * escala,
                    frente_x + sentido * 3 * escala, tub_y + 11 * escala,
                    frente_x - sentido * 14 * escala, tub_y + 10 * escala,
                )
                self.canvas.coords(
                    tubarao["dentes"],
                    frente_x + sentido * 10 * escala, tub_y + 5 * escala,
                    frente_x + sentido * 5 * escala, tub_y + 12 * escala,
                    frente_x, tub_y + 6 * escala,
                    frente_x - sentido * 6 * escala, tub_y + 13 * escala,
                    frente_x - sentido * 12 * escala, tub_y + 7 * escala,
                    frente_x - sentido * 17 * escala, tub_y + 11 * escala,
                )
                for chave in tubarao:
                    self.canvas.itemconfig(tubarao[chave], state="normal")
        else:
            cena["tubarao_x"] = None
            cena["tubarao_y"] = None
            for chave in tubarao:
                self.canvas.itemconfig(tubarao[chave], state="hidden")

        caranguejo = cena["caranguejo"]
        if intervalo_tubarao is None:
            cena["proximo_caranguejo"] = float("inf")

        if (not cena["caranguejo_ativo"]) and agora_ts >= cena["proximo_caranguejo"]:
            cena["caranguejo_ativo"] = True
            cena["caranguejo_inicio"] = agora_ts
            cena["caranguejo_duracao"] = random.uniform(16.0, 22.0)

        if cena["caranguejo_ativo"]:
            progresso_caranguejo = (agora_ts - cena["caranguejo_inicio"]) / max(1.0, cena["caranguejo_duracao"])
            if progresso_caranguejo >= 1.0:
                cena["caranguejo_ativo"] = False
                cena["caranguejo_x"] = None
                cena["caranguejo_y"] = None
                cena["proximo_caranguejo"] = (agora_ts + intervalo_tubarao) if intervalo_tubarao else float("inf")
                for chave in caranguejo:
                    self.canvas.itemconfig(caranguejo[chave], state="hidden")
            else:
                inicio_x = int(larg * 0.11)
                fim_x = int(larg * 0.19)
                crab_x = inicio_x + (fim_x - inicio_x) * progresso_caranguejo
                crab_y = int(alt * 0.80) - math.sin(progresso_caranguejo * math.pi * 3.2) * 5
                cena["caranguejo_x"] = crab_x
                cena["caranguejo_y"] = crab_y
                pulso = math.sin(agora_ts * 8)
                garra = 8 + pulso * 2
                passo = math.sin(agora_ts * 12) * 3

                self.canvas.coords(caranguejo["corpo"], crab_x - 16, crab_y - 9, crab_x + 16, crab_y + 9)
                self.canvas.coords(
                    caranguejo["garra_esq"],
                    crab_x - 12, crab_y - 2,
                    crab_x - 24 - garra, crab_y - 12,
                    crab_x - 18, crab_y + 4,
                )
                self.canvas.coords(
                    caranguejo["garra_dir"],
                    crab_x + 12, crab_y - 2,
                    crab_x + 24 + garra, crab_y - 12,
                    crab_x + 18, crab_y + 4,
                )
                self.canvas.coords(caranguejo["perna1"], crab_x - 8, crab_y + 5, crab_x - 18, crab_y + 12 + passo, crab_x - 26, crab_y + 16)
                self.canvas.coords(caranguejo["perna2"], crab_x - 2, crab_y + 6, crab_x - 8, crab_y + 16 - passo, crab_x - 14, crab_y + 20)
                self.canvas.coords(caranguejo["perna3"], crab_x + 2, crab_y + 6, crab_x + 8, crab_y + 16 + passo, crab_x + 14, crab_y + 20)
                self.canvas.coords(caranguejo["perna4"], crab_x + 8, crab_y + 5, crab_x + 18, crab_y + 12 - passo, crab_x + 26, crab_y + 16)
                self.canvas.coords(caranguejo["olho_esq"], crab_x - 6, crab_y - 4, crab_x - 8, crab_y - 13)
                self.canvas.coords(caranguejo["olho_dir"], crab_x + 6, crab_y - 4, crab_x + 8, crab_y - 13)
                self.canvas.coords(caranguejo["pupila_esq"], crab_x - 10, crab_y - 15, crab_x - 6, crab_y - 11)
                self.canvas.coords(caranguejo["pupila_dir"], crab_x + 6, crab_y - 15, crab_x + 10, crab_y - 11)
                for chave in caranguejo:
                    self.canvas.itemconfig(caranguejo[chave], state="normal")
        else:
            cena["caranguejo_x"] = None
            cena["caranguejo_y"] = None
            for chave in caranguejo:
                self.canvas.itemconfig(caranguejo[chave], state="hidden")

        tubarao_perto = (
            cena["tubarao_x"] is not None
            and abs(cena["tubarao_x"] - cena["base_x"]) < 310
            and abs(cena["tubarao_y"] - cena["base_y"]) < 165
        )
        tubarao_longe = (
            cena["tubarao_x"] is None
            or abs(cena["tubarao_x"] - cena["base_x"]) > 430
            or abs(cena["tubarao_y"] - cena["base_y"]) > 230
        )
        if tubarao_perto and cena["pescador_modo"] != "fugindo":
            cena["pescador_modo"] = "fugindo"
            cena["estado"] = "idle"
            cena["captura_tipo"] = ""
        elif tubarao_longe and cena["pescador_modo"] == "fugindo":
            cena["pescador_modo"] = "voltando"

        if cena["pescador_modo"] == "fugindo":
            alvo_x = cena["refugio_x"]
            alvo_y = cena["refugio_y"]
        elif cena["pescador_modo"] in {"indo_coqueiro", "no_coqueiro"}:
            alvo_x = cena["coqueiro_x"] + 18
            alvo_y = cena["coqueiro_y"]
        else:
            alvo_x = cena["base_x"]
            alvo_y = cena["base_y"]

        cena["pescador_x"] = self._aproximar_valor(cena["pescador_x"], alvo_x, 7.5)
        cena["pescador_y"] = self._aproximar_valor(cena["pescador_y"], alvo_y, 5.5)

        if (
            cena["pescador_modo"] == "voltando"
            and abs(cena["pescador_x"] - cena["base_x"]) < 4
            and abs(cena["pescador_y"] - cena["base_y"]) < 4
        ):
            cena["pescador_modo"] = "pescando"
        elif (
            cena["pescador_modo"] == "indo_coqueiro"
            and abs(cena["pescador_x"] - (cena["coqueiro_x"] + 18)) < 4
            and abs(cena["pescador_y"] - cena["coqueiro_y"]) < 4
        ):
            cena["pescador_modo"] = "no_coqueiro"
            cena["estado"] = "coco_balancando"
            cena["evento_inicio"] = agora_ts
            cena["evento_ate"] = agora_ts + 2.6
        elif (
            cena["pescador_modo"] == "voltando_coqueiro"
            and abs(cena["pescador_x"] - cena["base_x"]) < 4
            and abs(cena["pescador_y"] - cena["base_y"]) < 4
        ):
            cena["pescador_modo"] = "pescando"
            cena["estado"] = "idle"
            cena["acao_tipo"] = ""
            cena["proximo_evento"] = agora_ts + random.uniform(18, 32)

        gaivota = cena["gaivota"]
        gaivota_carregando = False
        if (
            (not cena["gaivota_ativa"])
            and cena["trofeu_tipo"] in {"peixe", "peixinho"}
            and agora_ts >= cena["gaivota_pouso"]
        ):
            cena["gaivota_ativa"] = True
            cena["gaivota_inicio"] = agora_ts
            cena["gaivota_duracao"] = random.uniform(6.8, 8.4)

        if cena["gaivota_ativa"]:
            progresso_gaivota = (agora_ts - cena["gaivota_inicio"]) / max(0.1, cena["gaivota_duracao"])
            if progresso_gaivota >= 1.0:
                cena["gaivota_ativa"] = False
                cena["gaivota_pouso"] = float("inf")
                if cena["trofeu_tipo"] in {"peixe", "peixinho"}:
                    cena["trofeu_tipo"] = ""
                    cena["trofeu_instante"] = 0.0
                self._ocultar_forma_captura(gaivota)
                for chave in ("sombra", "asa_esq", "asa_dir", "cauda", "corpo", "peito", "cabeca", "bico", "olho"):
                    self.canvas.itemconfig(gaivota[chave], state="hidden")
            else:
                ponto_inicio = (larg + 80, alt * 0.28)
                ponto_pouso = (cena["trofeu_x"] + 18, cena["trofeu_y"] - 34)
                ponto_saida = (-90, alt * 0.18)
                fase_pega = 0.48
                if progresso_gaivota < fase_pega:
                    trecho = progresso_gaivota / fase_pega
                    gaiv_x = ponto_inicio[0] + (ponto_pouso[0] - ponto_inicio[0]) * trecho
                    gaiv_y = ponto_inicio[1] + (ponto_pouso[1] - ponto_inicio[1]) * trecho - math.sin(trecho * math.pi) * 18
                else:
                    trecho = (progresso_gaivota - fase_pega) / (1 - fase_pega)
                    gaiv_x = ponto_pouso[0] + (ponto_saida[0] - ponto_pouso[0]) * trecho
                    gaiv_y = ponto_pouso[1] + (ponto_saida[1] - ponto_pouso[1]) * trecho - math.sin(trecho * math.pi) * 42
                    gaivota_carregando = cena["trofeu_tipo"] in {"peixe", "peixinho"}

                batida_asa = math.sin(agora_ts * 12) * 9
                # Silhueta de gaivota de perfil: asas, cauda bifurcada e peito claro.
                self.canvas.coords(gaivota["sombra"], gaiv_x - 21, gaiv_y - 4, gaiv_x + 19, gaiv_y + 10)
                self.canvas.coords(gaivota["corpo"], gaiv_x - 19, gaiv_y - 8, gaiv_x + 17, gaiv_y + 9)
                self.canvas.coords(gaivota["peito"], gaiv_x - 20, gaiv_y - 4, gaiv_x + 1, gaiv_y + 8)
                self.canvas.coords(gaivota["cabeca"], gaiv_x - 27, gaiv_y - 8, gaiv_x - 11, gaiv_y + 7)
                self.canvas.coords(gaivota["bico"], gaiv_x - 28, gaiv_y - 2, gaiv_x - 43, gaiv_y + 1, gaiv_x - 28, gaiv_y + 4)
                self.canvas.coords(gaivota["olho"], gaiv_x - 22, gaiv_y - 3, gaiv_x - 19, gaiv_y)
                self.canvas.coords(
                    gaivota["asa_esq"],
                    gaiv_x - 10, gaiv_y - 3,
                    gaiv_x + 5, gaiv_y - 26 - batida_asa,
                    gaiv_x + 28, gaiv_y - 18 - batida_asa * 0.4,
                    gaiv_x + 14, gaiv_y + 1,
                )
                self.canvas.coords(
                    gaivota["asa_dir"],
                    gaiv_x - 4, gaiv_y + 1,
                    gaiv_x + 13, gaiv_y + 18 + batida_asa * 0.35,
                    gaiv_x + 28, gaiv_y + 9,
                    gaiv_x + 12, gaiv_y - 1,
                )
                self.canvas.coords(
                    gaivota["cauda"],
                    gaiv_x + 14, gaiv_y - 3,
                    gaiv_x + 31, gaiv_y - 10,
                    gaiv_x + 25, gaiv_y,
                    gaiv_x + 33, gaiv_y + 7,
                    gaiv_x + 14, gaiv_y + 5,
                )
                for chave in ("sombra", "asa_esq", "asa_dir", "cauda", "corpo", "peito", "cabeca", "bico", "olho"):
                    self.canvas.itemconfig(gaivota[chave], state="normal")
                if gaivota_carregando:
                    self._desenhar_forma_captura(
                        {
                            "peixe_corpo": gaivota["carga_peixe_corpo"],
                            "peixe_cauda": gaivota["carga_peixe_cauda"],
                            "peixe_olho": gaivota["carga_peixe_olho"],
                        },
                        cena["trofeu_tipo"],
                        gaiv_x - 42,
                        gaiv_y + 8,
                        escala=0.55,
                    )
                else:
                    self._ocultar_forma_captura(gaivota)
        else:
            self._ocultar_forma_captura(gaivota)
            for chave in ("sombra", "asa_esq", "asa_dir", "cauda", "corpo", "peito", "cabeca", "bico", "olho"):
                self.canvas.itemconfig(gaivota[chave], state="hidden")

        estado = cena["estado"]
        pode_pescar = cena["pescador_modo"] == "pescando"
        if pode_pescar and estado == "idle" and agora_ts >= cena["proximo_evento"]:
            proxima_acao = random.choices(("pesca", "coco", "aviao"), weights=(6, 2, 2), k=1)[0]
            cena["evento_inicio"] = agora_ts
            cena["acao_tipo"] = proxima_acao
            if proxima_acao == "coco":
                cena["estado"] = "coco_caminhando"
                cena["pescador_modo"] = "indo_coqueiro"
                cena["evento_ate"] = float("inf")
            elif proxima_acao == "aviao":
                cena["estado"] = "aviao_observando"
                cena["evento_ate"] = agora_ts + 7.5
            else:
                cena["estado"] = "puxando"
                cena["evento_ate"] = agora_ts + random.uniform(2.8, 4.0)
                cena["captura_tipo"] = random.choice([
                    "peixe", "peixe", "peixinho",
                    "bota", "bigorna", "pneu",
                ])
        elif pode_pescar and estado == "puxando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "arremessando"
            cena["evento_inicio"] = agora_ts
            cena["evento_ate"] = agora_ts + random.uniform(1.8, 2.3)
        elif pode_pescar and estado == "arremessando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "idle"
            cena["trofeu_tipo"] = cena["captura_tipo"]
            cena["trofeu_instante"] = agora_ts
            cena["gaivota_pouso"] = (
                agora_ts + random.uniform(4.2, 7.0)
                if cena["trofeu_tipo"] in {"peixe", "peixinho"}
                else float("inf")
            )
            cena["captura_tipo"] = ""
            cena["evento_inicio"] = agora_ts
            cena["proximo_evento"] = agora_ts + random.uniform(16, 28)
            cena["acao_tipo"] = ""
        elif estado == "coco_balancando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "coco_recebendo"
            cena["evento_inicio"] = agora_ts
            cena["evento_ate"] = agora_ts + 1.5
        elif estado == "coco_recebendo" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "coco_bebendo"
            cena["evento_inicio"] = agora_ts
            cena["evento_ate"] = agora_ts + 5.0
        elif estado == "coco_bebendo" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "coco_retornando"
            cena["pescador_modo"] = "voltando_coqueiro"
            cena["evento_inicio"] = agora_ts
            cena["evento_ate"] = float("inf")
        elif pode_pescar and estado == "aviao_observando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "idle"
            cena["acao_tipo"] = ""
            cena["evento_inicio"] = agora_ts
            cena["proximo_evento"] = agora_ts + random.uniform(18, 32)

        estado = cena["estado"]
        deslocamento_coqueiro = 0.0
        if estado == "coco_balancando":
            fase_coqueiro = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / 2.6))
            deslocamento_coqueiro = math.sin(fase_coqueiro * math.pi * 7) * 4
        delta_coqueiro = deslocamento_coqueiro - cena["coqueiro_deslocamento"]
        if delta_coqueiro:
            for item in [cena["coqueiro_tronco"], *cena["marcas_tronco"], *cena["folhas"], *cena["cocos"]]:
                self.canvas.move(item, delta_coqueiro, 0)
        cena["coqueiro_deslocamento"] = deslocamento_coqueiro
        if cena["estado"] == "idle":
            progresso = math.sin(agora_ts * 1.4) * 0.5 + 0.5
        else:
            duracao_evento = max(0.1, cena["evento_ate"] - cena["evento_inicio"])
            progresso_evento = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / duracao_evento))
            if cena["estado"] == "puxando":
                progresso = progresso_evento
            else:
                progresso = 0.82

        base_x = cena["pescador_x"]
        base_y = cena["pescador_y"]
        andando = cena["pescador_modo"] in {"fugindo", "voltando", "indo_coqueiro", "voltando_coqueiro"}
        direcao_andando = -1 if cena["pescador_modo"] in {"fugindo", "indo_coqueiro"} else 1
        balanco = 0 if andando else math.sin(agora_ts * 1.1) * 3
        inclinacao = -20 if andando else (-12 - (18 * progresso if cena["estado"] != "idle" else 0))
        torso_topo_x = base_x - 6 + balanco
        torso_topo_y = base_y - (46 if andando else 42)
        torso_base_x = base_x + balanco
        torso_base_y = base_y
        ombro_x = torso_topo_x + 4
        ombro_y = torso_topo_y + 10
        if andando:
            passo_perna = math.sin(agora_ts * 10) * 7
            mao_x = base_x + 22 * direcao_andando
            mao_y = base_y - 20
            ponta_x = mao_x + 40 * direcao_andando
            ponta_y = mao_y - 14
            linha_x = ponta_x
            linha_y = ponta_y
        else:
            passo_perna = 0
            mao_x = base_x + 54 + progresso * 24
            mao_y = base_y - 26 - progresso * 18
            ponta_x = mao_x + 76
            ponta_y = mao_y + inclinacao
            linha_x = ponta_x + 14
            linha_y = max(horizonte + 18, ponta_y + 56 + (1 - progresso) * 34)

        if estado == "coco_balancando":
            fase_coco = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / 2.6))
            mao_x = base_x - 20 + math.sin(fase_coco * math.pi * 7) * 8
            mao_y = base_y - 38
            ponta_x = mao_x
            ponta_y = mao_y
            linha_x = mao_x
            linha_y = mao_y
        elif estado == "coco_recebendo":
            # Depois de balancar o tronco, ele estende a mao sob os cocos.
            mao_x = base_x + 28
            mao_y = torso_topo_y - 2
            ponta_x = mao_x
            ponta_y = mao_y
            linha_x = mao_x
            linha_y = mao_y
        elif estado == "coco_bebendo":
            mao_x = torso_topo_x + 15
            mao_y = torso_topo_y - 4
            ponta_x = mao_x
            ponta_y = mao_y
            linha_x = mao_x
            linha_y = mao_y
        elif estado == "aviao_observando":
            olhar = math.sin(agora_ts * 2.8) * 7
            mao_x = torso_topo_x + 10 + olhar
            mao_y = torso_topo_y - 10
            ponta_x = mao_x
            ponta_y = mao_y
            linha_x = mao_x
            linha_y = mao_y
        captura_x = linha_x + 18
        captura_y = linha_y - 10

        if pode_pescar and cena["estado"] == "arremessando":
            duracao_arremesso = max(0.1, cena["evento_ate"] - cena["evento_inicio"])
            progresso_arremesso = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / duracao_arremesso))
            origem_x = mao_x + 18
            origem_y = mao_y - 18
            destino_x = cena["trofeu_x"]
            destino_y = cena["trofeu_y"] - 4
            captura_x = origem_x + (destino_x - origem_x) * progresso_arremesso
            captura_y = origem_y + (destino_y - origem_y) * progresso_arremesso - math.sin(progresso_arremesso * math.pi) * 92
            linha_x = mao_x + 10
            linha_y = mao_y + 4

        pescador = cena["pescador"]
        trofeu_ilha = cena["trofeu_ilha"]
        self.canvas.coords(
            pescador["banco"],
            cena["base_x"] - 28, cena["base_y"] + 6,
            cena["base_x"] + 18, cena["base_y"] + 16,
        )
        self.canvas.coords(pescador["tronco"], torso_base_x, torso_base_y, torso_topo_x, torso_topo_y)
        self.canvas.coords(
            pescador["camisa"],
            torso_topo_x - 7, torso_topo_y + 1,
            torso_topo_x + 9, torso_topo_y + 4,
            torso_base_x + 8, torso_base_y - 5,
            torso_base_x - 7, torso_base_y - 3,
        )
        self.canvas.coords(pescador["colete"], torso_topo_x + 1, torso_topo_y + 5, torso_base_x + 1, torso_base_y - 5)
        self.canvas.coords(pescador["bermuda"], torso_base_x - 7, torso_base_y - 3, torso_base_x + 10, torso_base_y + 2)
        self.canvas.coords(pescador["perna1"], base_x - 2, base_y, base_x - 16 + passo_perna, base_y + 24)
        self.canvas.coords(pescador["perna2"], base_x + 6, base_y, base_x + 24 - passo_perna, base_y + 22)
        self.canvas.coords(pescador["sapato1"], base_x - 16 + passo_perna, base_y + 24, base_x - 23 + passo_perna, base_y + 25)
        self.canvas.coords(pescador["sapato2"], base_x + 24 - passo_perna, base_y + 22, base_x + 31 - passo_perna, base_y + 23)
        self.canvas.coords(pescador["braco"], ombro_x, ombro_y, mao_x, mao_y)
        self.canvas.coords(
            pescador["vara"],
            mao_x - 8, mao_y + 4,
            mao_x + 32, mao_y - 10,
            ponta_x, ponta_y,
        )
        self.canvas.coords(pescador["linha"], ponta_x, ponta_y, linha_x, linha_y)
        self.canvas.coords(pescador["cabeca"], torso_topo_x - 12, torso_topo_y - 20, torso_topo_x + 12, torso_topo_y + 4)
        if estado in {"coco_recebendo", "coco_bebendo"}:
            direcao_rosto = 1  # O coco esta na mao direita durante essa sequencia.
        else:
            direcao_rosto = -1 if cena["pescador_modo"] in {"fugindo", "indo_coqueiro", "no_coqueiro"} else 1
        orelha_x = torso_topo_x - direcao_rosto * 11
        self.canvas.coords(pescador["orelha"], orelha_x - 3, torso_topo_y - 8, orelha_x + 3, torso_topo_y - 1)
        self.canvas.coords(
            pescador["nariz"],
            torso_topo_x + direcao_rosto * 10, torso_topo_y - 10,
            torso_topo_x + direcao_rosto * 18, torso_topo_y - 7,
            torso_topo_x + direcao_rosto * 10, torso_topo_y - 4,
        )
        olho_x = torso_topo_x + direcao_rosto * 6
        self.canvas.coords(pescador["olho"], olho_x - 2, torso_topo_y - 12, olho_x + 2, torso_topo_y - 9)
        self.canvas.coords(
            pescador["barba"],
            torso_topo_x + direcao_rosto * 2, torso_topo_y - 1,
            torso_topo_x + direcao_rosto * 13, torso_topo_y - 4,
            torso_topo_x + direcao_rosto * 9, torso_topo_y + 5,
            torso_topo_x - direcao_rosto * 3, torso_topo_y + 3,
        )
        self.canvas.coords(
            pescador["chapeu"],
            torso_topo_x - 18, torso_topo_y - 10,
            torso_topo_x + 16, torso_topo_y - 10,
            torso_topo_x + 6, torso_topo_y - 22,
        )
        self.canvas.coords(pescador["chapeu_faixa"], torso_topo_x - 11, torso_topo_y - 12, torso_topo_x + 11, torso_topo_y - 12)

        acoes = cena["acoes"]
        for item in acoes.values():
            self.canvas.itemconfig(item, state="hidden")
        acao_especial = estado in {
            "coco_caminhando", "coco_balancando", "coco_recebendo", "coco_bebendo", "coco_retornando", "aviao_observando"
        }
        self.canvas.itemconfig(pescador["vara"], state="hidden" if acao_especial else "normal")
        self.canvas.itemconfig(pescador["linha"], state="hidden" if acao_especial else "normal")

        if estado == "coco_recebendo":
            fase_coco = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / 1.5))
            inicio_coco_x = mao_x + 3
            inicio_coco_y = int(alt * 0.57) + 8
            coco_x = inicio_coco_x + (mao_x - inicio_coco_x) * fase_coco
            coco_y = inicio_coco_y + (mao_y - inicio_coco_y) * fase_coco * fase_coco
            self.canvas.coords(acoes["coco_caindo"], coco_x - 10, coco_y - 10, coco_x + 10, coco_y + 10)
            self.canvas.itemconfig(acoes["coco_caindo"], state="normal")
        elif estado == "coco_bebendo":
            self.canvas.coords(acoes["coco_aberto"], mao_x - 11, mao_y - 6, mao_x + 11, mao_y + 14)
            self.canvas.coords(acoes["coco_polpa"], mao_x - 8, mao_y - 3, mao_x + 8, mao_y + 11)
            self.canvas.itemconfig(acoes["coco_aberto"], state="normal")
            self.canvas.itemconfig(acoes["coco_polpa"], state="normal")
        elif estado == "aviao_observando":
            fase_aviao = min(1.0, max(0.0, (agora_ts - cena["evento_inicio"]) / 7.5))
            aviao_x = larg * (0.08 + fase_aviao * 0.84)
            aviao_y = alt * 0.20 + math.sin(fase_aviao * math.pi * 2) * 12
            self.canvas.coords(
                acoes["aviao_corpo"],
                aviao_x - 17, aviao_y,
                aviao_x + 16, aviao_y - 4,
                aviao_x + 23, aviao_y,
                aviao_x + 16, aviao_y + 4,
            )
            self.canvas.coords(acoes["aviao_asa"], aviao_x - 3, aviao_y, aviao_x + 7, aviao_y - 13, aviao_x + 10, aviao_y + 2)
            self.canvas.coords(acoes["aviao_cauda"], aviao_x - 14, aviao_y - 2, aviao_x - 19, aviao_y - 10, aviao_x - 7, aviao_y - 3)
            olhar = math.sin(agora_ts * 2.8) * 7
            bino_x = torso_topo_x + 14 + olhar
            bino_y = torso_topo_y - 12
            self.canvas.coords(acoes["binoculo_esq"], bino_x - 10, bino_y - 6, bino_x, bino_y + 5)
            self.canvas.coords(acoes["binoculo_dir"], bino_x + 2, bino_y - 6, bino_x + 12, bino_y + 5)
            self.canvas.coords(acoes["binoculo_ponte"], bino_x, bino_y, bino_x + 2, bino_y)
            for chave in ("aviao_corpo", "aviao_asa", "aviao_cauda", "binoculo_esq", "binoculo_dir", "binoculo_ponte"):
                self.canvas.itemconfig(acoes[chave], state="normal")

        ripple_cx = linha_x
        ripple_cy = linha_y + 6
        for indice, item in enumerate(cena["ripples"]):
            if andando or acao_especial:
                self.canvas.itemconfig(item, state="hidden")
            else:
                raio = 10 + indice * 12 + math.sin(agora_ts * 3 + indice) * 2
                self.canvas.coords(
                    item,
                    ripple_cx - raio, ripple_cy - raio * 0.45,
                    ripple_cx + raio, ripple_cy + raio * 0.45,
                )
                self.canvas.itemconfig(item, state="normal")

        captura = pescador["captura"]
        self.canvas.itemconfig(captura, text="", state="hidden")
        if cena["trofeu_tipo"] and not (cena["gaivota_ativa"] and gaivota_carregando):
            self._desenhar_forma_captura(
                trofeu_ilha,
                cena["trofeu_tipo"],
                cena["trofeu_x"],
                cena["trofeu_y"],
                escala=0.92,
            )
        else:
            self._ocultar_forma_captura(trofeu_ilha)

        if cena["estado"] == "arremessando" and cena["captura_tipo"]:
            self._desenhar_forma_captura(
                {
                    "peixe_corpo": pescador["captura_peixe_corpo"],
                    "peixe_cauda": pescador["captura_peixe_cauda"],
                    "peixe_olho": pescador["captura_peixe_olho"],
                    "bota_cano": pescador["captura_bota_cano"],
                    "bota_sola": pescador["captura_bota_sola"],
                    "bigorna_base": pescador["captura_bigorna_base"],
                    "bigorna_topo": pescador["captura_bigorna_topo"],
                    "pneu_externo": pescador["captura_pneu_externo"],
                    "pneu_interno": pescador["captura_pneu_interno"],
                    "pneu_aro": pescador["captura_pneu_aro"],
                    "pneu_brilho": pescador["captura_pneu_brilho"],
                },
                cena["captura_tipo"],
                captura_x,
                captura_y,
                escala=0.95,
            )
        else:
            self._ocultar_forma_captura(
                {
                    "peixe_corpo": pescador["captura_peixe_corpo"],
                    "peixe_cauda": pescador["captura_peixe_cauda"],
                    "peixe_olho": pescador["captura_peixe_olho"],
                    "bota_cano": pescador["captura_bota_cano"],
                    "bota_sola": pescador["captura_bota_sola"],
                    "bigorna_base": pescador["captura_bigorna_base"],
                    "bigorna_topo": pescador["captura_bigorna_topo"],
                    "pneu_externo": pescador["captura_pneu_externo"],
                    "pneu_interno": pescador["captura_pneu_interno"],
                    "pneu_aro": pescador["captura_pneu_aro"],
                    "pneu_brilho": pescador["captura_pneu_brilho"],
                }
            )

    def _loop_animacao_pescador(self):
        """Executa o loop suave do fundo da ilha."""
        if not self._animacao_pescador_ativa:
            return
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return

        if not self._pescador_cena:
            self._animacao_pescador_ativa = False
            return

        self._atualizar_cena_pescador(datetime.now(), time.time())
        self._animacao_pescador_after = self.after(120, self._loop_animacao_pescador)

    def _parar_animacao_pescador(self):
        """Interrompe o loop da ilha do pescador."""
        self._animacao_pescador_ativa = False
        if self._animacao_pescador_after is not None:
            try:
                self.after_cancel(self._animacao_pescador_after)
            except Exception:
                pass
        self._animacao_pescador_after = None
        self._pescador_cena = {}

    def _desenhar_fundo_blocos(self, larg, alt, cor_destaque):
        """Cria uma paisagem original em blocos, com construtor e vila em evolucao."""
        tags = self._tags_fundo_blocos
        horizonte = int(alt * 0.58)
        estado_solar = self._obter_estado_solar(datetime.now(), larg, alt)
        ceu = []
        for indice in range(7):
            y1 = int(horizonte * indice / 7)
            y2 = int(horizonte * (indice + 1) / 7)
            ceu.append(self.canvas.create_rectangle(0, y1, larg, y2, fill=estado_solar["ceu_topo"], outline="", tags=tags))

        estrelas = []
        for indice in range(24):
            x = int(larg * ((indice * 37 % 91 + 4) / 100))
            y = int(horizonte * ((indice * 19 % 58 + 5) / 100))
            estrelas.append(self.canvas.create_rectangle(x, y, x + 3, y + 3, fill="#EAF6FF", outline="", state="hidden", tags=tags))

        sol_aura = self.canvas.create_rectangle(0, 0, 0, 0, fill="#FFF0A6", outline="", stipple="gray50", tags=tags)
        sol = self.canvas.create_rectangle(0, 0, 0, 0, fill=estado_solar["sol_cor"], outline="#FFF8C6", width=2, tags=tags)
        lua = self.canvas.create_rectangle(0, 0, 0, 0, fill="#D8E8F4", outline="#FFFFFF", width=1, state="hidden", tags=tags)

        morros = []
        bloco = max(24, int(larg / 60))
        for coluna in range(-1, int(larg / bloco) + 2):
            altura_morro = int((math.sin(coluna * 0.38) + math.sin(coluna * 0.13 + 1.5) + 2.2) * bloco * 1.7)
            x1 = coluna * bloco
            x2 = x1 + bloco + 1
            morros.append(self.canvas.create_rectangle(x1, horizonte - altura_morro, x2, horizonte, fill="#356B49", outline="#244D37", width=1, tags=tags))

        solo = self.canvas.create_rectangle(0, horizonte, larg, alt, fill="#7C5A38", outline="", tags=tags)
        grama = self.canvas.create_rectangle(0, horizonte - 10, larg, horizonte + 12, fill="#56A64B", outline="#2C6C36", width=2, tags=tags)
        detalhes_solo = []
        for indice in range(38):
            x = (indice * 79) % larg
            y = horizonte + 18 + (indice * 47) % max(40, int(alt * 0.35))
            detalhes_solo.append(self.canvas.create_rectangle(x, y, x + 16, y + 9, fill="#936B42", outline="", tags=tags))

        rio = self.canvas.create_polygon(
            int(larg * 0.89), horizonte,
            int(larg * 0.95), horizonte,
            int(larg * 0.96), alt,
            int(larg * 0.80), alt,
            fill="#2C91BE", outline="#8AD8F3", width=2, tags=tags,
        )
        ondas_rio = [self.canvas.create_line(0, 0, 0, 0, fill="#B8EEFF", width=2, tags=tags) for _ in range(5)]

        arvores = []
        for frac_x, escala in ((0.10, 1.0), (0.24, 0.78), (0.88, 1.08)):
            x = int(larg * frac_x)
            y = int(horizonte * 1.06)
            tronco = self.canvas.create_rectangle(x - 8 * escala, y - 58 * escala, x + 8 * escala, y, fill="#704526", outline="#3F2817", width=2, tags=tags)
            copa = []
            for dx, dy in ((0, -82), (-24, -56), (25, -58), (0, -42)):
                r = 26 * escala
                copa.append(self.canvas.create_rectangle(x + dx * escala - r, y + dy * escala - r, x + dx * escala + r, y + dy * escala + r, fill="#2F8744", outline="#1E592F", width=2, tags=tags))
            arvores.append({"tronco": tronco, "copa": copa})

        horta = []
        for indice in range(7):
            x = int(larg * 0.12) + indice * max(18, int(larg * 0.018))
            y = int(alt * 0.78) + (indice % 2) * 13
            terra = self.canvas.create_rectangle(x - 8, y, x + 10, y + 12, fill="#5D3E26", outline="#AD7A45", width=1, tags=tags)
            planta = self.canvas.create_rectangle(x - 3, y - 18, x + 4, y, fill="#75C94D", outline="#D4F58A", width=1, tags=tags)
            horta.append((terra, planta))

        casa_blocos = []
        casa_x = int(larg * 0.63)
        casa_y = int(alt * 0.75)
        posicoes_casa = [
            (indice * 28, 0) for indice in range(6)
        ] + [
            (0, -28), (28, -28), (56, -28), (84, -28), (112, -28), (140, -28),
            (0, -56), (28, -56), (112, -56), (140, -56),
            (0, -84), (28, -84), (56, -84), (84, -84), (112, -84), (140, -84),
        ]
        for dx, dy in posicoes_casa:
            casa_blocos.append(self.canvas.create_rectangle(casa_x + dx, casa_y + dy, casa_x + dx + 28, casa_y + dy + 28, fill="#B98750", outline="#5A3923", width=2, state="hidden", tags=tags))
        telhado = [
            self.canvas.create_rectangle(casa_x + 24, casa_y - 112, casa_x + 144, casa_y - 84, fill="#A9473F", outline="#5D2825", width=2, state="hidden", tags=tags),
            self.canvas.create_rectangle(casa_x + 52, casa_y - 140, casa_x + 116, casa_y - 112, fill="#A9473F", outline="#5D2825", width=2, state="hidden", tags=tags),
        ]
        porta = self.canvas.create_rectangle(casa_x + 62, casa_y - 56, casa_x + 88, casa_y, fill="#4B301E", outline="#D6A15B", width=2, state="hidden", tags=tags)
        blocos_extra = []
        oficina_x, oficina_y = int(larg * 0.42), int(alt * 0.80)
        for dx, dy in ([(indice * 24, 0) for indice in range(6)] + [(0, -24), (24, -24), (96, -24), (120, -24), (0, -48), (120, -48), (24, -72), (48, -72), (72, -72), (96, -72)]):
            blocos_extra.append(self.canvas.create_rectangle(oficina_x + dx, oficina_y + dy, oficina_x + dx + 24, oficina_y + dy + 24, fill="#8797A1", outline="#3E4D55", width=2, state="hidden", tags=tags))
        torre_x, torre_y = int(larg * 0.54), int(alt * 0.72)
        for dx, dy in [(0, -indice * 22) for indice in range(8)] + [(22, -indice * 22) for indice in range(5)]:
            blocos_extra.append(self.canvas.create_rectangle(torre_x + dx, torre_y + dy, torre_x + dx + 22, torre_y + dy + 22, fill="#A7B3BD", outline="#56636B", width=2, state="hidden", tags=tags))
        ponte_y = int(alt * 0.74)
        proporcao_ponte = (ponte_y - horizonte) / max(1, alt - horizonte)
        ponte_esq = int(larg * (0.89 - 0.09 * proporcao_ponte)) - 18
        ponte_dir = int(larg * (0.95 + 0.01 * proporcao_ponte)) + 18
        largura_ponte_bloco = (ponte_dir - ponte_esq) / 9
        for indice in range(9):
            ponte_x = ponte_esq + indice * largura_ponte_bloco
            blocos_extra.append(self.canvas.create_rectangle(ponte_x, ponte_y, ponte_x + largura_ponte_bloco + 1, ponte_y + 14, fill="#9A673B", outline="#56371F", width=2, state="hidden", tags=tags))
        mirante_blocos = []
        mirante_x, mirante_y = int(larg * 0.31), int(alt * 0.74)
        for dx, dy in ((0, 0), (24, 0), (48, 0), (0, -24), (48, -24), (0, -48), (48, -48), (24, -72), (24, -96)):
            mirante_blocos.append(self.canvas.create_rectangle(mirante_x + dx, mirante_y + dy, mirante_x + dx + 24, mirante_y + dy + 24, fill="#C48D4A", outline="#65411F", width=2, state="hidden", tags=tags))
        obras = [
            {"nome": "a cabana", "blocos": casa_blocos, "alvo": (larg * 0.59, alt * 0.78)},
            {"nome": "a oficina", "blocos": blocos_extra[:16], "alvo": (larg * 0.41, alt * 0.80)},
            {"nome": "a torre", "blocos": blocos_extra[16:29], "alvo": (larg * 0.53, alt * 0.70)},
            {"nome": "a ponte", "blocos": blocos_extra[29:], "alvo": ((ponte_esq + ponte_dir) / 2, alt * 0.74)},
            {"nome": "o mirante", "blocos": mirante_blocos, "alvo": (larg * 0.31, alt * 0.74)},
        ]
        acabamentos = [
            [],
            [
                self.canvas.create_rectangle(oficina_x - 8, oficina_y - 96, oficina_x + 152, oficina_y - 72, fill="#53646D", outline="#2D3C43", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(oficina_x + 57, oficina_y - 48, oficina_x + 84, oficina_y, fill="#394851", outline="#BBC8CF", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(oficina_x + 16, oficina_y - 53, oficina_x + 43, oficina_y - 31, fill="#8DD4F2", outline="#315A70", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(oficina_x + 104, oficina_y - 53, oficina_x + 131, oficina_y - 31, fill="#8DD4F2", outline="#315A70", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(oficina_x + 116, oficina_y - 119, oficina_x + 132, oficina_y - 96, fill="#6D4C37", outline="#3F2D22", width=2, state="hidden", tags=tags),
            ],
            [
                self.canvas.create_rectangle(torre_x - 8, torre_y - 190, torre_x + 52, torre_y - 168, fill="#7A4A3A", outline="#462A22", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(torre_x + 5, torre_y - 145, torre_x + 17, torre_y - 125, fill="#91DAF1", outline="#445D6B", width=1, state="hidden", tags=tags),
                self.canvas.create_rectangle(torre_x + 5, torre_y - 95, torre_x + 17, torre_y - 75, fill="#91DAF1", outline="#445D6B", width=1, state="hidden", tags=tags),
                self.canvas.create_line(torre_x + 22, torre_y - 190, torre_x + 22, torre_y - 224, fill="#513326", width=3, state="hidden", tags=tags),
                self.canvas.create_rectangle(torre_x + 22, torre_y - 224, torre_x + 45, torre_y - 210, fill="#E9B93F", outline="#9C6821", width=1, state="hidden", tags=tags),
            ],
            [
                self.canvas.create_line(ponte_esq, ponte_y - 8, ponte_dir, ponte_y - 8, fill="#684322", width=3, state="hidden", tags=tags),
                self.canvas.create_line(ponte_esq + 8, ponte_y - 8, ponte_esq + 8, ponte_y - 28, fill="#684322", width=3, state="hidden", tags=tags),
                self.canvas.create_line(ponte_dir - 8, ponte_y - 8, ponte_dir - 8, ponte_y - 28, fill="#684322", width=3, state="hidden", tags=tags),
            ],
            [
                self.canvas.create_rectangle(mirante_x - 10, mirante_y - 120, mirante_x + 82, mirante_y - 96, fill="#A9473F", outline="#5D2825", width=2, state="hidden", tags=tags),
                self.canvas.create_rectangle(mirante_x + 23, mirante_y - 67, mirante_x + 49, mirante_y - 43, fill="#8DD4F2", outline="#5A3923", width=2, state="hidden", tags=tags),
                self.canvas.create_line(mirante_x + 36, mirante_y - 120, mirante_x + 36, mirante_y - 150, fill="#543926", width=3, state="hidden", tags=tags),
                self.canvas.create_rectangle(mirante_x + 36, mirante_y - 150, mirante_x + 60, mirante_y - 136, fill="#E9B93F", outline="#9C6821", width=1, state="hidden", tags=tags),
            ],
        ]

        construtor = {
            "cabeca": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F0B27A", outline="#8C5634", width=1, tags=tags),
            "cabelo": self.canvas.create_rectangle(0, 0, 0, 0, fill="#4C3025", outline="", tags=tags),
            "corpo": self.canvas.create_rectangle(0, 0, 0, 0, fill="#3F91C8", outline="#1F5075", width=1, tags=tags),
            "perna1": self.canvas.create_rectangle(0, 0, 0, 0, fill="#273B68", outline="", tags=tags),
            "perna2": self.canvas.create_rectangle(0, 0, 0, 0, fill="#273B68", outline="", tags=tags),
            "braco": self.canvas.create_line(0, 0, 0, 0, fill="#F0B27A", width=6, tags=tags),
            "picareta": self.canvas.create_line(0, 0, 0, 0, fill="#6E4B2C", width=3, tags=tags),
            "bloco": self.canvas.create_rectangle(0, 0, 0, 0, fill="#B98750", outline="#5A3923", width=2, state="hidden", tags=tags),
        }
        animais = []
        for cor, contorno in (("#F5F2E9", "#A4A09A"), ("#FFFFFF", "#C8D3DA")):
            animais.append({
                "corpo": self.canvas.create_rectangle(0, 0, 0, 0, fill=cor, outline=contorno, width=1, tags=tags),
                "cabeca": self.canvas.create_rectangle(0, 0, 0, 0, fill=cor, outline=contorno, width=1, tags=tags),
                "pernas": self.canvas.create_line(0, 0, 0, 0, fill="#4D4035", width=3, tags=tags),
            })
        pedras = [
            self.canvas.create_rectangle(int(larg * 0.47) + dx, int(alt * 0.77) + dy, int(larg * 0.47) + dx + 20, int(alt * 0.77) + dy + 16, fill="#71808A", outline="#3D4B54", width=2, tags=tags)
            for dx, dy in ((0, 0), (22, -9), (43, 3))
        ]
        particulas = [
            self.canvas.create_rectangle(0, 0, 0, 0, fill="#D5E2E7", outline="", state="hidden", tags=tags)
            for _ in range(7)
        ]
        rotulo_acao = self.canvas.create_text(0, 0, text="", font=("Segoe UI", 12, "bold"), fill="#FFF0B0", anchor="s", state="hidden", tags=tags)
        balao = {
            "corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#F06D50", outline="#FFE3A3", width=2, state="hidden", tags=tags),
            "cesta": self.canvas.create_rectangle(0, 0, 0, 0, fill="#9A673B", outline="#F3C77A", width=1, state="hidden", tags=tags),
            "cordas": self.canvas.create_line(0, 0, 0, 0, fill="#F4D6A1", width=1, state="hidden", tags=tags),
        }
        estrela_cadente = self.canvas.create_line(0, 0, 0, 0, fill="#FFF5C4", width=3, state="hidden", tags=tags)
        barra_y = alt - 116
        tamanho_slot, espaco_slot, quantidade_slots = 42, 5, 9
        barra_larg = quantidade_slots * tamanho_slot + (quantidade_slots + 1) * espaco_slot
        barra_x = (larg - barra_larg) / 2
        barra_fundo = self.canvas.create_rectangle(barra_x, barra_y, barra_x + barra_larg, barra_y + tamanho_slot + 10, fill="#29363D", outline="#A9B8BA", width=2, tags=tags)
        slots_barra, icones_barra = [], []
        cores_itens = ("#A87843", "#7D8D99", "#70B657", "#C79B54", "#62B7D4", "#D05C45", "#E6C24F", "#A778B8", "#EFF2E4")
        for indice, cor_item in enumerate(cores_itens):
            x_slot = barra_x + espaco_slot + indice * (tamanho_slot + espaco_slot)
            slots_barra.append(self.canvas.create_rectangle(x_slot, barra_y + 5, x_slot + tamanho_slot, barra_y + 5 + tamanho_slot, fill="#172127", outline="#75858A", width=2, tags=tags))
            if indice in (0, 1, 3):
                icones_barra.append(self.canvas.create_line(x_slot + 11, barra_y + 37, x_slot + 30, barra_y + 15, fill=cor_item, width=5, tags=tags))
                icones_barra.append(self.canvas.create_line(x_slot + 21, barra_y + 14, x_slot + 34, barra_y + 18, fill="#D7E2E1", width=4, tags=tags))
            elif indice == 4:
                icones_barra.append(self.canvas.create_rectangle(x_slot + 11, barra_y + 15, x_slot + 31, barra_y + 35, fill=cor_item, outline="#D8F4FF", width=1, tags=tags))
            else:
                icones_barra.append(self.canvas.create_rectangle(x_slot + 12, barra_y + 15, x_slot + 31, barra_y + 34, fill=cor_item, outline="#1C2930", width=1, tags=tags))
        selecao_barra = self.canvas.create_rectangle(0, 0, 0, 0, outline="#F4E7A6", width=3, tags=tags)
        barra_status = self.canvas.create_text(larg / 2, barra_y - 12, text="VILA: 0/5", font=("Segoe UI", 11, "bold"), fill="#F4E7A6", anchor="s", tags=tags)
        criatura = {
            "cabeca": self.canvas.create_rectangle(0, 0, 0, 0, fill="#5CAD48", outline="#1E4E2D", width=2, state="hidden", tags=tags),
            "corpo": self.canvas.create_rectangle(0, 0, 0, 0, fill="#4B9A42", outline="#1E4E2D", width=2, state="hidden", tags=tags),
            "perna1": self.canvas.create_rectangle(0, 0, 0, 0, fill="#387C37", outline="#1E4E2D", width=1, state="hidden", tags=tags),
            "perna2": self.canvas.create_rectangle(0, 0, 0, 0, fill="#387C37", outline="#1E4E2D", width=1, state="hidden", tags=tags),
            "olho1": self.canvas.create_rectangle(0, 0, 0, 0, fill="#172B22", outline="", state="hidden", tags=tags),
            "olho2": self.canvas.create_rectangle(0, 0, 0, 0, fill="#172B22", outline="", state="hidden", tags=tags),
            "boca": self.canvas.create_rectangle(0, 0, 0, 0, fill="#172B22", outline="", state="hidden", tags=tags),
        }
        explosao = [
            self.canvas.create_rectangle(0, 0, 0, 0, fill="#FFCC3D", outline="#F26430", width=1, state="hidden", tags=tags)
            for _ in range(18)
        ]

        self._blocos_cena = {
            "larg": larg, "alt": alt, "horizonte": horizonte, "ceu": ceu, "estrelas": estrelas,
            "sol": sol, "sol_aura": sol_aura, "lua": lua, "morros": morros, "solo": solo, "grama": grama,
            "detalhes_solo": detalhes_solo, "rio": rio, "ondas_rio": ondas_rio, "arvores": arvores, "horta": horta,
            "casa_blocos": casa_blocos, "blocos_extra": blocos_extra, "obras": obras, "acabamentos": acabamentos, "telhado": telhado, "porta": porta, "construtor": construtor, "animais": animais,
            "pedras": pedras, "particulas": particulas, "rotulo_acao": rotulo_acao, "balao": balao, "estrela_cadente": estrela_cadente, "criatura": criatura, "explosao": explosao,
            "barra": {"slots": slots_barra, "selecao": selecao_barra, "status": barra_status, "x": barra_x, "y": barra_y, "tamanho": tamanho_slot, "espaco": espaco_slot},
            "construtor_x": float(int(larg * 0.58)), "construtor_y": float(int(alt * 0.78)),
            "estado": "construindo", "inicio_acao": time.time(), "fim_acao": time.time(),
            "proxima_acao": time.time(), "progresso_obras": [12, 0, 0, 0, 0], "reparo_indice": None,
            "ataque_fase": "", "ataque_alvo": None, "ultimo_ataque": -1, "ataque_x": -80.0, "ataque_y": 0.0, "ataque_inicio": 0.0,
            "proximo_ataque": time.time() + random.uniform(12, 22),
            "evento_ambiente": "", "evento_inicio": 0.0, "evento_fim": 0.0,
            "proximo_evento_ambiente": time.time() + random.uniform(18, 32),
        }
        self._atualizar_cena_blocos(datetime.now(), time.time())
        self._animacao_blocos_ativa = True
        self._animacao_blocos_after = self.after(100, self._loop_animacao_blocos)

    def _atualizar_cena_blocos(self, agora, agora_ts):
        """Atualiza o ciclo do mundo de blocos sem usar imagens externas."""
        cena = self._blocos_cena
        if not cena:
            return
        larg, alt, horizonte = cena["larg"], cena["alt"], cena["horizonte"]
        solar = self._obter_estado_solar(agora, larg, alt)
        for indice, item in enumerate(cena["ceu"]):
            self.canvas.itemconfig(item, fill=self._mix_cor(solar["ceu_topo"], solar["ceu_base"], indice / 6))
        noite = not solar["sol_visivel"]
        for estrela in cena["estrelas"]:
            self.canvas.itemconfig(estrela, state="normal" if noite else "hidden")
        if solar["sol_visivel"]:
            sx, sy = solar["sol_x"], solar["sol_y"]
            self.canvas.coords(cena["sol_aura"], sx - 54, sy - 54, sx + 54, sy + 54)
            self.canvas.coords(cena["sol"], sx - 28, sy - 28, sx + 28, sy + 28)
            self.canvas.itemconfig(cena["sol_aura"], fill=solar["sol_cor"], state="normal")
            self.canvas.itemconfig(cena["sol"], fill=solar["sol_cor"], state="normal")
            self.canvas.itemconfig(cena["lua"], state="hidden")
        else:
            self.canvas.itemconfig(cena["sol"], state="hidden")
            self.canvas.itemconfig(cena["sol_aura"], state="hidden")
            self.canvas.coords(cena["lua"], int(larg * 0.80) - 20, int(horizonte * 0.20) - 20, int(larg * 0.80) + 20, int(horizonte * 0.20) + 20)
            self.canvas.itemconfig(cena["lua"], state="normal")

        for indice, onda in enumerate(cena["ondas_rio"]):
            y = horizonte + 28 + indice * 42
            deslocamento = math.sin(agora_ts * 1.5 + indice) * 12
            proporcao_rio = (y - horizonte) / max(1, alt - horizonte)
            margem_esq = larg * (0.89 - 0.09 * proporcao_rio)
            margem_dir = larg * (0.95 + 0.01 * proporcao_rio)
            self.canvas.coords(onda, margem_esq + 12 + deslocamento, y, margem_dir - 12 + deslocamento, y)

        obras = cena["obras"]
        progresso = cena["progresso_obras"]
        indice_obra = cena["reparo_indice"]
        reparando = indice_obra is not None
        if indice_obra is None:
            indice_obra = next((indice for indice, obra in enumerate(obras) if progresso[indice] < len(obra["blocos"])), None)

        if indice_obra is None:
            # A vila pronta continua viva, mas uma nova ameaca so e liberada apos o reparo anterior.
            ciclo = int(agora_ts / 5) % len(obras)
            alvo_x, alvo_y = obras[ciclo]["alvo"]
            texto_acao = "Inspecionando " + obras[ciclo]["nome"]
        else:
            alvo_x, alvo_y = obras[indice_obra]["alvo"]
            texto_acao = ("Consertando " if reparando else "Construindo ") + obras[indice_obra]["nome"]

        cena["construtor_x"] = self._aproximar_valor(cena["construtor_x"], alvo_x, 3.6)
        cena["construtor_y"] = self._aproximar_valor(cena["construtor_y"], alvo_y, 2.2)
        x, y = cena["construtor_x"], cena["construtor_y"]
        em_acao = indice_obra is not None and abs(x - alvo_x) < 5 and abs(y - alvo_y) < 4
        if em_acao and agora_ts >= cena["fim_acao"]:
            progresso[indice_obra] = min(len(obras[indice_obra]["blocos"]), progresso[indice_obra] + 1)
            cena["inicio_acao"] = agora_ts
            cena["fim_acao"] = agora_ts + 0.72
            if reparando and progresso[indice_obra] == len(obras[indice_obra]["blocos"]):
                cena["reparo_indice"] = None
                cena["proximo_ataque"] = agora_ts + random.uniform(12, 22)

        todas_concluidas = all(progresso[indice] == len(obra["blocos"]) for indice, obra in enumerate(obras))
        if not solar["sol_visivel"] and cena["ataque_fase"]:
            cena["ataque_fase"] = ""
            cena["proximo_ataque"] = agora_ts + random.uniform(12, 22)
        if (solar["sol_visivel"] and todas_concluidas and cena["reparo_indice"] is None
                and not cena["ataque_fase"] and agora_ts >= cena["proximo_ataque"]):
            alvos_possiveis = [indice for indice in range(len(obras)) if indice != cena["ultimo_ataque"]]
            cena["ataque_alvo"] = random.choice(alvos_possiveis)
            cena["ataque_fase"] = "aproximando"
            cena["ataque_x"] = -48.0
            cena["ataque_y"] = alt * 0.79

        if cena["ataque_fase"] == "aproximando":
            alvo_ataque_x, alvo_ataque_y = obras[cena["ataque_alvo"]]["alvo"]
            cena["ataque_x"] = self._aproximar_valor(cena["ataque_x"], alvo_ataque_x - 42, 2.5)
            cena["ataque_y"] = self._aproximar_valor(cena["ataque_y"], alvo_ataque_y, 1.6)
            if abs(cena["ataque_x"] - (alvo_ataque_x - 42)) < 4 and abs(cena["ataque_y"] - alvo_ataque_y) < 3:
                cena["ataque_fase"] = "explodindo"
                cena["ataque_inicio"] = agora_ts
        elif cena["ataque_fase"] == "explodindo" and agora_ts - cena["ataque_inicio"] > 1.0:
            alvo_destruido = cena["ataque_alvo"]
            dano = min(4, max(2, len(obras[alvo_destruido]["blocos"]) // 3))
            progresso[alvo_destruido] = max(0, progresso[alvo_destruido] - dano)
            cena["ultimo_ataque"] = alvo_destruido
            cena["reparo_indice"] = alvo_destruido
            cena["ataque_fase"] = ""
            cena["ataque_alvo"] = None
            cena["fim_acao"] = agora_ts

        passo = math.sin(agora_ts * 9) * 5 if not em_acao else 0
        c = cena["construtor"]
        escala = 0.72 + ((y - horizonte) / max(1, alt - horizonte)) * 0.38
        self.canvas.coords(c["cabeca"], x - 10 * escala, y - 48 * escala, x + 10 * escala, y - 28 * escala)
        self.canvas.coords(c["cabelo"], x - 11 * escala, y - 51 * escala, x + 11 * escala, y - 43 * escala)
        self.canvas.coords(c["corpo"], x - 13 * escala, y - 28 * escala, x + 13 * escala, y - 4 * escala)
        self.canvas.coords(c["perna1"], x - 11 * escala, y - 4 * escala, x - 2 * escala + passo * escala, y + 18 * escala)
        self.canvas.coords(c["perna2"], x + 2 * escala, y - 4 * escala, x + 11 * escala - passo * escala, y + 18 * escala)
        mao_x = x + 20 * escala
        mao_y = y - 20 * escala + (math.sin(agora_ts * 8) * 8 * escala if em_acao else 0)
        self.canvas.coords(c["braco"], x + 9 * escala, y - 23 * escala, mao_x, mao_y)
        self.canvas.coords(c["picareta"], mao_x - 6 * escala, mao_y - 12 * escala, mao_x + 11 * escala, mao_y + 13 * escala)
        if em_acao:
            self.canvas.coords(c["bloco"], mao_x + 4 * escala, mao_y - 8 * escala, mao_x + 20 * escala, mao_y + 8 * escala)
            self.canvas.itemconfig(c["bloco"], state="normal")
        else:
            self.canvas.itemconfig(c["bloco"], state="hidden")

        for indice, obra in enumerate(obras):
            for indice_bloco, bloco_obra in enumerate(obra["blocos"]):
                self.canvas.itemconfig(bloco_obra, state="normal" if indice_bloco < progresso[indice] else "hidden")
            obra_pronta = progresso[indice] == len(obra["blocos"])
            for acabamento in cena["acabamentos"][indice]:
                self.canvas.itemconfig(acabamento, state="normal" if obra_pronta else "hidden")
        telhado_visivel = progresso[0] >= len(cena["casa_blocos"]) - 3
        for item in cena["telhado"]:
            self.canvas.itemconfig(item, state="normal" if telhado_visivel else "hidden")
        self.canvas.itemconfig(cena["porta"], state="normal" if progresso[0] >= 12 else "hidden")

        concluidas = sum(progresso[indice] == len(obra["blocos"]) for indice, obra in enumerate(obras))
        barra = cena["barra"]
        slot_ativo = indice_obra if indice_obra is not None else int(agora_ts / 5) % len(barra["slots"])
        x_slot = barra["x"] + barra["espaco"] + slot_ativo * (barra["tamanho"] + barra["espaco"])
        self.canvas.coords(barra["selecao"], x_slot - 2, barra["y"] + 3, x_slot + barra["tamanho"] + 2, barra["y"] + barra["tamanho"] + 9)
        estado_vila = "REPARO" if cena["reparo_indice"] is not None else "VILA"
        self.canvas.itemconfig(barra["status"], text=f"{estado_vila}: {concluidas}/5")

        for indice, animal in enumerate(cena["animais"]):
            base_x = larg * (0.14 + indice * 0.24)
            animal_x = base_x + math.sin(agora_ts * (0.36 + indice * 0.07)) * 35
            animal_y = alt * (0.73 + indice * 0.025)
            self.canvas.coords(animal["corpo"], animal_x - 18, animal_y - 17, animal_x + 18, animal_y)
            self.canvas.coords(animal["cabeca"], animal_x + 15, animal_y - 20, animal_x + 29, animal_y - 5)
            self.canvas.coords(animal["pernas"], animal_x - 11, animal_y, animal_x - 11, animal_y + 11, animal_x + 10, animal_y, animal_x + 10, animal_y + 11)

        self.canvas.coords(cena["rotulo_acao"], x, y - 66)
        self.canvas.itemconfig(cena["rotulo_acao"], text=texto_acao, state="normal")
        for indice, particula in enumerate(cena["particulas"]):
            if em_acao:
                fase = agora_ts * 7 + indice * 1.7
                px = mao_x + math.cos(fase) * (8 + indice * 2)
                py = mao_y - abs(math.sin(fase)) * (7 + indice * 3)
                self.canvas.coords(particula, px - 3, py - 3, px + 3, py + 3)
                self.canvas.itemconfig(particula, fill="#FFF0B0", state="normal")
            else:
                self.canvas.itemconfig(particula, state="hidden")

        for item in cena["criatura"].values():
            self.canvas.itemconfig(item, state="hidden")
        for fragmento in cena["explosao"]:
            self.canvas.itemconfig(fragmento, state="hidden")
        if cena["ataque_fase"] == "aproximando":
            ax, ay = cena["ataque_x"], cena["ataque_y"]
            escala_ameaca = 0.72 + ((ay - horizonte) / max(1, alt - horizonte)) * 0.34
            criatura = cena["criatura"]
            self.canvas.coords(criatura["cabeca"], ax - 15 * escala_ameaca, ay - 67 * escala_ameaca, ax + 15 * escala_ameaca, ay - 39 * escala_ameaca)
            self.canvas.coords(criatura["corpo"], ax - 13 * escala_ameaca, ay - 40 * escala_ameaca, ax + 13 * escala_ameaca, ay - 13 * escala_ameaca)
            self.canvas.coords(criatura["perna1"], ax - 12 * escala_ameaca, ay - 13 * escala_ameaca, ax - 2 * escala_ameaca, ay + 10 * escala_ameaca)
            self.canvas.coords(criatura["perna2"], ax + 2 * escala_ameaca, ay - 13 * escala_ameaca, ax + 12 * escala_ameaca, ay + 10 * escala_ameaca)
            self.canvas.coords(criatura["olho1"], ax - 9 * escala_ameaca, ay - 60 * escala_ameaca, ax - 4 * escala_ameaca, ay - 54 * escala_ameaca)
            self.canvas.coords(criatura["olho2"], ax + 4 * escala_ameaca, ay - 60 * escala_ameaca, ax + 9 * escala_ameaca, ay - 54 * escala_ameaca)
            self.canvas.coords(criatura["boca"], ax - 5 * escala_ameaca, ay - 52 * escala_ameaca, ax + 5 * escala_ameaca, ay - 45 * escala_ameaca)
            for item in criatura.values():
                self.canvas.itemconfig(item, state="normal")
        elif cena["ataque_fase"] == "explodindo":
            alvo_x, alvo_y = obras[cena["ataque_alvo"]]["alvo"]
            fase_explosao = min(1.0, (agora_ts - cena["ataque_inicio"]) / 1.0)
            raio = 12 + fase_explosao * 58
            for indice, fragmento in enumerate(cena["explosao"]):
                angulo = (math.tau * indice / len(cena["explosao"])) + agora_ts * 2
                px = alvo_x + math.cos(angulo) * raio
                py = alvo_y - 28 + math.sin(angulo) * raio * 0.62
                tamanho = 8 + (indice % 3) * 4
                self.canvas.coords(fragmento, px - tamanho, py - tamanho, px + tamanho, py + tamanho)
                self.canvas.itemconfig(fragmento, state="normal")

        if not cena["evento_ambiente"] and agora_ts >= cena["proximo_evento_ambiente"]:
            cena["evento_ambiente"] = "estrela" if noite else "balao"
            cena["evento_inicio"] = agora_ts
            cena["evento_fim"] = agora_ts + (4.5 if noite else 13.0)
        if cena["evento_ambiente"] and agora_ts >= cena["evento_fim"]:
            cena["evento_ambiente"] = ""
            cena["proximo_evento_ambiente"] = agora_ts + random.uniform(35, 70)

        for item in cena["balao"].values():
            self.canvas.itemconfig(item, state="hidden")
        self.canvas.itemconfig(cena["estrela_cadente"], state="hidden")
        if cena["evento_ambiente"] == "balao":
            fase_balao = (agora_ts - cena["evento_inicio"]) / 13.0
            bx = larg * (-0.08 + fase_balao * 1.16)
            by = horizonte * 0.28 + math.sin(agora_ts * 1.3) * 9
            self.canvas.coords(cena["balao"]["corpo"], bx - 18, by - 26, bx + 18, by + 20)
            self.canvas.coords(cena["balao"]["cesta"], bx - 7, by + 29, bx + 7, by + 39)
            self.canvas.coords(cena["balao"]["cordas"], bx - 10, by + 16, bx - 5, by + 29, bx + 10, by + 16, bx + 5, by + 29)
            for item in cena["balao"].values():
                self.canvas.itemconfig(item, state="normal")
        elif cena["evento_ambiente"] == "estrela":
            fase_estrela = (agora_ts - cena["evento_inicio"]) / 4.5
            sx = larg * (0.18 + fase_estrela * 0.48)
            sy = horizonte * (0.12 + fase_estrela * 0.48)
            self.canvas.coords(cena["estrela_cadente"], sx - 42, sy - 24, sx, sy)
            self.canvas.itemconfig(cena["estrela_cadente"], state="normal")

    def _loop_animacao_blocos(self):
        if not self._animacao_blocos_ativa or not self._blocos_cena:
            return
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        self._atualizar_cena_blocos(datetime.now(), time.time())
        self._animacao_blocos_after = self.after(100, self._loop_animacao_blocos)

    def _parar_animacao_blocos(self):
        self._animacao_blocos_ativa = False
        if self._animacao_blocos_after is not None:
            try:
                self.after_cancel(self._animacao_blocos_after)
            except Exception:
                pass
        self._animacao_blocos_after = None
        self._blocos_cena = {}

    def _desenhar_fundo_voo(self, larg, alt, cor_destaque):
        """Cria um voo automatico sobre um rio, com desvio e abastecimento seguro."""
        return self._desenhar_fundo_voo_pixel(larg, alt, cor_destaque)
        tags = self._tags_fundo_voo
        ceu = self.canvas.create_rectangle(0, 0, larg, alt, fill="#7DB9E8", outline="", tags=tags)
        for indice in range(7):
            y = int(alt * indice / 7)
            self.canvas.create_rectangle(0, y, larg, y + int(alt / 7) + 1, fill=self._mix_cor("#75B8F0", "#D8F0FC", indice / 7), outline="", tags=tags)

        def margem_rio(y):
            centro = larg * (0.5 + math.sin(y / max(1, alt) * math.tau * 0.72) * 0.075)
            largura = larg * (0.17 + 0.28 * y / max(1, alt))
            return centro - largura / 2, centro + largura / 2

        pontos_esq = [margem_rio(y)[0] for y in range(0, alt + 1, max(1, int(alt / 18)))]
        pontos_dir = [margem_rio(y)[1] for y in range(0, alt + 1, max(1, int(alt / 18)))]
        ys = list(range(0, alt + 1, max(1, int(alt / 18))))
        poligono_rio = []
        for x, y in zip(pontos_esq, ys):
            poligono_rio.extend((x, y))
        for x, y in zip(reversed(pontos_dir), reversed(ys)):
            poligono_rio.extend((x, y))
        self.canvas.create_polygon(*poligono_rio, fill="#2D91C4", outline="#C5F1FF", width=3, tags=tags)

        for lado in (0.16, 0.84):
            for indice in range(11):
                x = int(larg * lado + math.sin(indice * 2.4) * larg * 0.05)
                y = int(alt * (0.08 + indice * 0.095))
                self.canvas.create_rectangle(x - 26, y - 20, x + 26, y + 22, fill="#4F9E45", outline="#276233", width=2, tags=tags)
                self.canvas.create_rectangle(x - 7, y + 18, x + 7, y + 40, fill="#755131", outline="#4A3020", width=1, tags=tags)

        ondas = [self.canvas.create_line(0, 0, 0, 0, fill="#B9ECFA", width=2, tags=tags) for _ in range(13)]
        trilhas = [self.canvas.create_line(0, 0, 0, 0, fill="#D8F5FF", width=2, tags=tags) for _ in range(5)]
        aviao = {
            "asa": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#E9F4F7", outline="#173D59", width=2, tags=tags),
            "corpo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F04C48", outline="#67282B", width=2, tags=tags),
            "cabine": self.canvas.create_rectangle(0, 0, 0, 0, fill="#8EE2F5", outline="#24546E", width=1, tags=tags),
            "cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F4D66B", outline="#775A22", width=1, tags=tags),
        }
        obstaculos = []
        for indice in range(9):
            obstaculos.append({
                "tipo": random.choice(("barco", "helicoptero", "balao", "nuvem")),
                "x": 0.0, "y": float(indice * alt / 9), "velocidade": random.uniform(1.5, 3.2),
                "fase": random.uniform(0, math.tau),
                "corpo": self.canvas.create_oval(0, 0, 0, 0, fill="#FFFFFF", outline="#274A61", width=2, tags=tags),
                "detalhe": self.canvas.create_line(0, 0, 0, 0, fill="#274A61", width=2, tags=tags),
                "extra": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F4D86A", outline="#604B22", width=1, tags=tags),
            })
        posto = {
            "ativo": False, "x": larg * 0.5, "y": -100.0,
            "plataforma": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F2A538", outline="#70431E", width=2, state="hidden", tags=tags),
            "bomba": self.canvas.create_rectangle(0, 0, 0, 0, fill="#D94D3F", outline="#5C2824", width=2, state="hidden", tags=tags),
            "letreiro": self.canvas.create_text(0, 0, text="REABASTECER", font=("Segoe UI", 11, "bold"), fill="#FFF5C8", state="hidden", tags=tags),
        }
        hud_y = alt - 86
        tanque = {
            "fundo": self.canvas.create_rectangle(larg / 2 - 160, hud_y, larg / 2 + 160, hud_y + 22, fill="#172832", outline="#DCE9E9", width=2, tags=tags),
            "nivel": self.canvas.create_rectangle(larg / 2 - 156, hud_y + 4, larg / 2 + 156, hud_y + 18, fill="#66D06C", outline="", tags=tags),
            "texto": self.canvas.create_text(larg / 2, hud_y - 7, text="COMBUSTIVEL 100%", font=("Segoe UI", 11, "bold"), fill="#F3F7E8", anchor="s", tags=tags),
            "estado": self.canvas.create_text(larg / 2, hud_y + 38, text="VOO AUTOMATICO", font=("Segoe UI", 9, "bold"), fill="#D6EDF2", anchor="n", tags=tags),
        }
        self._voo_cena = {
            "larg": larg, "alt": alt, "margem_rio": margem_rio, "ondas": ondas, "trilhas": trilhas, "aviao": aviao,
            "obstaculos": obstaculos, "posto": posto, "tanque": tanque, "aviao_x": larg * 0.5, "alvo_x": larg * 0.5,
            "aviao_y": alt * 0.76, "combustivel": 100.0, "reabastecendo": False, "ultimo_tempo": time.time(),
        }
        self._atualizar_cena_voo(time.time())
        self._animacao_voo_ativa = True
        self._animacao_voo_after = self.after(80, self._loop_animacao_voo)

    def _atualizar_cena_voo(self, agora_ts):
        """Atualiza a rota segura, os obstaculos e o abastecimento do voo."""
        cena = self._voo_cena
        if not cena:
            return
        larg, alt = cena["larg"], cena["alt"]
        margem_rio = cena["margem_rio"]
        aviao_y = cena["aviao_y"]
        esquerda, direita = margem_rio(aviao_y)
        centro = (esquerda + direita) / 2
        largura_segura = (direita - esquerda) * 0.34
        dt = min(0.2, max(0.02, agora_ts - cena["ultimo_tempo"]))
        cena["ultimo_tempo"] = agora_ts
        cena["combustivel"] = max(0.0, cena["combustivel"] - dt * 1.55)

        for indice, onda in enumerate(cena["ondas"]):
            y = ((agora_ts * 80 + indice * 98) % (alt + 120)) - 60
            esq, dir_ = margem_rio(y)
            self.canvas.coords(onda, esq + 16, y, dir_ - 16, y)
        for indice, trilha in enumerate(cena["trilhas"]):
            y = aviao_y + 42 + indice * 18
            self.canvas.coords(trilha, cena["aviao_x"] - 8 + indice * 3, y, cena["aviao_x"] + 8 - indice * 3, y + 20)

        alvo_x = centro
        posto = cena["posto"]
        if cena["combustivel"] < 28 and not posto["ativo"]:
            posto["ativo"] = True
            posto["x"] = centro + random.uniform(-largura_segura * 0.45, largura_segura * 0.45)
            posto["y"] = -70.0
        if posto["ativo"]:
            if cena["reabastecendo"]:
                posto["y"] = aviao_y
            else:
                posto["y"] += 7.0
            alvo_x = posto["x"]

        for obstaculo in cena["obstaculos"]:
            obstaculo["y"] += obstaculo["velocidade"] * (dt * 12)
            if obstaculo["y"] > alt + 90:
                obstaculo["tipo"] = random.choice(("barco", "helicoptero", "balao", "nuvem"))
                obstaculo["y"] = random.uniform(-260, -70)
                obstaculo["velocidade"] = random.uniform(1.5, 3.2)
                obstaculo["fase"] = random.uniform(0, math.tau)
            esq, dir_ = margem_rio(obstaculo["y"])
            obstaculo["x"] = (esq + dir_) / 2 + math.sin(agora_ts * 0.75 + obstaculo["fase"]) * (dir_ - esq) * 0.24
            if not posto["ativo"] and abs(obstaculo["y"] - aviao_y) < 100:
                deslocamento = (dir_ - esq) * 0.27
                alvo_x = centro - deslocamento if obstaculo["x"] >= cena["aviao_x"] else centro + deslocamento
            self._desenhar_obstaculo_voo(obstaculo)

        alvo_x = max(esquerda + 34, min(direita - 34, alvo_x))
        cena["alvo_x"] = alvo_x
        cena["aviao_x"] = self._aproximar_valor(cena["aviao_x"], alvo_x, 4.8)
        x, y = cena["aviao_x"], aviao_y
        aviao = cena["aviao"]
        inclinacao = max(-16, min(16, (alvo_x - x) * 0.24))
        self.canvas.coords(aviao["asa"], x - 48, y + 8, x - 12, y - 4, x + 50, y + 8, x + 14, y + 19)
        self.canvas.coords(aviao["corpo"], x - 13, y + 30, x - 13, y - 38, x, y - 57, x + 13, y - 38, x + 13, y + 30)
        self.canvas.coords(aviao["cabine"], x - 7, y - 29, x + 7, y - 12)
        self.canvas.coords(aviao["cauda"], x - 13, y + 18, x - 29 - inclinacao * 0.15, y + 35, x - 13, y + 29)

        if posto["ativo"]:
            px, py = posto["x"], posto["y"]
            self.canvas.coords(posto["plataforma"], px - 68, py - 10, px + 68, py + 12)
            self.canvas.coords(posto["bomba"], px - 13, py - 42, px + 13, py - 10)
            self.canvas.coords(posto["letreiro"], px, py - 58)
            for item in (posto["plataforma"], posto["bomba"], posto["letreiro"]):
                self.canvas.itemconfig(item, state="normal")
            if abs(py - y) < 35 and abs(px - x) < 32:
                cena["reabastecendo"] = True
                posto["y"] = y
                cena["combustivel"] = min(100.0, cena["combustivel"] + dt * 45)
            if cena["reabastecendo"] and cena["combustivel"] >= 99.5:
                posto["ativo"] = False
                cena["reabastecendo"] = False
        else:
            for item in (posto["plataforma"], posto["bomba"], posto["letreiro"]):
                self.canvas.itemconfig(item, state="hidden")

        combustivel = cena["combustivel"]
        cor_tanque = "#62D56B" if combustivel >= 45 else "#F0C451" if combustivel >= 25 else "#F06B4E"
        tanque = cena["tanque"]
        x1, y1, x2, y2 = self.canvas.coords(tanque["fundo"])
        self.canvas.coords(tanque["nivel"], x1 + 4, y1 + 4, x1 + 4 + (x2 - x1 - 8) * combustivel / 100, y2 - 4)
        self.canvas.itemconfig(tanque["nivel"], fill=cor_tanque)
        self.canvas.itemconfig(tanque["texto"], text=f"COMBUSTIVEL {int(combustivel):02d}%")
        estado = "REABASTECENDO" if cena["reabastecendo"] else "ROTA SEGURA"
        self.canvas.itemconfig(tanque["estado"], text=estado, fill=cor_tanque)

    def _desenhar_obstaculo_voo(self, obstaculo):
        """Desenha cada obstaculo com silhueta propria, mantendo o voo sem combate."""
        x, y, tipo = obstaculo["x"], obstaculo["y"], obstaculo["tipo"]
        corpo, detalhe, extra = obstaculo["corpo"], obstaculo["detalhe"], obstaculo["extra"]
        if tipo == "barco":
            self.canvas.coords(corpo, x - 28, y - 8, x + 28, y + 12)
            self.canvas.itemconfig(corpo, fill="#EAE8D7", outline="#263D4C")
            self.canvas.coords(detalhe, x, y - 30, x, y - 6)
            self.canvas.coords(extra, x - 3, y - 30, x + 19, y - 14)
        elif tipo == "helicoptero":
            self.canvas.coords(corpo, x - 22, y - 11, x + 23, y + 11)
            self.canvas.itemconfig(corpo, fill="#E6B74B", outline="#604519")
            self.canvas.coords(detalhe, x - 34, y - 22, x + 35, y - 22)
            self.canvas.coords(extra, x + 20, y - 4, x + 43, y + 3)
        elif tipo == "balao":
            self.canvas.coords(corpo, x - 17, y - 30, x + 17, y + 8)
            self.canvas.itemconfig(corpo, fill="#F06E59", outline="#78382B")
            self.canvas.coords(detalhe, x - 8, y + 8, x - 4, y + 20, x + 8, y + 8, x + 4, y + 20)
            self.canvas.coords(extra, x - 7, y + 20, x + 7, y + 30)
        else:
            self.canvas.coords(corpo, x - 38, y - 16, x + 38, y + 15)
            self.canvas.itemconfig(corpo, fill="#F2FAFF", outline="#B9DCEA")
            self.canvas.coords(detalhe, x - 54, y + 1, x + 54, y + 1)
            self.canvas.coords(extra, x - 9, y - 25, x + 15, y + 8)

    def _desenhar_fundo_voo_pixel(self, larg, alt, cor_destaque):
        """Cena arcade original, limpa e legivel, inspirada em voos verticais classicos."""
        tags = self._tags_fundo_voo
        hud_top = alt - max(135, int(alt * 0.15))
        rio_largo_esq, rio_largo_dir = int(larg * 0.23), int(larg * 0.77)
        rio_esq, rio_dir = int(larg * 0.365), int(larg * 0.635)
        self.canvas.create_rectangle(0, 0, larg, hud_top, fill="#4CB273", outline="", tags=tags)
        def margens_rio(y, deslocamento=0.0):
            # Secoes largas e estreitas percorrem a tela; nenhuma margem fica fixa em funil.
            mundo_y = y - deslocamento
            largura_relativa = 0.34 + 0.12 * ((math.sin(mundo_y / 190) + 1) / 2)
            largura_relativa = round(largura_relativa * 18) / 18
            centro = larg / 2 + math.sin(mundo_y / 430) * larg * 0.025
            return centro - larg * largura_relativa / 2, centro + larg * largura_relativa / 2

        def pontos_do_rio(deslocamento=0.0):
            passos_y = list(range(0, hud_top + 1, max(12, int(hud_top / 22))))
            pontos = []
            for y in passos_y:
                margem_esq, _ = margens_rio(y, deslocamento)
                pontos.extend((margem_esq, y))
            for y in reversed(passos_y):
                _, margem_dir = margens_rio(y, deslocamento)
                pontos.extend((margem_dir, y))
            return pontos

        rio_item = self.canvas.create_polygon(*pontos_do_rio(), fill="#355E9F", outline="", tags=tags)

        # Linhas claras da margem e faixa de estrada antes do painel de instrumentos.
        estrada_esq = self.canvas.create_rectangle(0, hud_top - 7, rio_esq, hud_top, fill="#77797A", outline="", tags=tags)
        estrada_dir = self.canvas.create_rectangle(rio_dir, hud_top - 7, larg, hud_top, fill="#77797A", outline="", tags=tags)
        faixa_esq = self.canvas.create_rectangle(0, hud_top - 15, rio_esq, hud_top - 7, fill="#EDE56A", outline="", tags=tags)
        faixa_dir = self.canvas.create_rectangle(rio_dir, hud_top - 15, larg, hud_top - 7, fill="#EDE56A", outline="", tags=tags)

        ondas = [self.canvas.create_line(0, 0, 0, 0, fill="#7B6BF0", width=2, tags=tags) for _ in range(8)]
        aviao = {
            "sombra": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#203D87", outline="", tags=tags),
            "corpo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F4C63E", outline="#8E6417", width=2, tags=tags),
            "asas": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#FFE96B", outline="#B77A1A", width=2, tags=tags),
            "cabine": self.canvas.create_rectangle(0, 0, 0, 0, fill="#BFEAF0", outline="#39728A", width=1, tags=tags),
            "cauda": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F04D4F", outline="#8C2D32", width=1, tags=tags),
            "motor_esq": self.canvas.create_rectangle(0, 0, 0, 0, fill="#D67B2D", outline="#80421D", width=1, tags=tags),
            "motor_dir": self.canvas.create_rectangle(0, 0, 0, 0, fill="#D67B2D", outline="#80421D", width=1, tags=tags),
        }
        obstaculos = []
        for indice in range(7):
            obstaculos.append({
                "indice": indice,
                "tipo": "resgate" if indice == 1 else random.choice(("barco", "helicoptero", "boia", "ilha")),
                "x": 0.0, "y": float(indice * hud_top / 7 - 40), "vel": random.uniform(2.3, 4.1),
                "fase": random.uniform(0, math.tau), "resgatado": False,
                "corpo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#101018", outline="#101018", tags=tags),
                "detalhe": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F14C57", outline="", tags=tags),
                "extra": self.canvas.create_line(0, 0, 0, 0, fill="#F7D94D", width=3, tags=tags),
                "sombra": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#243D78", outline="", tags=tags),
                "acento": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F8F3CF", outline="#1D314F", width=1, tags=tags),
                "esqui": self.canvas.create_line(0, 0, 0, 0, fill="#1A2630", width=2, tags=tags),
                "pessoa1": self.canvas.create_oval(0, 0, 0, 0, fill="#F1C596", outline="#593A2A", width=1, state="hidden", tags=tags),
                "pessoa2": self.canvas.create_oval(0, 0, 0, 0, fill="#C58455", outline="#593A2A", width=1, state="hidden", tags=tags),
                "boia_resgate": self.canvas.create_oval(0, 0, 0, 0, fill="#F6F0D7", outline="#E04C41", width=3, state="hidden", tags=tags),
                "boia_sinal": self.canvas.create_oval(0, 0, 0, 0, fill="#F58135", outline="#9D3D1E", width=2, state="hidden", tags=tags),
                "haste_boia": self.canvas.create_line(0, 0, 0, 0, fill="#8B4A22", width=3, state="hidden", tags=tags),
                "luz_boia": self.canvas.create_oval(0, 0, 0, 0, fill="#FFE45B", outline="#AD7321", width=1, state="hidden", tags=tags),
            })
        margem_objetos = []
        for indice in range(14):
            margem_objetos.append({
                "lado": -1 if indice % 2 == 0 else 1, "y": float(indice * hud_top / 7 - 45),
                "vel": random.uniform(1.8, 3.4), "tipo": "arvore" if indice % 3 else "casa",
                "x": random.uniform(larg * 0.06, larg * 0.18) if indice % 2 == 0 else random.uniform(larg * 0.82, larg * 0.94),
                "corpo": self.canvas.create_rectangle(0, 0, 0, 0, fill="#287841", outline="#17522D", width=2, tags=tags),
                "topo": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#72B94B", outline="#287841", width=1, tags=tags),
            })
        posto = {
            "ativo": False, "x": larg * 0.5, "y": -80.0,
            "corpo": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F7F5D8", outline="#C2B76B", width=2, state="hidden", tags=tags),
            "faixa1": self.canvas.create_rectangle(0, 0, 0, 0, fill="#F04DA0", outline="", state="hidden", tags=tags),
            "faixa2": self.canvas.create_rectangle(0, 0, 0, 0, fill="#4BC9E9", outline="", state="hidden", tags=tags),
            "texto": self.canvas.create_text(0, 0, text="F\nU\nE\nL", font=("Courier New", 9, "bold"), fill="#3141A9", state="hidden", tags=tags),
        }
        self.canvas.create_rectangle(0, hud_top, larg, alt, fill="#BDBDC2", outline="#77777C", width=2, tags=tags)
        painel_x1, painel_x2 = larg * 0.36, larg * 0.64
        painel_y1, painel_y2 = hud_top + 28, hud_top + 76
        self.canvas.create_rectangle(painel_x1, painel_y1, painel_x2, painel_y2, fill="#F2F1E2", outline="#171720", width=4, tags=tags)
        segmentos = []
        largura_segmento = (painel_x2 - painel_x1 - 42) / 8
        for indice in range(8):
            x1 = painel_x1 + 21 + indice * largura_segmento
            segmentos.append(self.canvas.create_rectangle(x1, painel_y1 + 20, x1 + largura_segmento - 4, painel_y2 - 10, fill="#28323A", outline="", tags=tags))
        tanque = {
            "nivel": self.canvas.create_line(painel_x1 + 20, painel_y1 + 13, painel_x2 - 20, painel_y1 + 13, fill="#171720", width=2, tags=tags),
            "ponteiro": self.canvas.create_polygon(0, 0, 0, 0, 0, 0, fill="#F1C92F", outline="#8F6918", width=1, tags=tags),
            "texto": self.canvas.create_text(larg / 2, hud_top + 18, text="COMBUSTÍVEL", font=("Courier New", 12, "bold"), fill="#25252A", anchor="s", tags=tags),
            "estado": self.canvas.create_text(larg / 2, hud_top + 88, text="ROTA SEGURA", font=("Courier New", 11, "bold"), fill="#4B429E", anchor="n", tags=tags),
            "e": self.canvas.create_text(painel_x1 + 14, painel_y1 + 22, text="E", font=("Courier New", 18, "bold"), fill="#202028", tags=tags),
            "f": self.canvas.create_text(painel_x2 - 14, painel_y1 + 22, text="F", font=("Courier New", 18, "bold"), fill="#202028", tags=tags),
            "x1": painel_x1, "x2": painel_x2, "y1": painel_y1, "y2": painel_y2, "segmentos": segmentos,
        }
        pontuacao_box = self.canvas.create_rectangle(larg * 0.07, hud_top + 24, larg * 0.28, hud_top + 82, fill="#26344A", outline="#F1D54A", width=3, tags=tags)
        pontuacao_id = self.canvas.create_text(larg * 0.175, hud_top + 53, text="RESGATES\n0000", font=("Courier New", 17, "bold"), fill="#FFF4A4", anchor="center", tags=tags)
        cabo_resgate = self.canvas.create_line(0, 0, 0, 0, fill="#F5E8A7", width=2, state="hidden", tags=tags)
        self._voo_cena = {
            "larg": larg, "alt": alt, "hud_top": hud_top, "rio_esq": rio_esq, "rio_dir": rio_dir, "margens_rio": margens_rio, "pontos_do_rio": pontos_do_rio, "rio_item": rio_item,
            "estrada_esq": estrada_esq, "estrada_dir": estrada_dir, "faixa_esq": faixa_esq, "faixa_dir": faixa_dir, "deslocamento_margem": 0.0, "ondas": ondas,
            "aviao": aviao, "obstaculos": obstaculos, "margem_objetos": margem_objetos, "posto": posto, "tanque": tanque,
            "aviao_x": larg * 0.5, "aviao_y": hud_top - 145, "aviao_y_base": hud_top - 145, "combustivel": 100.0,
            "reabastecendo": False, "manobra_ativa": False, "manobra_x": larg * 0.5, "manobra_y": hud_top - 145, "manobra_obstaculo": None, "manobra_inicio": 0.0,
            "pontuacao": 0, "pontuacao_id": pontuacao_id, "pontuacao_box": pontuacao_box, "cabo_resgate": cabo_resgate, "resgate_ativo": None, "resgate_inicio": 0.0, "resgate_alvo": None,
            "ultimo_tempo": time.time(),
        }
        self._atualizar_cena_voo_pixel(time.time())
        self._animacao_voo_ativa = True
        self._animacao_voo_after = self.after(80, self._loop_animacao_voo)

    def _atualizar_cena_voo_pixel(self, agora_ts):
        """Mantem uma partida demonstrativa, sempre com desvio e abastecimento bem-sucedidos."""
        cena = self._voo_cena
        if not cena:
            return
        larg = cena["larg"]
        rio_esq, rio_dir = cena["rio_esq"], cena["rio_dir"]
        margens_rio = cena["margens_rio"]
        aviao_y = cena["aviao_y"]
        dt = min(0.2, max(0.02, agora_ts - cena["ultimo_tempo"]))
        cena["ultimo_tempo"] = agora_ts
        cena["combustivel"] = max(0.0, cena["combustivel"] - dt * 1.45)
        cena["deslocamento_margem"] += dt * 72
        deslocamento_margem = cena["deslocamento_margem"]
        self.canvas.coords(cena["rio_item"], *cena["pontos_do_rio"](deslocamento_margem))
        margem_base_esq, margem_base_dir = margens_rio(cena["hud_top"], deslocamento_margem)
        self.canvas.coords(cena["estrada_esq"], 0, cena["hud_top"] - 7, margem_base_esq, cena["hud_top"])
        self.canvas.coords(cena["estrada_dir"], margem_base_dir, cena["hud_top"] - 7, larg, cena["hud_top"])
        self.canvas.coords(cena["faixa_esq"], 0, cena["hud_top"] - 15, margem_base_esq, cena["hud_top"] - 7)
        self.canvas.coords(cena["faixa_dir"], margem_base_dir, cena["hud_top"] - 15, larg, cena["hud_top"] - 7)
        rio_esq, rio_dir = margens_rio(aviao_y, deslocamento_margem)
        for indice, onda in enumerate(cena["ondas"]):
            y = ((agora_ts * 105 + indice * 110) % (cena["hud_top"] + 100)) - 50
            margem_esq, margem_dir = margens_rio(y, deslocamento_margem)
            self.canvas.coords(onda, margem_esq + 18, y, margem_dir - 18, y)

        # A rota normal se move continuamente nos dois eixos, sem periodos de espera.
        alvo_x = (rio_esq + rio_dir) / 2 + math.sin(agora_ts * 0.47) * (rio_dir - rio_esq) * 0.22
        altura_cruzeiro = cena["aviao_y_base"] + math.sin(agora_ts * 0.72) * 78
        alvo_y = altura_cruzeiro
        posto = cena["posto"]
        if cena["combustivel"] < 30 and not posto["ativo"]:
            posto["ativo"] = True
            posto["x"] = larg / 2 + random.choice((-1, 1)) * larg * 0.075
            posto["y"] = -65.0
        if posto["ativo"]:
            if cena["reabastecendo"]:
                posto["y"] = aviao_y - 46
            else:
                posto["y"] += 6.6
            alvo_x = posto["x"]

        obstaculo_proximo = None
        distancia_proxima = float("inf")
        bote_resgate = None
        distancia_resgate = float("inf")
        for obstaculo in cena["obstaculos"]:
            obstaculo["y"] += obstaculo["vel"] * dt * 11
            if obstaculo["y"] > cena["hud_top"] + 70:
                novo_tipo = "resgate" if random.random() < 0.18 else random.choice(("barco", "helicoptero", "boia", "ilha"))
                obstaculo.update(tipo=novo_tipo, y=random.uniform(-230, -50), vel=random.uniform(2.3, 4.1), fase=random.uniform(0, math.tau), resgatado=False)
            margem_esq, margem_dir = margens_rio(obstaculo["y"], deslocamento_margem)
            obstaculo["x"] = (margem_esq + margem_dir) / 2 + math.sin(agora_ts * (0.5 + obstaculo["vel"] * 0.08) + obstaculo["fase"]) * (margem_dir - margem_esq) * 0.27
            distancia = abs(obstaculo["y"] - aviao_y)
            if obstaculo["tipo"] == "resgate" and not obstaculo["resgatado"] and distancia < distancia_resgate:
                bote_resgate, distancia_resgate = obstaculo, distancia
            elif distancia < distancia_proxima:
                obstaculo_proximo, distancia_proxima = obstaculo, distancia
            self._desenhar_obstaculo_voo_pixel(obstaculo)

        buscando_resgate = not posto["ativo"] and bote_resgate is not None and distancia_resgate < 460
        if buscando_resgate:
            cena["resgate_alvo"] = bote_resgate["indice"]
            cena["manobra_ativa"] = False
            alvo_x = bote_resgate["x"]
            alvo_y = min(cena["hud_top"] - 58, bote_resgate["y"] + 96)
            if abs(bote_resgate["x"] - cena["aviao_x"]) < 70 and abs(bote_resgate["y"] - aviao_y) < 145:
                bote_resgate["resgatado"] = True
                cena["pontuacao"] += 100
                cena["resgate_ativo"] = bote_resgate["indice"]
                cena["resgate_inicio"] = agora_ts
                cena["resgate_alvo"] = None

        ameaca_antecipada = not buscando_resgate and not posto["ativo"] and obstaculo_proximo and distancia_proxima < 235
        if ameaca_antecipada and (not cena["manobra_ativa"] or cena["manobra_obstaculo"] != obstaculo_proximo["indice"]):
            cena["manobra_ativa"] = True
            cena["manobra_obstaculo"] = obstaculo_proximo["indice"]
            cena["manobra_inicio"] = agora_ts
            if obstaculo_proximo["x"] >= cena["aviao_x"]:
                cena["manobra_x"] = rio_esq + (rio_dir - rio_esq) * 0.13
            else:
                cena["manobra_x"] = rio_esq + (rio_dir - rio_esq) * 0.87
            if obstaculo_proximo["y"] < aviao_y:
                cena["manobra_y"] = min(cena["hud_top"] - 32, altura_cruzeiro + 92)
            else:
                cena["manobra_y"] = max(int(cena["hud_top"] * 0.48), altura_cruzeiro - 92)

        if cena["manobra_ativa"]:
            alvo_x, alvo_y = cena["manobra_x"], cena["manobra_y"]
            obstaculo_da_manobra = next((item for item in cena["obstaculos"] if item["indice"] == cena["manobra_obstaculo"]), None)
            if (not obstaculo_da_manobra or obstaculo_da_manobra["y"] > aviao_y + 125
                    or agora_ts - cena["manobra_inicio"] > 5.5):
                cena["manobra_ativa"] = False
                cena["manobra_obstaculo"] = None
        desvio_urgente = cena["manobra_ativa"]

        for objeto_margem in cena["margem_objetos"]:
            objeto_margem["y"] += objeto_margem["vel"] * dt * 13
            if objeto_margem["y"] > cena["hud_top"] + 45:
                objeto_margem["y"] = random.uniform(-180, -35)
                objeto_margem["vel"] = random.uniform(1.8, 3.4)
                objeto_margem["tipo"] = "arvore" if random.random() < 0.7 else "casa"
            self._desenhar_objeto_margem_voo(objeto_margem)

        alvo_x = max(rio_esq + 32, min(rio_dir - 32, alvo_x))
        cena["aviao_x"] = self._aproximar_valor(cena["aviao_x"], alvo_x, 11.5 if desvio_urgente else 5.5)
        cena["aviao_y"] = self._aproximar_valor(cena["aviao_y"], alvo_y, 8.0 if desvio_urgente else 3.0)
        x = cena["aviao_x"]
        aviao_y = cena["aviao_y"]
        aviao = cena["aviao"]
        inclinacao = max(-15, min(15, (alvo_x - x) * 0.13))
        self.canvas.coords(aviao["sombra"], x - 33, aviao_y + 27, x - 9, aviao_y + 33, x + 34, aviao_y + 27, x + 12, aviao_y + 35, x - 12, aviao_y + 35)
        self.canvas.tag_lower(aviao["sombra"], aviao["asas"])
        self.canvas.coords(aviao["asas"], x - 38, aviao_y + 9 + inclinacao, x - 13, aviao_y - 2 + inclinacao * 0.45, x - 9, aviao_y - 14, x + 9, aviao_y - 14, x + 13, aviao_y - 2 - inclinacao * 0.45, x + 38, aviao_y + 9 - inclinacao, x + 13, aviao_y + 16 - inclinacao * 0.3, x - 13, aviao_y + 16 + inclinacao * 0.3)
        self.canvas.coords(aviao["corpo"], x - 8, aviao_y + 25, x - 8, aviao_y - 24, x, aviao_y - 39, x + 8, aviao_y - 24, x + 8, aviao_y + 25)
        self.canvas.coords(aviao["cabine"], x - 5, aviao_y - 25, x + 5, aviao_y - 10)
        self.canvas.coords(aviao["cauda"], x - 8, aviao_y + 17, x - 18, aviao_y + 31, x + 18, aviao_y + 31, x + 8, aviao_y + 17)
        self.canvas.coords(aviao["motor_esq"], x - 28, aviao_y + 4 + inclinacao * 0.45, x - 15, aviao_y + 14 + inclinacao * 0.25)
        self.canvas.coords(aviao["motor_dir"], x + 15, aviao_y + 4 - inclinacao * 0.45, x + 28, aviao_y + 14 - inclinacao * 0.25)

        resgate_em_curso = next((item for item in cena["obstaculos"] if item["indice"] == cena["resgate_ativo"]), None)
        if resgate_em_curso and agora_ts - cena["resgate_inicio"] < 1.8:
            self.canvas.coords(cena["cabo_resgate"], x, aviao_y + 22, resgate_em_curso["x"], resgate_em_curso["y"] - 10)
            self.canvas.itemconfig(cena["cabo_resgate"], state="normal")
        else:
            self.canvas.itemconfig(cena["cabo_resgate"], state="hidden")
            if cena["resgate_ativo"] is not None:
                cena["resgate_ativo"] = None
        self.canvas.itemconfig(cena["pontuacao_id"], text=f"RESGATES\n{cena['pontuacao']:04d}")

        if posto["ativo"]:
            px, py = posto["x"], posto["y"]
            self.canvas.coords(posto["corpo"], px - 20, py - 32, px + 20, py + 32)
            self.canvas.coords(posto["faixa1"], px - 20, py - 19, px + 20, py - 7)
            self.canvas.coords(posto["faixa2"], px - 20, py + 7, px + 20, py + 19)
            self.canvas.coords(posto["texto"], px, py)
            for item in (posto["corpo"], posto["faixa1"], posto["faixa2"], posto["texto"]):
                self.canvas.itemconfig(item, state="normal")
            if abs(px - x) < 26 and abs(py - (aviao_y - 46)) < 33:
                cena["reabastecendo"] = True
                cena["combustivel"] = min(100.0, cena["combustivel"] + dt * 48)
            if cena["reabastecendo"] and cena["combustivel"] >= 99.5:
                posto["ativo"] = False
                cena["reabastecendo"] = False
        else:
            for item in (posto["corpo"], posto["faixa1"], posto["faixa2"], posto["texto"]):
                self.canvas.itemconfig(item, state="hidden")

        tanque, combustivel = cena["tanque"], cena["combustivel"]
        segmentos_ativos = max(0, min(8, math.ceil(combustivel / 12.5)))
        for indice, segmento in enumerate(tanque["segmentos"]):
            if indice < segmentos_ativos:
                cor_segmento = "#5DAF4E" if indice >= 3 else "#F0CA3A"
                self.canvas.itemconfig(segmento, fill=cor_segmento)
            else:
                self.canvas.itemconfig(segmento, fill="#34383C")
        ponteiro_x = tanque["x1"] + 21 + (tanque["x2"] - tanque["x1"] - 42) * combustivel / 100
        self.canvas.coords(tanque["ponteiro"], ponteiro_x - 7, tanque["y1"] + 4, ponteiro_x + 7, tanque["y1"] + 4, ponteiro_x, tanque["y1"] + 16)
        estado = "REABASTECENDO" if cena["reabastecendo"] else "ATENCAO: COMBUSTIVEL" if combustivel < 30 else "ROTA SEGURA"
        cor = "#ED4A4A" if combustivel < 30 else "#4B429E"
        self.canvas.itemconfig(tanque["estado"], text=estado, fill=cor)

    def _desenhar_obstaculo_voo_pixel(self, obstaculo):
        """Formas grandes em pixel art para leitura a distancia."""
        x, y, tipo = obstaculo["x"], obstaculo["y"], obstaculo["tipo"]
        corpo, detalhe, extra = obstaculo["corpo"], obstaculo["detalhe"], obstaculo["extra"]
        sombra, acento, esqui = obstaculo["sombra"], obstaculo["acento"], obstaculo["esqui"]
        pessoa1, pessoa2 = obstaculo["pessoa1"], obstaculo["pessoa2"]
        boia_resgate = obstaculo["boia_resgate"]
        boia_sinal = obstaculo["boia_sinal"]
        haste_boia, luz_boia = obstaculo["haste_boia"], obstaculo["luz_boia"]
        self.canvas.tag_lower(sombra, corpo)
        self.canvas.itemconfig(esqui, state="normal" if tipo == "helicoptero" else "hidden")
        self.canvas.itemconfig(pessoa1, state="hidden")
        self.canvas.itemconfig(pessoa2, state="hidden")
        self.canvas.itemconfig(boia_resgate, state="hidden")
        self.canvas.itemconfig(boia_sinal, state="hidden")
        self.canvas.itemconfig(haste_boia, state="hidden")
        self.canvas.itemconfig(luz_boia, state="hidden")
        if tipo == "resgate":
            self.canvas.coords(sombra, x - 55, y + 18, x + 55, y + 18, x + 40, y + 27, x - 40, y + 27)
            self.canvas.itemconfig(sombra, fill="#294A92")
            self.canvas.coords(corpo, x - 52, y + 4, x - 37, y - 12, x + 37, y - 12, x + 52, y + 4, x + 35, y + 19, x - 35, y + 19)
            self.canvas.itemconfig(corpo, fill="#F07A3E", outline="#8A3824")
            self.canvas.coords(detalhe, x - 39, y + 2, x - 27, y - 7, x + 27, y - 7, x + 39, y + 2, x + 25, y + 10, x - 25, y + 10)
            self.canvas.itemconfig(detalhe, fill="#F7E9B1")
            self.canvas.coords(extra, x - 48, y + 15, x + 48, y + 15)
            self.canvas.coords(acento, x - 5, y - 31, x + 5, y - 20)
            self.canvas.itemconfig(acento, fill="#E34D45")
            if not obstaculo["resgatado"]:
                self.canvas.coords(pessoa1, x - 25, y - 22, x - 11, y - 8)
                self.canvas.coords(pessoa2, x + 11, y - 22, x + 25, y - 8)
                self.canvas.coords(boia_resgate, x - 8, y - 7, x + 8, y + 9)
                self.canvas.itemconfig(pessoa1, state="normal")
                self.canvas.itemconfig(pessoa2, state="normal")
                self.canvas.itemconfig(boia_resgate, state="normal")
        elif tipo == "barco":
            self.canvas.coords(sombra, x - 43, y + 14, x + 42, y + 14, x + 30, y + 21, x - 30, y + 21)
            self.canvas.itemconfig(sombra, fill="#294A92")
            self.canvas.coords(corpo, x - 40, y + 6, x - 29, y - 5, x + 27, y - 5, x + 40, y + 6, x + 26, y + 14, x - 27, y + 14)
            self.canvas.itemconfig(corpo, fill="#3A2520", outline="#1B1516")
            self.canvas.coords(detalhe, x - 17, y - 18, x + 12, y - 18, x + 21, y - 5, x - 23, y - 5)
            self.canvas.itemconfig(detalhe, fill="#F0E9C9")
            self.canvas.coords(extra, x + 2, y - 31, x + 2, y - 18)
            self.canvas.coords(acento, x - 10, y - 15, x + 7, y - 7)
            self.canvas.itemconfig(acento, fill="#72C7E2")
        elif tipo == "helicoptero":
            # Perfil completo: fuselagem, cabine, cauda, rotor e esquis.
            self.canvas.coords(sombra, x - 38, y - 3, x - 8, y - 5, x - 3, y + 3, x - 37, y + 6)
            self.canvas.itemconfig(sombra, fill="#1D3F36")
            self.canvas.coords(corpo, x - 9, y - 15, x + 18, y - 15, x + 30, y - 5, x + 22, y + 10, x - 12, y + 10, x - 21, y + 1)
            self.canvas.itemconfig(corpo, fill="#3D8B59", outline="#173E2A")
            self.canvas.coords(detalhe, x + 5, y - 12, x + 19, y - 12, x + 24, y - 4, x + 14, y + 4, x + 3, y + 1)
            self.canvas.itemconfig(detalhe, fill="#9DE2D2")
            self.canvas.coords(extra, x - 31, y - 27, x + 35, y - 27)
            self.canvas.coords(acento, x - 42, y - 9, x - 34, y - 1)
            self.canvas.itemconfig(acento, fill="#E8D454")
            self.canvas.coords(esqui, x - 10, y + 15, x + 23, y + 15)
        elif tipo == "ilha":
            self.canvas.coords(sombra, x - 40, y + 16, x + 40, y + 16, x + 27, y + 23, x - 28, y + 23)
            self.canvas.itemconfig(sombra, fill="#294A92")
            self.canvas.coords(corpo, x - 35, y + 12, x - 23, y - 3, x + 20, y - 7, x + 38, y + 9, x + 25, y + 17, x - 27, y + 17)
            self.canvas.itemconfig(corpo, fill="#D7B65A", outline="#8A6A2C")
            self.canvas.coords(detalhe, x - 16, y - 24, x, y - 35, x + 16, y - 24, x + 10, y - 7, x - 10, y - 7)
            self.canvas.itemconfig(detalhe, fill="#2A7A42")
            self.canvas.coords(extra, x, y - 36, x, y - 8)
            self.canvas.coords(acento, x - 4, y - 30, x + 4, y - 9)
            self.canvas.itemconfig(acento, fill="#684426")
        else:
            # Marcador nautico: base flutuante, corpo cilindrico, luz e reflexo na agua.
            self.canvas.coords(sombra, x - 30, y + 19, x + 30, y + 19, x + 20, y + 26, x - 20, y + 26)
            self.canvas.itemconfig(sombra, fill="#294A92")
            self.canvas.coords(boia_sinal, x - 18, y - 2, x + 18, y + 22)
            self.canvas.itemconfig(boia_sinal, state="normal")
            self.canvas.coords(corpo, x - 11, y - 16, x + 11, y - 16, x + 14, y - 3, x + 11, y + 16, x - 11, y + 16, x - 14, y - 3)
            self.canvas.itemconfig(corpo, fill="#F8F5D8", outline="#9D3D1E")
            self.canvas.coords(detalhe, x - 13, y - 3, x + 13, y - 3, x + 13, y + 6, x - 13, y + 6)
            self.canvas.itemconfig(detalhe, fill="#E65343")
            self.canvas.coords(acento, x - 4, y - 14, x + 4, y - 5)
            self.canvas.itemconfig(acento, fill="#2F84B3")
            self.canvas.coords(haste_boia, x, y - 26, x, y - 16)
            self.canvas.coords(luz_boia, x - 5, y - 32, x + 5, y - 22)
            self.canvas.coords(extra, x - 32, y + 27, x + 32, y + 27, x - 20, y + 32, x + 20, y + 32)
            self.canvas.itemconfig(haste_boia, state="normal")
            self.canvas.itemconfig(luz_boia, state="normal")

    def _desenhar_objeto_margem_voo(self, objeto):
        """Desenha referencias laterais para reforcar a sensacao de deslocamento."""
        x, y = objeto["x"], objeto["y"]
        if objeto["tipo"] == "casa":
            self.canvas.coords(objeto["corpo"], x - 16, y - 10, x + 16, y + 16)
            self.canvas.itemconfig(objeto["corpo"], fill="#D9C486", outline="#6A5734")
            self.canvas.coords(objeto["topo"], x - 21, y - 10, x, y - 28, x + 21, y - 10)
            self.canvas.itemconfig(objeto["topo"], fill="#D95147", outline="#7B342F")
        else:
            self.canvas.coords(objeto["corpo"], x - 5, y - 7, x + 5, y + 21)
            self.canvas.itemconfig(objeto["corpo"], fill="#6C4728", outline="#4A2E1B")
            self.canvas.coords(objeto["topo"], x, y - 31, x - 20, y - 5, x - 10, y + 1, x, y - 11, x + 10, y + 1, x + 20, y - 5)
            self.canvas.itemconfig(objeto["topo"], fill="#247A3D", outline="#15522A")

    def _loop_animacao_voo(self):
        if not self._animacao_voo_ativa or not self._voo_cena:
            return
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        self._atualizar_cena_voo_pixel(time.time())
        self._animacao_voo_after = self.after(80, self._loop_animacao_voo)

    def _parar_animacao_voo(self):
        self._animacao_voo_ativa = False
        if self._animacao_voo_after is not None:
            try:
                self.after_cancel(self._animacao_voo_after)
            except Exception:
                pass
        self._animacao_voo_after = None
        self._voo_cena = {}

    def _desenhar_painel_central(self, larg, alt):
        """Painel único que contém relógio + todas as informações do evento."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")
        cores_texto = self.config_dados.get("cores_texto", {})
        cor_evento = cores_texto.get("evento", "#FB1874")
        cor_data = cores_texto.get("data", "#FB1874")
        cor_professores = cores_texto.get("professores", "#FB1874")
        cor_mensagem = cores_texto.get("mensagem", "#E6EEF5")
        escala_texto = min(1.22, max(1.08, larg / 1600))

        # Painel: 78% da largura, 78% da altura, centralizado
        painel_larg = min(960, int(larg * 0.78))
        painel_alt  = int(alt  * 0.78)
        px = (larg - painel_larg) // 2
        py = (alt  - painel_alt)  // 2
        cx = larg // 2

        painel_transparente = bool(self.fundo_path)

        if not painel_transparente:
            # Sombra
            self._retangulo_arredondado(
                px+8, py+8, px+painel_larg+8, py+painel_alt+8,
                raio=18, fill="#000000", alpha_simulado=True)
            # Fundo escuro
            self._retangulo_arredondado(
                px, py, px+painel_larg, py+painel_alt,
                raio=18, fill="#0A1628")

        if not painel_transparente:
            # Borda
            self._retangulo_arredondado(
                px, py, px+painel_larg, py+painel_alt,
                raio=18, fill="", outline=cor_destaque, espessura=2)
            # Faixa no topo
            self.canvas.create_rectangle(
                px+18, py, px+painel_larg-18, py+5,
                fill=cor_destaque, outline="")

        # Logo (canto superior direito do painel)
        logo_path_rel = self.config_dados.get("logo", "")
        if PIL_DISPONIVEL and logo_path_rel:
            logo_path = BASE_DIR / logo_path_rel
            if logo_path.exists():
                try:
                    logo_pil = Image.open(logo_path).convert("RGBA")
                    logo_h = max(120, int(alt * 0.20))
                    logo_w = int(logo_pil.width * logo_h / logo_pil.height)
                    logo_pil = logo_pil.resize((logo_w, logo_h), Image.LANCZOS)
                    self._img_logo = ImageTk.PhotoImage(logo_pil)
                    self.canvas.create_image(
                        px + painel_larg - 20, py + 20,
                        image=self._img_logo, anchor="ne")
                except Exception as e:
                    print(f"[AVISO] Logo: {e}")

        # ── RELÓGIO ─────────────────────────────────────────────────────
        tam_clock_padrao = max(56, min(84, int(alt * 0.095 * escala_texto)))
        y_clock_padrao = py + int(painel_alt * 0.20)
        relogio_compacto = self.config_dados.get("posicao_relogio", "Central (grande)") == "Inferior direito (compacto)"
        if relogio_compacto:
            # Ainda discreto, mas com leitura confortavel em projetores e TVs.
            tam_clock = max(34, min(52, int(alt * 0.050)))
            x_clock = larg - max(150, int(larg * 0.12))
            y_clock = alt - max(112, int(alt * 0.125))
        else:
            tam_clock = tam_clock_padrao
            x_clock = cx
            y_clock = y_clock_padrao

        self.canvas.create_text(x_clock+6, y_clock+6,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#000000", anchor="center", tags=("clock_shadow", "intermitencia_hora"))
        self.canvas.create_text(x_clock+2, y_clock+2,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#2ECC71", anchor="center", tags=("clock_glow", "intermitencia_hora"))
        self.canvas.create_text(x_clock, y_clock,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#7DFF9B", anchor="center", tags=("clock_main", "intermitencia_hora"))

        # ── INFORMAÇÕES ─────────────────────────────────────────────────
        # Começa abaixo do centro do relógio + espaço proporcional ao tamanho da fonte
        y = y_clock_padrao + int(tam_clock_padrao * 0.82)
        tam_evento = int(30 * escala_texto)
        tam_espaco = int(22 * escala_texto)
        tam_data = int(18 * escala_texto)
        tam_prof = int(20 * escala_texto)
        tam_msg = int(19 * escala_texto)

        # Separador
        self._linha_decorativa(cx, y, int(painel_larg * 0.5), cor_destaque)
        y += int(20 * escala_texto)

        # Nome do evento
        evento = self.config_dados.get("evento", "Evento CIEBP")
        self.canvas.create_text(cx+2, y+2, text=evento.upper(),
            font=("Segoe UI", tam_evento, "bold"), fill="#000000", anchor="n", tags="intermitencia_evento")
        self.canvas.create_text(cx, y, text=evento.upper(),
            font=("Segoe UI", tam_evento, "bold"), fill=cor_evento, anchor="n", tags="intermitencia_evento")
        y += int(tam_evento * 1.5)

        # Nome do espaço
        self.canvas.create_text(cx+2, y+2, text=self.espaco_atual,
            font=("Segoe UI", tam_espaco, "bold"), fill="#000000", anchor="n", tags="intermitencia_legenda")
        self.canvas.create_text(cx, y, text=self.espaco_atual,
            font=("Segoe UI", tam_espaco, "bold"), fill=cor_destaque, anchor="n", tags="intermitencia_legenda")
        y += int(tam_espaco * 1.45)

        # Data
        data = self.config_dados.get("data", datetime.now().strftime("%d/%m/%Y"))
        self.canvas.create_text(cx+1, y+1, text=f"\U0001f4c5  {data}",
            font=("Segoe UI", tam_data, "bold"), fill="#000000", anchor="n", tags="intermitencia_data")
        self.canvas.create_text(cx, y, text=f"\U0001f4c5  {data}",
            font=("Segoe UI", tam_data, "bold"), fill=cor_data, anchor="n", tags="intermitencia_data")
        y += int(tam_data * 1.7)

        # Professores
        professores = self.config_dados.get("professores", [])
        if professores:
            texto_prof = "  ·  ".join(p.strip() for p in professores if p.strip())
            fonte_prof = tkfont.Font(family="Segoe UI", size=tam_prof, weight="bold")
            linhas_prof = self._quebrar_texto(texto_prof, fonte_prof, painel_larg - 120)
            linha_gap = max(8, int(tam_prof * 0.35))
            altura_linha = int(tam_prof * 1.2)

            for indice, linha_prof in enumerate(linhas_prof):
                texto_linha = linha_prof
                if indice == 0:
                    texto_linha = f"Professores: {linha_prof}"

                self.canvas.create_text(
                    cx + 1, y + 1,
                    text=texto_linha,
                    font=("Segoe UI", tam_prof, "bold"),
                    fill="#000000",
                    anchor="n",
                    tags="intermitencia_professores",
                )
                self.canvas.create_text(
                    cx, y,
                    text=texto_linha,
                    font=("Segoe UI", tam_prof, "bold"),
                    fill=cor_professores,
                    anchor="n",
                    tags="intermitencia_professores",
                )
                y += altura_linha
                if indice < len(linhas_prof) - 1:
                    y += linha_gap

            y += int(12 * escala_texto)

        # Separador
        self._linha_decorativa(cx, y, int(painel_larg * 0.4), cor_destaque, espessura=1)
        y += int(18 * escala_texto)
        if self.fundo_path == "__pescador__":
            y += int(60 * escala_texto) + 80

        # Mensagem
        mensagem = self.config_dados.get("mensagem_boas_vindas", "Bem-vindo(a) ao CIEBP!")
        sombra_msg = self.canvas.create_text(cx+2, y+2, text=mensagem,
            font=("Segoe UI", tam_msg), fill="#000000",
            anchor="n", width=painel_larg - 50, tags="intermitencia_mensagem")
        texto_msg_id = self.canvas.create_text(cx, y, text=mensagem,
            font=("Segoe UI", tam_msg), fill=cor_mensagem,
            anchor="n", width=painel_larg - 50, tags=("texto_mensagem", "intermitencia_mensagem"))

    def _desenhar_rodape(self, larg, alt):
        """Desenha o rodapé com o nome institucional."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")
        tam_rodape = max(12, min(16, int(larg / 150)))

        # Faixa colorida no rodapé
        self.canvas.create_rectangle(0, alt - 4, larg, alt, fill=cor_destaque, outline="")

        # Texto institucional
        self.canvas.create_text(
            larg // 2,
            alt - 28,
            text="CIEBP  ·  Centro de Inovação da Escola Básica Paulista",
            font=("Segoe UI", tam_rodape),
            fill="#8EA4B8",
            anchor="center"
        )

    def _desenhar_info_radio(self, larg, alt):
        """
        Exibe discretamente no canto inferior esquerdo
        o nome da rádio e o nível de volume.
        """
        tam_info = max(11, min(15, int(larg / 155)))
        if self.radio_nome == "Sem rádio":
            texto_radio = "♪  Sem rádio"
        else:
            texto_radio = f"♪  {self.radio_nome}"

        if self.mudo:
            texto_vol = "🔇  Mudo"
        else:
            texto_vol = f"🔊  Volume: {self.volume}%"

        self.canvas.create_text(
            22, alt - 48,
            text=texto_radio,
            font=("Segoe UI", tam_info),
            fill="#8EA4B8",
            anchor="w",
            tags="info_radio"
        )
        self.canvas.create_text(
            22, alt - 28,
            text=texto_vol,
            font=("Segoe UI", tam_info),
            fill="#8EA4B8",
            anchor="w",
            tags="info_volume"
        )

    # ------------------------------------------------------------------
    # ELEMENTOS GRÁFICOS AUXILIARES
    # ------------------------------------------------------------------

    def _retangulo_arredondado(self, x1, y1, x2, y2, raio=15,
                                fill="#0A1628", outline="", espessura=1,
                                alpha_simulado=False, tags=""):
        """
        Desenha um retângulo com cantos arredondados no canvas.
        Tkinter não possui suporte nativo, então simulamos com
        arcos e retângulos sobrepostos.
        """
        if alpha_simulado:
            fill = "#040B14"

        if fill:
            # Corpo central
            self.canvas.create_rectangle(x1 + raio, y1, x2 - raio, y2,
                                          fill=fill, outline="", tags=tags)
            self.canvas.create_rectangle(x1, y1 + raio, x2, y2 - raio,
                                          fill=fill, outline="", tags=tags)
        if fill:
            # Cantos arredondados
            self.canvas.create_arc(x1, y1, x1 + 2*raio, y1 + 2*raio,
                                    start=90, extent=90, fill=fill, outline="", tags=tags)
            self.canvas.create_arc(x2 - 2*raio, y1, x2, y1 + 2*raio,
                                    start=0, extent=90, fill=fill, outline="", tags=tags)
            self.canvas.create_arc(x1, y2 - 2*raio, x1 + 2*raio, y2,
                                    start=180, extent=90, fill=fill, outline="", tags=tags)
            self.canvas.create_arc(x2 - 2*raio, y2 - 2*raio, x2, y2,
                                    start=270, extent=90, fill=fill, outline="", tags=tags)

        if outline:
            # Bordas arredondadas
            self.canvas.create_arc(x1, y1, x1 + 2*raio, y1 + 2*raio,
                                    start=90, extent=90,
                                    outline=outline, style="arc", width=espessura, tags=tags)
            self.canvas.create_arc(x2 - 2*raio, y1, x2, y1 + 2*raio,
                                    start=0, extent=90,
                                    outline=outline, style="arc", width=espessura, tags=tags)
            self.canvas.create_arc(x1, y2 - 2*raio, x1 + 2*raio, y2,
                                    start=180, extent=90,
                                    outline=outline, style="arc", width=espessura, tags=tags)
            self.canvas.create_arc(x2 - 2*raio, y2 - 2*raio, x2, y2,
                                    start=270, extent=90,
                                    outline=outline, style="arc", width=espessura, tags=tags)
            # Linhas retas das bordas
            self.canvas.create_line(x1 + raio, y1, x2 - raio, y1,
                                     fill=outline, width=espessura, tags=tags)
            self.canvas.create_line(x1 + raio, y2, x2 - raio, y2,
                                     fill=outline, width=espessura, tags=tags)
            self.canvas.create_line(x1, y1 + raio, x1, y2 - raio,
                                     fill=outline, width=espessura, tags=tags)
            self.canvas.create_line(x2, y1 + raio, x2, y2 - raio,
                                     fill=outline, width=espessura, tags=tags)

    def _linha_decorativa(self, cx, cy, metade_larg, cor, espessura=2):
        """Desenha uma linha horizontal decorativa centrada em (cx, cy)."""
        self.canvas.create_line(
            cx - metade_larg, cy,
            cx + metade_larg, cy,
            fill=cor, width=espessura, dash=(4, 6)
        )

    def _quebrar_texto(self, texto, fonte, largura_max):
        """Quebra o texto em linhas com base na largura máxima medida pela fonte."""
        palavras = texto.split()
        if not palavras:
            return [""]

        linhas = []
        linha_atual = palavras[0]

        for palavra in palavras[1:]:
            candidato = f"{linha_atual} {palavra}"
            if fonte.measure(candidato) <= largura_max:
                linha_atual = candidato
            else:
                linhas.append(linha_atual)
                linha_atual = palavra

        linhas.append(linha_atual)
        return linhas

    def _obter_alerta_horario(self, agora):
        """Retorna os dados do alerta se o horário estiver em uma janela de aviso."""
        janelas = [
            {
                "id": "almoco",
                "hora": 12,
                "minuto": 0,
                "titulo": "Almoço em 5 minutos",
                "mensagem": "Vamos nos organizar para a pausa do meio-dia.",
            },
            {
                "id": "saida_1",
                "hora": 17,
                "minuto": 0,
                "titulo": "Saída do 1º professor em 5 minutos",
                "mensagem": "Encaminhe os últimos avisos para o encerramento do primeiro turno.",
            },
            {
                "id": "saida_2",
                "hora": 18,
                "minuto": 0,
                "titulo": "Saída do 2º professor em 5 minutos",
                "mensagem": "Estamos nos aproximando do fechamento das atividades do segundo professor.",
            },
        ]

        if self.alerta_teste:
            for janela in janelas:
                if janela["id"] == self.alerta_teste:
                    return {
                        **janela,
                        "minutos_restantes": 5,
                        "modo_teste": True,
                    }

        minutos_agora = agora.hour * 60 + agora.minute
        for janela in janelas:
            alvo = janela["hora"] * 60 + janela["minuto"]
            inicio = alvo - 5
            if inicio <= minutos_agora < alvo:
                restante = alvo - minutos_agora
                return {
                    **janela,
                    "minutos_restantes": restante,
                    "modo_teste": False,
                }
        return None

    def _iniciar_animacao_alerta(self, alerta):
        """Ativa o ciclo de confetes e foguetes do aviso."""
        self._animacao_alerta_ativa = True
        self._ultimo_alerta_animado = alerta["id"]
        self._confetes = []
        self._foguetes = []
        self._estouros = []
        self._ultimo_confete = 0.0
        self._ultimo_foguete = 0.0
        self._contador_foguetes = 0
        if self._animacao_alerta_after:
            try:
                self.after_cancel(self._animacao_alerta_after)
            except Exception:
                pass
            self._animacao_alerta_after = None
        self._loop_animacao_alerta()

    def _parar_animacao_alerta(self):
        """Encerra a animação do aviso e limpa os elementos do canvas."""
        self._animacao_alerta_ativa = False
        self._ultimo_alerta_animado = None
        self._confetes = []
        self._foguetes = []
        self._estouros = []
        if self._animacao_alerta_after:
            try:
                self.after_cancel(self._animacao_alerta_after)
            except Exception:
                pass
            self._animacao_alerta_after = None
        try:
            self.canvas.delete(self._efeitos_tags)
        except Exception:
            pass

    def _sincronizar_animacao_alerta(self, alerta):
        """Mantém a animação alinhada ao alerta ativo."""
        if not alerta:
            if self._animacao_alerta_ativa:
                self._parar_animacao_alerta()
            return

        if (not self._animacao_alerta_ativa) or (self._ultimo_alerta_animado != alerta["id"]):
            self._iniciar_animacao_alerta(alerta)

    def _criar_confete(self, larg):
        cores = ["#FFEB3B", "#00E5FF", "#FF6D00", "#FF4081", "#69F0AE", "#7C4DFF"]
        return {
            "x": random.randint(30, max(31, larg - 30)),
            "y": random.randint(-80, -10),
            "vx": random.uniform(-1.6, 1.6),
            "vy": random.uniform(3.2, 5.6),
            "tam": random.randint(6, 12),
            "cor": random.choice(cores),
        }

    def _criar_foguete(self, larg, alt):
        esquerda = (self._contador_foguetes % 2) == 0
        origem_x = int(larg * (0.14 if esquerda else 0.86))
        alvo_x = random.randint(int(larg * 0.22), int(larg * 0.78))
        alvo_y = random.randint(int(alt * 0.12), int(alt * 0.32))
        return {
            "x": origem_x,
            "y": alt + 20,
            "vx": (alvo_x - origem_x) / 28,
            "vy": -random.uniform(16.0, 20.0),
            "alvo_y": alvo_y,
            "cor": random.choice(["#FFD740", "#FF5252", "#40C4FF", "#69F0AE"]),
        }

    def _criar_estouro(self, x, y):
        cores = ["#FFD740", "#FF5252", "#40C4FF", "#69F0AE", "#FF80AB", "#FFFF8D"]
        particulas = []
        for indice in range(18):
            angulo = (math.tau / 18) * indice + random.uniform(-0.08, 0.08)
            velocidade = random.uniform(3.8, 7.2)
            particulas.append({
                "x": x,
                "y": y,
                "vx": math.cos(angulo) * velocidade,
                "vy": math.sin(angulo) * velocidade,
                "vida": random.randint(14, 22),
                "cor": random.choice(cores),
                "tam": random.randint(3, 6),
            })
        return particulas

    def _loop_animacao_alerta(self):
        """Atualiza a animação de confetes e foguetes enquanto o alerta estiver ativo."""
        if not self._animacao_alerta_ativa:
            return
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return

        larg = self.winfo_width()
        alt = self.winfo_height()
        if larg < 100 or alt < 100:
            self._animacao_alerta_after = self.after(120, self._loop_animacao_alerta)
            return

        agora = time.time()
        if agora - self._ultimo_confete >= 0.12:
            for _ in range(4):
                self._confetes.append(self._criar_confete(larg))
            self._ultimo_confete = agora

        if agora - self._ultimo_foguete >= 1.35:
            self._foguetes.append(self._criar_foguete(larg, alt))
            self._contador_foguetes += 1
            self._ultimo_foguete = agora

        novos_foguetes = []
        for foguete in self._foguetes:
            foguete["x"] += foguete["vx"]
            foguete["y"] += foguete["vy"]
            foguete["vy"] += 0.28
            if foguete["y"] <= foguete["alvo_y"] or foguete["vy"] >= -1.5:
                self._estouros.extend(self._criar_estouro(foguete["x"], foguete["y"]))
            else:
                novos_foguetes.append(foguete)
        self._foguetes = novos_foguetes

        novos_confetes = []
        for confete in self._confetes:
            confete["x"] += confete["vx"]
            confete["y"] += confete["vy"]
            confete["vy"] += 0.08
            if confete["y"] < alt + 30:
                novos_confetes.append(confete)
        self._confetes = novos_confetes

        novos_estouros = []
        for particula in self._estouros:
            particula["x"] += particula["vx"]
            particula["y"] += particula["vy"]
            particula["vy"] += 0.12
            particula["vida"] -= 1
            if particula["vida"] > 0:
                novos_estouros.append(particula)
        self._estouros = novos_estouros

        self.canvas.delete(self._efeitos_tags)

        for confete in self._confetes:
            x1 = confete["x"] - confete["tam"] / 2
            y1 = confete["y"] - confete["tam"] / 2
            x2 = confete["x"] + confete["tam"] / 2
            y2 = confete["y"] + confete["tam"] / 2
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                fill=confete["cor"],
                outline="",
                tags=self._efeitos_tags,
            )

        for foguete in self._foguetes:
            self.canvas.create_line(
                foguete["x"], foguete["y"] + 18,
                foguete["x"], foguete["y"] + 2,
                fill="#FFF3E0",
                width=3,
                tags=self._efeitos_tags,
            )
            self.canvas.create_oval(
                foguete["x"] - 5, foguete["y"] - 5,
                foguete["x"] + 5, foguete["y"] + 5,
                fill=foguete["cor"],
                outline="",
                tags=self._efeitos_tags,
            )

        for particula in self._estouros:
            tam = particula["tam"]
            self.canvas.create_oval(
                particula["x"] - tam, particula["y"] - tam,
                particula["x"] + tam, particula["y"] + tam,
                fill=particula["cor"],
                outline="",
                tags=self._efeitos_tags,
            )

        self._animacao_alerta_after = self.after(90, self._loop_animacao_alerta)

    def _atualizar_alerta_horario(self, agora):
        """Desenha ou remove o aviso animado de horários críticos."""
        self.canvas.delete("alerta_horario")
        alerta = self._obter_alerta_horario(agora)
        if not alerta:
            self._sincronizar_animacao_alerta(None)
            return
        self._sincronizar_animacao_alerta(alerta)
        if not alerta.get("modo_teste"):
            return

        larg = self.winfo_width()
        alt = self.winfo_height()
        if larg < 100 or alt < 100:
            return

        pulso = (agora.second % 2) == 0
        largura_box = min(920, int(larg * 0.58))
        altura_box = 108
        x1 = (larg - largura_box) // 2
        y1 = int(alt * 0.79)
        x2 = x1 + largura_box
        y2 = y1 + altura_box
        cor_principal = "#FF7043" if pulso else "#FF8A65"
        cor_fundo = "#1B2432" if pulso else "#142033"
        cor_texto = "#FFF4EE"
        cor_texto_sec = "#FFD8CB"

        self._retangulo_arredondado(
            x1 + 6, y1 + 6, x2 + 6, y2 + 6,
            raio=22, fill="#000000", alpha_simulado=True, tags="alerta_horario"
        )
        self._retangulo_arredondado(
            x1, y1, x2, y2,
            raio=22, fill=cor_fundo, outline=cor_principal, espessura=3, tags="alerta_horario"
        )

        self.canvas.create_text(
            larg // 2,
            y1 + 28,
            text=alerta["titulo"],
            font=("Segoe UI", 20, "bold"),
            fill=cor_texto,
            anchor="center",
            tags="alerta_horario",
        )
        self.canvas.create_text(
            larg // 2,
            y1 + 58,
            text="Modo de teste do aviso." if alerta.get("modo_teste") else f"Faltam {alerta['minutos_restantes']} minuto(s).",
            font=("Segoe UI", 15, "bold"),
            fill=cor_principal,
            anchor="center",
            tags="alerta_horario",
        )
        self.canvas.create_text(
            larg // 2,
            y1 + 84,
            text=alerta["mensagem"],
            font=("Segoe UI", 12),
            fill=cor_texto_sec,
            anchor="center",
            tags="alerta_horario",
        )
        for desloc in (26, largura_box - 26):
            self.canvas.create_oval(
                x1 + desloc - 7, y1 + 20 - 7,
                x1 + desloc + 7, y1 + 20 + 7,
                fill=cor_principal,
                outline="",
                tags="alerta_horario",
            )

    # ------------------------------------------------------------------
    # RELÓGIO EM TEMPO REAL
    # ------------------------------------------------------------------

    def _verificar_agendamentos_banner(self, agora):
        """Coloca em fila as mensagens cujo horario chegou, uma vez por dia."""
        for indice, agendamento in enumerate(self.config_dados.get("agendamentos_banner", [])[:8]):
            horario = str(agendamento.get("horario", ""))
            mensagem = str(agendamento.get("mensagem", "")).strip()
            try:
                hora, minuto = (int(valor) for valor in horario.split(":", 1))
            except (TypeError, ValueError):
                continue
            if not (0 <= hora <= 23 and 0 <= minuto <= 59 and mensagem):
                continue
            chave = (agora.date().isoformat(), indice, hora, minuto)
            if agora.hour == hora and agora.minute == minuto and chave not in self._agendamentos_disparados:
                self._agendamentos_disparados.add(chave)
                self._fila_banners_agendados.append({"mensagem": mensagem, "horario": horario})

        if not self._banner_agendado_ativo and self._fila_banners_agendados:
            self._iniciar_banner_agendado(self._fila_banners_agendados.pop(0))

    def _iniciar_banner_agendado(self, agendamento):
        """Inicia uma passagem de aviao com banner no topo da tela."""
        self._banner_agendado_ativo = {
            **agendamento,
            "inicio": time.time(),
            "duracao": 30.0,
        }
        self._loop_banner_agendado()

    def _loop_banner_agendado(self):
        """Desenha o aviao e seu banner em movimento suave."""
        ativo = self._banner_agendado_ativo
        if not ativo:
            return
        agora_ts = time.time()
        progresso = (agora_ts - ativo["inicio"]) / ativo["duracao"]
        self.canvas.delete("banner_agendado")
        if progresso >= 1.0:
            self._banner_agendado_ativo = None
            if self._fila_banners_agendados:
                self._iniciar_banner_agendado(self._fila_banners_agendados.pop(0))
            return

        larg = self.winfo_width()
        alt = self.winfo_height()
        if larg < 100 or alt < 100:
            self._banner_agendado_after = self.after(60, self._loop_banner_agendado)
            return

        mensagem = ativo["mensagem"]
        fonte = tkfont.Font(family="Segoe UI", size=max(27, min(44, int(larg / 42))), weight="bold")
        largura_texto = fonte.measure(mensagem)
        largura_banner = min(int(larg * 0.78), max(640, largura_texto + 120))
        altura_banner = 118
        y_banner = max(78, int(alt * 0.105))
        aviao_x = -300 + (larg + largura_banner + 600) * progresso
        banner_x2 = aviao_x - 132
        banner_x1 = banner_x2 - largura_banner
        banner_y1 = y_banner - altura_banner // 2
        banner_y2 = y_banner + altura_banner // 2
        ondulacao = math.sin(agora_ts * 7.0) * 7
        ondulacao_meio = math.sin(agora_ts * 7.0 + 1.8) * 5

        # Cabos que prendem o banner ao aviao.
        self.canvas.create_line(aviao_x - 48, y_banner - 20, banner_x2, banner_y1 + 18 + ondulacao, fill="#C9E7F0", width=3, tags="banner_agendado")
        self.canvas.create_line(aviao_x - 48, y_banner + 20, banner_x2, banner_y2 - 18 + ondulacao, fill="#C9E7F0", width=3, tags="banner_agendado")
        sombra_banner = (
            banner_x1 + 7, banner_y1 + ondulacao + 7,
            (banner_x1 + banner_x2) / 2, banner_y1 + ondulacao_meio + 7,
            banner_x2 + 7, banner_y1 + ondulacao + 7,
            banner_x2 + 7, banner_y2 + ondulacao + 7,
            (banner_x1 + banner_x2) / 2, banner_y2 - ondulacao_meio + 7,
            banner_x1 + 7, banner_y2 - ondulacao + 7,
        )
        tecido_banner = (
            banner_x1, banner_y1 + ondulacao,
            (banner_x1 + banner_x2) / 2, banner_y1 + ondulacao_meio,
            banner_x2, banner_y1 + ondulacao,
            banner_x2, banner_y2 + ondulacao,
            (banner_x1 + banner_x2) / 2, banner_y2 - ondulacao_meio,
            banner_x1, banner_y2 - ondulacao,
        )
        self.canvas.create_polygon(*sombra_banner, fill="#071522", outline="", tags="banner_agendado")
        self.canvas.create_polygon(*tecido_banner, fill="#F5FAFC", outline="#00BCD4", width=4, smooth=True, splinesteps=16, tags="banner_agendado")
        self.canvas.create_text(
            (banner_x1 + banner_x2) / 2,
            y_banner,
            text=mensagem,
            font=fonte,
            fill="#0A3140",
            width=largura_banner - 62,
            justify="center",
            tags="banner_agendado",
        )
        # Aviao de perfil com detalhes suficientes para leitura a distancia.
        self.canvas.create_polygon(
            aviao_x - 64, y_banner - 7,
            aviao_x + 42, y_banner - 16,
            aviao_x + 76, y_banner - 4,
            aviao_x + 90, y_banner + 3,
            aviao_x + 70, y_banner + 12,
            aviao_x - 58, y_banner + 14,
            fill="#EAF5FA", outline="#46697A", width=2, tags="banner_agendado",
        )
        self.canvas.create_polygon(
            aviao_x - 18, y_banner - 8,
            aviao_x + 30, y_banner - 48,
            aviao_x + 53, y_banner - 43,
            aviao_x + 17, y_banner - 4,
            fill="#BED8E3", outline="#46697A", width=2, tags="banner_agendado",
        )
        self.canvas.create_polygon(
            aviao_x - 28, y_banner + 8,
            aviao_x + 36, y_banner + 35,
            aviao_x + 52, y_banner + 27,
            aviao_x + 4, y_banner + 5,
            fill="#9CC4D4", outline="#46697A", width=2, tags="banner_agendado",
        )
        self.canvas.create_polygon(
            aviao_x - 55, y_banner - 7,
            aviao_x - 86, y_banner - 34,
            aviao_x - 37, y_banner - 13,
            fill="#B7D1DC", outline="#46697A", width=2, tags="banner_agendado",
        )
        self.canvas.create_polygon(
            aviao_x + 10, y_banner - 17,
            aviao_x + 31, y_banner - 35,
            aviao_x + 48, y_banner - 16,
            fill="#174B62", outline="#78B9D0", width=1, tags="banner_agendado",
        )
        self.canvas.create_line(aviao_x + 6, y_banner + 13, aviao_x + 3, y_banner + 28, fill="#46697A", width=2, tags="banner_agendado")
        self.canvas.create_line(aviao_x + 37, y_banner + 10, aviao_x + 42, y_banner + 25, fill="#46697A", width=2, tags="banner_agendado")
        self.canvas.create_oval(aviao_x - 2, y_banner + 25, aviao_x + 12, y_banner + 35, fill="#25343C", outline="#A4BAC4", width=1, tags="banner_agendado")
        self.canvas.create_oval(aviao_x + 34, y_banner + 22, aviao_x + 48, y_banner + 32, fill="#25343C", outline="#A4BAC4", width=1, tags="banner_agendado")
        self.canvas.create_line(aviao_x + 89, y_banner - 22, aviao_x + 89, y_banner + 24, fill="#E5F4F8", width=3, tags="banner_agendado")
        self.canvas.create_oval(aviao_x + 84, y_banner - 2, aviao_x + 94, y_banner + 8, fill="#F0B93F", outline="#805C14", width=1, tags="banner_agendado")
        self._banner_agendado_after = self.after(45, self._loop_banner_agendado)

    def _parar_banner_agendado(self):
        """Cancela a animacao do banner ao sair da tela."""
        self._banner_agendado_ativo = None
        self._fila_banners_agendados = []
        if self._banner_agendado_after:
            try:
                self.after_cancel(self._banner_agendado_after)
            except Exception:
                pass
            self._banner_agendado_after = None
        self.canvas.delete("banner_agendado")

    def _periodo_intermitencia(self, valor):
        """Converte a opcao do formulario em segundos; zero mantem o texto fixo."""
        periodos = {
            "1 segundo": 1, "2 segundos": 2, "5 segundos": 5,
            "10 segundos": 10, "30 segundos": 30,
            "1 minuto": 60, "2 minutos": 120, "5 minutos": 300,
        }
        return periodos.get(str(valor), 0)

    def _atualizar_intermitencias_texto(self, agora_ts):
        """Mostra ou oculta grupos de texto conforme os intervalos configurados."""
        configuracoes = self.config_dados.get("intermitencia_textos", {})
        grupos = {
            "hora": "intermitencia_hora",
            "evento": "intermitencia_evento",
            "legenda": "intermitencia_legenda",
            "data": "intermitencia_data",
            "professores": "intermitencia_professores",
            "mensagem": "intermitencia_mensagem",
        }
        for chave, tag in grupos.items():
            opcao = configuracoes.get(chave, "Desligado")
            periodo = self._periodo_intermitencia(opcao)
            visivel = opcao != "Ocultar" and (periodo == 0 or int(agora_ts / periodo) % 2 == 0)
            try:
                self.canvas.itemconfig(tag, state="normal" if visivel else "hidden")
            except tk.TclError:
                pass

    def _atualizar_opacidades_texto(self):
        """Aplica intensidade visual por meio de tramas do Canvas."""
        opacidades = self.config_dados.get("opacidade_textos", {})
        grupos = {
            "hora": "intermitencia_hora",
            "evento": "intermitencia_evento",
            "legenda": "intermitencia_legenda",
            "data": "intermitencia_data",
            "professores": "intermitencia_professores",
            "mensagem": "intermitencia_mensagem",
        }
        tramas = {"100%": "", "75%": "gray75", "50%": "gray50", "25%": "gray25"}
        for chave, tag in grupos.items():
            try:
                self.canvas.itemconfig(tag, stipple=tramas.get(opacidades.get(chave, "100%"), ""))
            except tk.TclError:
                pass

    def _loop_relogio(self):
        """Atualiza o relógio a cada 1 segundo via itemconfig (sem recriar o canvas)."""
        agora_dt = datetime.now()
        agora = agora_dt.strftime("%H:%M:%S")
        try:
            self.canvas.itemconfig("clock_main", text=agora)
            self.canvas.itemconfig("clock_glow", text=agora)
            self.canvas.itemconfig("clock_shadow", text=agora)
        except Exception:
            pass
        self._atualizar_intermitencias_texto(time.time())
        self._atualizar_opacidades_texto()
        self._atualizar_alerta_horario(agora_dt)
        self._verificar_agendamentos_banner(agora_dt)
        self.after(1000, self._loop_relogio)

    # ------------------------------------------------------------------
    # CONTROLE DE RÁDIO
    # ------------------------------------------------------------------

    def _iniciar_radio(self):
        """
        Inicia a reprodução do áudio.
        Arquivo local -> pygame, PowerShell, player do Windows (COM) ou MCI como último fallback.
        """
        url = self.config_dados.get("radios", {}).get(self.radio_nome, "")
        loop = False
        if url == "__local__":
            url = self.config_dados.get("audio_local", "").strip()
            if url and not Path(url).is_absolute():
                url = str(BASE_DIR / url)
            loop = True
        if not url:
            return

        self._url_radio_atual = url
        self._fila_mci = queue.Queue()

        if loop and PYGAME_DISPONIVEL:
            self._backend_audio = "pygame"
            self._iniciar_radio_pygame(url, loop=True)
        elif PS_DISPONIVEL:
            self._backend_audio = "ps"
            self._iniciar_radio_ps(url, self.volume, loop=loop)
        elif loop and WMP_DISPONIVEL:
            self._backend_audio = "wmp"
            self._iniciar_radio_wmp(url, loop=True)
        else:
            self._backend_audio = "mci"
            self._iniciar_radio_mci(url, loop=loop)

    def _iniciar_radio_pygame(self, path, loop=False):
        """Backend pygame mixer para arquivos locais."""
        try:
            if not pygame.mixer.get_init():
                pygame.mixer.init()
            pygame.mixer.music.load(path)
            pygame.mixer.music.set_volume(0 if self.mudo else self.volume / 100)
            pygame.mixer.music.play(-1 if loop else 0)
        except Exception as e:
            print(f"[ERRO] pygame áudio: {e}")
            if PS_DISPONIVEL:
                self._backend_audio = "ps"
                self._iniciar_radio_ps(path, self.volume, loop=loop)
            elif WMP_DISPONIVEL:
                self._backend_audio = "wmp"
                self._iniciar_radio_wmp(path, loop=loop)
            else:
                self._backend_audio = "mci"
                self._iniciar_radio_mci(path, loop=loop)

    def _iniciar_radio_wmp(self, path, loop=False):
        """
        Backend Windows Media Player via COM.
        Costuma ser mais estável para MP3 local do que o MCI.
        """
        volume_inicial = 0 if self.mudo else self.volume

        def _worker():
            player = None
            try:
                pythoncom.CoInitialize()
                player = win32com.client.Dispatch("WMPlayer.OCX")
                player.settings.autoStart = False
                player.settings.volume = max(0, min(100, int(volume_inicial)))
                player.settings.mute = bool(self.mudo)
                player.URL = path
                player.controls.play()

                while True:
                    try:
                        cmd, valor = self._fila_mci.get_nowait()
                        if cmd == "parar":
                            try:
                                player.controls.stop()
                            finally:
                                try:
                                    player.close()
                                except Exception:
                                    pass
                            return
                        elif cmd == "volume":
                            player.settings.volume = max(0, min(100, int(valor)))
                        elif cmd == "mudo":
                            player.settings.mute = bool(valor)
                    except queue.Empty:
                        pass

                    if loop and player.playState == 1:
                        player.controls.currentPosition = 0
                        player.controls.play()

                    time.sleep(0.2)
            except Exception as e:
                print(f"[ERRO] WMP COM: {e}")
                if PS_DISPONIVEL:
                    self._backend_audio = "ps"
                    self._iniciar_radio_ps(path, self.volume, loop=loop)
                else:
                    self._backend_audio = "mci"
                    self._iniciar_radio_mci(path, loop=loop)
            finally:
                try:
                    if player is not None:
                        player.controls.stop()
                except Exception:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        self._player = threading.Thread(target=_worker, daemon=True)
        self._player.start()

    def _iniciar_radio_mci(self, path, loop=False):
        """
        Backend MCI via winmm.dll — presente em TODOS os Windows sem instalar nada.
        Suporta MP3, WAV, WMA e outros formatos de áudio locais.
        """
        import ctypes
        self._parar_mci = False
        volume_inicial = 0 if self.mudo else self.volume
        alias = self._mci_alias

        def _worker():
            try:
                winmm = ctypes.WinDLL("winmm")

                def mci(cmd):
                    buf = ctypes.create_unicode_buffer(512)
                    winmm.mciSendStringW(cmd, buf, 512, 0)
                    return buf.value.strip()

                mci(f'open "{path}" type mpegvideo alias {alias}')
                mci(f'setaudio {alias} volume to {volume_inicial * 10}')
                if loop:
                    mci(f'play {alias} repeat')
                else:
                    mci(f'play {alias}')

                while True:
                    # Processa comandos da fila (volume, mudo, parar)
                    try:
                        cmd, valor = self._fila_mci.get_nowait()
                        if cmd == "parar":
                            mci(f'stop {alias}')
                            mci(f'close {alias}')
                            return
                        elif cmd == "volume":
                            mci(f'setaudio {alias} volume to {valor * 10}')
                        elif cmd == "mudo":
                            mci(f'setaudio {alias} volume to {0 if valor else self.volume * 10}')
                    except queue.Empty:
                        pass

                    time.sleep(0.3)
            except Exception as e:
                print(f"[ERRO] MCI: {e}")
                if WMP_DISPONIVEL:
                    self._backend_audio = "wmp"
                    self._iniciar_radio_wmp(path, loop=loop)
                elif PS_DISPONIVEL:
                    self._backend_audio = "ps"
                    self._iniciar_radio_ps(path, self.volume, loop=loop)

        self._player = threading.Thread(target=_worker, daemon=True)
        self._player.start()

    def _iniciar_radio_ps(self, url, volume, loop=False):
        """
        Backend PowerShell usando WMPlayer.OCX via COM (built-in no Windows 10/11).
        Suporta streams HTTP e arquivos locais (MP3, WAV, WMA, etc.).
        """
        vol_int = max(0, min(100, int(volume)))
        loop_flag = "$true" if loop else "$false"
        ps_cmd = (
            f"$wmp = New-Object -ComObject wmplayer.ocx; "
            f"$wmp.settings.autoStart = $false; "
            f"$wmp.settings.volume = {vol_int}; "
            f"$wmp.settings.setMode('loop', {loop_flag}); "
            f"$wmp.URL = \"{url}\"; "
            f"Start-Sleep -Milliseconds 400; "
            f"$wmp.controls.play(); "
            f"while ($true) {{ Start-Sleep -Seconds 3600 }}"
        )
        try:
            self._proc_radio = subprocess.Popen(
                ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                 '-Sta',
                 '-WindowStyle', 'Hidden', '-Command', ps_cmd],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception as e:
            print(f"[ERRO] PowerShell rádio: {e}")

    def _parar_radio(self):
        """Para a reprodução do áudio (todos os backends)."""
        if self._backend_audio == "pygame":
            try:
                pygame.mixer.music.stop()
                pygame.mixer.quit()
            except Exception:
                pass
        # MCI
        try:
            self._fila_mci.put(("parar", None))
        except Exception:
            pass
        # PowerShell
        if self._proc_radio:
            try:
                self._proc_radio.terminate()
                self._proc_radio = None
            except Exception:
                pass

    def _atualizar_info_radio(self):
        """Atualiza o texto de rádio e volume exibido discretamente na tela."""
        # Remove textos antigos e redesenha
        larg = self.winfo_width()
        alt = self.winfo_height()
        self.canvas.delete("info_radio")
        self.canvas.delete("info_volume")
        self._desenhar_info_radio(larg, alt)

    # ------------------------------------------------------------------
    # ATALHOS DE TECLADO
    # ------------------------------------------------------------------

    def _sair_tela_cheia(self, event=None):
        """ESC: para a rádio e fecha a janela."""
        self._parar_animacao_alerta()
        self._parar_banner_agendado()
        self._parar_animacao_pescador()
        self._parar_animacao_blocos()
        self._parar_animacao_voo()
        self._parar_radio()
        self.destroy()

    def _volume_mais(self, event=None):
        """Seta para cima: aumenta o volume em 5%."""
        if self.mudo:
            self.mudo = False
        self.volume = min(100, self.volume + 5)
        self._aplicar_volume()
        self._atualizar_info_radio()

    def _volume_menos(self, event=None):
        """Seta para baixo: diminui o volume em 5%."""
        self.volume = max(0, self.volume - 5)
        if self.volume == 0:
            self.mudo = True
        self._aplicar_volume()
        self._atualizar_info_radio()

    def _alternar_mudo(self, event=None):
        """M: alterna entre mudo e som."""
        self.mudo = not self.mudo
        self._aplicar_volume()
        self._atualizar_info_radio()

    def _aplicar_volume(self):
        """Aplica volume/mudo no backend de áudio ativo."""
        if self._backend_audio == "pygame":
            try:
                pygame.mixer.music.set_volume(0 if self.mudo else self.volume / 100)
            except Exception:
                pass
        elif self._backend_audio in {"wmp", "mci"}:
            try:
                if self.mudo:
                    self._fila_mci.put(("mudo", True))
                else:
                    self._fila_mci.put(("mudo", False))
                    self._fila_mci.put(("volume", self.volume))
            except Exception:
                pass
        elif self._backend_audio == "ps":
            if self._proc_radio:
                try:
                    self._proc_radio.terminate()
                except Exception:
                    pass
                self._proc_radio = None
            if not self.mudo and self._url_radio_atual:
                self._iniciar_radio_ps(self._url_radio_atual, self.volume)


# =============================================================================
# PONTO DE ENTRADA
# =============================================================================

if __name__ == "__main__":
    config = carregar_config()
    launcher = JanelaConfig(config)
    launcher.mainloop()
