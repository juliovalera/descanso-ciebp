# =============================================================================
# CIEBP - Descanso de Tela Institucional
# Versão: 1.0
# Descrição: Tela de recepção profissional para eventos do CIEBP.
#            Exibe informações do evento, relógio em tempo real e
#            toca rádio online enquanto aguarda o início da atividade.
# =============================================================================

import tkinter as tk
from tkinter import ttk, font as tkfont
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
        self.var_audio_local = tk.StringVar(value=config.get("audio_local", ""))
        self.var_fundo = tk.StringVar(value=config.get("fundo_padrao", "Padrão do espaço"))
        self.var_alerta_teste = tk.StringVar(value="Aviso de almoço")

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
        opcoes_fundo.append("Ilha do Pescador")
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

        espaco = self.espaco_escolhido.get()
        radio = self.radio_escolhida.get()
        fundo = self.var_fundo.get()
        if fundo == "Padrão do espaço":
            fundo_path = ""
        elif fundo == "Ilha do Pescador":
            fundo_path = "__pescador__"
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
        self._tags_fundo_pescador = "fundo_pescador"
        self._animacao_pescador_ativa = False
        self._animacao_pescador_after = None
        self._pescador_cena = {}

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
        coqueiro_tronco = self.canvas.create_line(
            int(larg * 0.21), int(alt * 0.73),
            int(larg * 0.24), int(alt * 0.56),
            fill="#6A3D1A", width=8, smooth=True, tags=tags
        )
        folhas = []
        for pontos in [
            (0.24, 0.56, 0.19, 0.51, 0.16, 0.45),
            (0.24, 0.56, 0.24, 0.49, 0.19, 0.43),
            (0.24, 0.56, 0.29, 0.49, 0.34, 0.46),
            (0.24, 0.56, 0.31, 0.55, 0.36, 0.55),
        ]:
            folhas.append(self.canvas.create_line(
                int(larg * pontos[0]), int(alt * pontos[1]),
                int(larg * pontos[2]), int(alt * pontos[3]),
                int(larg * pontos[4]), int(alt * pontos[5]),
                fill="#2ECC71", width=4, smooth=True, tags=tags
            ))

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
            "tronco": self.canvas.create_line(0, 0, 0, 0, fill="#F6F1E9", width=8, tags=tags),
            "perna1": self.canvas.create_line(0, 0, 0, 0, fill="#213547", width=6, tags=tags),
            "perna2": self.canvas.create_line(0, 0, 0, 0, fill="#213547", width=6, tags=tags),
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
        }

        ondas = []
        for frac, amp in ((0.52, 0.16), (0.67, 0.13), (0.8, 0.11)):
            ondas.append(self.canvas.create_line(0, 0, 0, 0, fill="#89D4FF", width=2, smooth=True, tags=tags))

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
            "folhas": folhas,
            "trapiche": trapiche,
            "ondas": ondas,
            "ripples": agua_ripples,
            "pescador": pescador,
            "base_x": int(larg * 0.4),
            "base_y": int(alt * 0.69),
            "estado": "idle",
            "evento_ate": time.time() + random.uniform(14, 22),
            "proximo_evento": time.time() + random.uniform(18, 30),
            "captura_tipo": "",
        }

        self._atualizar_cena_pescador(datetime.now(), time.time())
        self._animacao_pescador_ativa = True
        self._animacao_pescador_after = self.after(120, self._loop_animacao_pescador)

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

        fase_nuvem = agora_ts * 0.008
        for indice, nuvem in enumerate(cena["nuvens"]):
            desloc_x = math.sin(fase_nuvem + indice * 1.7) * 26
            desloc_y = math.cos(fase_nuvem * 0.7 + indice) * 6
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

        estado = cena["estado"]
        if estado == "idle" and agora_ts >= cena["proximo_evento"]:
            cena["estado"] = "puxando"
            cena["evento_ate"] = agora_ts + random.uniform(2.8, 4.0)
            cena["captura_tipo"] = random.choice(["peixe", "bota", "peixinho"])
        elif estado == "puxando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "mostrando"
            cena["evento_ate"] = agora_ts + random.uniform(2.4, 3.2)
        elif estado == "mostrando" and agora_ts >= cena["evento_ate"]:
            cena["estado"] = "idle"
            cena["captura_tipo"] = ""
            cena["proximo_evento"] = agora_ts + random.uniform(16, 28)

        if cena["estado"] == "idle":
            progresso = math.sin(agora_ts * 1.4) * 0.5 + 0.5
        elif cena["estado"] == "puxando":
            progresso = min(1.0, max(0.0, 1 - ((cena["evento_ate"] - agora_ts) / 3.4)))
        else:
            progresso = 1.0

        base_x = cena["base_x"]
        base_y = cena["base_y"]
        balanco = math.sin(agora_ts * 1.1) * 3
        inclinacao = -12 - (18 * progresso if cena["estado"] != "idle" else 0)
        torso_topo_x = base_x - 6 + balanco
        torso_topo_y = base_y - 42
        torso_base_x = base_x + balanco
        torso_base_y = base_y
        ombro_x = torso_topo_x + 4
        ombro_y = torso_topo_y + 10
        mao_x = base_x + 54 + progresso * 24
        mao_y = base_y - 26 - progresso * 18
        ponta_x = mao_x + 76
        ponta_y = mao_y + inclinacao
        linha_x = ponta_x + 14
        linha_y = max(horizonte + 18, ponta_y + 56 + (1 - progresso) * 34)
        captura_x = linha_x + 18
        captura_y = linha_y - 10

        if cena["estado"] == "mostrando":
            balanco_captura = math.sin(agora_ts * 4.2) * 6
            captura_x = mao_x + 42 + balanco_captura
            captura_y = mao_y - 34
            linha_x = captura_x
            linha_y = captura_y - 16

        pescador = cena["pescador"]
        self.canvas.coords(pescador["banco"], base_x - 28, base_y + 6, base_x + 18, base_y + 16)
        self.canvas.coords(pescador["tronco"], torso_base_x, torso_base_y, torso_topo_x, torso_topo_y)
        self.canvas.coords(pescador["perna1"], base_x - 2, base_y, base_x - 16, base_y + 24)
        self.canvas.coords(pescador["perna2"], base_x + 6, base_y, base_x + 24, base_y + 22)
        self.canvas.coords(pescador["braco"], ombro_x, ombro_y, mao_x, mao_y)
        self.canvas.coords(
            pescador["vara"],
            mao_x - 8, mao_y + 4,
            mao_x + 32, mao_y - 10,
            ponta_x, ponta_y,
        )
        self.canvas.coords(pescador["linha"], ponta_x, ponta_y, linha_x, linha_y)
        self.canvas.coords(pescador["cabeca"], torso_topo_x - 12, torso_topo_y - 20, torso_topo_x + 12, torso_topo_y + 4)
        self.canvas.coords(
            pescador["chapeu"],
            torso_topo_x - 18, torso_topo_y - 10,
            torso_topo_x + 16, torso_topo_y - 10,
            torso_topo_x + 6, torso_topo_y - 22,
        )

        ripple_cx = linha_x
        ripple_cy = linha_y + 6
        for indice, item in enumerate(cena["ripples"]):
            raio = 10 + indice * 12 + math.sin(agora_ts * 3 + indice) * 2
            self.canvas.coords(
                item,
                ripple_cx - raio, ripple_cy - raio * 0.45,
                ripple_cx + raio, ripple_cy + raio * 0.45,
            )

        captura = pescador["captura"]
        if cena["estado"] == "mostrando":
            if cena["captura_tipo"] == "bota":
                self.canvas.coords(
                    pescador["captura_bota_cano"],
                    captura_x - 14, captura_y - 30,
                    captura_x + 6, captura_y - 30,
                    captura_x + 6, captura_y + 2,
                    captura_x - 14, captura_y + 2,
                )
                self.canvas.coords(
                    pescador["captura_bota_sola"],
                    captura_x - 18, captura_y + 2,
                    captura_x + 12, captura_y + 2,
                    captura_x + 18, captura_y + 12,
                    captura_x - 18, captura_y + 12,
                )
                self.canvas.itemconfig(pescador["captura_bota_cano"], state="normal")
                self.canvas.itemconfig(pescador["captura_bota_sola"], state="normal")
                self.canvas.itemconfig(pescador["captura_peixe_corpo"], state="hidden")
                self.canvas.itemconfig(pescador["captura_peixe_cauda"], state="hidden")
                self.canvas.itemconfig(pescador["captura_peixe_olho"], state="hidden")
            elif cena["captura_tipo"] == "peixinho":
                meio = 12
                altura = 8
                self.canvas.coords(
                    pescador["captura_peixe_corpo"],
                    captura_x - meio, captura_y - altura,
                    captura_x + meio, captura_y + altura,
                )
                self.canvas.coords(
                    pescador["captura_peixe_cauda"],
                    captura_x + meio - 2, captura_y,
                    captura_x + meio + 14, captura_y - 10,
                    captura_x + meio + 14, captura_y + 10,
                )
                self.canvas.coords(
                    pescador["captura_peixe_olho"],
                    captura_x - 6, captura_y - 2,
                    captura_x - 2, captura_y + 2,
                )
                self.canvas.itemconfig(pescador["captura_peixe_corpo"], fill="#9BE7FF", outline="#EAFBFF", state="normal")
                self.canvas.itemconfig(pescador["captura_peixe_cauda"], fill="#69D8F6", outline="#EAFBFF", state="normal")
                self.canvas.itemconfig(pescador["captura_peixe_olho"], state="normal")
                self.canvas.itemconfig(pescador["captura_bota_cano"], state="hidden")
                self.canvas.itemconfig(pescador["captura_bota_sola"], state="hidden")
            else:
                meio = 18
                altura = 12
                self.canvas.coords(
                    pescador["captura_peixe_corpo"],
                    captura_x - meio, captura_y - altura,
                    captura_x + meio, captura_y + altura,
                )
                self.canvas.coords(
                    pescador["captura_peixe_cauda"],
                    captura_x + meio - 2, captura_y,
                    captura_x + meio + 18, captura_y - 14,
                    captura_x + meio + 18, captura_y + 14,
                )
                self.canvas.coords(
                    pescador["captura_peixe_olho"],
                    captura_x - 8, captura_y - 3,
                    captura_x - 3, captura_y + 3,
                )
                self.canvas.itemconfig(pescador["captura_peixe_corpo"], fill="#7DFF9B", outline="#E5FFEF", state="normal")
                self.canvas.itemconfig(pescador["captura_peixe_cauda"], fill="#55D68C", outline="#E5FFEF", state="normal")
                self.canvas.itemconfig(pescador["captura_peixe_olho"], state="normal")
                self.canvas.itemconfig(pescador["captura_bota_cano"], state="hidden")
                self.canvas.itemconfig(pescador["captura_bota_sola"], state="hidden")
            self.canvas.itemconfig(captura, text="", state="hidden")
        else:
            self.canvas.itemconfig(captura, text="", state="hidden")
            self.canvas.itemconfig(pescador["captura_peixe_corpo"], state="hidden")
            self.canvas.itemconfig(pescador["captura_peixe_cauda"], state="hidden")
            self.canvas.itemconfig(pescador["captura_peixe_olho"], state="hidden")
            self.canvas.itemconfig(pescador["captura_bota_cano"], state="hidden")
            self.canvas.itemconfig(pescador["captura_bota_sola"], state="hidden")

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

    def _desenhar_painel_central(self, larg, alt):
        """Painel único que contém relógio + todas as informações do evento."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")
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
        tam_clock = max(56, min(84, int(alt * 0.095 * escala_texto)))
        y_clock = py + int(painel_alt * 0.20)

        self.canvas.create_text(cx+6, y_clock+6,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#000000", anchor="center", tags="clock_shadow")
        self.canvas.create_text(cx+2, y_clock+2,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#2ECC71", anchor="center", tags="clock_glow")
        self.canvas.create_text(cx, y_clock,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#7DFF9B", anchor="center", tags="clock_main")

        # ── INFORMAÇÕES ─────────────────────────────────────────────────
        # Começa abaixo do centro do relógio + espaço proporcional ao tamanho da fonte
        y = y_clock + int(tam_clock * 0.82)
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
            font=("Segoe UI", tam_evento, "bold"), fill="#000000", anchor="n")
        self.canvas.create_text(cx, y, text=evento.upper(),
            font=("Segoe UI", tam_evento, "bold"), fill="#FB1874", anchor="n")
        y += int(tam_evento * 1.5)

        # Nome do espaço
        self.canvas.create_text(cx+2, y+2, text=self.espaco_atual,
            font=("Segoe UI", tam_espaco, "bold"), fill="#000000", anchor="n")
        self.canvas.create_text(cx, y, text=self.espaco_atual,
            font=("Segoe UI", tam_espaco, "bold"), fill=cor_destaque, anchor="n")
        y += int(tam_espaco * 1.45)

        # Data
        data = self.config_dados.get("data", datetime.now().strftime("%d/%m/%Y"))
        self.canvas.create_text(cx+1, y+1, text=f"\U0001f4c5  {data}",
            font=("Segoe UI", tam_data, "bold"), fill="#000000", anchor="n")
        self.canvas.create_text(cx, y, text=f"\U0001f4c5  {data}",
            font=("Segoe UI", tam_data, "bold"), fill="#FB1874", anchor="n")
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
                )
                self.canvas.create_text(
                    cx, y,
                    text=texto_linha,
                    font=("Segoe UI", tam_prof, "bold"),
                    fill="#FB1874",
                    anchor="n",
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
            anchor="n", width=painel_larg - 50)
        texto_msg_id = self.canvas.create_text(cx, y, text=mensagem,
            font=("Segoe UI", tam_msg), fill="#E6EEF5",
            anchor="n", width=painel_larg - 50, tags="texto_mensagem")

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
        self._atualizar_alerta_horario(agora_dt)
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
        self._parar_animacao_pescador()
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
