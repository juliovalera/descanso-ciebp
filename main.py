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
        self.espaco_escolhido = tk.StringVar(value=config.get("espaco_padrao", "Hub de Inovação"))
        self.radio_escolhida = tk.StringVar(value=config.get("radio_padrao", "Melody FM"))

        # Variáveis dos campos editáveis do evento
        self.var_evento = tk.StringVar(value=config.get("evento", "Evento CIEBP"))
        self.var_data = tk.StringVar(value=config.get("data", datetime.now().strftime("%d/%m/%Y")))
        self.var_professores = tk.StringVar(value=", ".join(config.get("professores", [])))
        self.var_mensagem = tk.StringVar(value=config.get("mensagem_boas_vindas", ""))
        self.var_audio_local = tk.StringVar(value=config.get("audio_local", ""))
        self.var_fundo = tk.StringVar(value=config.get("fundo_padrao", "Padrão do espaço"))

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
        larg = 520

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
            text="Descanso de Tela Institucional  ·  Centro de Inovação da Escola Básica Paulista",
            font=("Segoe UI", 8),
            fg="#546E7A",
            bg="#0D1B2A",
            wraplength=440,
            justify="center"
        ).pack(pady=(1, 0))

        # Separador
        tk.Frame(F, bg="#1E3A5F", height=1).pack(fill="x", padx=30, pady=8)

        # --- Configuração do evento ---
        frame_ev = tk.Frame(F, bg="#0D1B2A")
        frame_ev.pack(fill="x", padx=30, pady=(0, 4))

        tk.Label(frame_ev, text="Configuração do evento",
                 font=("Segoe UI", 9, "bold"), fg="#B0BEC5",
                 bg="#0D1B2A", anchor="w").pack(fill="x", pady=(0, 2))

        def _campo(rotulo, var, parent=frame_ev):
            """Cria um par rótulo + campo de texto estilizado."""
            tk.Label(parent, text=rotulo, font=("Segoe UI", 8),
                     fg="#546E7A", bg="#0D1B2A", anchor="w").pack(fill="x")
            e = tk.Entry(parent, textvariable=var, font=("Segoe UI", 10),
                         bg="#132336", fg="#ECEFF1", insertbackground="#ECEFF1",
                         relief="flat", bd=0, highlightthickness=1,
                         highlightbackground="#1E3A5F",
                         highlightcolor="#00BCD4")
            e.pack(fill="x", ipady=3, pady=(0, 3))
            e.bind("<MouseWheel>", _on_wheel)

        _campo("Nome do evento", self.var_evento)
        _campo("Data  (dd/mm/aaaa)", self.var_data)
        _campo("Professores  (separe por vírgula)", self.var_professores)
        _campo("Mensagem de boas-vindas", self.var_mensagem)

        # Separador
        tk.Frame(F, bg="#1E3A5F", height=1).pack(fill="x", padx=30, pady=(2, 6))

        # --- Seleção de Espaço ---
        frame_espaco = tk.Frame(F, bg="#0D1B2A")
        frame_espaco.pack(fill="x", padx=30, pady=(0, 6))

        tk.Label(
            frame_espaco,
            text="Espaço do evento:",
            font=("Segoe UI", 10, "bold"),
            fg="#B0BEC5",
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
            text="Imagem de fundo:",
            font=("Segoe UI", 10, "bold"),
            fg="#B0BEC5",
            bg="#0D1B2A",
            anchor="w"
        ).pack(fill="x", pady=(0, 3))

        base_dir = BASE_DIR
        opcoes_fundo = ["Padrão do espaço"] + [
            f"Fundo {i}" for i in range(1, 7)
            if (base_dir / "assets" / f"fundo{i}.png").exists()
        ]
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
            text="Rádio de fundo:",
            font=("Segoe UI", 10, "bold"),
            fg="#B0BEC5",
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

        tk.Label(self._frame_mp3, text="Arquivo de áudio (MP3, WAV, WMA...):",
                 font=("Segoe UI", 8), fg="#546E7A",
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
        tk.Button(
            frame_mp3_row, text="Procurar…",
            font=("Segoe UI", 9), fg="#FFFFFF", bg="#1E3A5F",
            activebackground="#00BCD4", activeforeground="#FFFFFF",
            relief="flat", cursor="hand2", bd=0, padx=10,
            command=self._selecionar_mp3
        ).pack(side="right", padx=(4, 0))

        # --- Status do áudio ---
        frame_ast = tk.Frame(F, bg="#0D1B2A")
        frame_ast.pack(fill="x", padx=30, pady=(4, 2))
        if WMP_DISPONIVEL:
            _cor_a = "#4CAF50"
            _txt_a = "✔  Controle de volume disponível"
        elif PS_DISPONIVEL:
            _cor_a = "#FF9800"
            _txt_a = "⚠  Áudio via PowerShell  —  pip install pywin32 para controle de volume"
        else:
            _cor_a = "#F44336"
            _txt_a = "✖  Sem áudio  —  execute: pip install pywin32"
        tk.Label(frame_ast, text=_txt_a, font=("Segoe UI", 8),
                 fg=_cor_a, bg="#0D1B2A", anchor="w",
                 wraplength=462, justify="left").pack(fill="x")

        # --- Dica de atalhos ---
        frame_dica = tk.Frame(F, bg="#0F2336", bd=0)
        frame_dica.pack(fill="x", padx=30, pady=(4, 6))
        tk.Label(
            frame_dica,
            text="Atalhos:  ESC → sair da tela cheia   ⇕ → volume   M → mudo/som",
            font=("Segoe UI", 8),
            fg="#607D8B",
            bg="#0F2336",
            justify="left",
            padx=12,
            pady=4
        ).pack(fill="x")

        # --- Botão Sobre ---
        frame_sobre = tk.Frame(F, bg="#0D1B2A")
        frame_sobre.pack(fill="x", padx=30, pady=(0, 8))
        tk.Button(
            frame_sobre,
            text="ℹ  Sobre este projeto",
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
            text="▶  Iniciar Descanso de Tela",
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
        # Atualiza config em memória com os valores digitados
        self.config_dados["evento"] = self.var_evento.get().strip() or "Evento CIEBP"
        self.config_dados["data"] = self.var_data.get().strip()
        profs_raw = self.var_professores.get()
        self.config_dados["professores"] = [p.strip() for p in profs_raw.split(",") if p.strip()]
        self.config_dados["mensagem_boas_vindas"] = self.var_mensagem.get().strip()

        espaco = self.espaco_escolhido.get()
        radio = self.radio_escolhida.get()
        fundo = self.var_fundo.get()
        if fundo == "Padrão do espaço":
            fundo_path = ""
        else:
            num = fundo.split()[-1]  # "Fundo 3" -> "3"
            fundo_path = f"assets/fundo{num}.png"
        if radio == "Música local (MP3)":
            self.config_dados["audio_local"] = self.var_audio_local.get().strip()

        # Persiste todas as escolhas para próxima abertura
        self.config_dados["espaco_padrao"] = espaco
        self.config_dados["radio_padrao"] = radio
        self.config_dados["fundo_padrao"] = fundo
        salvar_config(self.config_dados)

        self.destroy()
        app = DescansoTela(self.config_dados, espaco, radio, fundo_path)
        app.mainloop()

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
        tk.Label(janela, text="✦  CIEBP Descanso de Tela  ✦",
                 font=("Segoe UI", 15, "bold"), fg="#00BCD4",
                 bg="#0D1B2A").pack(pady=(28, 4))

        tk.Label(janela, text="Centro de Inovação da Escola Básica Paulista",
                 font=("Segoe UI", 9), fg="#546E7A",
                 bg="#0D1B2A").pack()

        # Separador
        tk.Frame(janela, bg="#1E3A5F", height=1).pack(fill="x", padx=40, pady=16)

        # Texto do projeto
        texto_projeto = (
            "Este projeto nasceu com dois propósitos que se complementam:\n\n"
            "🐍  Demonstrar na prática o poder do Python — linguagem dominante\n"
            "      no mercado atual — como ferramenta de criação real, não apenas\n"
            "      como exercício acadêmico.\n\n"
            "🎓  Apoiar os eventos pedagógicos do CIEBP, oferecendo uma solução\n"
            "      profissional, bonita e funcional para as atividades de formação\n"
            "      e orientação técnica dos professores da rede."
        )
        tk.Label(janela, text=texto_projeto,
                 font=("Segoe UI", 10), fg="#CFD8DC", bg="#0D1B2A",
                 justify="left", wraplength=440, padx=20).pack()

        # Separador
        tk.Frame(janela, bg="#1E3A5F", height=1).pack(fill="x", padx=40, pady=16)

        # Créditos
        tk.Label(janela, text="Criado por",
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

    def __init__(self, config, espaco, radio_nome, fundo_path=""):
        super().__init__()

        self.config_dados = config
        self.espaco_atual = espaco
        self.radio_nome = radio_nome
        self.fundo_path = fundo_path  # caminho relativo ou "" para usar o do espaço

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

        self.canvas.delete("all")
        self._desenhar_fundo(larg, alt)
        self._desenhar_painel_central(larg, alt)  # inclui relógio e logo
        self._desenhar_rodape(larg, alt)
        self._desenhar_info_radio(larg, alt)

    def _desenhar_fundo(self, larg, alt):
        """Carrega e desenha a imagem de fundo com overlay escuro."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")

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

    def _desenhar_painel_central(self, larg, alt):
        """Painel único que contém relógio + todas as informações do evento."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")

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
        tam_clock = max(48, min(72, int(alt * 0.09)))
        y_clock = py + int(painel_alt * 0.20)

        self.canvas.create_text(cx+3, y_clock+3,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill="#000000", anchor="center", tags="clock_shadow")
        self.canvas.create_text(cx, y_clock,
            text="00:00:00",
            font=("Segoe UI", tam_clock, "bold"),
            fill=cor_destaque, anchor="center", tags="clock_main")

        # ── INFORMAÇÕES ─────────────────────────────────────────────────
        # Começa abaixo do centro do relógio + espaço proporcional ao tamanho da fonte
        y = y_clock + int(tam_clock * 0.75)

        # Separador
        self._linha_decorativa(cx, y, int(painel_larg * 0.5), cor_destaque)
        y += 24

        # Nome do evento
        evento = self.config_dados.get("evento", "Evento CIEBP")
        self.canvas.create_text(cx+2, y+2, text=evento.upper(),
            font=("Segoe UI", 30, "bold"), fill="#000000", anchor="n")
        self.canvas.create_text(cx, y, text=evento.upper(),
            font=("Segoe UI", 30, "bold"), fill="#FFFFFF", anchor="n")
        y += 52

        # Nome do espaço
        self.canvas.create_text(cx+2, y+2, text=self.espaco_atual,
            font=("Segoe UI", 22), fill="#000000", anchor="n")
        self.canvas.create_text(cx, y, text=self.espaco_atual,
            font=("Segoe UI", 22), fill=cor_destaque, anchor="n")
        y += 40

        # Data
        data = self.config_dados.get("data", datetime.now().strftime("%d/%m/%Y"))
        self.canvas.create_text(cx, y, text=f"\U0001f4c5  {data}",
            font=("Segoe UI", 18), fill="#B0BEC5", anchor="n")
        y += 36

        # Professores
        professores = self.config_dados.get("professores", [])
        if professores:
            texto_prof = "  \u00b7  ".join(professores)
            self.canvas.create_text(cx+1, y+1,
                text=f"\U0001f464  {texto_prof}", font=("Segoe UI", 17),
                fill="#000000", anchor="n")
            self.canvas.create_text(cx, y,
                text=f"\U0001f464  {texto_prof}", font=("Segoe UI", 17),
                fill="#CFD8DC", anchor="n", width=painel_larg - 60)
            y += 38

        # Separador
        self._linha_decorativa(cx, y, int(painel_larg * 0.4), cor_destaque, espessura=1)
        y += 20

        # Mensagem
        mensagem = self.config_dados.get("mensagem_boas_vindas", "Bem-vindo(a) ao CIEBP!")
        self.canvas.create_text(cx+1, y+1, text=mensagem,
            font=("Segoe UI", 19, "italic"), fill="#000000",
            anchor="n", width=painel_larg - 80)
        self.canvas.create_text(cx, y, text=mensagem,
            font=("Segoe UI", 19, "italic"), fill="#90A4AE",
            anchor="n", width=painel_larg - 80, tags="texto_mensagem")

    def _desenhar_rodape(self, larg, alt):
        """Desenha o rodapé com o nome institucional."""
        espaco_info = self.config_dados.get("espacos", {}).get(self.espaco_atual, {})
        cor_destaque = espaco_info.get("cor_destaque", "#00BCD4")

        # Faixa colorida no rodapé
        self.canvas.create_rectangle(0, alt - 4, larg, alt, fill=cor_destaque, outline="")

        # Texto institucional
        self.canvas.create_text(
            larg // 2,
            alt - 28,
            text="CIEBP  ·  Centro de Inovação da Escola Básica Paulista",
            font=("Segoe UI", 11),
            fill="#546E7A",
            anchor="center"
        )

    def _desenhar_info_radio(self, larg, alt):
        """
        Exibe discretamente no canto inferior esquerdo
        o nome da rádio e o nível de volume.
        """
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
            font=("Segoe UI", 10),
            fill="#455A64",
            anchor="w",
            tags="info_radio"
        )
        self.canvas.create_text(
            22, alt - 28,
            text=texto_vol,
            font=("Segoe UI", 10),
            fill="#455A64",
            anchor="w",
            tags="info_volume"
        )

    # ------------------------------------------------------------------
    # ELEMENTOS GRÁFICOS AUXILIARES
    # ------------------------------------------------------------------

    def _retangulo_arredondado(self, x1, y1, x2, y2, raio=15,
                                fill="#0A1628", outline="", espessura=1,
                                alpha_simulado=False):
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
                                          fill=fill, outline="")
            self.canvas.create_rectangle(x1, y1 + raio, x2, y2 - raio,
                                          fill=fill, outline="")
        if fill:
            # Cantos arredondados
            self.canvas.create_arc(x1, y1, x1 + 2*raio, y1 + 2*raio,
                                    start=90, extent=90, fill=fill, outline="")
            self.canvas.create_arc(x2 - 2*raio, y1, x2, y1 + 2*raio,
                                    start=0, extent=90, fill=fill, outline="")
            self.canvas.create_arc(x1, y2 - 2*raio, x1 + 2*raio, y2,
                                    start=180, extent=90, fill=fill, outline="")
            self.canvas.create_arc(x2 - 2*raio, y2 - 2*raio, x2, y2,
                                    start=270, extent=90, fill=fill, outline="")

        if outline:
            # Bordas arredondadas
            self.canvas.create_arc(x1, y1, x1 + 2*raio, y1 + 2*raio,
                                    start=90, extent=90,
                                    outline=outline, style="arc", width=espessura)
            self.canvas.create_arc(x2 - 2*raio, y1, x2, y1 + 2*raio,
                                    start=0, extent=90,
                                    outline=outline, style="arc", width=espessura)
            self.canvas.create_arc(x1, y2 - 2*raio, x1 + 2*raio, y2,
                                    start=180, extent=90,
                                    outline=outline, style="arc", width=espessura)
            self.canvas.create_arc(x2 - 2*raio, y2 - 2*raio, x2, y2,
                                    start=270, extent=90,
                                    outline=outline, style="arc", width=espessura)
            # Linhas retas das bordas
            self.canvas.create_line(x1 + raio, y1, x2 - raio, y1,
                                     fill=outline, width=espessura)
            self.canvas.create_line(x1 + raio, y2, x2 - raio, y2,
                                     fill=outline, width=espessura)
            self.canvas.create_line(x1, y1 + raio, x1, y2 - raio,
                                     fill=outline, width=espessura)
            self.canvas.create_line(x2, y1 + raio, x2, y2 - raio,
                                     fill=outline, width=espessura)

    def _linha_decorativa(self, cx, cy, metade_larg, cor, espessura=2):
        """Desenha uma linha horizontal decorativa centrada em (cx, cy)."""
        self.canvas.create_line(
            cx - metade_larg, cy,
            cx + metade_larg, cy,
            fill=cor, width=espessura, dash=(4, 6)
        )

    # ------------------------------------------------------------------
    # RELÓGIO EM TEMPO REAL
    # ------------------------------------------------------------------

    def _loop_relogio(self):
        """Atualiza o relógio a cada 1 segundo via itemconfig (sem recriar o canvas)."""
        agora = datetime.now().strftime("%H:%M:%S")
        try:
            self.canvas.itemconfig("clock_main", text=agora)
            self.canvas.itemconfig("clock_shadow", text=agora)
        except Exception:
            pass
        self.after(1000, self._loop_relogio)

    # ------------------------------------------------------------------
    # CONTROLE DE RÁDIO
    # ------------------------------------------------------------------

    def _iniciar_radio(self):
        """
        Inicia a reprodução do áudio.
        Arquivo local  -> MCI/winmm.dll (nativo em todo Windows) ou PS fallback.
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
        self._backend_audio = "mci"
        self._iniciar_radio_mci(url, loop=loop)

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

                    # Verifica se terminou para reiniciar (loop)
                    status = mci(f'status {alias} mode')
                    if loop and status == 'stopped':
                        mci(f'seek {alias} to start')
                        mci(f'play {alias}')

                    time.sleep(0.3)
            except Exception as e:
                print(f"[ERRO] MCI: {e}")
                if PS_DISPONIVEL:
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
        if loop:
            # Arquivo local: reinicia quando o playState volta a 1 (parado)
            ps_cmd = (
                f'$wmp = New-Object -ComObject wmplayer.ocx; '
                f'$wmp.settings.autoStart = $true; '
                f'$wmp.settings.volume = {vol_int}; '
                f'$wmp.URL = "{url}"; '
                f'while ($true) {{ Start-Sleep -Milliseconds 800; '
                f'if ($wmp.playState -eq 1) {{ $wmp.controls.play() }} }}'
            )
        else:
            ps_cmd = (
                f'$wmp = New-Object -ComObject wmplayer.ocx; '
                f'$wmp.settings.autoStart = $true; '
                f'$wmp.settings.volume = {vol_int}; '
                f'$wmp.URL = "{url}"; '
                f'Start-Sleep 86400'
            )
        try:
            self._proc_radio = subprocess.Popen(
                ['powershell', '-ExecutionPolicy', 'Bypass',
                 '-WindowStyle', 'Hidden', '-Command', ps_cmd],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception as e:
            print(f"[ERRO] PowerShell rádio: {e}")

    def _parar_radio(self):
        """Para a reprodução do áudio (todos os backends)."""
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
        if self._backend_audio == "mci":
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
