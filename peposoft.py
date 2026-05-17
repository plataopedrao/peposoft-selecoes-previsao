import math
import os
import random
import tkinter as tk
import unicodedata
from ctypes import windll
from tkinter import ttk, messagebox

# ----------------------------------------------------------------------
# ÁUDIO (MCI / winmm — toca MP3, alias único interrompe o áudio anterior)
# ----------------------------------------------------------------------
SFX_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sfx")
_SFX_ALIAS = "peposoft_sfx"


def _mci(cmd):
    try:
        windll.winmm.mciSendStringW(cmd, None, 0, None)
    except Exception:
        pass


def play_sfx(filename):
    path = os.path.join(SFX_DIR, filename)
    if not os.path.isfile(path):
        return
    _mci(f'close {_SFX_ALIAS}')
    _mci(f'open "{path}" alias {_SFX_ALIAS}')
    _mci(f'play {_SFX_ALIAS} from 0')

# ----------------------------------------------------------------------
# DADOS
# ----------------------------------------------------------------------
selecoes = {
    "Brasil":          {"forca": 95, "flag": "🇧🇷", "cor": "#FFDF00"},
    "Argentina":       {"forca": 94, "flag": "🇦🇷", "cor": "#75AADB"},
    "França":          {"forca": 93, "flag": "🇫🇷", "cor": "#0055A4"},
    "Inglaterra":      {"forca": 91, "flag": "🏴󠁧󠁢󠁥󠁮󠁧󠁿", "cor": "#FFFFFF"},
    "Portugal":        {"forca": 90, "flag": "🇵🇹", "cor": "#006600"},
    "Espanha":         {"forca": 89, "flag": "🇪🇸", "cor": "#AA151B"},
    "Alemanha":        {"forca": 88, "flag": "🇩🇪", "cor": "#FFCE00"},
    "Holanda":         {"forca": 87, "flag": "🇳🇱", "cor": "#F36C21"},
    "Itália":          {"forca": 86, "flag": "🇮🇹", "cor": "#008C45"},
    "Uruguai":         {"forca": 84, "flag": "🇺🇾", "cor": "#7CB9E8"},
    "Croácia":         {"forca": 83, "flag": "🇭🇷", "cor": "#FF0000"},
    "Bélgica":         {"forca": 82, "flag": "🇧🇪", "cor": "#FAE042"},
    "Colômbia":        {"forca": 82, "flag": "🇨🇴", "cor": "#FCD116"},
    "Marrocos":        {"forca": 81, "flag": "🇲🇦", "cor": "#C1272D"},
    "México":          {"forca": 80, "flag": "🇲🇽", "cor": "#006847"},
    "Suíça":           {"forca": 80, "flag": "🇨🇭", "cor": "#FF0000"},
    "Estados Unidos":  {"forca": 79, "flag": "🇺🇸", "cor": "#3C3B6E"},
    "Japão":           {"forca": 78, "flag": "🇯🇵", "cor": "#BC002D"},
    "Sérvia":          {"forca": 77, "flag": "🇷🇸", "cor": "#C6363C"},
    "Coreia do Sul":   {"forca": 76, "flag": "🇰🇷", "cor": "#003478"},
}

NARRACOES_GOL = [
    "GOOOOL! Que finalização!",
    "Bola na rede! Jogada impecável!",
    "É goooool! O estádio explode!",
    "Pegou de primeira e fez!",
    "Cabeçada certeira, é gol!",
    "Chute de fora da área, no ângulo!",
    "Driblou e finalizou, é gol!",
]

EVENTOS_NEUTROS = [
    "Posse de bola no meio-campo.",
    "Falta marcada.",
    "Escanteio cobrado, defesa afasta.",
    "Tiro de meta.",
    "Cartão amarelo no zagueiro.",
    "Lateral cobrada rapidamente.",
    "Goleiro defende firme!",
    "Cruzamento na área, afastado.",
    "Chute pra fora, quase!",
    "Substituição no banco.",
]

# ----------------------------------------------------------------------
# TEMA
# ----------------------------------------------------------------------
BG          = "#0b1320"
BG_CARD     = "#152238"
BG_CARD_2   = "#1c2e4a"
ACCENT      = "#f2aa4c"
ACCENT_2    = "#22d3ee"
TXT         = "#f8fafc"
TXT_MUTED   = "#94a3b8"
WIN         = "#22c55e"
LOSE        = "#ef4444"
DRAW        = "#eab308"

# ----------------------------------------------------------------------
# LÓGICA
# ----------------------------------------------------------------------
def calcular_chance_gol(forca_atacante, forca_defensor):
    diff = forca_atacante - forca_defensor
    base = 0.025 + diff * 0.0015
    return max(0.005, min(base, 0.08))


def probabilidades(f1, f2):
    total = f1 + f2
    return f1 / total * 100, f2 / total * 100


# ----------------------------------------------------------------------
# KICKFORM — estatísticas dos últimos 10 jogos (resultados times.xlsx)
# ----------------------------------------------------------------------
XLSX_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "resultados times", "resultados times.xlsx",
)
# Referências do KickForm (Bundesliga ~ 1.66 em casa, 1.20 fora)
LIGA_HOME_AVG = 1.66
LIGA_AWAY_AVG = 1.20


def _norm(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFD", s.strip())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.upper()


def _carregar_kickform():
    """Lê o xlsx, retorna (stats_por_time, media_gols_casa, media_gols_fora)."""
    try:
        import openpyxl
    except ImportError:
        return {}, LIGA_HOME_AVG, LIGA_AWAY_AVG
    if not os.path.isfile(XLSX_PATH):
        return {}, LIGA_HOME_AVG, LIGA_AWAY_AVG
    try:
        wb = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
        ws = wb.active
        jogos = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            line = row[0]
            if not isinstance(line, str):
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 6:
                continue
            try:
                sel = _norm(parts[0])
                mand_n = _norm(parts[1])
                gm = int(parts[2])
                gv = int(parts[3])
                vis_n = _norm(parts[4])
            except ValueError:
                continue
            if sel == mand_n:
                pro, contra, casa = gm, gv, True
            elif sel == vis_n:
                pro, contra, casa = gv, gm, False
            else:
                continue
            jogos.setdefault(sel, []).append((pro, contra, casa))
        wb.close()
        stats = {}
        home_pro_total = home_n = away_pro_total = away_n = 0
        for nome, lst in jogos.items():
            n = len(lst)
            casa = [(p, c) for p, c, h in lst if h]
            fora = [(p, c) for p, c, h in lst if not h]
            for p, c, h in lst:
                if h:
                    home_pro_total += p
                    home_n += 1
                else:
                    away_pro_total += p
                    away_n += 1
            stats[nome] = {
                "n": n,
                "ataque": sum(p for p, _, _ in lst) / n,
                "defesa": sum(c for _, c, _ in lst) / n,
                "ataque_casa": (sum(p for p, _ in casa) / len(casa)) if casa else None,
                "defesa_casa": (sum(c for _, c in casa) / len(casa)) if casa else None,
                "ataque_fora": (sum(p for p, _ in fora) / len(fora)) if fora else None,
                "defesa_fora": (sum(c for _, c in fora) / len(fora)) if fora else None,
            }
        media_casa = home_pro_total / home_n if home_n else LIGA_HOME_AVG
        media_fora = away_pro_total / away_n if away_n else LIGA_AWAY_AVG
        return stats, media_casa, media_fora
    except Exception:
        return {}, LIGA_HOME_AVG, LIGA_AWAY_AVG


KICKFORM, LIGA_HOME_OBS, LIGA_AWAY_OBS = _carregar_kickform()


def _sample_poisson(mu):
    """Amostragem direta de Poisson (algoritmo de Knuth)."""
    if mu <= 0:
        return 0
    L = math.exp(-mu)
    k = 0
    p = 1.0
    while True:
        k += 1
        p *= random.random()
        if p <= L:
            return k - 1


def sortear_minutos_gols(n_a, n_b):
    """
    Distribui n_a + n_b gols entre os 90 minutos sem colisão, com leve viés
    pro 2º tempo (peso 1.15 vs 1.0) — reflete a distribuição real do futebol.
    """
    if n_a + n_b == 0:
        return [], []
    pesos = [1.0 if m < 46 else 1.15 for m in range(1, 91)]
    candidatos = list(range(1, 91))
    minutos = []
    for _ in range(min(n_a + n_b, 90)):
        idx = random.choices(range(len(candidatos)), weights=pesos, k=1)[0]
        minutos.append(candidatos.pop(idx))
        pesos.pop(idx)
    random.shuffle(minutos)
    return sorted(minutos[:n_a]), sorted(minutos[n_a:n_a + n_b])


def xg_partida(time_casa, time_fora):
    """
    KickForm: xG da partida casa × fora.
    Combina ataque/defesa específicos de mando × visitante das últimas 10
    partidas, com fallback baseado em 'forca' se o time não tem dados.
    """
    h = KICKFORM.get(_norm(time_casa))
    a = KICKFORM.get(_norm(time_fora))
    if h and a:
        atk_h = h["ataque_casa"] if h["ataque_casa"] is not None else h["ataque"]
        def_h = h["defesa_casa"] if h["defesa_casa"] is not None else h["defesa"]
        atk_a = a["ataque_fora"] if a["ataque_fora"] is not None else a["ataque"]
        def_a = a["defesa_fora"] if a["defesa_fora"] is not None else a["defesa"]
        # Fatores de força relativos à média OBSERVADA da liga
        f_atk_h = atk_h / LIGA_HOME_OBS
        f_def_a = def_a / LIGA_HOME_OBS
        f_atk_a = atk_a / LIGA_AWAY_OBS
        f_def_h = def_h / LIGA_AWAY_OBS
        xg_h = f_atk_h * f_def_a * LIGA_HOME_OBS
        xg_a = f_atk_a * f_def_h * LIGA_AWAY_OBS
    else:
        f1 = selecoes[time_casa]["forca"]
        f2 = selecoes[time_fora]["forca"]
        diff = (f1 - f2) / 40.0
        xg_h = LIGA_HOME_OBS * (1.0 + diff)
        xg_a = LIGA_AWAY_OBS * (1.0 - diff)
    xg_h = max(0.25, min(xg_h, 5.0))
    xg_a = max(0.20, min(xg_a, 5.0))
    return xg_h, xg_a


def _poisson(k, mu):
    if mu <= 0:
        return 1.0 if k == 0 else 0.0
    return math.exp(-mu) * (mu ** k) / math.factorial(k)


def chances_resultado(xg_h, xg_a, max_g=8):
    """Probabilidade de vitória/empate via produto de Poisson independente."""
    p_h = p_e = p_a = 0.0
    for gh in range(max_g):
        for ga in range(max_g):
            p = _poisson(gh, xg_h) * _poisson(ga, xg_a)
            if gh > ga:
                p_h += p
            elif gh < ga:
                p_a += p
            else:
                p_e += p
    return p_h, p_e, p_a


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
class Simulador:
    def __init__(self, root):
        self.root = root
        self.root.title("⚽  Simulador de Seleções")
        self.root.geometry("820x980")
        self.root.configure(bg=BG)
        self.root.minsize(780, 900)

        self.simulando = False
        self.animando_gol = False
        self.atk_atual = "a"
        self.gols_a = 0
        self.gols_b = 0
        self.minuto = 0
        self.stats = {"posse_a": 0, "posse_b": 0, "chutes_a": 0, "chutes_b": 0}

        self._estilo()
        self._cabecalho()
        self._seletor()
        self._barra_forca()
        self._placar()
        self._campo()
        self._narracao()
        self._botoes()

        self._atualizar_preview()

    # ---------- estilo ttk ----------
    def _estilo(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Mod.TCombobox",
            fieldbackground=BG_CARD_2,
            background=BG_CARD_2,
            foreground=TXT,
            arrowcolor=ACCENT,
            bordercolor=BG_CARD,
            lightcolor=BG_CARD,
            darkcolor=BG_CARD,
            selectbackground=BG_CARD_2,
            selectforeground=TXT,
            padding=8,
        )
        self.root.option_add("*TCombobox*Listbox.background", BG_CARD_2)
        self.root.option_add("*TCombobox*Listbox.foreground", TXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "black")
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 11))

    # ---------- cabeçalho ----------
    def _cabecalho(self):
        topo = tk.Frame(self.root, bg=BG)
        topo.pack(fill="x", pady=(18, 6))

        tk.Label(
            topo, text="⚽  SIMULADOR DE SELEÇÕES",
            font=("Segoe UI", 22, "bold"),
            bg=BG, fg=ACCENT,
        ).pack()

        tk.Label(
            topo, text="Escolha as seleções e veja a partida acontecer minuto a minuto",
            font=("Segoe UI", 10),
            bg=BG, fg=TXT_MUTED,
        ).pack(pady=(2, 0))

    # ---------- seletor de times ----------
    def _seletor(self):
        card = tk.Frame(self.root, bg=BG_CARD, padx=18, pady=14)
        card.pack(fill="x", padx=24, pady=10)

        nomes = list(selecoes.keys())

        col1 = tk.Frame(card, bg=BG_CARD)
        col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
        tk.Label(col1, text="MANDANTE", font=("Segoe UI", 9, "bold"),
                 bg=BG_CARD, fg=ACCENT_2).pack(anchor="w")
        self.combo1 = ttk.Combobox(col1, values=nomes, font=("Segoe UI", 12),
                                   state="readonly", style="Mod.TCombobox")
        self.combo1.pack(fill="x", pady=(4, 0))
        self.combo1.current(0)
        self.combo1.bind("<<ComboboxSelected>>", lambda e: self._atualizar_preview())

        tk.Label(card, text="×", font=("Segoe UI", 22, "bold"),
                 bg=BG_CARD, fg=ACCENT).pack(side="left", padx=10)

        col2 = tk.Frame(card, bg=BG_CARD)
        col2.pack(side="left", expand=True, fill="x", padx=(8, 0))
        tk.Label(col2, text="VISITANTE", font=("Segoe UI", 9, "bold"),
                 bg=BG_CARD, fg=ACCENT_2).pack(anchor="e")
        self.combo2 = ttk.Combobox(col2, values=nomes, font=("Segoe UI", 12),
                                   state="readonly", style="Mod.TCombobox")
        self.combo2.pack(fill="x", pady=(4, 0))
        self.combo2.current(1)
        self.combo2.bind("<<ComboboxSelected>>", lambda e: self._atualizar_preview())

    # ---------- barra de probabilidade ----------
    def _barra_forca(self):
        wrapper = tk.Frame(self.root, bg=BG)
        wrapper.pack(fill="x", padx=24, pady=(2, 6))

        legenda = tk.Frame(wrapper, bg=BG)
        legenda.pack(fill="x")
        self.lbl_prob_a = tk.Label(legenda, text="", font=("Segoe UI", 9, "bold"),
                                   bg=BG, fg=TXT)
        self.lbl_prob_a.pack(side="left")
        self.lbl_prob_b = tk.Label(legenda, text="", font=("Segoe UI", 9, "bold"),
                                   bg=BG, fg=TXT)
        self.lbl_prob_b.pack(side="right")

        self.canvas_barra = tk.Canvas(wrapper, height=10, bg=BG_CARD,
                                      highlightthickness=0)
        self.canvas_barra.pack(fill="x", pady=(4, 0))

        self.lbl_xg = tk.Label(wrapper, text="", font=("Segoe UI", 9),
                               bg=BG, fg=TXT_MUTED)
        self.lbl_xg.pack(pady=(4, 0))

    # ---------- placar ----------
    def _placar(self):
        self.card_placar = tk.Frame(self.root, bg=BG_CARD_2, padx=18, pady=18)
        self.card_placar.pack(fill="x", padx=24, pady=(10, 6))

        self.frame_a = tk.Frame(self.card_placar, bg=BG_CARD_2)
        self.frame_a.pack(side="left", expand=True)
        self.flag_a = tk.Label(self.frame_a, text="🏳️", font=("Segoe UI Emoji", 38),
                               bg=BG_CARD_2, fg=TXT)
        self.flag_a.pack()
        self.nome_a = tk.Label(self.frame_a, text="—", font=("Segoe UI", 12, "bold"),
                               bg=BG_CARD_2, fg=TXT)
        self.nome_a.pack()

        meio = tk.Frame(self.card_placar, bg=BG_CARD_2)
        meio.pack(side="left", padx=24)
        self.lbl_minuto = tk.Label(meio, text="00'", font=("Segoe UI", 11, "bold"),
                                   bg=BG_CARD_2, fg=ACCENT)
        self.lbl_minuto.pack()
        self.lbl_placar = tk.Label(meio, text="0  :  0",
                                   font=("Segoe UI", 36, "bold"),
                                   bg=BG_CARD_2, fg=TXT)
        self.lbl_placar.pack()
        self.lbl_status = tk.Label(meio, text="aguardando início",
                                   font=("Segoe UI", 9),
                                   bg=BG_CARD_2, fg=TXT_MUTED)
        self.lbl_status.pack()
        self.frame_meio = meio

        self.frame_b = tk.Frame(self.card_placar, bg=BG_CARD_2)
        self.frame_b.pack(side="left", expand=True)
        self.flag_b = tk.Label(self.frame_b, text="🏳️", font=("Segoe UI Emoji", 38),
                               bg=BG_CARD_2, fg=TXT)
        self.flag_b.pack()
        self.nome_b = tk.Label(self.frame_b, text="—", font=("Segoe UI", 12, "bold"),
                               bg=BG_CARD_2, fg=TXT)
        self.nome_b.pack()

    # ---------- campo 2D ----------
    def _campo(self):
        wrapper = tk.Frame(self.root, bg=BG)
        wrapper.pack(fill="x", padx=24, pady=(6, 6))

        tk.Label(wrapper, text="CAMPO",
                 font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=ACCENT_2).pack(anchor="w")

        self.campo_w = 720
        self.campo_h = 300
        holder = tk.Frame(wrapper, bg=BG_CARD, padx=6, pady=6)
        holder.pack(pady=(4, 0))
        self.canvas_campo = tk.Canvas(
            holder, width=self.campo_w, height=self.campo_h,
            bg="#2d7a3e", highlightthickness=2,
            highlightbackground="black",
        )
        self.canvas_campo.pack()

        self._desenhar_campo()
        self._criar_jogadores()
        self._criar_botao_play()

    def _desenhar_campo(self):
        c = self.canvas_campo
        w, h = self.campo_w, self.campo_h
        c.delete("linha")
        m = 12
        for i in range(8):
            cor = "#286f37" if i % 2 == 0 else "#2d7a3e"
            c.create_rectangle(m + i*(w-2*m)/8, m,
                               m + (i+1)*(w-2*m)/8, h-m,
                               fill=cor, outline="", tags="linha")
        c.create_rectangle(m, m, w-m, h-m,
                           outline="white", width=2, tags="linha")
        c.create_line(w/2, m, w/2, h-m, fill="white", width=2, tags="linha")
        c.create_oval(w/2-38, h/2-38, w/2+38, h/2+38,
                      outline="white", width=2, tags="linha")
        c.create_oval(w/2-3, h/2-3, w/2+3, h/2+3,
                      fill="white", outline="", tags="linha")
        pa_w, pa_h = 70, 140
        c.create_rectangle(m, h/2-pa_h/2, m+pa_w, h/2+pa_h/2,
                           outline="white", width=2, tags="linha")
        c.create_rectangle(w-m-pa_w, h/2-pa_h/2, w-m, h/2+pa_h/2,
                           outline="white", width=2, tags="linha")
        ga_w, ga_h = 26, 70
        c.create_rectangle(m, h/2-ga_h/2, m+ga_w, h/2+ga_h/2,
                           outline="white", width=2, tags="linha")
        c.create_rectangle(w-m-ga_w, h/2-ga_h/2, w-m, h/2+ga_h/2,
                           outline="white", width=2, tags="linha")
        gh = 50
        c.create_rectangle(m-8, h/2-gh/2, m, h/2+gh/2,
                           outline="white", fill="#dddddd", width=1, tags="linha")
        c.create_rectangle(w-m, h/2-gh/2, w-m+8, h/2+gh/2,
                           outline="white", fill="#dddddd", width=1, tags="linha")

    def _formacao(self, lado):
        w, h = self.campo_w, self.campo_h
        if lado == "a":
            gk_x, def_x, mid_x, fwd_x = w*0.06, w*0.20, w*0.34, w*0.44
        else:
            gk_x, def_x, mid_x, fwd_x = w*0.94, w*0.80, w*0.66, w*0.56
        return [
            (gk_x,  h*0.50),
            (def_x, h*0.18), (def_x, h*0.40),
            (def_x, h*0.60), (def_x, h*0.82),
            (mid_x, h*0.20), (mid_x, h*0.42),
            (mid_x, h*0.58), (mid_x, h*0.80),
            (fwd_x, h*0.36), (fwd_x, h*0.64),
        ]

    def _criar_jogadores(self):
        self.flash_overlay = self.canvas_campo.create_rectangle(
            0, 0, self.campo_w, self.campo_h,
            fill="#ffff66", outline="", state="hidden",
            stipple="gray25", tags="flash",
        )
        self.pos_base_a = self._formacao("a")
        self.pos_base_b = self._formacao("b")
        self.jogadores_a = []
        self.jogadores_b = []
        r = 7
        for x, y in self.pos_base_a:
            d = self.canvas_campo.create_oval(
                x-r, y-r, x+r, y+r,
                fill="white", outline="black", width=2, tags="jogador",
            )
            self.jogadores_a.append(d)
        for x, y in self.pos_base_b:
            d = self.canvas_campo.create_oval(
                x-r, y-r, x+r, y+r,
                fill="white", outline="black", width=2, tags="jogador",
            )
            self.jogadores_b.append(d)
        self.bola_pos = [self.campo_w/2, self.campo_h/2]
        self.bola = self.canvas_campo.create_oval(
            self.bola_pos[0]-5, self.bola_pos[1]-5,
            self.bola_pos[0]+5, self.bola_pos[1]+5,
            fill="white", outline="black", width=1, tags="bola",
        )

    def _criar_botao_play(self):
        cx, cy = self.campo_w/2, self.campo_h/2
        r = 34
        self.play_bg = self.canvas_campo.create_oval(
            cx-r, cy-r, cx+r, cy+r,
            fill="#000000", outline="white", width=3, tags="play_btn",
        )
        self.play_tri = self.canvas_campo.create_polygon(
            cx-11, cy-15, cx-11, cy+15, cx+17, cy,
            fill="white", outline="white", tags="play_btn",
        )
        self.canvas_campo.tag_bind("play_btn", "<Button-1>",
                                   lambda e: self.iniciar())
        self.canvas_campo.tag_bind("play_btn", "<Enter>",
                                   lambda e: (self.canvas_campo.config(cursor="hand2"),
                                              self.canvas_campo.itemconfig(self.play_bg, fill="#1a1a1a")))
        self.canvas_campo.tag_bind("play_btn", "<Leave>",
                                   lambda e: (self.canvas_campo.config(cursor=""),
                                              self.canvas_campo.itemconfig(self.play_bg, fill="#000000")))

    def _mostrar_play(self, visivel):
        state = "normal" if visivel else "hidden"
        self.canvas_campo.itemconfig("play_btn", state=state)

    def _resetar_posicoes(self):
        r = 7
        for dot, (x, y) in zip(self.jogadores_a, self.pos_base_a):
            self.canvas_campo.coords(dot, x-r, y-r, x+r, y+r)
        for dot, (x, y) in zip(self.jogadores_b, self.pos_base_b):
            self.canvas_campo.coords(dot, x-r, y-r, x+r, y+r)
        self.bola_pos = [self.campo_w/2, self.campo_h/2]
        self.canvas_campo.coords(
            self.bola,
            self.bola_pos[0]-5, self.bola_pos[1]-5,
            self.bola_pos[0]+5, self.bola_pos[1]+5,
        )

    def _mover_jogador(self, dot, base, peso_ball=0.03):
        r = 7
        x1, y1, x2, y2 = self.canvas_campo.coords(dot)
        cx, cy = (x1+x2)/2, (y1+y2)/2
        bx, by = base
        bola_x, bola_y = self.bola_pos
        dx = (bx - cx) * 0.15 + (bola_x - cx) * peso_ball + random.uniform(-2.5, 2.5)
        dy = (by - cy) * 0.15 + (bola_y - cy) * peso_ball + random.uniform(-2.5, 2.5)
        nx = max(14, min(self.campo_w-14, cx+dx))
        ny = max(14, min(self.campo_h-14, cy+dy))
        self.canvas_campo.coords(dot, nx-r, ny-r, nx+r, ny+r)

    def _animar_jogadores(self):
        if self.animando_gol:
            return
        self._mover_jogador(self.jogadores_a[0], self.pos_base_a[0], peso_ball=0.0)
        self._mover_jogador(self.jogadores_b[0], self.pos_base_b[0], peso_ball=0.0)
        for dot, base in zip(self.jogadores_a[1:], self.pos_base_a[1:]):
            self._mover_jogador(dot, base)
        for dot, base in zip(self.jogadores_b[1:], self.pos_base_b[1:]):
            self._mover_jogador(dot, base)
        target_x = self.campo_w * 0.78 if self.atk_atual == "a" else self.campo_w * 0.22
        target_y = self.campo_h / 2 + random.uniform(-70, 70)
        self.bola_pos[0] += (target_x - self.bola_pos[0]) * 0.10 + random.uniform(-5, 5)
        self.bola_pos[1] += (target_y - self.bola_pos[1]) * 0.10 + random.uniform(-4, 4)
        self.bola_pos[0] = max(18, min(self.campo_w-18, self.bola_pos[0]))
        self.bola_pos[1] = max(18, min(self.campo_h-18, self.bola_pos[1]))
        self.canvas_campo.coords(
            self.bola,
            self.bola_pos[0]-5, self.bola_pos[1]-5,
            self.bola_pos[0]+5, self.bola_pos[1]+5,
        )

    def _animar_gol(self, lado, cor):
        self.animando_gol = True
        goal_x = self.campo_w - 14 if lado == "a" else 14
        goal_y = self.campo_h / 2
        self._mover_bola_ate(goal_x, goal_y, 12, lambda: self._flash_campo(cor))

    def _mover_bola_ate(self, tx, ty, passos, callback):
        if passos <= 0:
            if callback:
                callback()
            return
        dx = (tx - self.bola_pos[0]) / passos
        dy = (ty - self.bola_pos[1]) / passos
        self.bola_pos[0] += dx
        self.bola_pos[1] += dy
        self.canvas_campo.coords(
            self.bola,
            self.bola_pos[0]-5, self.bola_pos[1]-5,
            self.bola_pos[0]+5, self.bola_pos[1]+5,
        )
        self.root.after(28, lambda: self._mover_bola_ate(tx, ty, passos-1, callback))

    def _flash_campo(self, cor, count=8):
        if count <= 0:
            self.canvas_campo.itemconfig(self.flash_overlay, state="hidden")
            self.animando_gol = False
            self.bola_pos = [self.campo_w/2, self.campo_h/2]
            self.canvas_campo.coords(
                self.bola,
                self.bola_pos[0]-5, self.bola_pos[1]-5,
                self.bola_pos[0]+5, self.bola_pos[1]+5,
            )
            return
        if count % 2 == 0:
            self.canvas_campo.itemconfig(self.flash_overlay, fill=cor, state="normal")
        else:
            self.canvas_campo.itemconfig(self.flash_overlay, state="hidden")
        self.root.after(110, lambda: self._flash_campo(cor, count-1))

    # ---------- narração ----------
    def _narracao(self):
        wrapper = tk.Frame(self.root, bg=BG)
        wrapper.pack(fill="both", expand=True, padx=24, pady=(6, 6))

        tk.Label(wrapper, text="NARRAÇÃO AO VIVO",
                 font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=ACCENT_2).pack(anchor="w")

        caixa = tk.Frame(wrapper, bg=BG_CARD)
        caixa.pack(fill="both", expand=True, pady=(4, 0))

        self.txt_narracao = tk.Text(
            caixa, bg=BG_CARD, fg=TXT,
            font=("Consolas", 10),
            relief="flat", padx=12, pady=10,
            wrap="word", state="disabled",
            height=5,
        )
        self.txt_narracao.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(caixa, command=self.txt_narracao.yview,
                              bg=BG_CARD, troughcolor=BG_CARD,
                              activebackground=ACCENT)
        scroll.pack(side="right", fill="y")
        self.txt_narracao.config(yscrollcommand=scroll.set)

        self.txt_narracao.tag_config("gol",   foreground=WIN, font=("Consolas", 10, "bold"))
        self.txt_narracao.tag_config("apito", foreground=ACCENT, font=("Consolas", 10, "bold"))

    # ---------- botões ----------
    def _botoes(self):
        bar = tk.Frame(self.root, bg=BG)
        bar.pack(fill="x", padx=24, pady=(4, 18))

        self.btn = tk.Button(
            bar, text="▶  SIMULAR PARTIDA",
            font=("Segoe UI", 13, "bold"),
            bg=ACCENT, fg="black", activebackground="#ffc46b",
            activeforeground="black",
            relief="flat", padx=22, pady=12,
            cursor="hand2",
            command=self.iniciar,
        )
        self.btn.pack(side="left", expand=True, fill="x", padx=(0, 6))

        self.btn_resumo = tk.Button(
            bar, text="📊  Resumo",
            font=("Segoe UI", 11, "bold"),
            bg=BG_CARD_2, fg=TXT, activebackground=BG_CARD,
            activeforeground=TXT,
            relief="flat", padx=18, pady=12,
            cursor="hand2",
            command=self.mostrar_resumo,
            state="disabled",
        )
        self.btn_resumo.pack(side="left", padx=(6, 0))

    # ------------------------------------------------------------------
    # PREVIEW (pré-partida)
    # ------------------------------------------------------------------
    def _atualizar_preview(self):
        t1, t2 = self.combo1.get(), self.combo2.get()
        if not t1 or not t2:
            return

        d1, d2 = selecoes[t1], selecoes[t2]

        self.flag_a.config(text=d1["flag"])
        self.flag_b.config(text=d2["flag"])
        self.nome_a.config(text=t1.upper(), fg=d1["cor"])
        self.nome_b.config(text=t2.upper(), fg=d2["cor"])

        if hasattr(self, "jogadores_a"):
            for dot in self.jogadores_a:
                self.canvas_campo.itemconfig(dot, fill=d1["cor"])
            for dot in self.jogadores_b:
                self.canvas_campo.itemconfig(dot, fill=d2["cor"])

        if not self.simulando:
            self.lbl_placar.config(text="0  :  0", fg=TXT)
            self.lbl_minuto.config(text="00'")
            self.lbl_status.config(text="aguardando início", fg=TXT_MUTED)

        xg_h, xg_a = xg_partida(t1, t2)
        p_h, p_e, p_a = chances_resultado(xg_h, xg_a)
        self.lbl_prob_a.config(text=f"{t1}  {p_h*100:.0f}%", fg=d1["cor"])
        self.lbl_prob_b.config(text=f"{p_a*100:.0f}%  {t2}", fg=d2["cor"])
        em_dados = "KickForm" if (_norm(t1) in KICKFORM and _norm(t2) in KICKFORM) else "fallback (força)"
        self.lbl_xg.config(
            text=f"xG ({em_dados}):  {xg_h:.2f}  ×  {xg_a:.2f}    empate {p_e*100:.0f}%"
        )

        self.root.update_idletasks()
        self.canvas_barra.delete("all")
        w = self.canvas_barra.winfo_width() or 700
        h = 10
        cut1 = int(w * p_h)
        cut2 = int(w * (p_h + p_e))
        self.canvas_barra.create_rectangle(0, 0, cut1, h, fill=d1["cor"], outline="")
        self.canvas_barra.create_rectangle(cut1, 0, cut2, h, fill=DRAW, outline="")
        self.canvas_barra.create_rectangle(cut2, 0, w, h, fill=d2["cor"], outline="")

    # ------------------------------------------------------------------
    # SIMULAÇÃO
    # ------------------------------------------------------------------
    def iniciar(self):
        if self.simulando:
            return

        t1, t2 = self.combo1.get(), self.combo2.get()
        if t1 == t2:
            messagebox.showwarning("Ops!", "Escolha duas seleções diferentes.")
            return

        self.xg_a, self.xg_b = xg_partida(t1, t2)
        # Pré-amostra o placar final via Poisson (mantém a distribuição
        # exatamente igual à previsão KickForm ao longo de muitas simulações)
        gols_a_total = _sample_poisson(self.xg_a)
        gols_b_total = _sample_poisson(self.xg_b)
        self.min_gols_a, self.min_gols_b = sortear_minutos_gols(
            gols_a_total, gols_b_total
        )

        self.simulando = True
        self.btn.config(state="disabled", text="● AO VIVO",
                        bg=LOSE, fg="white")
        self.btn_resumo.config(state="disabled")
        self.combo1.config(state="disabled")
        self.combo2.config(state="disabled")

        self.gols_a = 0
        self.gols_b = 0
        self.minuto = 0
        self.stats = {"posse_a": 0, "posse_b": 0, "chutes_a": 0, "chutes_b": 0}
        self.animando_gol = False
        self._resetar_posicoes()
        self._mostrar_play(False)
        self.lbl_placar.config(text="0  :  0", fg=TXT)
        self.lbl_status.config(text="bola rolando", fg=WIN)

        self.txt_narracao.config(state="normal")
        self.txt_narracao.delete("1.0", "end")
        self.txt_narracao.config(state="disabled")

        self._narrar(f"🏟️  Início de jogo: {t1} × {t2}", tag="apito")
        self.root.after(450, self._tick)

    def _tick(self):
        if self.minuto >= 90:
            self._fim()
            return

        self.minuto += 1
        self.lbl_minuto.config(text=f"{self.minuto:02d}'")

        if self.minuto == 45:
            self._narrar("⏸️  Fim do primeiro tempo.", tag="apito")

        t1, t2 = self.combo1.get(), self.combo2.get()

        # Posse baseada na razão de xG (time mais perigoso fica com mais bola)
        posse_a = self.xg_a / (self.xg_a + self.xg_b)
        if random.random() < posse_a:
            self.stats["posse_a"] += 1
            atk_lado, atk_time = "a", t1
        else:
            self.stats["posse_b"] += 1
            atk_lado, atk_time = "b", t2

        self.atk_atual = atk_lado
        self._animar_jogadores()

        # Gols pré-determinados via Poisson(xG) no início do jogo
        if self.minuto in self.min_gols_a:
            self.stats["chutes_a"] += 1
            self._gol("a", t1)
        elif self.minuto in self.min_gols_b:
            self.stats["chutes_b"] += 1
            self._gol("b", t2)
        elif random.random() < 0.20:
            self.stats[f"chutes_{atk_lado}"] += 1
            if random.random() < 0.5:
                self._narrar(f"  {self.minuto:02d}'  Chute de {atk_time}, defesa do goleiro!")
        elif random.random() < 0.12:
            self._narrar(f"  {self.minuto:02d}'  {random.choice(EVENTOS_NEUTROS)}")

        delay = random.randint(70, 140)
        self.root.after(delay, self._tick)

    def _gol(self, lado, time):
        if lado == "a":
            self.gols_a += 1
        else:
            self.gols_b += 1
        cor = selecoes[time]["cor"]

        self.lbl_placar.config(text=f"{self.gols_a}  :  {self.gols_b}", fg=cor)
        self._flash_placar(cor)
        self._animar_gol(lado, cor)
        play_sfx("gol efeito.mp3")

        msg = f"⚽ {self.minuto:02d}'  {time.upper()} — {random.choice(NARRACOES_GOL)}"
        self._narrar(msg, tag="gol")

    def _flash_placar(self, cor):
        widgets = [self.card_placar, self.frame_a, self.frame_b, self.frame_meio,
                   self.flag_a, self.nome_a, self.flag_b, self.nome_b,
                   self.lbl_minuto, self.lbl_placar, self.lbl_status]
        for w in widgets:
            try:
                w.config(bg=cor)
            except tk.TclError:
                pass
        self.root.after(220, self._restaurar_placar)

    def _restaurar_placar(self):
        widgets = [self.card_placar, self.frame_a, self.frame_b, self.frame_meio,
                   self.flag_a, self.nome_a, self.flag_b, self.nome_b,
                   self.lbl_minuto, self.lbl_placar, self.lbl_status]
        for w in widgets:
            try:
                w.config(bg=BG_CARD_2)
            except tk.TclError:
                pass

    def _fim(self):
        t1, t2 = self.combo1.get(), self.combo2.get()
        self._narrar("⏹️  Fim de jogo!", tag="apito")

        if self.gols_a > self.gols_b:
            self.lbl_status.config(text=f"🏆 {t1.upper()} VENCEU", fg=WIN)
            self._narrar(f"🏆 Vitória de {t1} por {self.gols_a} a {self.gols_b}!", tag="apito")
        elif self.gols_b > self.gols_a:
            self.lbl_status.config(text=f"🏆 {t2.upper()} VENCEU", fg=WIN)
            self._narrar(f"🏆 Vitória de {t2} por {self.gols_b} a {self.gols_a}!", tag="apito")
        else:
            self.lbl_status.config(text="🤝 EMPATE", fg=DRAW)
            self._narrar(f"🤝 Empate em {self.gols_a} a {self.gols_b}.", tag="apito")

        self.simulando = False
        self.btn.config(state="normal", text="▶  SIMULAR DE NOVO",
                        bg=ACCENT, fg="black")
        self.btn_resumo.config(state="normal")
        self.combo1.config(state="readonly")
        self.combo2.config(state="readonly")
        self._mostrar_play(True)

    # ------------------------------------------------------------------
    def _narrar(self, texto, tag=None):
        self.txt_narracao.config(state="normal")
        self.txt_narracao.insert("end", texto + "\n", tag or "")
        self.txt_narracao.see("end")
        self.txt_narracao.config(state="disabled")

    def mostrar_resumo(self):
        t1, t2 = self.combo1.get(), self.combo2.get()
        total_posse = self.stats["posse_a"] + self.stats["posse_b"] or 1
        pa = self.stats["posse_a"] / total_posse * 100
        pb = self.stats["posse_b"] / total_posse * 100

        msg = (
            f"PLACAR FINAL\n"
            f"{t1}  {self.gols_a}  x  {self.gols_b}  {t2}\n\n"
            f"Posse de bola\n"
            f"  {t1}: {pa:.0f}%\n"
            f"  {t2}: {pb:.0f}%\n\n"
            f"Finalizações\n"
            f"  {t1}: {self.stats['chutes_a']}\n"
            f"  {t2}: {self.stats['chutes_b']}\n"
        )
        messagebox.showinfo("📊 Resumo da partida", msg)


# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = Simulador(root)
    root.bind("<Configure>", lambda e: app._atualizar_preview())
    root.mainloop()
